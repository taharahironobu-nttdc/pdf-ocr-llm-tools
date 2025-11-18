import base64
import glob
import os
import sys
from pathlib import Path

import requests

# 利用可能なVisionモデル
VISION_MODELS = {
    # 無料モデル (推奨)
    "qwen-72b-free": "qwen/qwen2.5-vl-72b-instruct:free",
    "qwen-32b-free": "qwen/qwen2.5-vl-32b-instruct:free",
    "llama-vision-free": "meta-llama/llama-3.2-11b-vision-instruct:free",
    "kimi-vl-free": "moonshotai/kimi-vl-a3b-thinking:free",
    "gemma-27b-free": "google/gemma-3-27b-it:free",
    "gemma-12b-free": "google/gemma-3-12b-it:free",
    # 有料モデル
    "llama-3.2-90b": "meta-llama/llama-3.2-90b-vision-instruct",
    "gpt-4o": "openai/gpt-4o",
    "claude-3.5-sonnet": "anthropic/claude-3.5-sonnet",
}


def encode_image(image_path):
    """画像をBase64エンコード"""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")


def get_image_mime_type(image_path):
    """ファイル拡張子からMIMEタイプを取得"""
    ext = Path(image_path).suffix.lower()
    mime_types = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
        ".bmp": "image/bmp",
    }
    return mime_types.get(ext, "image/png")


def ocr_with_openrouter(image_path, model, api_key, custom_prompt=None):
    """OpenRouterを使ってOCR実行"""
    try:
        base64_image = encode_image(image_path)
        mime_type = get_image_mime_type(image_path)

        # デフォルトのプロンプト
        if custom_prompt is None:
            custom_prompt = """この画像からテキストを正確に抽出してください。

要件:
- レイアウト、表、箇条書きなどの構造を保持
- 出力はマークダウン形式
- 余計な説明は不要、テキストのみを出力
- 日本語の場合は日本語で、英語の場合は英語で出力"""

        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "HTTP-Referer": "https://github.com/pdf-converter",
                "X-Title": "Image OCR Tool",
            },
            json={
                "model": model,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": custom_prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:{mime_type};base64,{base64_image}"
                                },
                            },
                        ],
                    }
                ],
                "max_tokens": 4000,
                "temperature": 0.1,
            },
            timeout=120,
        )

        if response.status_code == 200:
            result = response.json()
            return result["choices"][0]["message"]["content"]
        else:
            error_msg = response.json().get("error", {})
            print(f"OpenRouter APIエラー: {response.status_code}")
            print(f"詳細: {error_msg}")
            return None

    except requests.exceptions.RequestException as e:
        print(f"リクエストエラー: {e}")
        return None
    except Exception as e:
        print(f"OCRエラー: {e}")
        return None


def process_single_image(
    image_path, model, api_key, output_file=None, custom_prompt=None
):
    """単一画像をOCR処理"""
    print(f"\n画像を処理中: {image_path}")
    print(f"モデル: {model}\n")

    text = ocr_with_openrouter(image_path, model, api_key, custom_prompt)

    if text:
        # 出力ファイル名を決定
        if output_file is None:
            base_name = Path(image_path).stem
            output_file = f"{base_name}_ocr.txt"

        # テキストファイルに保存
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(text)

        print(f"\n✓ OCR成功!")
        print(f"出力ファイル: {output_file}")
        print(f"\n--- 抽出されたテキスト (最初の500文字) ---")
        print(text[:500])
        if len(text) > 500:
            print("...(省略)...")
        return True
    else:
        print("\n✗ OCR失敗")
        return False


def process_multiple_images(
    image_patterns, model, api_key, output_dir=None, custom_prompt=None
):
    """複数画像をOCR処理"""
    # パターンから画像ファイルを取得
    image_files = []
    for pattern in image_patterns:
        if os.path.isdir(pattern):
            # ディレクトリの場合、すべての画像を取得
            pattern = os.path.join(pattern, "*")

        files = glob.glob(pattern)
        image_files.extend(
            [
                f
                for f in files
                if Path(f).suffix.lower()
                in [".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"]
            ]
        )

    image_files = sorted(list(set(image_files)))  # 重複削除とソート

    if not image_files:
        print("エラー: 画像ファイルが見つかりません")
        return False

    # 出力ディレクトリの作成
    if output_dir is None:
        output_dir = "ocr_results"
    os.makedirs(output_dir, exist_ok=True)

    print(f"\n=== 画像OCR処理 ===")
    print(f"画像数: {len(image_files)}")
    print(f"モデル: {model}")
    print(f"出力先: {output_dir}/\n")

    results = []
    for i, image_path in enumerate(image_files, 1):
        print(f"--- [{i}/{len(image_files)}] {Path(image_path).name} ---")

        output_file = os.path.join(output_dir, f"{Path(image_path).stem}_ocr.txt")
        text = ocr_with_openrouter(image_path, model, api_key, custom_prompt)

        if text:
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(text)
            results.append(
                {
                    "image": image_path,
                    "output": output_file,
                    "text": text,
                    "success": True,
                }
            )
            print(f"✓ 成功: {output_file}")
        else:
            results.append({"image": image_path, "success": False})
            print(f"✗ 失敗")
        print()

    # サマリー
    success_count = sum(1 for r in results if r["success"])
    print(f"\n=== 処理完了 ===")
    print(f"成功: {success_count}/{len(results)}")
    print(f"出力ディレクトリ: {output_dir}/")

    # すべてのテキストを結合したファイルも作成
    combined_file = os.path.join(output_dir, "all_combined.txt")
    with open(combined_file, "w", encoding="utf-8") as f:
        for i, result in enumerate(results, 1):
            if result["success"]:
                f.write(f"=== {Path(result['image']).name} ===\n\n")
                f.write(result["text"])
                f.write(f"\n\n{'=' * 50}\n\n")
    print(f"結合ファイル: {combined_file}")

    return success_count == len(results)


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="画像をOpenRouterでOCR処理",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # 単一画像をOCR
  python image_ocr.py image.png

  # 複数画像を指定
  python image_ocr.py image1.png image2.jpg image3.png

  # ワイルドカードで複数指定
  python image_ocr.py "document_images/*.png"

  # ディレクトリ内のすべての画像
  python image_ocr.py document_images/

  # モデルを指定
  python image_ocr.py image.png --model qwen-72b-free

  # 出力先を指定
  python image_ocr.py image.png -o result.txt

  # カスタムプロンプト
  python image_ocr.py receipt.png --prompt "この領収書から金額と日付を抽出してください"
""",
    )
    parser.add_argument(
        "images", nargs="+", help="画像ファイル、ディレクトリ、またはパターン"
    )
    parser.add_argument("-o", "--output", help="出力ファイル名（単一画像の場合）")
    parser.add_argument(
        "--output-dir",
        help="出力ディレクトリ（複数画像の場合、デフォルト: ocr_results）",
    )
    parser.add_argument(
        "--model",
        default="qwen-32b-free",
        choices=list(VISION_MODELS.keys()),
        help="使用するモデル (デフォルト: qwen-72b-free)",
    )
    parser.add_argument("--api-key", help="OpenRouter APIキー")
    parser.add_argument("--prompt", help="カスタムプロンプト")
    parser.add_argument(
        "--list-models", action="store_true", help="利用可能なモデルを表示"
    )

    args = parser.parse_args()

    if args.list_models:
        print("\n利用可能なモデル:")
        for short_name, full_name in VISION_MODELS.items():
            free = "🆓" if "free" in short_name else "💰"
            print(f"  {free} {short_name:20s} -> {full_name}")
        return

    # APIキーの取得
    api_key = args.api_key or os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        print("エラー: OPENROUTER_API_KEYが設定されていません")
        print("\n以下のいずれかの方法で設定してください:")
        print("1. export OPENROUTER_API_KEY='your-api-key'")
        print("2. python image_ocr.py image.png --api-key your-api-key")
        sys.exit(1)

    model_full_name = VISION_MODELS[args.model]

    # 画像ファイルのチェック
    all_images = []
    for pattern in args.images:
        if os.path.isfile(pattern):
            all_images.append(pattern)
        elif os.path.isdir(pattern) or "*" in pattern or "?" in pattern:
            # ディレクトリまたはワイルドカード
            all_images.extend(args.images)
            break

    # 単一画像か複数画像かで処理を分ける
    if len(args.images) == 1 and os.path.isfile(args.images[0]):
        # 単一画像
        process_single_image(
            args.images[0], model_full_name, api_key, args.output, args.prompt
        )
    else:
        # 複数画像
        process_multiple_images(
            args.images, model_full_name, api_key, args.output_dir, args.prompt
        )


if __name__ == "__main__":
    main()
