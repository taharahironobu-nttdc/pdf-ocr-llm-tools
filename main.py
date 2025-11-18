import os
import sys


def convert_with_pdf2docx(input_pdf, output_docx):
    """pdf2docxを使った変換（レイアウト保持に優れる）"""
    try:
        from pdf2docx import Converter

        print(f"pdf2docxで変換中: {input_pdf} -> {output_docx}")
        cv = Converter(input_pdf)
        cv.convert(output_docx)
        cv.close()
        print("✓ 変換成功!")
        return True
    except Exception as e:
        print(f"✗ pdf2docxでの変換失敗: {e}")
        return False


def convert_with_pymupdf(input_pdf, output_docx):
    """PyMuPDFを使った変換（制限を無視できる）"""
    try:
        import fitz  # PyMuPDF
        from docx import Document

        print(f"PyMuPDFで変換中: {input_pdf} -> {output_docx}")
        doc = Document()
        pdf = fitz.open(input_pdf)

        for page_num, page in enumerate(pdf, 1):
            print(f"  ページ {page_num}/{len(pdf)} 処理中...")
            text = page.get_text()
            if text.strip():  # 空白ページをスキップ
                doc.add_paragraph(text)
                doc.add_page_break()

        pdf.close()
        doc.save(output_docx)
        print("✓ 変換成功!")
        return True
    except Exception as e:
        print(f"✗ PyMuPDFでの変換失敗: {e}")
        return False


def unlock_pdf(input_pdf, output_pdf):
    """PDFの制限を解除"""
    try:
        from pypdf import PdfReader, PdfWriter

        print(f"PDF制限解除中: {input_pdf} -> {output_pdf}")
        reader = PdfReader(input_pdf)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        with open(output_pdf, "wb") as f:
            writer.write(f)

        print("✓ 制限解除成功!")
        return True
    except Exception as e:
        print(f"✗ 制限解除失敗: {e}")
        return False


def main():
    if len(sys.argv) < 2:
        print("使い方: python convert_pdf.py input.pdf [output.docx]")
        sys.exit(1)

    input_pdf = sys.argv[1]

    if not os.path.exists(input_pdf):
        print(f"エラー: ファイルが見つかりません: {input_pdf}")
        sys.exit(1)

    # 出力ファイル名を決定
    if len(sys.argv) >= 3:
        output_docx = sys.argv[2]
    else:
        base_name = os.path.splitext(input_pdf)[0]
        output_docx = f"{base_name}.docx"

    print(f"\n=== PDF to Word 変換 ===")
    print(f"入力: {input_pdf}")
    print(f"出力: {output_docx}\n")

    # まずpdf2docxを試す（レイアウトが綺麗）
    if convert_with_pdf2docx(input_pdf, output_docx):
        return

    # 失敗したらPyMuPDFを試す（制限に強い）
    print("\n別の方法を試します...")
    if convert_with_pymupdf(input_pdf, output_docx):
        return

    # それでも失敗したら制限解除を試みる
    print("\nPDF制限解除を試みます...")
    unlocked_pdf = "unlocked_temp.pdf"
    if unlock_pdf(input_pdf, unlocked_pdf):
        print("\n制限解除したPDFで再変換...")
        if convert_with_pdf2docx(unlocked_pdf, output_docx) or convert_with_pymupdf(
            unlocked_pdf, output_docx
        ):
            os.remove(unlocked_pdf)  # 一時ファイル削除
            return
        os.remove(unlocked_pdf)

    print("\n✗ 全ての変換方法が失敗しました")
    sys.exit(1)


if __name__ == "__main__":
    main()
