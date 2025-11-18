TODO
ZscalerRootCA.cerをwindowsのcetmgの証明書->信頼されたroot証明機関からbase64でexportしておく

WSL内の作業
sudo mv -iv /tmp/*.cer /usr/share/ca-certificates/
sudo vi /etc/ca-certificates.conf

ca-certificates.conf
の最後の行にZscalerRootCA.cerを追加する


sudo update-ca-certificates

export OPENROUTER_API_KEY="your-api-key"
TODO
your-api-keyはOpenRouterのKey -> createKey自分で用意すること。金額設定は０ドルでよい

export SSL_CERT_FILE=/etc/ssl/certs/ca-certificates.crt
export REQUESTS_CA_BUNDLE=/etc/ssl/certs/ca-certificates.crt


pip3 install pdf2docx PyMuPDF python-docx pypdf requests


 python3 main.py your.pdf
 pdfがwordで開けるように変換される。ただし表示崩れる

 python3 image.py your.pdf
 pdfのページごとにpngに変換される（ページごとに個別にocrするための準備）

 python3 image_ocr.py 対象ディレクトリ/対象.png
 llmがocrしてマークダウンに変換する

-> LLMにコンテキストを渡して作図なり要件のまとめを行う
