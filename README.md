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


