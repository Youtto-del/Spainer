from urllib import request
file_url = 'https://eproc1g.tjrs.jus.br/eproc/controlador.php?acao=acessar_documento&doc=11647454802004143188223224923&evento=11647454802004143188223785201&key=086dceeee7cfd7413930e6b7b54465795999f16e0558cc479e9d8db26347e30e&hash=61c7374aaf35944208596efe91209b10'
file = 'arquivolocal.pdf'
request.urlretrieve(file_url , file)