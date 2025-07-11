# Desenvolvimento

Instruções para desenvolvimento.

## Instalação dos pacotes

```pip install -r requirements.txt```

## Gerando um executável empacotado

pyinstaller --onefile --noconsole --icon="./sources/itshare.ico" --name RecSUS ./recsus.py