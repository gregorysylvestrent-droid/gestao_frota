# gestao_frota

## Abrir o dashboard como App Web

Além do menu dentro da planilha, o dashboard pode ser publicado como um App Web do Google Apps Script para abrir por link direto:

1. Abra a planilha e acesse **Extensões > Apps Script**.
2. Clique em **Implantar > Nova implantação**.
3. Selecione o tipo **App da Web**.
4. Em **Executar como**, escolha o usuário que tem acesso à planilha.
5. Em **Quem pode acessar**, escolha o público desejado.
6. Clique em **Implantar** e copie a URL gerada.

Depois de publicado, o menu **📊 Dashboard Frota > Ver link do App Web** também mostra o link direto da implantação atual.

> Se o script for usado fora da planilha original, preencha `SPREADSHEET_ID` no `Code.gs` com o ID da planilha de dados.
