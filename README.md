# Painel YMS — Guarapuava

Painel HTML que lê ao vivo a aba "Programação diária" da planilha e mostra só os caminhões que ainda estão no pátio (Entrada preenchida, Saida vazia). Atualiza sozinho a cada 30s, com cores: amarelo (15–20min), laranja (20–25min), vermelho (25min+).

Não tenho um conector de GitHub nesta sessão, então não consigo criar o repositório por você — os passos abaixo são rápidos.

**Atualização importante:** a Mercadolibre SRL bloqueia a opção "Qualquer pessoa com o link" no Google Drive/Sheets — não dá para tornar a planilha pública. Por isso a arquitetura mudou: a planilha continua 100% restrita ao domínio, e um **Web App do Apps Script** (implantado com acesso "Qualquer pessoa da Mercadolibre SRL") funciona como ponte entre o painel público no GitHub e os dados privados. Só quem estiver logado com conta @mercadolibre.com consegue ver os dados no painel.

## Passo 1 — Implantar o Web App do Apps Script

Siga o **Passo 6** do arquivo `Guia_Controle_YMS_Guarapuava.md` (adicionar as funções `doGet`/`buildPayload` ao script já existente na planilha, e implantar como Web App). No final você vai ter uma URL assim:
```
https://script.google.com/macros/s/AKfycb.../exec
```

## Passo 2 — Colar a URL no index.html

Abra `index.html` e troque a linha:
```javascript
var WEBAPP_URL = 'COLE_AQUI_A_URL_DO_WEB_APP/exec';
```
pela URL que você copiou no Passo 1.

## Passo 3 — Subir o arquivo atualizado

Você já tem o repositório e o Pages ativo em `https://frmarques-spr5.github.io/yms-spr5/`. Só precisa subir este `index.html` novo por cima do antigo: no repositório `yms-spr5` no GitHub, abra o `index.html`, clique no ícone de lápis (editar), apague o conteúdo, cole o conteúdo do novo `index.html` (deste pacote) → `Commit changes`.

Em 1-2 minutos o GitHub Pages atualiza sozinho o link já existente.

## Se o painel aparecer com erro / sem dados

- Confirme se o Passo 1 (implantar o Web App) e o Passo 2 (colar a URL) foram feitos.
- Quem for abrir o painel precisa estar logado com uma conta **@mercadolibre.com** no navegador (a autenticação do Web App é por sessão do Google, não por senha no painel).
- Confirme se o nome da aba na planilha continua `Programação diária` (se renomear, edite `SHEET_MAIN` no Apps Script).
- Abra o painel e veja o console do navegador (F12) se o erro persistir.
