# Script para a planilha "YMS - Guarapuava" (conta pessoal, sem nome de motorista)

⚠️ **Antes de mais nada:** essa planilha (`1Go3UQPBkJQO28eLSQ1LaQ0vR7K7ZKitBzb4dS895W-I`) já foi copiada com os nomes de motorista preenchidos nas abas "Programação diária", "Já Saiu" e "Base". O script abaixo impede que **novos** nomes sejam gravados, mas não apaga o que já está lá. Para isso:

- Na aba "Programação diária": selecione a coluna D (motorista) e apague o conteúdo das células com dado.
- Na aba "Já Saiu": mesma coisa, apague a coluna D.
- Na aba "Base": apague a coluna C (Nome do Driver) inteira, ou a aba toda se não for mais usar.

## O que muda no código

- `preencherDadosDaPlaca` não escreve mais Motorista — só Tipo de veículo e Transportadora.
- `buscarNaBase` não lê mais a coluna de nome, e os dados agora começam na **linha 2** da Base (não na 3).
- `buildPayload`/`doGet` não incluem motorista no JSON — o painel só recebe Placa, Vaga, Tipo, Transportadora, Entrada.

Recomendo manter a arquitetura de Web App (em vez de deixar a planilha pública): mesmo estando numa conta pessoal sem a restrição do Workspace, o Web App só expõe os campos que você decidir (curados), enquanto deixar a planilha "Qualquer pessoa com o link" exporia a aba Base, o histórico completo de "Já Saiu" e qualquer outra coisa que você adicionar depois. Dá pra manter essa planilha como "Restrito" mesmo e só o Web App fica acessível.

## Código completo (cole em Extensões > Apps Script desta planilha)

```javascript
var SHEET_MAIN  = 'Programação diária';
var SHEET_BASE  = 'Base';
var SHEET_DONE  = 'Já Saiu';

var COL_PLACA          = 2;  // B
// COL_MOTORISTA removido de propósito — não gravamos mais nome de motorista
var COL_TIPO_VEICULO   = 5;  // E
var COL_TRANSPORTADORA = 6;  // F
var COL_ENTRADA        = 7;  // G
var COL_SAIDA          = 8;  // H
var COL_STATUS         = 9;  // I
var ULTIMA_COLUNA_PARA_PINTAR = 10; // pinta de A até J

var BASE_LINHA_CABECALHO = 1; // nesta planilha o cabeçalho da Base é a linha 1 (dados começam na 2)

function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_MAIN) return;

  var range = e.range;
  var col = range.getColumn();
  var numCols = range.getNumColumns();
  var firstRow = Math.max(range.getRow(), 2); // ignora cabeçalho
  var lastRow = range.getRow() + range.getNumRows() - 1;
  if (lastRow < firstRow) return;

  // 1) Edição tocou a coluna Placa -> preenche Entrada + busca Tipo/Transportadora na Base
  if (colunaTocada(col, numCols, COL_PLACA)) {
    for (var row = firstRow; row <= lastRow; row++) {
      preencherDadosDaPlaca(sheet, row);
    }
  }

  // 2) Atualiza cor/mensagem de status das linhas tocadas nesta edição
  for (var row4 = firstRow; row4 <= lastRow && row4 <= sheet.getLastRow(); row4++) {
    atualizarLinha(sheet, row4);
  }

  // 3) Edição tocou a coluna Saida -> move a linha para "Já Saiu" (por último, porque apaga linhas)
  if (colunaTocada(col, numCols, COL_SAIDA)) {
    var linhasParaMover = [];
    for (var row2 = firstRow; row2 <= lastRow; row2++) {
      if (sheet.getRange(row2, COL_SAIDA).getValue()) linhasParaMover.push(row2);
    }
    linhasParaMover.sort(function(a, b) { return b - a; }); // de baixo para cima
    linhasParaMover.forEach(function(row3) {
      moverParaJaSaiu(sheet, row3);
    });
  }
}

function colunaTocada(col, numCols, colAlvo) {
  return col <= colAlvo && (col + numCols - 1) >= colAlvo;
}

function preencherDadosDaPlaca(sheet, row) {
  var placaCell = sheet.getRange(row, COL_PLACA);
  var placa = String(placaCell.getValue() || '').trim().toUpperCase();
  if (!placa) return;

  var entradaCell = sheet.getRange(row, COL_ENTRADA);
  if (!entradaCell.getValue()) {
    entradaCell.setValue(new Date());
    entradaCell.setNumberFormat('hh:mm');
  }

  var dados = buscarNaBase(placa);
  if (dados) {
    sheet.getRange(row, COL_TIPO_VEICULO).setValue(dados.tipo);
    sheet.getRange(row, COL_TRANSPORTADORA).setValue(dados.transportadora);
  }
}

// Só lê Placa (col A), Tipo (col D) e Transportadora (col E) da Base — não toca na coluna de nome
function buscarNaBase(placaNormalizada) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseSheet = ss.getSheetByName(SHEET_BASE);
  if (!baseSheet) return null;

  var primeiraLinhaDado = BASE_LINHA_CABECALHO + 1;
  var lastRow = baseSheet.getLastRow();
  if (lastRow < primeiraLinhaDado) return null;

  var dados = baseSheet.getRange(primeiraLinhaDado, 1, lastRow - primeiraLinhaDado + 1, 5).getValues(); // A:E
  for (var i = 0; i < dados.length; i++) {
    var placaBase = String(dados[i][0] || '').trim().toUpperCase();
    if (placaBase === placaNormalizada) {
      return {
        tipo: dados[i][3],           // coluna D da Base = TIPO DE VEICULO
        transportadora: dados[i][4]  // coluna E da Base = TRANSPORTADORA
      };
    }
  }
  return null;
}

// Calcula minutos no pátio e aplica cor + mensagem na linha.
// Regras: 15-20min amarelo "⚠️" | 20-25min laranja "Atenção⚠️" | 25min+ vermelho "Possível ofensora⚠️"
function atualizarLinha(sheet, row) {
  var entrada = sheet.getRange(row, COL_ENTRADA).getValue();
  var saida = sheet.getRange(row, COL_SAIDA).getValue();
  var statusCell = sheet.getRange(row, COL_STATUS);
  var linhaRange = sheet.getRange(row, 1, 1, ULTIMA_COLUNA_PARA_PINTAR);

  if (!entrada || saida) return;
  if (!(entrada instanceof Date)) return;

  var minutos = (new Date().getTime() - entrada.getTime()) / 60000;
  if (minutos < 0 || minutos > 24 * 60) return; // segurança contra dado antigo sem data real

  var cor = null, mensagem = '', corTexto = '#000000';
  if (minutos >= 25) {
    cor = '#FF0000'; mensagem = 'Possível ofensora⚠️'; corTexto = '#FFFFFF';
  } else if (minutos >= 20) {
    cor = '#FF9900'; mensagem = 'Atenção⚠️';
  } else if (minutos >= 15) {
    cor = '#FFFF00'; mensagem = '⚠️';
  }

  statusCell.setValue(mensagem);
  linhaRange.setBackground(cor);
  linhaRange.setFontColor(corTexto);
}

// Roda em todas as linhas — chamada pelo gatilho de tempo (a cada 1 minuto)
function atualizarStatusPatio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  var lastRow = sheet.getLastRow();
  for (var row = 2; row <= lastRow; row++) {
    atualizarLinha(sheet, row);
  }
}

function moverParaJaSaiu(sheet, row) {
  var ss = sheet.getParent();
  var doneSheet = ss.getSheetByName(SHEET_DONE);

  if (!doneSheet) {
    doneSheet = ss.insertSheet(SHEET_DONE);
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    doneSheet.getRange(1, 1, 1, sheet.getLastColumn()).setValues(headerRange.getValues());
    doneSheet.getRange(1, sheet.getLastColumn() + 1).setValue('Registrado em');
  }

  var rowRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  var valores = rowRange.getValues()[0];
  valores[COL_MOTORISTA_INDEX_ZERO] = ''; // garante que nome não seja copiado, se ainda houver algo digitado manualmente
  valores.push(new Date());

  doneSheet.appendRow(valores);
  sheet.deleteRow(row);
}
var COL_MOTORISTA_INDEX_ZERO = 3; // coluna D, índice 0-based, usado só para zerar antes de arquivar

function criarTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'atualizarStatusPatio') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('atualizarStatusPatio')
    .timeBased()
    .everyMinutes(1)
    .create();
}

// ====== WEB APP (ponte para o painel HTML) ======
function doGet(e) {
  var callback = e.parameter.callback;
  var json = JSON.stringify(buildPayload());
  if (callback) {
    var out = ContentService.createTextOutput(callback + '(' + json + ');');
    out.setMimeType(ContentService.MimeType.JAVASCRIPT);
    return out;
  }
  var out2 = ContentService.createTextOutput(json);
  out2.setMimeType(ContentService.MimeType.JSON);
  return out2;
}

function buildPayload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var dados = [];
  if (lastRow >= 2) {
    var valores = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    valores.forEach(function (row) {
      var placa = row[COL_PLACA - 1];
      var entrada = row[COL_ENTRADA - 1];
      var saida = row[COL_SAIDA - 1];
      if (!placa) return;
      if (!(entrada instanceof Date)) return;
      if (saida) return;

      dados.push({
        placa: String(placa),
        vaga: row[2] != null ? String(row[2]) : '',
        tipo: row[COL_TIPO_VEICULO - 1] || '',
        transportadora: row[COL_TRANSPORTADORA - 1] || '',
        entradaISO: entrada.toISOString()
        // motorista propositalmente omitido
      });
    });
  }
  return { caminhoes: dados, geradoEm: new Date().toISOString() };
}
```

## Passos

1. `Extensões > Apps Script` nesta planilha → apague o conteúdo → cole o código acima → salve.
2. Rode `criarTrigger` uma vez (▶ Executar, selecionando essa função) para o status atualizar sozinho a cada minuto.
3. `Implantar > Nova implantação > App da Web` → Executar como "Eu" → Quem tem acesso: como essa é conta pessoal, pode escolher **"Qualquer pessoa"** (não precisa mais restringir a domínio) → Implantar → copie a URL `/exec`.
4. No `index.html`, troque `WEBAPP_URL` por essa nova URL, e remova a coluna/campo "Motorista" da tabela e dos cards (posso fazer essa parte do `index.html` agora se você confirmar que é pra essa planilha nova que o painel vai apontar).
