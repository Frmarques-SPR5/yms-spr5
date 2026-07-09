# Controle de YMS — GESTÃO YMS - Guarapuava (v3)

## Novidade v3 — 3 faixas de tempo, cor e mensagem controladas pelo script

Antes a cor vinha de uma fórmula (`I`/`J`) + formatação condicional nativa do Sheets. Agora, a pedido, tudo isso passou para dentro do script — ele mesmo calcula os minutos, escreve a mensagem na coluna Status e pinta a linha, a cada edição e também a cada 1 minuto (via gatilho de tempo, veja Passo 5).

| Tempo no pátio | Cor da linha | Mensagem na coluna Status |
|---|---|---|
| até 15 min | sem cor | (vazio) |
| 15 a 20 min | 🟡 amarelo | `⚠️` |
| 20 a 25 min | 🟠 laranja | `Atenção⚠️` |
| a partir de 25 min | 🔴 vermelho | `Possível ofensora⚠️` |

(Você mencionou "de 25 a 30 min" para o vermelho — não me disse o que acontece depois dos 30, então deixei o vermelho valendo a partir de 25 min sem limite superior. Se quiser uma 4ª faixa depois dos 30 min, me diga qual cor/mensagem.)

As fórmulas antigas em I/J e as regras de formatação condicional antigas (`ATRASADO` / `ESTOURANDO O TEMPO`) podem continuar na planilha sem problema — como o texto que elas procuram não vai mais aparecer, elas simplesmente não fazem nada. Pode apagá-las se quiser deixar mais limpo, mas não é obrigatório.

---

# Guia completo (v2, ainda válido para os passos 1 a 4 abaixo)

Não tenho um conector com permissão de escrita no Google Sheets, então não consigo aplicar isso direto na planilha — mas o script abaixo faz tudo automaticamente depois de colado uma única vez.

## Por que Motorista/Tipo de veículo/Transportadora não preencheram

Na v1 eu tinha sugerido fórmulas `VLOOKUP` para essas colunas. Se ficaram em branco, o motivo mais provável é que a Placa digitada na "Programação diária" não é idêntica à Placa cadastrada na "Base" (espaço extra, letra maiúscula/minúscula, ou "O" no lugar de "0" — a Base tem algumas inconsistências assim, ex.: `BAAOG70` vs `BAA0G70`).

Na v2 eu troquei o `VLOOKUP` por lógica dentro do próprio script, com `trim()` + `toUpperCase()` — isso resolve espaço e caixa alta/baixa. Erro de letra "O" por "0" na placa não tem solução automática; se sobrar algum motorista em branco, é provável que seja isso — vale revisar a Base.

## O que o script v2 faz

1. **Placa preenchida (coluna B)** → grava a Entrada (coluna G) com o horário atual (uma vez só, não sobrescreve) e busca na aba Base o Motorista (D), Tipo de veículo (E) e Transportadora (F), gravando como valor fixo.
2. **Saida preenchida (coluna H)** → tira a linha inteira da "Programação diária" e joga para uma aba nova chamada **"Já Saiu"**, com todos os dados congelados (incluindo Status e Tempo finais). A aba "Já Saiu" é criada automaticamente na primeira vez que isso acontecer, e vai acumulando o histórico — nunca é limpa pelo script.

A coluna **J (Tempo)** continua sendo fórmula (só para exibir o tempo decorrido). A coluna **I (Status)**, na v3, passou a ser escrita pelo próprio script (não precisa de fórmula nela — se você colocar uma, o script vai sobrescrever, sem problema).

## Passo 1 — Ativar recálculo automático (se ainda não fez)

`Arquivo > Configurações > Cálculo > Recálculo` → **"Em alteração e a cada minuto"**.

## Passo 2 — Fórmula de Tempo (opcional, só para exibir o tempo decorrido)

Célula **J2** (Tempo no pátio), arrastar até a linha 741:
```
=IF($G2="","",IF($H2="",NOW()-$G2,$H2-$G2))
```

Não coloque fórmula em D, E, F ou I — o script cuida de tudo isso (Motorista/Tipo/Transportadora via busca na Base, Status via cor/mensagem).

## Passo 3 — Lista suspensa de Placa (opcional, evita erro de digitação)

Coluna **B2:B741** → `Dados > Validação de dados` → Lista de um intervalo → `Base!$A$3:$A$1000`.

## Passo 4 — Substituir o script no Apps Script

`Extensões > Apps Script` → apague o script antigo → cole o código abaixo no lugar → salve (💾).

```javascript
var SHEET_MAIN  = 'Programação diária';
var SHEET_BASE  = 'Base';
var SHEET_DONE  = 'Já Saiu';

var COL_PLACA          = 2;  // B
var COL_MOTORISTA      = 4;  // D
var COL_TIPO_VEICULO   = 5;  // E
var COL_TRANSPORTADORA = 6;  // F
var COL_ENTRADA        = 7;  // G
var COL_SAIDA          = 8;  // H
var COL_STATUS         = 9;  // I
var ULTIMA_COLUNA_PARA_PINTAR = 10; // pinta de A até J

function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_MAIN) return;

  var range = e.range;
  var col = range.getColumn();
  var numCols = range.getNumColumns();
  var firstRow = Math.max(range.getRow(), 2); // ignora cabeçalho
  var lastRow = range.getRow() + range.getNumRows() - 1;
  if (lastRow < firstRow) return;

  // 1) Edição tocou a coluna Placa -> preenche Entrada + busca na Base
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
    sheet.getRange(row, COL_MOTORISTA).setValue(dados.motorista);
    sheet.getRange(row, COL_TIPO_VEICULO).setValue(dados.tipo);
    sheet.getRange(row, COL_TRANSPORTADORA).setValue(dados.transportadora);
  }
}

function buscarNaBase(placaNormalizada) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseSheet = ss.getSheetByName(SHEET_BASE);
  if (!baseSheet) return null;

  var lastRow = baseSheet.getLastRow();
  if (lastRow < 3) return null;

  var dados = baseSheet.getRange(3, 1, lastRow - 2, 5).getValues(); // A3:E(last)
  for (var i = 0; i < dados.length; i++) {
    var placaBase = String(dados[i][0] || '').trim().toUpperCase();
    if (placaBase === placaNormalizada) {
      return {
        motorista: dados[i][2],
        tipo: dados[i][3],
        transportadora: dados[i][4]
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

  // sem entrada, ou já saiu (essa linha vai ser movida/removida): não colore
  if (!entrada || saida) return;
  if (!(entrada instanceof Date)) return;

  var minutos = (new Date().getTime() - entrada.getTime()) / 60000;

  // segurança contra dado antigo sem data real (só hora), que dá número absurdo
  if (minutos < 0 || minutos > 24 * 60) return;

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
  var valores = rowRange.getValues()[0]; // pega o VALOR calculado (Status/Tempo), não a fórmula
  valores.push(new Date());

  doneSheet.appendRow(valores);
  sheet.deleteRow(row);
}

// Execute esta função MANUALMENTE UMA VEZ (▶ Executar, selecionando "criarTrigger")
// para instalar o gatilho que roda atualizarStatusPatio a cada 1 minuto.
function criarTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'atualizarStatusPatio') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('atualizarStatusPatio')
    .timeBased()
    .everyMinutes(1)
    .create();
}

// ====== WEB APP (ponte para o painel HTML no GitHub Pages) ======
// A planilha continua PRIVADA (restrita ao domínio). Este endpoint roda com a
// autoridade de quem implantou (você), então consegue ler a planilha mesmo
// sendo restrita, e devolve só os dados de quem ainda está no pátio.
// Implante como Web App com acesso "Qualquer pessoa da Mercadolibre SRL"
// (veja Passo 6 abaixo) — assim só quem logar com conta @mercadolibre.com
// consegue chamar esse endereço.
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
      if (!placa) return;              // sem placa, ignora
      if (!(entrada instanceof Date)) return; // sem entrada real, ignora
      if (saida) return;               // já saiu, não mostra no painel

      dados.push({
        placa: String(placa),
        vaga: row[2] != null ? String(row[2]) : '',
        motorista: row[COL_MOTORISTA - 1] || '',
        tipo: row[COL_TIPO_VEICULO - 1] || '',
        transportadora: row[COL_TRANSPORTADORA - 1] || '',
        entradaISO: entrada.toISOString()
      });
    });
  }
  return { caminhoes: dados, geradoEm: new Date().toISOString() };
}
```

## Passo 6 — Implantar o Web App (para o painel HTML conseguir ler os dados)

No editor do Apps Script:

1. `Implantar` (botão azul, canto superior direito) → `Nova implantação`
2. Clique na engrenagem ⚙️ ao lado de "Selecionar tipo" → escolha **App da Web**
3. Configurar:
   - **Executar como**: Eu (`freddy.marques@mercadolivre.com`)
   - **Quem pode acessar**: **Qualquer pessoa da Mercadolibre SRL** (essa opção existe justamente para não expor a planilha na internet pública — só quem estiver logado com conta @mercadolibre.com consegue chamar o endereço)
4. `Implantar` → autorize se pedir
5. Copie a **URL do app da Web** que aparece (algo como `https://script.google.com/macros/s/AKfycb.../exec`)

Guarde essa URL — ela vai no `index.html` (ver Passo 7, no arquivo do painel).

> Se essa opção "Qualquer pessoa da Mercadolibre SRL" não aparecer disponível, o admin do Workspace também bloqueou implantação de Web Apps por domínio — me avise, aí a rota que resta é publicar o painel no Grid em vez do GitHub.

## Passo 5 — Instalar o gatilho de tempo (só assim a cor muda mesmo sem ninguém editar a planilha)

No editor do Apps Script: no menu de funções (topo, ao lado de ▶ Executar), selecione **`criarTrigger`** → clique em ▶ Executar → autorize se pedir. Isso roda uma vez e instala um gatilho permanente que chama `atualizarStatusPatio` a cada minuto — assim a linha vai de amarelo → laranja → vermelho mesmo que ninguém toque na planilha.

## O que muda no dia a dia

- Você digita a placa → Entrada, Motorista, Tipo de veículo e Transportadora preenchem sozinhos.
- Passa 15 min → linha fica amarela (`⚠️`). Passa 20 min → laranja (`Atenção⚠️`). Passa 25 min → vermelha (`Possível ofensora⚠️`). Isso atualiza a cada minuto sozinho, mesmo sem ninguém tocar na planilha (Passo 5).
- Você preenche a Saida → a linha some da "Programação diária" (fila principal fica só com quem ainda está no pátio) e aparece na aba **"Já Saiu"** com todos os dados congelados (inclusive se ela furou o SLA ou não) e a data/hora em que foi registrada a saída.

## Se algum Motorista/Tipo/Transportadora continuar em branco

Confira a placa correspondente na aba Base — provavelmente tem uma inconsistência de digitação (letra "O" em vez de zero, espaço a mais, placa duplicada com dados diferentes). O script já ignora espaço e caixa alta/baixa, então o que sobrar é erro de caractere mesmo na Base.

## Aviso na coluna K (só um lembrete visual, não muda a automação)

A coluna B continua sendo o campo real de Placa (é nela que o script reage). Para deixar claro que a coluna K não deve receber outra coisa além da placa, coloque um dos dois:

- **Texto direto na célula K1**: `⚠️ Preencher somente a placa`
- **Nota ao passar o mouse** (não ocupa espaço visual): clique na célula K1 → `Inserir > Nota` → escreva `Preencher somente a placa aqui. Não digite nome, transportadora ou outro texto.` → um triângulo vermelho aparece no canto da célula e a mensagem aparece ao passar o mouse.
