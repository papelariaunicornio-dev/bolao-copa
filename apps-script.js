// Google Apps Script — Bolão Copa do Mundo – Papelaria Unicórnio
// Cole este código em script.google.com vinculado à sua planilha

const SHEET_NAME = 'Palpites';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      new Date(),          // A - Data/hora do envio
      data.pedido,         // B - Número do pedido
      data.jogo,           // C - Jogo (ex: Brasil x Haiti)
      data.golsBrasil,     // D - Gols Brasil
      data.golsHaiti,      // E - Gols Haiti
      data.placar,         // F - Placar formatado (ex: 3 x 1)
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Cabeçalhos
    sheet.appendRow(['Data/Hora', 'Pedido', 'Jogo', 'Gols Brasil', 'Gols Haiti', 'Placar']);
    sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 150);
  }

  return sheet;
}

// Teste manual — execute esta função para verificar se o script está funcionando
function testeManual() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        pedido: 'U123456',
        jogo: 'Brasil x Haiti',
        golsBrasil: 3,
        golsHaiti: 0,
        placar: '3 x 0',
        timestamp: new Date().toISOString(),
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
