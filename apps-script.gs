// =======================================================
// COLE ESTE CÓDIGO NO EDITOR DO APPS SCRIPT
// (Extensões > Apps Script dentro da planilha)
// =======================================================

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var respostas = sheet.getSheetByName("Respostas");
  var motoristas = sheet.getSheetByName("Motoristas");

  var data = JSON.parse(e.postData.contents);
  var id = String(data.id || "");
  var horaVisualizacao = data.horaVisualizacao || "";
  var horaVerificacao = data.horaVerificacao || "";
  var dataStr = data.data || "";

  // Busca o nome do motorista na aba "Motoristas"
  var nome = "";
  if (motoristas) {
    var motoristasData = motoristas.getDataRange().getValues();
    for (var i = 1; i < motoristasData.length; i++) {
      if (String(motoristasData[i][0]).trim() === id.trim()) {
        nome = motoristasData[i][1];
        break;
      }
    }
  }
  if (!nome) nome = "ID: " + id;

  // Adiciona linha na aba "Respostas"
  respostas.appendRow([id, nome, dataStr, horaVisualizacao, horaVerificacao]);

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", msg: "Verificação Sábado Shopee ativa" }))
    .setMimeType(ContentService.MimeType.JSON);
}
