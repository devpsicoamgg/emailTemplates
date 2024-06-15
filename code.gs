function enviarCorreos() {
  // var spreadsheetId =  
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  var data = sheet.getDataRange().getValues();
  

  var headers = data[0];
  var PARA_INDEX = headers.indexOf("PARA");
  var CC_INDEX = headers.indexOf("CC");
  var CCO_INDEX = headers.indexOf("CCO");
  var ASUNTO_INDEX = headers.indexOf("ASUNTO");
  var NOMBRE_INDEX = headers.indexOf("NOMBRE");
  var TEMPLATE_ID_INDEX = headers.indexOf("TEMPLATE_ID");

  for (var i = 1; i < data.length; i++) {
    var para = data[i][PARA_INDEX];
    var cc = data[i][CC_INDEX];
    var cco = data[i][CCO_INDEX];
    var asunto = data[i][ASUNTO_INDEX];
    var nombre = data[i][NOMBRE_INDEX];
    var templateId = data[i][TEMPLATE_ID_INDEX];
    
    if (para && templateId) {
      var templateFile = DriveApp.getFileById(templateId);
      var template = templateFile.getAs('text/html').getDataAsString();

      var htmlBody = template.replace("{{NOMBRE}}", nombre);
      
      MailApp.sendEmail({
        to: para,
        cc: cc,
        bcc: cco,
        subject: asunto,
        htmlBody: htmlBody
      });
      Logger.log("Correo enviado a: " + para);
    }
  }
}
