function enviarCorreos() {
  const env = env_();
  const datosSheet = SpreadsheetApp.openById(env.ID_DATABASE).getSheetByName(env.SHEET_NAME);
  const datosRange = datosSheet.getDataRange();
  const datosValues = datosRange.getValues();
  
  // Asume que la primera fila contiene los encabezados
  const headers = datosValues[0];
  const PARA_INDEX = headers.indexOf("PARA");
  const CC_INDEX = headers.indexOf("CC");
  const CCO_INDEX = headers.indexOf("CCO");
  const ASUNTO_INDEX = headers.indexOf("ASUNTO");
  const NOMBRE_INDEX = headers.indexOf("NOMBRE");
  const TEMPLATE_ID_INDEX = headers.indexOf("TEMPLATE_ID");

  for (let i = 1; i < datosValues.length; i++) {
    const para = datosValues[i][PARA_INDEX];
    const cc = datosValues[i][CC_INDEX];
    const cco = datosValues[i][CCO_INDEX];
    const asunto = datosValues[i][ASUNTO_INDEX];
    const nombre = datosValues[i][NOMBRE_INDEX];
    const templateId = datosValues[i][TEMPLATE_ID_INDEX];
    
    console.log("Para:", para);
    console.log("CC:", cc);
    console.log("CCO:", cco);
    console.log("Asunto:", asunto);
    console.log("Nombre:", nombre);
    console.log("Template ID:", templateId);

    if (para && templateId) {
      // Seleccionar la plantilla basada en TEMPLATE_ID
      let templateFile;
      if (templateId === 'TEMPLATE_HTML_ID_1') {
        templateFile = 'mailTemplateOne';
      } else if (templateId === 'TEMPLATE_HTML_ID_2') {
        templateFile = 'mailTemplateTwo';
      } else {
        continue; // Si el TEMPLATE_ID no coincide con ninguna plantilla, salta esta fila
      }
      
      const template = HtmlService.createTemplateFromFile(templateFile);
      const emailData = {
        nombreCompleto: capitalize(nombre)
      };
      template.data = emailData;
      const htmlBody = template.evaluate().getContent();

console.log("Template:", templateFile);
console.log("Email Data:", emailData);
console.log("HTML Body:", htmlBody); 
       
      GmailApp.sendEmail(para, asunto, '', {
        cc: cc,
        bcc: cco,
        htmlBody: htmlBody
      });
      
      console.log("Correo enviado a: " + para);
    }
  }
}

function capitalize(string) {
  return string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
}
