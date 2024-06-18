function enviarCorreos() {
  const env = env_();
  const datosSheet = SpreadsheetApp.openById(env.ID_DATABASE).getSheetByName(env.SHEET_NAME);
  const datosRange = datosSheet.getDataRange();
  const datosValues = datosRange.getValues();
  
  const headers = datosValues[0];
  const PARA_INDEX = headers.indexOf("PARA");
  const CC_INDEX = headers.indexOf("CC");
  const CCO_INDEX = headers.indexOf("CCO");
  const ASUNTO_INDEX = headers.indexOf("ASUNTO");
  const ENCABEZADO_INDEX = headers.indexOf("ENCABEZADO");
  const EMPRESA_PARA_INDEX = headers.indexOf("EMPRESA_PARA");
  const TEMPLATE_ID_INDEX = headers.indexOf("TEMPLATE_ID");
  const TITULO_CAJA_UNO_INDEX = headers.indexOf("TITULO_CAJA_UNO");
  const MENSAJE_CAJA_UNO_INDEX = headers.indexOf("MENSAJE_CAJA_UNO");
  const TITULO_CAJA_DOS_INDEX = headers.indexOf("TITULO_CAJA_DOS");
  const MENSAJE_CAJA_DOS_INDEX = headers.indexOf("MENSAJE_CAJA_DOS");
  const TITULO_CAJA_TRES_INDEX = headers.indexOf("TITULO_CAJA_TRES");
  const MENSAJE_CAJA_TRES_INDEX = headers.indexOf("MENSAJE_CAJA_TRES");
  const LINK_PDF_INDEX = headers.indexOf("LINK_PDF"); 

  for (let i = 1; i < datosValues.length; i++) {
    const para = datosValues[i][PARA_INDEX];
    const cc = datosValues[i][CC_INDEX];
    const cco = datosValues[i][CCO_INDEX];
    const asunto = datosValues[i][ASUNTO_INDEX];
    const encabezado = datosValues[i][ENCABEZADO_INDEX];
    const empresaPara = datosValues[i][EMPRESA_PARA_INDEX];
    const templateId = datosValues[i][TEMPLATE_ID_INDEX];
    const tituloCaja1 = datosValues[i][TITULO_CAJA_UNO_INDEX];
    const mensajeCaja1 = datosValues[i][MENSAJE_CAJA_UNO_INDEX];
    const tituloCaja2 = datosValues[i][TITULO_CAJA_DOS_INDEX];
    const mensajeCaja2 = datosValues[i][MENSAJE_CAJA_DOS_INDEX];
    const tituloCaja3 = datosValues[i][TITULO_CAJA_TRES_INDEX];
    const mensajeCaja3 = datosValues[i][MENSAJE_CAJA_TRES_INDEX];
    const linkPDF = datosValues[i][LINK_PDF_INDEX];

    if (para && templateId) {
      let templateFile;
      if (templateId === 'TEMPLATE_HTML_ID_1') {
        templateFile = 'mailTemplateOne';
      } else if (templateId === 'TEMPLATE_HTML_ID_2') {
        templateFile = 'mailTemplateTwo';
      } else {
        continue; 
      }
      
      const template = HtmlService.createTemplateFromFile(templateFile);
      const emailData = {
        nombreCompleto: encabezado,
        empresaPara: empresaPara,
        tituloCaja1: tituloCaja1,
        mensajeCaja1: mensajeCaja1,
        tituloCaja2: tituloCaja2,
        mensajeCaja2: mensajeCaja2,
        tituloCaja3: tituloCaja3,
        mensajeCaja3: mensajeCaja3
      };
      template.data = emailData;
      const htmlBody = template.evaluate().getContent();
      
      console.log("Correo enviado exitosamente a: " + para, "Con el asunto: ", asunto); 

      const pdfBlob = getPDFBlob(linkPDF);
      
      GmailApp.sendEmail(para, asunto, '', {
        cc: cc,
        bcc: cco,
        htmlBody: htmlBody,
        attachments: [pdfBlob]
      });
    }
  }
}

function getPDFBlob(linkPDF) {
  const fileId = getIdFromUrl(linkPDF);
  const file = DriveApp.getFileById(fileId);
  const pdfBlob = file.getBlob();
  return pdfBlob;
}

function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}
