// Code.gs
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Envío de Correos')
    .addItem('Abrir Interfaz', 'showUI')
    .addToUi();
}

function showUI() {
  const html = HtmlService.createHtmlOutputFromFile('emailUI')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gestor de Envíos de Correos');
}

// Obtener listado de hojas
function getSheets() {
  return SpreadsheetApp.getActive().getSheets().map(sheet => ({
    name: sheet.getName(),
    id: sheet.getSheetId()
  }));
}

// Obtener columnas de una hoja
function getColumns(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// Sistema de plantillas
function saveTemplate(template) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const templates = JSON.parse(scriptProperties.getProperty('templates') || '{}');
  templates[template.name] = {
    subject: template.subject,
    body: template.body,
    created: new Date().toISOString(),
    author: Session.getActiveUser().getEmail()
  };
  scriptProperties.setProperty('templates', JSON.stringify(templates));
  return true;
}

function getTemplates() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const templates = JSON.parse(scriptProperties.getProperty('templates') || '{}');
  return Object.keys(templates).map(name => ({
    name: name,
    created: templates[name].created,
    author: templates[name].author
  }));
}

function getTemplate(templateName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const templates = JSON.parse(scriptProperties.getProperty('templates') || '{}');
  return templates[templateName] || null;
}

// Procesamiento principal de correos
function sendEmails(config) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(config.sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const logColumn = getLogColumnPosition(sheet);
    
    const filteredData = getFilteredData(config, data);
    
    if (filteredData.length === 0) {
      return { error: 'No se encontraron destinatarios con los criterios actuales' };
    }
    
    filteredData.forEach((row, index) => {
      const emailInfo = composeEmail(row, headers, config);
      
      if (!isValidEmail(emailInfo.to)) {
        console.log(`Email inválido en fila ${index + 1}: ${emailInfo.to}`);
        return;
      }
      
      MailApp.sendEmail({
        to: emailInfo.to,
        cc: emailInfo.cc,
        bcc: emailInfo.bcc,
        subject: emailInfo.subject,
        body: emailInfo.body
      });
      
      logSentEmail(sheet, row, headers, logColumn, emailInfo.to);
    });
    
    return { success: true, count: filteredData.length };
  } catch (e) {
    return { error: e.message };
  }
}

// Funciones de soporte
function getFilteredData(config, data) {
  const headers = data[0];
  return data.filter((row, index) => {
    if (index === 0) return false;
    
    return config.conditions.every(condition => {
      const colIndex = headers.indexOf(condition.column);
      if (colIndex === -1) return false;
      
      try {
        return checkCondition(row[colIndex], condition.condition);
      } catch (e) {
        console.error(`Error en condición: ${condition.column} ${condition.condition}`);
        return false;
      }
    });
  });
}

function checkCondition(value, condition) {
  const operatorMap = {
    '==': (a, b) => a.toString().trim().toLowerCase() === b.toLowerCase(),
    '!=': (a, b) => a.toString().trim().toLowerCase() !== b.toLowerCase(),
    '>': (a, b) => !isNaN(a) && !isNaN(b) ? Number(a) > Number(b) : false,
    '<': (a, b) => !isNaN(a) && !isNaN(b) ? Number(a) < Number(b) : false,
    '>=': (a, b) => !isNaN(a) && !isNaN(b) ? Number(a) >= Number(b) : false,
    '<=': (a, b) => !isNaN(a) && !isNaN(b) ? Number(a) <= Number(b) : false,
    'contains': (a, b) => a.toString().toLowerCase().includes(b.toLowerCase()),
    'startsWith': (a, b) => a.toString().toLowerCase().startsWith(b.toLowerCase()),
    'endsWith': (a, b) => a.toString().toLowerCase().endsWith(b.toLowerCase())
  };
  
  const match = condition.match(/(==|!=|>|<|>=|<=|contains|startsWith|endsWith)\s*(.*)/);
  if (!match) return false;
  
  const [_, operator, operand] = match;
  const cleanOperand = operand.replace(/['"]/g, '');
  
  return operatorMap[operator](value.toString(), cleanOperand);
}

function composeEmail(row, headers, config) {
  return {
    to: row[headers.indexOf(config.emailColumn)],
    cc: config.cc,
    bcc: config.bcc,
    subject: fillTemplate(config.subject, row, headers),
    body: fillTemplate(config.body, row, headers)
  };
}

function fillTemplate(template, row, headers) {
  return template.replace(/\${(.*?)}/g, (match, colName) => {
    const index = headers.indexOf(colName);
    return index !== -1 ? row[index] : match;
  });
}

function getLogColumnPosition(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const logHeader = 'Estado Envío';
  
  if (!headers.includes(logHeader)) {
    sheet.getRange(1, headers.length + 1).setValue(logHeader);
    return headers.length + 1;
  }
  return headers.indexOf(logHeader) + 1;
}

function logSentEmail(sheet, row, headers, logColumn, recipient) {
  const timestamp = new Date();
  const logData = [
    `Enviado por: ${Session.getActiveUser().getEmail()}`,
    `Destinatario: ${recipient}`,
    `Fecha: ${timestamp.toLocaleDateString()}`,
    `Hora: ${timestamp.toLocaleTimeString()}`,
    `Registro: ${timestamp.toISOString()}`
  ].join('\n');
  
  const rowNumber = sheet.getDataRange().getValues().findIndex(r => r.join() === row.join()) + 1;
  sheet.getRange(rowNumber, logColumn).setValue(logData);
}

// Funciones para vista previa y conteo
function getEmailCount(config) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(config.sheetName);
  const data = sheet.getDataRange().getValues();
  return getFilteredData(config, data).length;
}

function getEmailPreview(config) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(config.sheetName);
  const data = sheet.getDataRange().getValues();
  const filteredData = getFilteredData(config, data);
  
  if (filteredData.length === 0) {
    return { error: 'No se encontraron destinatarios con los criterios actuales' };
  }
  
  const firstRow = filteredData[0];
  const headers = data[0];
  
  return {
    to: firstRow[headers.indexOf(config.emailColumn)],
    cc: config.cc,
    bcc: config.bcc,
    subject: fillTemplate(config.subject, firstRow, headers),
    body: fillTemplate(config.body, firstRow, headers)
  };
}

// Función de validación de email
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}