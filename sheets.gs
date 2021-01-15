function onOpen() {
SpreadsheetApp.getUi()
  .createMenu('Farmacia')
  .addItem('Processar novo período de agendamentos', 'userActionResetByRangesAddresses')
  .addSeparator()
  .addItem('Enviar uma cópia para o e-mail', 'sendEmail')
  .addSeparator()
  .addItem('Adicionar novo Estoque', 'stockadd_')
  .addToUi();
}

function stockadd_(){
  var stockaddsheet = SpreadsheetApp.getActiveSheet();
  var newStock = SpreadsheetApp.getUi().prompt("Por favor! Digite a nova quantidade de Estoque").getResponseText();
  
  stockaddsheet.getRange(3, 9).setFormula('='+newStock+'-COUNTIF(E3:E;"REALIZADO")');
}

function userActionResetByRangesAddresses(){
  var data = new Date();
  var today = data.getDay();   
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangeday = ['A3:A63'];
  var rangesAddressesList = ['C3:C63',
                             'D3:D63',
                             'E3:E63',
                             'F3:F63',
                             'G3:G63',
                             'H3:H63'];
  
  if(today){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('ATENÇÃO!', 'Deseja realmente continuar?\n\nA continuidade resultará na exclusão dos dados preenchidos e um novo calculo da data do dia atual.', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      Logger.log('The user clicked "Yes."');
    } else {
      Logger.log('The user clicked "No."');
      return
    }
    
  }
  
  stock_(sheet);
  resetByRangesList_(sheet,
                     rangesAddressesList,
                     rangeday
                    );
}

function stock_(sheet){
  var dailystock
  var dailyuse
  
  dailystock = sheet.getRange(3, 9).getValue();
  dailyuse   = sheet.getRange(3, 10).getValue();
  
  sheet.getRange(3, 9).setFormula('='+dailystock+'-COUNTIF(E3:E;"REALIZADO")');
  sheet.getRange(3, 10).setFormula('='+dailyuse+'+COUNTIF(E3:E;"REALIZADO")');

  if(dailystock < 0){
    sheet.getRange(3, 9).setFormula('=0-COUNTIF(E3:E;"REALIZADO")');
  } else{
    return;
  }
}

function resetByRangesList_(
                     sheet, 
                     rangesAddressesList,
                     rangeday){
  var date = new Date()
  
  //Limpa o conteúdo
  sheet.getRangeList(rangesAddressesList).clearContent();
  
  //Seta o dia atual
  sheet.getRangeList(rangeday).setValue(date.toLocaleDateString());
  
  
}


function sendEmail() {
  var farma = SpreadsheetApp.getUi().prompt("Digite o nome da farma conforme a ABA da planilha informado abaixo").getResponseText();
  var response =  SpreadsheetApp.getUi().prompt("Digite o email para onde será enviado a cópia").getResponseText();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(farma);

  const h1 = ws.getRange("A1").getValue();
  const headers = ws.getRange("A2:G2").getValues();

  const date = headers[0][0];
  const hour = headers[0][1];
  const name = headers[0][2];
  const contact = headers[0][3];
  const status = headers[0][4];
  const sector = headers[0][5];
  const responsible = headers[0][6];

  const lr = ws.getLastRow();

  const tableRangeValues = ws.getRange(3, 1, lr-1,7).getDisplayValues();

  const htmlTemplate = HtmlService.createTemplateFromFile("email");
  
  htmlTemplate.farma = farma;
  htmlTemplate.h1 = h1;
  htmlTemplate.date = date;
  htmlTemplate.hour = hour;
  htmlTemplate.name = name;
  htmlTemplate.contact = contact;
  htmlTemplate.status = status;
  htmlTemplate.sector = sector;
  htmlTemplate.responsible = responsible;
  htmlTemplate.tableRangeValues = tableRangeValues;

  const htmlForEmail = htmlTemplate.evaluate().getContent();

  GmailApp.sendEmail(response, "Planilha - Teste Exame de Covid-19", "Cópia de agendamento", { htmlBody: htmlForEmail });
}