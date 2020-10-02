function onOpen() {
SpreadsheetApp.getUi()
  .createMenu('ScriptsAgendamento')
  .addItem('Processar novo período de agendamentos', 'userActionResetByRangesAddresses')
  .addItem('Enviar uma cópia para o e-mail', 'sendEmail')
  .addToUi();
}
function userActionResetByRangesAddresses(){
  var data = new Date();
  var today = data.getDay();
  var monday = 1;
  
  if(today != monday){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('ATENÇÃO!', 'Identificamos que hoje não é uma Segunda-Feira. Deseja continuar?\n\nA continuidade resultará na exclusão dos dados preenchidos e um novo calculo de datas.', ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES) {
      Logger.log('The user clicked "Yes."');
    } else {
      Logger.log('The user clicked "No."');
      return
    }
    
  }else {
    //
  }
   
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangesWeekListSeg = ['A3:A19'];
  var rangesWeekListTer = ['A21:A37'];
  var rangesWeekListQua = ['A39:A55'];
  var rangesWeekListQui = ['A57:A73'];
  var rangesWeekListSex = ['A75:A91'];
  var rangesWeekListSab = ['A93:A109'];
  var rangesWeekListDom = ['A111:A127'];
  var rangesAddressesList = ['C3:C145',
                             'D3:D145',
                             'E3:E145'];
  resetByRangesList_(sheet,
                     rangesAddressesList,
                     rangesWeekListSeg,
                     rangesWeekListTer,
                     rangesWeekListQua,
                     rangesWeekListQui,
                     rangesWeekListSex,
                     rangesWeekListSab,
                     rangesWeekListDom
                    );
}

function resetByRangesList_(
                     sheet, 
                     rangesAddressesList, 
                     rangesWeekListSeg,
                     rangesWeekListTer,
                     rangesWeekListQua,
                     rangesWeekListQui,
                     rangesWeekListSex,
                     rangesWeekListSab,
                     rangesWeekListDom){
  
  var date = new Date()
  //Limpa o conteúdo
  sheet.getRangeList(rangesAddressesList).clearContent();
  //Seta um dia apos o outro
  sheet.getRangeList(rangesWeekListSeg).setValue(date.toLocaleDateString());
  date.setDate(date.getDate() + 1);
  sheet.getRangeList(rangesWeekListTer).setValue(date.toLocaleDateString());
  date.setDate(date.getDate() + 1);
  sheet.getRangeList(rangesWeekListQua).setValue(date.toLocaleDateString());
  date.setDate(date.getDate() + 1);
  sheet.getRangeList(rangesWeekListQui).setValue(date.toLocaleDateString());
  date.setDate(date.getDate() + 1);
  sheet.getRangeList(rangesWeekListSex).setValue(date.toLocaleDateString());
  date.setDate(date.getDate() + 1);
  sheet.getRangeList(rangesWeekListSab).setValue(date.toLocaleDateString());
  date.setDate(date.getDate() + 1);
  sheet.getRangeList(rangesWeekListDom).setValue(date.toLocaleDateString());
}

function sendEmail() {
  var farma = SpreadsheetApp.getUi().prompt("Digite o nome conforme a ABA da planilha informado abaixo").getResponseText();
  var response =  SpreadsheetApp.getUi().prompt("Digite o email para onde será enviado a cópia").getResponseText();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(farma);

  const h1 = ws.getRange("A1").getValue();
  const headers = ws.getRange("A2:E2").getValues();

  const date = headers[0][0];
  const hour = headers[0][1];
  const name = headers[0][2];
  const cpf = headers[0][3];
  const contact = headers[0][4];

  const lr = ws.getLastRow();

  const tableRangeValues = ws.getRange(3, 1, lr-1,5).getDisplayValues();

  const htmlTemplate = HtmlService.createTemplateFromFile("email");
  
  htmlTemplate.farma = farma;
  htmlTemplate.h1 = h1;
  htmlTemplate.date = date;
  htmlTemplate.hour = hour;
  htmlTemplate.name = name;
  htmlTemplate.cpf = cpf;
  htmlTemplate.contact = contact;
  htmlTemplate.tableRangeValues = tableRangeValues;

  const htmlForEmail = htmlTemplate.evaluate().getContent();

  GmailApp.sendEmail(response, "Planilha - Agendamento", "Cópia de agendamento", { htmlBody: htmlForEmail });
}
