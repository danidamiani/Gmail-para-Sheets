function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("DD´s Menu")
  .addItem('Buscar e-mails', 'getGmailEmails')
  .addItem('Caverna', 'caverna')
  .addToUi();
}


// Caputa os e-mails, através de uma categoria criada no Gmail, e os lista na primeira página da planilha
function getGmailEmails(){
  // Referência a Guia da planilha que irá receber os e-mails
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página1");
  
  // Ativa a guia
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet1);

  // Limpa a guia
  sheet1.getRange(1, 1, sheet1.getMaxRows(), sheet1.getMaxColumns()).activate();
  sheet1.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});  
  
  // Se posiciona na célula A1
  sheet1.getRange('A1').activate();

  // Interage com o gmail pegando os 500 e-mails mais recentes presentes na categoria criada dentro do gmail
  var label = GmailApp.getUserLabelByName('SAC');
  var threads = label.getThreads();
  
  // Looping nos e-mails
  for (var i=0 ; i<threads.length-1; i++) {
	// Captura o item do e-mail (contém inclusive as respostas)
    var messages = threads[i].getMessages();
	
	// captura o títudo do e-mail
    var subject = messages[0].getSubject();
	
	// captura o corpo do e-mail e elimina tags html presente n mesmo através da função eliminaTagsHtml
    var body = eliminaTagsHtml(messages[0].getPlainBody());
	
	// captura a data do e-mail
    var dataTime = messages[0].getDate();
	
	// se o corpo tiver mais de 100 caracteres, elimia os posteriores
    if (body.length>100) body = body.substring(0,99);
	
	// Acrescenta a linha na sheet com o e-mail correspondente
    sheet1.appendRow([i, subject,body, dataTime ]);
  }
}

// Função para remove os tags HTML de um string
function eliminaTagsHtml (param1) {
	if (param1==null) param1="";
	var resposta = "";
	var flag=true;
	for (var i=0;i<param1.length;i++){
		var letra = param1.charAt(i);
		var codAsc = param1.charCodeAt(i);
		if (codAsc == 60) flag=false;			// LETRA "<"
		if (flag==true){
			if (codAsc == 10) {
				resposta = resposta + " ";
			} else if (codAsc == 13){
				resposta = resposta;
			} else if (codAsc == 8211){
				resposta = resposta + "-";
			} else {
				resposta = resposta + letra;
			}
		}
		if (codAsc == 62) flag=true;			// LETRA ">"
	}
	return resposta;
}

// Serve  apenas para verificar o número de itens que retorna na funação getThreads
function caverna() {
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página2");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet2);
  var label = GmailApp.getUserLabelByName('SAC');
  var threads = label.getThreads();
  sheet2.appendRow([threads.length]);
}
