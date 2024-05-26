function onEdit() {
atualizaData()
}
 
function atualizaData() {
  var ActiveSheet= SpreadsheetApp.getActiveSheet();
  var capture = ActiveSheet.getActiveCell();
  
  if(ActiveSheet.getName() == 'Planilha Geral ' && capture.getColumn() == 2 ) {
    var add = capture.offset(0, 32);
    var data = new Date(); 
    data = Utilities.formatDate(data, "UTC-03:00", "dd/MM/yyyy");
    add.setValue(data);
  };
}



//Automação de Preenchimento
// Variáveis de ambiente
var mainWorksheet = "Planilha Geral";
var titleMenu = 'Preenchimento Miautomático de Miautualizações';
var textInsertMessages = 'Inserir Mensagens';
var instructionsTextMessages = ['Copie a mensagem de Atualização do Whatsapp e cole aqui', 'Somente atualizações de Local, Status, Gênero, Raça e Cor estão disponíveis. Deixe em branco e aperte OK quando finalizar'];

// Informações de Log
var logWorksheetTitle = "Logs";
var logCols = ['Data e Hora', 'Log', 'Mensagem Original'];
var logErrorTextColor = '#f652a0'; // meio avermelhado
var logRegularTextColor ='#4c5270' // azul escuro

// Validações de Parâmetros
var validLocations = ["Cafofinho", "Hospital Externo", "Petz", "Protetor Parceiro", "Quarentena Externa", "Sede", "Pendente", "Ronron Cat Café"];
var validStatuses = ["adotado", "Adotado", "arquivado", "Arquivado", "ced", "CED", "disponível", "Disponível", "estrelinha", "Estrelinha", "indisponível", "Indisponível", "lt eterno", "LT Eterno", "LT eterno", "pendente", "Pendente"];
var validSexes = ["Macho", "Fêmea", "Pendente", "N/A"];
var validRaces = ["SRD", "Persa", "Pendente", "N/A"];
var validColors = ["Amarelo e Branco", "Amarelo", "Bege e Branco", "Bege e Cinza", "Bege e Escaminha", "Bege e Marrom", "Bege e Tigrado",
                            "Bege, Branco e Marrom", "Bege, Cinza e Marrom", "Bege, Cinza e Preto", "Bege, Escaminha e Marrom", "Bege, Marrom e Branco",
                            "Bege, Marrom e Preto", "Bege", "Blue Point", "Branco e Amarelo", "Branco e Bege", "Branco e Cinza", "Branco e Escaminha",
                            "Branco e Laranja", "Branco e Marrom", "Branco e Preto", "Branco e Siberiano", "Branco e Tigrado", "Branco e Tricolor",
                            "Branco, Bege e Cinza", "Branco, Bege e Marrom", "Branco, Cinza e Bege", "Branco, Marrom e Preto", "Branco, Marrom e Tigrado",
                            "Branco, Preto e Marrom", "Branco", "Caramelo e Preto", "Caramelo", "Cinza e Bege", "Cinza e Branco", "Cinza e Escaminha",
                            "Cinza e Tigrado", "Cinza Escuro", "Cinza, Branco e Marrom", "Cinza", "Escaminha e Bege", "Escaminha", "Laranja e Branco",
                            "Laranja Tigrado e Branco", "Laranja, Branco e Preto", "Laranja", "Marrom e Bege", "Marrom e Branco", "Marrom e Preto",
                            "Marrom, Bege e Branco", "Marrom", "Preto e Branco", "Preto e Laranja", "Preto e Marrom", "Preto Fumaça",
                            "Preto, Amarelo e Branco", "Preto, Branco e Marrom", "Preto, Marrom e Bege", "Preto", "Red Point", "Siamês Siberiano",
                            "Siberiano", "Tigrado e Branco", "Tigrado e Cinza", "Tigrado e Tricolor", "Tigrado", "Tricolor e Tigrado", "Tricolor",
                            "Pendente", "N/A"];
var validBoolList = ["Sim", "Não", "Pendente", "N/A"];
var validFivFelvList =["Positivo", "Negativo", "Pendente", "N/A"];
var validVaccineList = ["V3", "V4", "V5", "Pendente", "N/A"];
var validProfileList = ["Arisco","Assustado", "Brincalhão", "Carinhoso", "Dócil", "Dorminhoco", "Neutro", "Temperamental", "Tímido",
                        "Tranquilo", "Pendente","N/A"];
var validReturnedList = ["Sim", "Não", "N/A"];

var colNumDict ={
                    "location": 1,
                    "status": 2,
                    "azureCode": 3,            // deprecated
                    "code": 4,
                    "name": 5,
                    "entryNGO": 6,
                    "exitNGO": 7,
                    "timeSpanAtNGO": 8,      // not used in any formula, this is calculated directly on the spreadsheet
                    "vaccinationCard": 9,
                    "birthdate": 10,
                    "age": 11,                  // not used in any formula, this is calculated directly on the spreadsheet
                    "sex": 12,
                    "race": 13,
                    "color": 14,
                    "neuterDate": 15,
                    "fiv": 16,
                    "felv": 17,
                    "fivFelvTestDate": 18,
                    "vaccine2ndDoseDate": 19,
                    "vaccinationType": 20,
                    "vaccineRenewalStatus": 21,
                    "rabiesDate": 22,
                    "rabiesRenewalStatus": 23,
                    "lifeStory": 24,
                    "family": 25,
                    "notes": 26,
                    "microchip": 27,
                    "interactionAnimals": 28,
                    "interactionHumans": 29,
                    "profile": 30,
                    "returnedStatus": 31,
                    "returnedDate": 32,
                    "returnedReason": 33,
                    "healthNotes": 34,
                    "adminNotes": 35
}

// Global array to store messages
var messages = [];



/**
 * Adds a custom menu to the Google Sheets UI.
 */


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(titleMenu)
    .addItem(textInsertMessages, 'showPrompt')
    .addToUi();
}


/**
 * Shows a prompt to enter messages. Collects messages until a blank message is entered.
 */
function showPrompt() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    instructionsTextMessages[0],
    instructionsTextMessages[1],
    ui.ButtonSet.OK_CANCEL
  );

  // Process the input
  if (result.getSelectedButton() == ui.Button.OK) {
    var message = result.getResponseText();
    if (message) {
      messages.push(message);
      Logger.log("Message added: "+ message);
      // Prompt again
      showPrompt();
    } else {
      // No more messages, process all collected messages
      processMessages();
    }
  }
}


/**
 * Logs a message to the "Logs" sheet.
 * 
 * @param {string} logMessage - The log message to be recorded.
 * @param {string} originalMessage - The original message that was processed.
 * @param {boolean} isError - Whether the log is an error message.
 */
function logToSheet(logMessage = "Sem mensagem de Log", originalMessage = "Sem mensagem disponível", isError = false) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(logWorksheetTitle);
  
  if (!logSheet) {
    // Create the Logs sheet if it doesn't exist
    logSheet = ss.insertSheet(logWorksheetTitle);
    logSheet.appendRow(logCols);
  }

  // Append the log message with a timestamp and the original message
  var timestamp = new Date();
  var row = logSheet.appendRow([timestamp, logMessage, originalMessage || '']);

  // Apply highlight text color if it is an error log
  var lastRow = logSheet.getLastRow();
  if (isError) {
    logSheet.getRange(lastRow, 1, 1, logSheet.getLastColumn()).setFontColor(logErrorTextColor);
  }
  else {
    logSheet.getRange(lastRow, 1, 1, logSheet.getLastColumn()).setFontColor(logRegularTextColor);
  }
}

/**
 * Processes all collected messages by calling processMessage on each one.
 */
function processMessages() {
  messages.forEach(function(message) {
    processMessage(message);
  });
  // Clear messages after processing
  messages = [];
  // Switch to Logs sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(logWorksheetTitle);
  if (logSheet) {
    ss.setActiveSheet(logSheet);
  }
}

/**
 * Processes a message to update specific fields in the "Planilha Geral" sheet.
 * 
 * @param {string} message - The message text containing updates.
 */
function processMessage(message) {
  // Define a common positive lookahead assertion for properties
  var lookahead = '(?=\\sLocal|\\sStatus|\\sCód\\.\\sSimplesvet|\\sNome|\\sNovo\\sNome|\\sEntrada\\sONG|\\sSaída\\sONG|\\sCarteirinha\\sde\\sVacinação|\\sData nasc\\.|\\sGênero|\\sRaça|\\sCor|\\sCastração|\\sFIV|\\sFELV|\\sData\\sTeste\\sFIV\\se\\sFELV|\\sData\\s2ª\\sDose\\sVacina|\\sTipo\\sde\\sVacina|\\sRenovação\\sVacina|\\sRenovação\\s\\sVacina|\\sRaiva|\\sRenovação\\sRaiva|\\sHistória|\\sFamília|\\sObservação|\\sMicrochip|\\sInterações\\scom\\soutros\\sanimais|\\sInteração\\scom\\sHumanos|\\sPerfil|\\sDevolvido|\\sDevolução|\\sData\\sda\\sDevolução|\\sMotivo\\sda\\sDevolução|\\sMotivo|\\sObs\\.\\sSaúde|\\sObs\\.\\sAdministrativo|\\s\\w+:|$)';
  
  // Define property patterns with the common positive lookahead assertion
  //var nomePattern = new RegExp('Nome:\\s*([^]+?)(?=\\s(?:Novo\\sNome:|Local:|Raça:|Cor:|Cód\\.\\sSimplesvet:|Sexo?:|Status:|$))');
  var nomePattern = new RegExp('Nome:\\s*([^]+?)' + lookahead);
  var codigoPattern = new RegExp('Cód\\. Simplesvet:\\s*(\\d+)' + lookahead);
  var statusPattern = new RegExp('Status:\\s*(' + validStatuses.join('|') + ')');
  var locationPattern = new RegExp('Local:\\s*(' + validLocations.join('|') + ')');
  var racePattern = new RegExp('Raça:\\s*(' + validRaces.join('|') + ')');
  var colorPattern = new RegExp('Cor:\\s*(' + validColors.join('|') + ')');
  var sexPattern = new RegExp('Gênero:\\s*(' + validSexes.join('|') + ')');
  var vaccinationPattern = new RegExp('Carteirinha de Vacinação:\\s*(' + validBoolList.join('|') + ')');
  var novoNomePattern = new RegExp('Novo nome:\\s*([^]+?)' + lookahead);
  var fivPattern = new RegExp('FIV:\\s*(' + validFivFelvList.join('|') + ')');
  var felvPattern = new RegExp('FELV:\\s*(' + validFivFelvList.join('|') + ')');
  var vaccineTypePattern = new RegExp('Tipo de Vacina:\\s*(' + validVaccineList.join('|') + ')');
  var lifeStoryPattern = new RegExp('História:\\s*([^]+?)' + lookahead);
  var familyPattern = new RegExp('Família:\\s*([^]+?)' + lookahead);
  var notesPattern = new RegExp('Observação:\\s*([^]+?)' + lookahead);
  var microchipPattern = new RegExp('Microchip:\\s*([^]+?)' + lookahead);
  var interactionAnimalsPattern = new RegExp('Interações com outros animais:\\s*(' + validBoolList.join('|') + ')');
  var interactionHumansPattern = new RegExp('Interação com Humanos:\\s*(' + validBoolList.join('|') + ')');
  var profilePattern = new RegExp('Perfil:\\s*(' + validProfileList.join('|') + ')');
  var reasonReturnedPattern = new RegExp('Motivo:\\s*([^]+?)' + lookahead);
  var healthNotesPattern = new RegExp('Obs\\. Saúde:\\s*([^]+?)' + lookahead);
  var adminNotesPattern = new RegExp('Obs\\. Administrativo:\\s*([^]+?)' + lookahead);
  
  var nomeMatch = message.match(nomePattern);
  var codigoMatch = message.match(codigoPattern);
  var locationMatch = message.match(locationPattern);
  var sexMatch = message.match(sexPattern);
  var raceMatch = message.match(racePattern);
  var colorMatch = message.match(colorPattern);
  var statusMatch = message.match(statusPattern);
  var vaccinationMatch = message.match(vaccinationPattern);
  var novoNomeMatch = message.match(novoNomePattern);
  var fivMatch = message.match(fivPattern);
  var felvMatch = message.match(felvPattern);
  var vaccineTypeMatch = message.match(vaccineTypePattern);
  var lifeStoryMatch = message.match(lifeStoryPattern);
  var familyMatch = message.match(familyPattern);
  var notesMatch = message.match(notesPattern);
  var microchipMatch = message.match(microchipPattern);
  var interactionAnimalsMatch = message.match(interactionAnimalsPattern);
  var interactionHumansMatch = message.match(interactionHumansPattern);
  var profileMatch = message.match(profilePattern);
  var reasonReturnedMatch = message.match(reasonReturnedPattern);
  var healthNotesMatch = message.match(healthNotesPattern);
  var adminNotesMatch = message.match(adminNotesPattern);
  
  if (!nomeMatch || !codigoMatch) {
    logToSheet('Nome e/ou Cód Simplesvet não encontrados: nome: ' + nomeMatch[1] + ", codigo: " + codigoMatch[1], message, true);
    Logger.log('Nome e/ou Cód Simplesvet não encontrados: ' + message);
    return;
  }

  var nome = nomeMatch[1].trim();
  var codigo = codigoMatch[1].trim();


  // Access the "Planilha Geral" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(mainWorksheet);
  if (!sheet) {
    logToSheet('Planilha não encontrada: '+ mainWorksheet, message, true);
    Logger.log('Planilha não encontrada: '+ mainWorksheet + message);
    return;
  }

  var data = sheet.getDataRange().getValues();
  // Find the row with the matching code in column D
  for (var i = 0; i < data.length; i++) {
    if (data[i][colNumDict["code"]-1] == codigo) {    //indices start from 0 in getDataRange(), but not in getRange()!
      //Check if name is correct
      if (data[i][colNumDict["name"]-1] != nome){
        logToSheet('Nome na mensagem não bate com nome na planilha. Atualização: ' + nome + ", Mensagem: " + data[i][colNumDict["name"]-1], message, true);
        Logger.log('Nome na mensagem não bate com nome na planilha. Atualização: ' + nome + ", Mensagem: " + data[i][colNumDict["name"]-1] + " : " + message);
        return;
      }
      
      if (locationMatch) {
        updateField(
                      parameterMatch = locationMatch, message = message, listValidOptions = validLocations, rowNumber = i, columnNumber = colNumDict["location"],
                      errorLogText = "Local inválido", successLogText = "Local atualizado na linha", sheet = sheet
                      )
        
      }

      if (statusMatch){
        updateField(
                      parameterMatch = statusMatch, message = message, listValidOptions = validStatuses, rowNumber = i, columnNumber = colNumDict["status"],
                      errorLogText = "Status inválido", successLogText = "Status atualizado na linha", sheet = sheet
                      )
        
      }

      if (sexMatch) {
        updateField(
                      parameterMatch = sexMatch, message = message, listValidOptions = validSexes, rowNumber = i, columnNumber = colNumDict["sex"],
                      errorLogText = "Gênero inválido", successLogText = "Gênero atualizado na linha", sheet = sheet
                      )
        
      }
      if (raceMatch) {
        updateField(
                      parameterMatch = raceMatch, message = message, listValidOptions = validRaces, rowNumber = i, columnNumber = colNumDict["race"],
                      errorLogText = "Raça inválida", successLogText = "Raça atualizada na linha", sheet = sheet
                      )

      }
      if (colorMatch) {
        updateField(
                      parameterMatch = colorMatch, message = message, listValidOptions = validColors, rowNumber = i, columnNumber = colNumDict["color"],
                      errorLogText = "Cor inválida", successLogText = "Cor atualizada na linha", sheet = sheet
                      )
        
      }

      if (vaccinationMatch) {
        updateField(
                      parameterMatch = vaccinationMatch, message = message, listValidOptions = validBoolList, rowNumber = i, columnNumber = colNumDict["vaccinationCard"],
                      errorLogText = "Vacinação inválida", successLogText = "Vacinação atualizada na linha", sheet = sheet
                      )
      }

      if (novoNomeMatch) {
        updateField(
                      parameterMatch = novoNomeMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["name"],
                      errorLogText = "Novo nome inválido", successLogText = "Nome atualizado na linha", sheet = sheet
                      )
      }

      if (fivMatch) {
        updateField(
                      parameterMatch = fivMatch, message = message, listValidOptions = validFivFelvList, rowNumber = i, columnNumber = colNumDict["fiv"],
                      errorLogText = "FIV inválido", successLogText = "FIV atualizado na linha", sheet = sheet
                      )
      }

      if (felvMatch) {
        updateField(
                      parameterMatch = felvMatch, message = message, listValidOptions = validFivFelvList, rowNumber = i, columnNumber = colNumDict["felv"],
                      errorLogText = "FELV inválido", successLogText = "FELV atualizado na linha", sheet = sheet
                      )
      }

      if (vaccineTypeMatch) {
        updateField(
                      parameterMatch = vaccineTypeMatch, message = message, listValidOptions = validVaccineList, rowNumber = i, columnNumber = colNumDict["vaccinationType"],
                      errorLogText = "Tipo de Vacina inválido", successLogText = "Tipo de Vacina atualizado na linha", sheet = sheet
                      )
      }

      if (lifeStoryMatch) {
        updateField(
                      parameterMatch = lifeStoryMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["lifeStory"],
                      errorLogText = "História inválida", successLogText = "História atualizada na linha", sheet = sheet
                      )
      }

      if (familyMatch) {
        updateField(
                      parameterMatch = familyMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["family"],
                      errorLogText = "Família inválida", successLogText = "Família atualizada na linha", sheet = sheet
                      )
      }

      if (notesMatch) {
        updateField(
                      parameterMatch = notesMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["notes"],
                      errorLogText = "Observação inválida", successLogText = "Observação atualizada na linha", sheet = sheet
                      )
      }
      if (microchipMatch) {
        updateField(
                      parameterMatch = microchipMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["microchip"],
                      errorLogText = "Microchip inválido", successLogText = "Microchip atualizado na linha", sheet = sheet
                      )
      }
      if (interactionAnimalsMatch) {
        updateField(
                      parameterMatch = interactionAnimalsMatch, message = message, listValidOptions = validBoolList, rowNumber = i, columnNumber = colNumDict["interactionAnimals"],
                      errorLogText = "Interações com outros Animais inválido", successLogText = "Interações com outros Animais atualizado na linha", sheet = sheet
                      )
      }

      if (interactionHumansMatch) {
        updateField(
                      parameterMatch = interactionHumansMatch, message = message, listValidOptions = validBoolList, rowNumber = i, columnNumber = colNumDict["interactionHumans"],
                      errorLogText = "Interação com Humanos inválido", successLogText = "Interação com Humanos atualizado na linha", sheet = sheet
                      )
      }
      if (profileMatch) {
        updateField(
                      parameterMatch = profileMatch, message = message, listValidOptions = validProfileList, rowNumber = i, columnNumber = colNumDict["profile"],
                      errorLogText = "Perfil inválido", successLogText = "Perfil atualizado na linha", sheet = sheet
                      )                
      }

      if (reasonReturnedMatch) {
        updateField(
                      parameterMatch = reasonReturnedMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["returnedReason"],
                      errorLogText = "Motivo inválido", successLogText = "Motivo atualizado na linha", sheet = sheet
                      )
        returnedStatusMatch = ["Sim","Sim"];      // weird, I know. This way to simulate other match variables
        updateField(
                      parameterMatch = returnedStatusMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["returnedStatus"],
                      errorLogText = "Devolução inválida", successLogText = "Devolução atualizada na linha", sheet = sheet
                      )         
      }

      if (healthNotesMatch) {
        updateField(
                      parameterMatch = healthNotesMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["healthNotes"],
                      errorLogText = "Obs Saúde inválido", successLogText = "Obs Saúde atualizado na linha", sheet = sheet
                      )                
      }

      if (adminNotesMatch) {
        updateField(
                      parameterMatch = adminNotesMatch, message = message, listValidOptions = [], rowNumber = i, columnNumber = colNumDict["adminNotes"],
                      errorLogText = "Perfil inválido", successLogText = "Perfil atualizado na linha", sheet = sheet
                      )                
      }

      logToSheet('Atualização concluída', message, false);
      return;
    }
  }

  logToSheet('Código não encontrado', message, true);
}


function updateField(
                    parameterMatch, message, listValidOptions = [], rowNumber, columnNumber, errorLogText = "Erro", successLogText = "Sucesso",
                    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainWorksheet)
                    ){
  var parameter = parameterMatch[1].trim();
  // Capitalize the first letter of the color
  parameter = parameter.charAt(0).toUpperCase() + parameter.slice(1);

  if (parameter == "Lt eterno") {
          parameter = "LT Eterno"
  }

  // Validate parameter
  if (listValidOptions.length > 0 && !listValidOptions.includes(parameter)){
    logToSheet(errorLogText + ": " + parameter, message, true);
    Logger.log(errorLogText + ": " + parameter + ": " + message);
    return;
  }
  sheet.getRange(rowNumber + 1, columnNumber).setValue(parameter);
  logToSheet(successLogText + " " + (rowNumber + 1), message, false);
  Logger.log(successLogText + " " + (rowNumber + 1) + ": " + message);
}



















