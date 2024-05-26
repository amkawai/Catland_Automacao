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
var main_worksheet = "Planilha Geral";
var title_menu = 'Preenchimento Miautomático de Miautualizações';
var text_insert_messages = 'Inserir Mensagens';
var instructions_text_messages = ['Copie a mensagem de Atualização do Whatsapp e cole aqui', 'Somente atualizações de Local, Status, Gênero, Raça e Cor estão disponíveis. Deixe em branco e aperte OK quando finalizar'];

// Informações de Log
var log_worksheet_title = "Logs";
var log_cols = ['Data e Hora', 'Log', 'Mensagem Original'];
var log_error_text_color = '#f652a0'; // meio avermelhado
var log_regular_text_color ='#4c5270' // azul escuro

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
var valid_bool_list = ["Sim", "Não", "Pendente", "N/A"];
var valid_fiv_felv_list =["Positivo", "Negativo", "Pendente", "N/A"];
var valid_vaccine_list = ["V3", "V4", "V5", "Pendente", "N/A"];
var valid_profile_list = ["Arisco","Assustado", "Brincalhão", "Carinhoso", "Dócil", "Dorminhoco", "Neutro", "Temperamental", "Tímido",
                        "Tranquilo", "Pendente","N/A"];
var valid_returned_list = ["Sim", "Não", "N/A"];

var col_num_dict ={
                    "location": 1,
                    "status": 2,
                    "azure_code": 3,            // deprecated
                    "code": 4,
                    "name": 5,
                    "entry NGO": 6,
                    "exit NGO": 7,
                    "time_span_at_NGO": 8,      // not used in any formula, this is calculated directly on the spreadsheet
                    "vaccination_card": 9,
                    "birthdate": 10,
                    "age": 11,                  // not used in any formula, this is calculated directly on the spreadsheet
                    "sex": 12,
                    "race": 13,
                    "color": 14,
                    "neuter_date": 15,
                    "fiv": 16,
                    "felv": 17,
                    "fiv_felv_test_date": 18,
                    "vaccine_2nd_dose_date": 19,
                    "vaccinationType": 20,
                    "vaccine_renewal_status": 21,
                    "rabies_date": 22,
                    "rabies_renewal_status": 23,
                    "lifeStory": 24,
                    "family": 25,
                    "notes": 26,
                    "microchip": 27,
                    "interaction_animals": 28,
                    "interaction_humans": 29,
                    "profile": 30,
                    "returned_status": 31,
                    "returned_date": 32,
                    "returned_reason": 33,
                    "health_notes": 34,
                    "admin_notes": 35
}

// Global array to store messages
var messages = [];



/**
 * Adds a custom menu to the Google Sheets UI.
 */


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(title_menu)
    .addItem(text_insert_messages, 'showPrompt')
    .addToUi();
}


/**
 * Shows a prompt to enter messages. Collects messages until a blank message is entered.
 */
function showPrompt() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    instructions_text_messages[0],
    instructions_text_messages[1],
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
  var logSheet = ss.getSheetByName(log_worksheet_title);
  
  if (!logSheet) {
    // Create the Logs sheet if it doesn't exist
    logSheet = ss.insertSheet(log_worksheet_title);
    logSheet.appendRow(log_cols);
  }

  // Append the log message with a timestamp and the original message
  var timestamp = new Date();
  var row = logSheet.appendRow([timestamp, logMessage, originalMessage || '']);

  // Apply highlight text color if it is an error log
  var lastRow = logSheet.getLastRow();
  if (isError) {
    logSheet.getRange(lastRow, 1, 1, logSheet.getLastColumn()).setFontColor(log_error_text_color);
  }
  else {
    logSheet.getRange(lastRow, 1, 1, logSheet.getLastColumn()).setFontColor(log_regular_text_color);
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
  var logSheet = ss.getSheetByName(log_worksheet_title);
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
  var lookahead = '(?=\\sLocal|\\sStatus|\\sCód\\.\\sSimplesvet|\\sNome|\\sNovo\\sNome|\\sEntrada\\sONG|\\sSaída\\sONG|\\sCarteirinha\\sde\\sVacinação|\\sData nasc\\.|\\sGênero|\\sRaça|\\sCor|\\sCastração|\\sFIV|\\sFELV|\\sData\\sTeste\\sFIV\\se\\sFELV|\\sData\\s2ª\\sDose\\sVacina|\\sTipo\\sde\\sVacina|\\sRenovação\\sVacina|\\sRenovação\\s\\sVacina|\\sRaiva|\\sRenovação\\sRaiva|\\sHistória|\\sFamília|\\sObservação|\\sMicrochip|\\sInterações\\scom\\soutros\\sanimais|\\sInteração\\scom\\sHumanos|\\sPerfil|\\sDevolvido|\\sData\\sda\\sDevolução|\\sMotivo\\sda\\sDevolução|\\sObs\\.\\sSaúde|\\sObs\\.\\sAdministrativo|\\s\\w+:|$)';
  
  // Define property patterns with the common positive lookahead assertion
  //var nomePattern = new RegExp('Nome:\\s*([^]+?)(?=\\s(?:Novo\\sNome:|Local:|Raça:|Cor:|Cód\\.\\sSimplesvet:|Sexo?:|Status:|$))');
  var nomePattern = new RegExp('Nome:\\s*([^]+?)' + lookahead);
  var codigoPattern = new RegExp('Cód\\. Simplesvet:\\s*(\\d+)' + lookahead);
  var statusPattern = new RegExp('Status:\\s*(' + validStatuses.join('|') + ')');
  var locationPattern = new RegExp('Local:\\s*(' + validLocations.join('|') + ')');
  var racePattern = new RegExp('Raça:\\s*(' + validRaces.join('|') + ')');
  var colorPattern = new RegExp('Cor:\\s*(' + validColors.join('|') + ')');
  var sexPattern = new RegExp('Gênero:\\s*(' + validSexes.join('|') + ')');
  var vaccinationPattern = new RegExp('Carteirinha de Vacinação:\\s*(' + valid_bool_list.join('|') + ')');
  var novoNomePattern = new RegExp('Novo nome:\\s*([^]+?)' + lookahead);
  var fivPattern = new RegExp('FIV:\\s*(' + valid_fiv_felv_list.join('|') + ')');
  var felvPattern = new RegExp('FELV:\\s*(' + valid_fiv_felv_list.join('|') + ')');
  var vaccineTypePattern = new RegExp('Tipo de Vacina:\\s*(' + valid_vaccine_list.join('|') + ')');
  var lifeStoryPattern = new RegExp('História:\\s*([^]+?)' + lookahead);
  var familyPattern = new RegExp('Família:\\s*([^]+?)' + lookahead);

  
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
  
  if (!nomeMatch || !codigoMatch) {
    logToSheet('Nome e/ou Cód Simplesvet não encontrados: nome: ' + nomeMatch[1] + ", codigo: " + codigoMatch[1], message, true);
    Logger.log('Nome e/ou Cód Simplesvet não encontrados: ' + message);
    return;
  }

  var nome = nomeMatch[1].trim();
  var codigo = codigoMatch[1].trim();


  // Access the "Planilha Geral" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(main_worksheet);
  if (!sheet) {
    logToSheet('Planilha não encontrada: '+ main_worksheet, message, true);
    Logger.log('Planilha não encontrada: '+ main_worksheet + message);
    return;
  }

  var data = sheet.getDataRange().getValues();
  // Find the row with the matching code in column D
  for (var i = 0; i < data.length; i++) {
    if (data[i][col_num_dict["code"]-1] == codigo) {    //indices start from 0 in getDataRange(), but not in getRange()!
      //Check if name is correct
      if (data[i][col_num_dict["name"]-1] != nome){
        logToSheet('Nome na mensagem não bate com nome na planilha. Atualização: ' + nome + ", Mensagem: " + data[i][col_num_dict["name"]-1], message, true);
        Logger.log('Nome na mensagem não bate com nome na planilha. Atualização: ' + nome + ", Mensagem: " + data[i][col_num_dict["name"]-1] + " : " + message);
        return;
      }
      
      if (locationMatch) {
        update_field(
                      parameterMatch = locationMatch, message = message, list_valid_options = validLocations, row_number = i, column_number = col_num_dict["location"],
                      error_log_text = "Local inválido", success_log_text = "Local atualizado na linha", sheet = sheet
                      )
        
      }

      if (statusMatch){
        update_field(
                      parameterMatch = statusMatch, message = message, list_valid_options = validStatuses, row_number = i, column_number = col_num_dict["status"],
                      error_log_text = "Status inválido", success_log_text = "Status atualizado na linha", sheet = sheet
                      )
        
      }

      if (sexMatch) {
        update_field(
                      parameterMatch = sexMatch, message = message, list_valid_options = validSexes, row_number = i, column_number = col_num_dict["sex"],
                      error_log_text = "Gênero inválido", success_log_text = "Gênero atualizado na linha", sheet = sheet
                      )
        
      }
      if (raceMatch) {
        update_field(
                      parameterMatch = raceMatch, message = message, list_valid_options = validRaces, row_number = i, column_number = col_num_dict["race"],
                      error_log_text = "Raça inválida", success_log_text = "Raça atualizada na linha", sheet = sheet
                      )

      }
      if (colorMatch) {
        update_field(
                      parameterMatch = colorMatch, message = message, list_valid_options = validColors, row_number = i, column_number = col_num_dict["color"],
                      error_log_text = "Cor inválida", success_log_text = "Cor atualizada na linha", sheet = sheet
                      )
        
      }

      if (vaccinationMatch) {
        update_field(
                      parameterMatch = vaccinationMatch, message = message, list_valid_options = valid_bool_list, row_number = i, column_number = col_num_dict["vaccination_card"],
                      error_log_text = "Vacinação inválida", success_log_text = "Vacinação atualizada na linha", sheet = sheet
                      )
      }

      if (novoNomeMatch) {
        update_field(
                      parameterMatch = novoNomeMatch, message = message, list_valid_options = [], row_number = i, column_number = col_num_dict["name"],
                      error_log_text = "Novo nome inválido", success_log_text = "Nome atualizado na linha", sheet = sheet
                      )
      }

      if (fivMatch) {
        update_field(
                      parameterMatch = fivMatch, message = message, list_valid_options = valid_fiv_felv_list, row_number = i, column_number = col_num_dict["fiv"],
                      error_log_text = "FIV inválido", success_log_text = "FIV atualizado na linha", sheet = sheet
                      )
      }

      if (felvMatch) {
        update_field(
                      parameterMatch = felvMatch, message = message, list_valid_options = valid_fiv_felv_list, row_number = i, column_number = col_num_dict["felv"],
                      error_log_text = "FELV inválido", success_log_text = "FELV atualizado na linha", sheet = sheet
                      )
      }

      if (vaccineTypeMatch) {
        update_field(
                      parameterMatch = vaccineTypeMatch, message = message, list_valid_options = valid_vaccine_list, row_number = i, column_number = col_num_dict["vaccinationType"],
                      error_log_text = "Tipo de Vacina inválido", success_log_text = "Tipo de Vacina atualizado na linha", sheet = sheet
                      )
      }

      if (lifeStoryMatch) {
        update_field(
                      parameterMatch = lifeStoryMatch, message = message, list_valid_options = [], row_number = i, column_number = col_num_dict["lifeStory"],
                      error_log_text = "História inválida", success_log_text = "História atualizada na linha", sheet = sheet
                      )
      }

      if (familyMatch) {
        update_field(
                      parameterMatch = familyMatch, message = message, list_valid_options = [], row_number = i, column_number = col_num_dict["family"],
                      error_log_text = "Família inválida", success_log_text = "Família atualizada na linha", sheet = sheet
                      )
      }
    

      logToSheet('Atualização concluída', message, false);
      return;
    }
  }

  logToSheet('Código não encontrado', message, true);
}


function update_field(
                    parameterMatch, message, list_valid_options = [], row_number, column_number, error_log_text = "Erro", success_log_text = "Sucesso",
                    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(main_worksheet)
                    ){
  var parameter = parameterMatch[1].trim();
  // Capitalize the first letter of the color
  parameter = parameter.charAt(0).toUpperCase() + parameter.slice(1);

  if (parameter == "Lt eterno") {
          parameter = "LT Eterno"
  }

  // Validate parameter
  if (list_valid_options.length > 0 && !list_valid_options.includes(parameter)){
    logToSheet(error_log_text + ": " + parameter, message, true);
    Logger.log(error_log_text + ": " + parameter + ": " + message);
    return;
  }
  sheet.getRange(row_number + 1, column_number).setValue(parameter);
  logToSheet(success_log_text + " " + (row_number + 1), message, false);
  Logger.log(success_log_text + " " + (row_number + 1) + ": " + message);
}



















