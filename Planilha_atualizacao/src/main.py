
from src.helper.settings import Settings
from src.helper.file_helper import FileHelper
from src.helper.updater import Updater
from src.helper.google_sheets_connector import GoogleSheetsConnector
from src.helper.logger import Logger

def main():
    n_updates = 0
    n_valid_updates = 0
    n_invalid_updates = 0

    

    # Load settings
    settings = Settings("config/settings.json")
    Logger.log(level="INFO", message="Configurações carregadas.")

    # Removing old log file:
    FileHelper.remove_file_if_exists(settings.log_file_path)
    Logger.log(level="INFO", message="Arquivo de log antigo removido.")
    Logger.initialize_log()

    # Initialize FileHelper and Updater
    file_helper = FileHelper(settings)
    updater = Updater(settings)

    # Read files from the input folder
    Logger.log(level="INFO", message = f"Lendo arquivos da pasta {settings.input_folder_path}...")
    file_helper.read_files()
    file_contents = file_helper.get_contents()
    Logger.log(level="INFO", message = f"Arquivos da pasta `{settings.input_folder_path}` lidos com sucesso.")
    

    # Parse updates from file contents
    Logger.log(level="INFO", message = f"Processando mensagens...")
    for content in file_contents:
        Logger.log(level="INFO", message = f"Mensagem em análise: \n{content}\n")
        updates = updater.parse_updates(content)
        if updates:
            Logger.log(level="INFO", message=f"Atualizações encontradas:{updates}")
        else:
            Logger.log(level="WARNING", message=f"\n###Nenhuma atualização encontrada\n\n")
    
    # Connect to Google Sheets
    Logger.log(level="INFO", message=f"Conectando ao Google Planilhas...")
    try:
        gs_connector = GoogleSheetsConnector(
                                        credentials_path=settings.credentials_path,
                                        spreadsheet_id=settings.google_sheets_id,
                                        sheet_name=settings.google_sheets_tab)
        Logger.log(level="INFO", message=f"Conexão ao Google Planilhas realizada.")
    except Exception as e:
        Logger.log(level="ERROR", message = f"\n###Erro ao conectar ao Google Planilhas:\n\n{e}{Logger.add_horizontal_line()}")
    
    # Checking if names and ids match (this is a common error)
    Logger.log(level="INFO", message=f"Verificando se ids e nomes estão corretos...")
    valid_updates = []
    invalid_updates = []
    for update in updates:
        name_id_check = gs_connector.validate_cat_data(cat_id=update["Cód. Simplesvet"], cat_name= update["Nome"], id_col=settings.simplesvet_col, name_col=settings.name_col)
        if not name_id_check:
            invalid_updates.append(update)
            Logger.log(level = "ERROR", message=f"\n###Nome e id não batem com planilha:\n\n {update['Mensagem Original']}\n\n")
            n_invalid_updates += 1
        if name_id_check:
            valid_updates.append(update)
            Logger.log(level="INFO", message = f"Esta atualização será realizada: {update}")
    Logger.log(level="INFO", message = f"Estas atualizações serão realizadas: {valid_updates}")
    Logger.log(level="WARNING", message = f"\n###Estas atualizações NÃO serão realizadas:\n\n{invalid_updates}\n\n")

    n_updates = len(updates)
    
    # Updating items
    for valid_update in valid_updates:
        row_number = gs_connector.get_row_from_data(pattern = valid_update["Cód. Simplesvet"], col = settings.simplesvet_col)
        if "Local" in valid_update.keys():
            if valid_update["Local"] in settings.valid_local_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.local_col, value=valid_update["Local"])
                Logger.log(level = "INFO", message = f"Local atualizado: {valid_update['Cód. Simplesvet']}: {valid_update['Local']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Local inválido:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Status" in valid_update.keys():
            if valid_update["Status"] in settings.valid_status_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.status_col, value=valid_update["Status"])
                Logger.log(level = "INFO", message = f"Status atualizado: {valid_update['Cód. Simplesvet']}: {valid_update['Status']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Status inválido:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Novo nome" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.name_col, value=valid_update["Novo nome"])
                Logger.log(level = "INFO", message = f"Nome atualizado: {valid_update['Cód. Simplesvet']}: {valid_update['Novo nome']}")
                n_valid_updates += 1

        if "Entrada ONG" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.ngo_entry_col, value = valid_update["Entrada ONG"])
                Logger.log(level = "INFO", message = f"Entrada ONG atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Entrada ONG']}")
                n_valid_updates += 1

        if "Saída ONG" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.ngo_exit_col, value = valid_update["Saída ONG"])
                Logger.log(level = "INFO", message = f"Saída ONG atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Saída ONG']}")
                n_valid_updates += 1

        if "Carteirinha de Vacinação" in valid_update.keys():
            if valid_update["Carteirinha de Vacinação"] in settings.valid_bool_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.vaccination_card_col, value=valid_update["Carteirinha de Vacinação"])
                Logger.log(level = "INFO", message = f"Carteirinha de Vacinação atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Carteirinha de Vacinação']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Carteirinha de Vacinação inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Data nasc." in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.birthday_col, value = valid_update["Data nasc."])
                Logger.log(level = "INFO", message = f"Data nasc. atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Data nasc.']}")
                n_valid_updates += 1

        if "Gênero" in valid_update.keys():
            if valid_update["Gênero"] in settings.valid_sex_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.sex_col, value=valid_update["Gênero"])
                Logger.log(level = "INFO", message = f"Gênero atualizado: {valid_update['Cód. Simplesvet']}: {valid_update['Gênero']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Gênero inválido:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Raça" in valid_update.keys():
            if valid_update["Raça"] in settings.valid_race_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.race_col, value=valid_update["Raça"])
                Logger.log(level = "INFO", message = f"Raça atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Raça']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\#nRaça inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1
        
        if "Cor" in valid_update.keys():
            if valid_update["Cor"] in settings.valid_color_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.color_col, value=valid_update["Cor"])
                Logger.log(level = "INFO", message = f"Cor atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Cor']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Cor inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Castração" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.date_neuter_col, value = valid_update["Castração"])
                Logger.log(level = "INFO", message = f"Castração atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Castração']}")
                n_valid_updates += 1

        if "FIV" in valid_update.keys():
            if valid_update["FIV"] in settings.valid_fiv_felv_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.FIV_col, value=valid_update["FIV"])
                Logger.log(level = "INFO", message = f"FIV atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['FIV']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###FIV inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1
        
        if "FELV" in valid_update.keys():
            if valid_update["FELV"] in settings.valid_fiv_felv_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.FELV_col, value=valid_update["FELV"])
                Logger.log(level = "INFO", message = f"FELV atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['FELV']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###FELV inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Data Teste FIV e FELV" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.date_FIV_FELV_col, value = valid_update["Data Teste FIV e FELV"])
                Logger.log(level = "INFO", message = f"Data Teste FIV e FELV atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Data Teste FIV e FELV']}")
                n_valid_updates += 1

        if "Data 2ª Dose Vacina" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.date_2nd_dose_vaccine_col, value = valid_update["Data 2ª Dose Vacina"])
                Logger.log(level = "INFO", message = f"Data 2ª Dose Vacina atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Data 2ª Dose Vacina']}")
                n_valid_updates += 1

        if "Tipo de Vacina" in valid_update.keys():
            if valid_update["Tipo de Vacina"] in settings.valid_vaccine_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.vaccine_type_col, value=valid_update["Tipo de Vacina"])
                Logger.log(level = "INFO", message = f"Tipo de Vacina atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Tipo de Vacina']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Tipo de Vacina inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Renovação Vacina" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.date_vaccine_renew_col, value = valid_update["Renovação Vacina"])
                Logger.log(level = "INFO", message = f"Renovação Vacina atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Renovação Vacina']}")
                n_valid_updates += 1

        if "Raiva" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.date_rabies_vaccine_col, value = valid_update["Raiva"])
                Logger.log(level = "INFO", message = f"Raiva atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Raiva']}")
                n_valid_updates += 1
        
        if "Renovação Raiva" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.date_rabies_renew_col, value = valid_update["Renovação Raiva"])
                Logger.log(level = "INFO", message = f"Renovação Raiva atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Renovação Raiva']}")
                n_valid_updates += 1
        
        if "História" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.history_col, value = valid_update["História"])
                Logger.log(level = "INFO", message = f"História atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['História']}")
                n_valid_updates += 1
        
        if "Família" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.family_col, value = valid_update["Família"])
                Logger.log(level = "INFO", message = f"Família atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Família']}")
                n_valid_updates += 1

        if "Observação" in valid_update.keys():
            if True:
                gs_connector.append_data_to_cell(row_number = row_number, col_letter=settings.notes_col, additional_value = valid_update["Observação"])
                Logger.log(level = "INFO", message = f"Observação atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Observação']}")
                n_valid_updates += 1

        if "Microchip" in valid_update.keys():
            if valid_update["Microchip"] in settings.valid_bool_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.microchip_col, value=valid_update["Microchip"])
                Logger.log(level = "INFO", message = f"Microchip atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Microchip']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Microchip inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Interações com outros animais" in valid_update.keys():
            if valid_update["Interações com outros animais"] in settings.valid_bool_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.interaction_animals_col, value=valid_update["Interações com outros animais"])
                Logger.log(level = "INFO", message = f"Interações com outros animais atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Interações com outros animais']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Interações com outros animais inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Interação com Humanos" in valid_update.keys():
            if valid_update["Interação com Humanos"] in settings.valid_bool_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.interaction_humans_col, value=valid_update["Interação com Humanos"])
                Logger.log(level = "INFO", message = f"Interação com Humanos atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Interação com Humanos']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Interação com Humanos inválida:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1


        if "Perfil" in valid_update.keys():
            if valid_update["Perfil"] in settings.valid_profile_list:
                gs_connector.update_value(row_number = row_number, col_letter=settings.profile_col, value=valid_update["Perfil"])
                Logger.log(level = "INFO", message = f"Perfil atualizado: {valid_update['Cód. Simplesvet']}: {valid_update['Perfil']}")
                n_valid_updates += 1
            else:
                Logger.log(level = "ERROR", message = f"\n###Perfil inválido:\n\n{update['Mensagem Original']}\n\n")
                n_invalid_updates += 1

        if "Devolução" in valid_update.keys():
            if True:
                gs_connector.update_value(row_number = row_number, col_letter=settings.returned_col, value = "Sim")
                gs_connector.update_value(row_number = row_number, col_letter=settings.reason_returned_col, value = valid_update["Motivo"])
                gs_connector.update_value(row_number = row_number, col_letter=settings.date_returned_col, value = str(valid_update["Devolução"]))
                Logger.log(level = "INFO", message = f"Devolução atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Devolução']} {valid_update['Motivo']}")
                n_valid_updates += 1

        if "Observação saúde" in valid_update.keys():
            if True:
                gs_connector.append_data_to_cell(row_number = row_number, col_letter=settings.health_notes_col, additional_value = valid_update["Observação saúde"])
                Logger.log(level = "INFO", message = f"Observação saúde atualizada: {valid_update['Cód. Simplesvet']}: {valid_update['Observação saúde']}")
                n_valid_updates += 1

    #Finishing Log
    Logger.add_horizontal_line()
    Logger.write_to_log_file(message=f'### Resultados\n')
    Logger.write_to_log_file(message=f'-Alterações realizadas: {n_valid_updates}.\n')
    Logger.write_to_log_file(message=f'-Alterações com problemas: {n_invalid_updates}.\n')
    Logger.write_to_log_file(message=f'Terminado: ')
    Logger.write_timestamp()
    FileHelper.convert_md_to_html(md_file_path = settings.log_file_path, html_file_path= "errors.html")
    FileHelper.open_file_in_browser(url = "errors.html")

    # Cleaning up txt file
    #file_helper.remove_files()
    #file_helper.create_empty_txt_file()


if __name__ == "__main__":
    main()
