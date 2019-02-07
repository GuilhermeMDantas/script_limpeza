# Esboço para o script de limpeza do Access

import os, os.path                              # Diretórios e arquivos
import shutil                                   # Cópia de segurança antes dos procedimentos
import win32com.client                          # Execução do compact and repair
import time
import glob                                     # Auxilia a execução do compact and repair
import zipfile                                  # Zipar/Compactar
import logging                                  # Logs
from datetime import datetime                   # Salva a data nos arquivos
import smtplib                                  # Email
from email.mime.multipart import MIMEMultipart  # Email
from email.mime.text import MIMEText            # Email
from email.mime.base import MIMEBase            # Email
from email import encoders                      # Email
#import socks                                    # proxy

#                   To do
#   1 - TIRAR AS CREDENCIAIS DO MEIO DO SCRIPT !!!
#   2 - Definir os error codes
#   3 - Enviar os emails com os erros

#########################################
#                                       #
#   Globais para lidar com as funções   #
#                                       #
#########################################

flag = True
while flag:
    access_db_path = input('Digite o caminho completo dos arquivos (ex.: C:\\Users\\Bruno\\Documents): ')
    try:
        os.chdir(access_db_path)    # Muda para o diretório com os arquivos
    except OSError:
        logging.error('O diretório digitado não existe')
        continue
    flag = False

data_de_hoje = str(datetime.now().date().strftime("%d-%m-%Y"))  # Data de Hoje
bk_str = '_BACKUP_'                                             # Praticidade
delete_copy_fail = False                                        # Caso falhe a remoção das cópias

#####################################################
#                                                   #
#           Configuração para os logs               #
#                                                   #
#####################################################

def log_config():
    logging.basicConfig(
            filename= data_de_hoje + ' cleaning.log',
            filemode='w+',
            format='%(asctime)s - %(levelname)-8s:: %(message)s',
            datefmt='%d/%m/%Y %H:%M:%S',
            level=logging.INFO)

    logging.debug('logging.config set')
    logging.info('Caminho definido para {}'.format(access_db_path))


#############################
#                           #
#   Código começa aqui      #
#                           #
#############################

def main():    
    
    # verifica se há arquivos .accdb no diretório
    logging.info('Encontrando arquivos...')
    file_list = get_files()
    

    if not file_list:
        logging.error('Não há arquivos .accdb ou .mdb no diretório')
        return 1  # Identificador de erro
    logging.info('Arquivos adicionados na lista')
    
            
    #####################################################
    #                                                   #
    #   Checa se algum arquivo NÃO pode ser manipulado  #
    #                                                   #
    #####################################################

    # Checa se algum dos arquivos na lista está sendo usado por outro programa
    logging.info('Verificando por arquivos bloqueados...')
    if isBlocked(file_list):
        logging.error('Access já está aberto')
        return 2  # Identificador de erro

    logging.info('Nenhum arquivo está bloqueado')


    #############################################
    #                                           #
    #   Faz uma cópia dos arquivos encontrados  #
    #                                           #
    #############################################

    logging.info('Fazendo cópia dos arquivos encontrados...')
    if copia(file_list):
        logging.info('Cópia dos arquivos feita')
    else:
        logging.error('Não foi possível fazer cópias dos arquivos')
        return 3  # Identificador de erro


    #############################
    #                           #
    #   Compact and Repair      #
    #                           #
    #############################

    # Executa o compact and Repair do Access
    logging.info('Começando processo de limpeza...')
    if compact_repair():
        logging.info('Os arquivos foram limpos com sucesso')
    else:
        logging.critical("Ocorreu um erro durante o processo de limpeza")
        logging.critical("Os Arquivos serão substituídos pelas cópias")        
        
        # Pega todos as cópias feitas antes do processo de limpeza e joga nessa lista
        logging.info('Procurando pelos backups...')
        lista_de_backup = get_files(True)  # get_backup = True
        
        # O script NÃO deveria ter chegado até aqui caso não houvessem backups, mas caso isso aconteça
        # Ele irá garantir que há backups que possam ser usados
        # E caso não irá abortar para garantir que os arquivos, mesmo corrompidos, não sejam deletados
        if not lista_de_backup:
            logging.critical('Não foram encontrados backups')
            logging.critical('Para evitar maiores perdas, o script irá abortar sem substituir os arquivos')
            return 4
        logging.info('Backups encontrados')
        
        # Remove TODOS os arquivos não _BACKUP_ e .accdb da pasta e substitui pelas cópias
        logging.info('Deletando arquivos corrompidos...')
        for file in file_list:
            
            # Primeiro pula os backups
            if file.endswith(bk_str):
                continue

            # Deleta o arquivo
            logging.debug('Deleting file \'{}\''.format(file))
            os.remove(file)
            logging.debug('Deleted')

            # Renomeia o backup para o nome original do arquivo
            for backup in lista_de_backup:
                if backup.replace(bk_str, '') == file:
                    logging.debug('Renaming \'{}\' to \'{}\''.format(backup, file))
                    os.rename(backup, file)
        
        logging.info('Arquivos corrompidos deletados.')
        logging.info('Arquivos corrompidos foram substituídos por seus backups')
        return 5  # Identificador do erro
    

    #####################################################################
    #                                                                   #
    #   Exclui as cópias feitas após uma limpeza concluída com sucesso  #
    #                                                                   #
    #####################################################################

    # Pega as cópias
    logging.info('Pegando cópias...')
    copies_list = get_files(True)  # get_backup = True
    
    # Necessário porque se não o python cria uma nova variavel
    global delete_copy_fail

    if not copies_list:
        delete_copy_fail = True
    else:
        logging.info('Cópias encontradas')
    
    # Exclui as cópias
    if not delete_copy_fail:
        logging.info('Deletando cópias...')
        if delete_copies(copies_list):
            logging.info('Cópias deletadas')
        else:
             logging.error('Houve um erro durante a exclusão das cópias')
             delete_copy_fail = True
    else:
        logging.warning('Devido à um erro, não foi possivel encontrar as cópias')
        
        
    
    #####################################################################
    #                                                                   #
    #   Adiciona os arquivos, agora limpos, na lista para compactar     #
    #                                                                   #
    #####################################################################

    # Limpa a lista
    logging.info('Limpando lista...')
    file_list.clear()
    logging.info('Lista limpa')

    # Adiciona os arquivos limpos
    logging.info('Criando nova lista...')
    file_list = get_files()

    if not file_list:
        logging.error('Devido à um erro não foi possível encontrar os arquivos')
        return 6  # Identificador de erro
    logging.info('Lista criada')
    

    #################################
    #                               #
    #   Zipa os arquivos limpos     #
    #                               #
    #################################

    logging.info('Compactando arquivos...')
    return_value = zipar(file_list)    
    
    if return_value == -1:
        logging.critical('Devido à um erro, não foi possível compactar os arquivos')
        return 7  # Identificador de erro
    elif return_value == 1:
        logging.error("Durante o processo de backup alguma database foi aberta e não foi possível compacta-lás")
        return 8  # Identificador de erro
    else:
        logging.info('Todos os arquivos foram compactados com sucesso')
        
    return 0  # Script rodou sem erros

    

#################################################################
#                                                               #
#                       Função get_files()                      #
#   adiciona os arquivos do diretório numa lista e retorna      #
#                                                               #
#################################################################

def get_files(get_backup = False):
    logging.debug('get_files()')

    # Adiciona os arquivos na lista
    # Ignora arquivos _BACKUP_

    if not get_backup: # Pegar arquivos NÃO backup
        logging.debug('get regular files')
        file_list = [file
                    for file in os.listdir()
                        # Para garantir que é um arquivo .accdb ou .mdb
                        if file.endswith('.accdb') or file.endswith('.mdb')
                            # Para garantir que NÃO é um _BACKUP_
                            if not file.endswith(bk_str + '.accdb') and not file.endswith(bk_str + '.mdb')
                    ]
    else: # Pegar backups
        logging.debug('get backups')
        file_list = [backup
                     for backup in os.listdir()
                        # Para garantir que é um arquivo .accdb
                        if backup.endswith('.accdb') or backup.endswith('.mdb')
                            # Para garantir que É um _BACKUP_
                            if backup.endswith(bk_str + '.accdb') or backup.endswith(bk_str + '.mdb')
                     ]

    # Sem {extension} no diretorio
    # Em outras palavras, lista vazia
    if not file_list:
        logging.debug('List is empty')
        logging.debug('end of get_files()')
        return None

    # Logging, quais arquivos foram achados/adicionados na lista
    for file in file_list:
        logging.debug('Got \'{}\''.format(file))

    # retorna lista > com arquivos <
    logging.debug('end of get_files()')
    return file_list


#########################################################
#                                                       #
#                   Função copia()                      #
#   Faz uma cópia dos arquivos que vão ser manipulados  #
#                                                       #
#########################################################

def copia(file_list):
    logging.debug('copia()')
    
    try:
        for file in file_list:
            # Irá ignorar o arquivo se uma cópia dele já existir
            if file.endswith('.accdb'):
                new_file = file.replace('.accdb', bk_str + '.accdb')
                
                if os.path.exists(file.replace('.accdb', bk_str + '.accdb')):
                    logging.debug('Backup for \'{}\' already exists. Skipping'.format(file))
                    continue

            elif file.endswith('.mdb'):
                new_file = file.replace('.mdb', bk_str + '.mdb')
                
                if os.path.exists(file.replace('.mdb', bk_str + '.mdb')):
                    logging.debug('Backup for \'{}\' already exists. Skipping'.format(file))
                    continue
        
            # Remove o .accdb ou .mdb do final, adiciona '_BACKUP_.extensão' no arquivo
            logging.debug('Making copy of \'{}\''.format(file))
            if file.endswith('.accdb'):
                shutil.copyfile(file, new_file)
                logging.debug('Copy \'{}\' created'.format(file.replace('.accdb', bk_str + '.accdb')))
            elif file.endswith('.mdb'):
                shutil.copyfile(file, new_file)
                logging.debug('Copy \'{}\' created'.format(file.replace('.mdb', bk_str + '.mdb')))
            
    except Exception:
        # Se de alguma forma, o if acima nao conseguir pegar algum arquivo repetido
        # A função irá retornar false devido a algum erro na hora da cópia
        logging.debug('end of copia()')
        logging.exception('ALGO DEU ERRADO NA CRIAÇÃO DAS CÓPIAS')
        return False

    logging.debug('end of copia()')
    return True


#################################################################
#                                                               #
#                   Função isBlocked()                          #
#                   Retorna True ou False                       #
#   De acordo com se o arquivo está aberto (em uso) ou não      #
#                                                               #
#################################################################

def isBlocked(file_list):
    logging.debug('isBlocked()')

    # Cria instância do Access
    logging.debug('Creating Access instance')
    access_instance = win32com.client.Dispatch('Access.Application')
    
    # A função que é chamada no main()
    for file in file_list:

        # Coloca o caminho completo do arquivo
        file = access_db_path + '\\' + file

        # Chama a blocked check(file) para testar se o arquivo tem senha/pode ser aberto
        logging.debug('calling blocked_check()')
        if blocked_check(file, access_instance):
            # Access não está aberto
            continue
        else:
            # Access já está aberto
            # Desvincula a variavel do Access Object
            access_instance = None
            return False
    

    # Fecha a instância do Access que foi aberta na checagem
    logging.debug('Closing Access instance')
    access_instance.Application.Quit(2)
    logging.debug('Closed')
    # Desvincula a variavel do Access Object
    logging.debug('Unlinking variable to the Access Object')
    access_instance = None
    logging.debug('Unlinked')

    logging.debug('end of isBlocked()')
    return False


#########################################
#                                       #
#       Função blocked_check()          #
#   Função suplementar à isBlocked()    #
#                                       #
#########################################

def blocked_check(file, access_instance):
    logging.debug('blocked_check()')

    try:
        # Tenta abrir a database
        logging.debug('Checking if \'{}\' is blocked...'.format(file))        
        access_instance.Application.OpenCurrentDatabase(file)
        logging.debug('Database sucessfully opened.')
        
        logging.debug('Closing database...')
        access_instance.Application.CloseCurrentDatabase()
        logging.debug('Closed')
    except Exception:
        logging.debug('end of blocked_check()')
        logging.exception('Access está aberto')
        return False

    logging.debug('end of blocked_check()')
    return True


#########################################################
#                                                       #
#               Função compact_repair()                 #
#   Executa a limpeza e reparo dos arquivos do Access   #
#                                                       #
#########################################################

def compact_repair():
    logging.debug('compact_repair()')
    
    # 'Abre' o Access
    logging.debug('Creating Access instance')
    db = win32com.client.Dispatch('Access.Application')
    logging.debug('Created')
    
    # Executa o compact and repair em todos arquivos .accdb
    logging.debug('STARTING .ACCDB CLEANING')
    for file in glob.glob(access_db_path + '\\*.accdb'):

        # Ira ignorar as cópias feitas anteriormente
        if file.endswith(bk_str + '.accdb'):
            logging.debug('SKIPPING FILE \'{}\''.format(file.replace(access_db_path + '\\', '')))
            continue

        # Backup obrigatório do arquivo para se executar o compact and repair
        logging.debug('Making tmp file')
        new_file = file.replace('.accdb', 'teste.accdb')
        #logging.debug('tmp file is \'{}\''.format(new_file.replace(access_db_path + '\\', '')))

        # Compact and Repair
        try:
            logging.debug('Trying to repair \'{}\''.format(file.replace(access_db_path + '\\', '')))
            db.CompactRepair(file, new_file, False)
            logging.debug('Repaired')
        except Exception:
            logging.debug('end of compact_repair()')
            logging.exception('ALGO DEU ERRADO DURANTE A LIMPEZA DAS DATABASES .accdb')
            return False

        # Substitui o arquivo compactado com o original
        # E deleta o arquivo criado no processo
        logging.debug('Deleting tmp copy')
        shutil.copyfile(new_file, file)
        os.remove(new_file)
        logging.debug('Deleted')

    # Executa o compact and repair em todos os arquivos .mdb
    logging.debug('STARTING .MDB CLEANING')
    for file in glob.glob(access_db_path + '\\*.mdb'):

        # Ira ignorar as cópias feitas anteriormente
        if file.endswith(bk_str + '.mdb'):
            logging.debug('SKIPPING FILE \'{}\''.format(file.replace(access_db_path + '\\', '')))
            continue

        # Backup obrigatório do arquivo para se executar o compact and repair
        logging.debug('Making tmp file')
        new_file = file.replace('.mdb', 'teste.mdb')
        #logging.debug('tmp file is \'{}\''.format(new_file.replace(access_db_path + '\\', '')))

        # Compact and Repair
        try:
            logging.debug('Trying to repair \'{}\''.format(file.replace(access_db_path + '\\', '')))
            db.CompactRepair(file, new_file, False)
            logging.debug('Repaired')
        except Exception:
            logging.debug('end of compact_repair()')
            logging.exception('ALGO DEU ERRADO DURANTE A LIMPEZA DAS DATABASES .mdb')
            return False

        # Substitui o arquivo compactado com o original
        # E deleta o arquivo criado no processo
        logging.debug('Deleting tmp copy')
        shutil.copyfile(new_file, file)
        os.remove(new_file)
        logging.debug('Deleted')
        

    # 'Fecha' o Access
    logging.debug('"Closing" Access instance')
    db = None
    logging.debug('Closed')

    logging.debug('end of compact_repair()')
    return True


#############################################################
#                                                           #
#                   Função delete_copies()                  #
#   Deleta as cópias feitas antes da limpeza dos arquivos   #
#                                                           #
#############################################################

def delete_copies(copies_list):
    logging.debug('delete_copies()')

    for copy in copies_list:
        try:
            logging.debug('Trying to delete \'{}\''.format(copy))
            os.remove(copy)
            logging.debug('Deleted')
        except Exception:
            logging.debug('end of delete_copies()')
            logging.exception('ALGO DEU ERRADO DURANTE A EXCLUSÃO DOS BACKUPS')
            return False

    logging.debug('end of delete_copies()')
    return True
    

#########################################################################################
#                                                                                       #
#                                   Função zipar()                                      #
#                       Zipa os arquivos passados como parâmetros                       #
#   IMPORTANTE: Arquivos perdem as permissões quando extraidos ex.: somente leitura     #
#                                                                                       #
#########################################################################################

def zipar(file_list):
    logging.debug('zipar()')
    
    # Zipa os arquivos
    with zipfile.ZipFile("Backup databases " + str(datetime.now().date().strftime("%d-%m-%Y")) + ".zip", 'w') as backup:

        # Testa novamente se algum arquivo está bloqueado antes de zipar
        logging.info('Verificando por arquivos bloqueados...')
        if isBlocked(file_list):
            logging.info('Algum arquivo foi aberto...')
            logging.debug('end of zipar()')
            return 1

        logging.info('Nenhum arquivo está bloqueado')

        try:
            for file in file_list:
                logging.debug('Trying to compact \'{}\''.format(file))
                backup.write(file)
                logging.debug('Compacted')
        except Exception:
            logging.debug('end of zipar()')
            logging.exception('EXCEPTION OCCURED')
            return -1

    logging.debug('end of zipar()')
    return 0


#########################
#   Work in Progress    #
#########################
    
def send_mail(assunto, body, log = None, fail = False):
    logging.debug('send_mail()')
    
    # logging.warning("O Script encontrou um erro e irá mandar um email contendo as informações do erro")
    if fail:
        body += '. Também houve um erro durante a exclusão das cópias dos arquivos'

    de = 'exemplo@gmail.com'
    para = 'exemplo@gmail.com'
    corpo = body
    
    # Definição do email
    email = MIMEMultipart()

    # De
    email['From'] = de
    logging.debug('From set to \'{}\''.format(de))
    # Para
    email['To'] = para
    logging.debug('To set to \'{}\''.format(para))
    # Assunto
    email['Subject'] = assunto
    logging.debug('Subject set to \'{}\''.format(assunto))


    # Adiciona o corpo do email
    email.attach(MIMEText(corpo))
    logging.debug('Body set')

    # Caso não haja .log para ser anexado
    # Esse bloco será ignorado
    if log:
        # anexo do .log
        anexo = open(log, 'r')

        # Anexa o .log no email
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(anexo.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % log)

        # Adiciona o anexo no email
        email.attach(part)
    
    mensagem = email.as_string()
    
    # proxy
    #socks.setdefaultproxy(socks.TYPE, 'ip', port, True, 'login', 'pass')
    #socks.wrapmodule(smtplib)

    # Conecta no host
    server = smtplib.SMTP()
    server.set_debuglevel(1)
    #server.connect('C70v40i.rede.sp')
    #server.login('login', 'pass')
    #server.ehlo()
    #server.starttls()

    # Envia o Email
    #server.sendmail(de, para, mensagem)

    # Sai do host
    #server.quit()

    return


if __name__ == '__main__':
    log_config()

    logging.info('Começando script...')
    return_value = main()
    

    log = data_de_hoje + ' cleaning.log'


#    if return_value == 1:
#        send_mail('SCRIPT DE LIMPEZA: Diretório vazio', 'O diretório digitado no início do script não contem arquivos .accdb', log)
#    elif return_value == 2:
#        send_mail('SCRIPT DE LIMPEZA: Arquivo bloqueado', 'Durante o check de arquivos bloqueados o script encontrou um arquivo bloqueado e abortou', log)
#    elif return_value == 3:
#        send_mail('SCRIPT DE LIMPEZA: Não foi possível criar um backup', 'Durante a cópia dos arquivos houve um erro e não foi possível criar uma cópia', log)
#    elif return_value == 4:
#        send_mail('SCRIPT DE LIMPEZA: URGENTE: ERRO DURANTE O COMPACT AND REPAIR', 'Houve um erro durante a execução do compact and repair e ele não pode completar.\nDurante o processo de substituição dos arquivos corrompidos por seus backups, as cópias não foram encontradas e o script abortou para evitar maiores perdas', log)
#    elif return_value == 5:
#        send_mail('SCRIPT DE LIMPEZA: Erro durante o compact and repair', 'Houve um erro durante o processo de limpeza dos arquivos e ele teve que abortar inesperadamente. Os arquivos foram substituidos por suas cópias feitas antes da limpeza dos arquivos', log)
#    elif return_value == 6:
#        send_mail('SCRIPT DE LIMPEZA: Arquivos não foram encontrados', 'O script não conseguiu encontrar os arquivos após a limpeza e abortou', log)
#    elif return_value == 7:
#        send_mail('SCRIPT DE LIMPEZA: Não foi possível compactar os arquivos', 'Devido a algum erro, não foi possível compactar os arquivos limpos. O script abortou', log)
#    elif return_value == 8:
#        send_mail('SCRIPT DE LIMPEZA: Algum arquivo ficou bloqueado durante a compactação', 'Durante a compactação do script, algum arquivo ficou bloqueado (aberto) e não foi possivel compacta-los', log)
#    else:
#        logging.info('O script conseguiu completar sem nenhum problema')

logging.info('=' * 20 + ' Fim do Script ' + '=' * 20)
logging.shutdown()
