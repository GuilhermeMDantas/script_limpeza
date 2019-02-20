# Esboço para o script de limpeza do Access
import os, os.path                              # Diretórios e arquivos
import csv                                      # Ler o .csv para pegar os caminhos e arquivos
import sys                                      # Para o script
import shutil                                   # Cópia de segurança antes dos procedimentos
import win32com.client                          # Execução do compact and repair
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

# .csv com os caminhos e arquivos
with open(r'', newline='', encoding='utf8') as csvfile:
    leitor = csv.reader(csvfile)
    dados_csv = list(leitor)

    # Caminhos
    # Pega o primeiro elemento (caminho) de cada linha
    db_path = [linha[0] for linha in dados_csv]

    # Pega o segundo elemento (arquivo) de cada linha
    arquivos = [linha[1] for linha in dados_csv]

    # Pega o terceiro elemento (caminho de bk) de cada linha
    bk_path = [linha[2] for linha in dados_csv]

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
            level=logging.DEBUG)

    logging.debug('logging.config set')
    logging.info('Caminho definido para {}'.format(db_path))


#############################
#                           #
#   Código começa aqui      #
#                           #
#############################

def main():
            
    #####################################################
    #                                                   #
    #   Checa se algum arquivo NÃO pode ser manipulado  #
    #                                                   #
    #####################################################
    
    # Checa se algum dos arquivos na lista está sendo usado por outro programa
    logging.info('Verificando por arquivos bloqueados...')
    if isBlocked():
        logging.error('Access já está aberto')
        return 2  # Identificador de erro

    logging.info('Nenhum arquivo está bloqueado')


    #############################################
    #                                           #
    #   Faz uma cópia dos arquivos encontrados  #
    #                                           #
    #############################################

    logging.info('Fazendo cópia dos arquivos encontrados...')
    if copia():
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
        return 5  # Identificador do erro
    
    #####################################################################
    #                                                                   #
    #   Exclui as cópias feitas após uma limpeza concluída com sucesso  #
    #                                                                   #
    #####################################################################
    
    # Pega as cópias
    logging.info('Pegando cópias...')
    
    # Necessário porque se não o python cria uma nova variavel
    global delete_copy_fail
    
    # Exclui as cópias
    logging.info('Deletando cópias...')
    if delete_copies():
        logging.info('Cópias deletadas')
    else:
        logging.error('Houve um erro durante a exclusão das cópias')
        delete_copy_fail = True


    #################################
    #                               #
    #   Zipa os arquivos limpos     #
    #                               #
    #################################

    logging.info('Compactando arquivos...')
    return_value = zipar()    
    
    if return_value == -1:
        logging.critical('Devido à um erro, não foi possível compactar os arquivos')
        return 7  # Identificador de erro
    elif return_value == 1:
        logging.error("Durante o processo de backup alguma database foi aberta e não foi possível compacta-lás")
        return 8  # Identificador de erro
    else:
        logging.info('Todos os arquivos foram compactados com sucesso')
        
    return 0  # Script rodou sem erros


#########################################################
#                                                       #
#                   Função copia()                      #
#   Faz uma cópia dos arquivos que vão ser manipulados  #
#                                                       #
#########################################################

def copia():
    logging.debug('copia()')

    # Iterador
    i = 0
    
    try:
        # Por caminho em caminhos
        for caminho in db_path:

            # Arquivos ordenados por linha
            arquivo = arquivos[i]

            # Pastas para bk ordenadas por linha
            pasta_bk = bk_path[i]

            # Definição do arquivo que vai ser salvo a cópia
            arquivo_bk = pasta_bk + '\\' + arquivo
            # arquivo_bk = arquivo_bk.replace('.accdb', bk_str + '.accdb')

            # Remove o .accdb ou .mdb do final, adiciona '_BACKUP_.extensão' no arquivo
            logging.debug('Making copy of \'{}\''.format(arquivo))
            if arquivo.endswith('.accdb'):
                shutil.copyfile(caminho + '\\' + arquivo, arquivo_bk)
                logging.debug('Copy of \'{}\' created at \'{}\''.format(arquivo, pasta_bk))
            elif arquivo.endswith('.mdb'):
                shutil.copyfile(caminho + '\\' + arquivo, arquivo_bk)
                logging.debug('Copy \'{}\' created \'{}\''.format(arquivo, pasta_bk))

            # Sai do loop e vai pra a próxima linha/caminho
            logging.debug('Increasing i by 1')
            i += 1
            
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

def isBlocked():
    logging.debug('isBlocked()')

    # Iterador
    i = 0

    # Cria instância do Access
    logging.debug('Creating Access instance')
    access_instance = win32com.client.Dispatch('Access.Application')
    
    # A função que é chamada no main()
    for arquivo in arquivos:

        # Coloca o caminho completo do arquivo
        arquivo = db_path[i] + '\\' + arquivo

        # Chama a blocked check(file) para testar se o arquivo tem senha/pode ser aberto
        logging.debug('calling blocked_check()')
        if not blocked_check(arquivo, access_instance):
            # Access já está aberto
            # Desvincula a variavel do Access Object
            access_instance = None

        # Próxima pasta
        i += 1

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

    # Iterador
    i = 0
    
    # 'Abre' o Access
    logging.debug('Creating Access instance')
    db = win32com.client.Dispatch('Access.Application')
    logging.debug('Created')
    
    # Executa o compact and repair em todos arquivos .accdb
    logging.debug('STARTING .ACCDB CLEANING')
    for caminho in db_path:

        # .accdb
        arquivo = arquivos[i]

        # pasta de bk
        pasta_bk = bk_path[i]

        # Arquivo completo
        file = caminho + '\\' + arquivo

        # Backup obrigatório do arquivo para se executar o compact and repair
        logging.debug('Making tmp file')
        tmp_file = arquivo.replace('.accdb', 'BK.accdb')
        tmp_file = pasta_bk + '\\' + tmp_file
        logging.debug('tmp file is \'{}\''.format(tmp_file))

        # Compact and Repair
        try:
            logging.debug('Trying to repair \'{}\''.format(arquivo))
            db.CompactRepair(file, tmp_file, False)
            logging.debug('Repaired')
        except Exception:
            logging.debug('end of compact_repair()')
            logging.exception('ALGO DEU ERRADO DURANTE A LIMPEZA DAS DATABASES .accdb')
            return False

        # Substitui o arquivo compactado com o original
        # E deleta o arquivo criado no proceso
        logging.debug('Deleting tmp copy')
        shutil.copyfile(tmp_file, file)
        os.remove(tmp_file)
        logging.debug('Deleted')

        # Proxima linha
        i += 1

    """
    # Executa o compact and repair em todos os arquivos .mdb
    logging.debug('STARTING .MDB CLEANING')
    for caminho in db_path:

        # .accdb
        arquivo = arquivos[i]

        # pasta de bk
        pasta_bk = bk_path[i]

        # Arquivo completo
        file = caminho + '\\' + arquivo

        # Backup obrigatório do arquivo para se executar o compact and repair
        logging.debug('Making tmp file')
        tmp_file = arquivo.replace('.accdb', 'BK.accdb')
        tmp_file = pasta_bk + '\\' + tmp_file
        logging.debug('tmp file is \'{}\''.format(tmp_file))
        #logging.debug('tmp file is \'{}\''.format(tmp_file.replace(db_path + '\\', '')))

        # Compact and Repair
        try:
            logging.debug('Trying to repair \'{}\''.format(file.replace(db_path + '\\', '')))
            db.CompactRepair(file, tmp_file, False)
            logging.debug('Repaired')
        except Exception:
            logging.debug('end of compact_repair()')
            logging.exception('ALGO DEU ERRADO DURANTE A LIMPEZA DAS DATABASES .mdb')
            return False

        # Substitui o arquivo compactado com o original
        # E deleta o arquivo criado no processo
        logging.debug('Deleting tmp copy')
        shutil.copyfile(tmp_file, file)
        os.remove(tmp_file)
        logging.debug('Deleted')
    """
        

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

def delete_copies():
    logging.debug('delete_copies()')

    # Iterador
    i = 0

    # por pasta de backup
    for caminho in bk_path:

        # nome do arquivo
        arquivo = arquivos[i]

        # caminho completo do arquivo
        copy = caminho + '\\' + arquivo

        try:
            logging.debug('Trying to delete \'{}\''.format(copy))
            os.remove(copy)
            logging.debug('Deleted')
        except Exception:
            logging.debug('end of delete_copies()')
            logging.exception('ALGO DEU ERRADO DURANTE A EXCLUSÃO DAS CÓPIAS')
            return False

        i += 1

    logging.debug('end of delete_copies()')
    return True
    

#########################################################################################
#                                                                                       #
#                                   Função zipar()                                      #
#                       Zipa os arquivos passados como parâmetros                       #
#   IMPORTANTE: Arquivos perdem as permissões quando extraidos ex.: somente leitura     #
#                                                                                       #
#########################################################################################

def zipar():
    logging.debug('zipar()')

    # Iterador
    i = 0
    
    # Zipa os arquivos
    for caminho in db_path:
        arquivo = arquivos[i]
        pasta_bk = bk_path[i]

        # caminho completo para o arquivo
        caminho_completo = caminho + '\\' + arquivo
        with zipfile.ZipFile("{}\\Backup {} -- {}".format(pasta_bk, arquivo.replace('.accdb', ''),str(datetime.now().date().strftime("%d-%m-%Y"))) + ".zip", 'w') as backup:

            try:
                logging.debug('Trying to compact \'{}\''.format(arquivo))
                # Syntax
                # filename = caminho_completo = C:\...
                # arcname = arquivo = nome_do.accdb
                # para zipar apenas o arquivos invés do caminho inteiro até ele
                backup.write(caminho_completo, arquivo)
                logging.debug('Compacted')
            except Exception:
                logging.debug('end of zipar()')
                logging.exception('EXCEPTION OCCURED')
                return -1

        i += 1

    logging.debug('end of zipar()')
    return 0


#########################
#   Work in Progress    #
#########################
    
def send_mail(assunto, body, log = None, copy_fail = False):
    logging.debug('send_mail()')
    
    if copy_fail:
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
    try:
        server.connect('')
    except Exception as e:
        print('e:  ' + str(e))
    print('here')
    #server.noop()
    #server.set_debuglevel(1)
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

    #send_mail('teste assunto', 'teste corpo', log)

    # Testar os return_value (error code) do código
    # return_value = 1
"""
    if return_value == 1:
        send_mail('SCRIPT DE LIMPEZA: Diretório vazio', 'O diretório digitado no início do script não contem arquivos .accdb', log)
    elif return_value == 2:
        send_mail('SCRIPT DE LIMPEZA: Arquivo bloqueado', 'Durante o check de arquivos bloqueados o script encontrou um arquivo bloqueado e abortou', log)
    elif return_value == 3:
        send_mail('SCRIPT DE LIMPEZA: Não foi possível criar um backup', 'Durante a cópia dos arquivos houve um erro e não foi possível criar uma cópia', log)
    elif return_value == 4:
        send_mail('SCRIPT DE LIMPEZA: URGENTE: ERRO DURANTE O COMPACT AND REPAIR', 'Houve um erro durante a execução do compact and repair e ele não pode completar.\nDurante o processo de substituição dos arquivos corrompidos por seus backups, as cópias não foram encontradas e o script abortou para evitar maiores perdas', log)
    elif return_value == 5:
        send_mail('SCRIPT DE LIMPEZA: Erro durante o compact and repair', 'Houve um erro durante o processo de limpeza dos arquivos e ele teve que abortar inesperadamente. Os arquivos foram substituidos por suas cópias feitas antes da limpeza dos arquivos', log)
    elif return_value == 6:
        send_mail('SCRIPT DE LIMPEZA: Arquivos não foram encontrados', 'O script não conseguiu encontrar os arquivos após a limpeza e abortou', log, delete_copy_fail)
    elif return_value == 7:
        send_mail('SCRIPT DE LIMPEZA: Não foi possível compactar os arquivos', 'Devido a algum erro, não foi possível compactar os arquivos limpos. O script abortou', log, delete_copy_fail)
    elif return_value == 8:
        send_mail('SCRIPT DE LIMPEZA: Algum arquivo ficou bloqueado durante a compactação', 'Durante a compactação do script, algum arquivo ficou bloqueado (aberto) e não foi possivel compacta-los', log, delete_copy_fail)
    else:
        logging.info('O script conseguiu completar sem nenhum problema')
"""
logging.info('=' * 20 + ' Fim do Script ' + '=' * 20)
logging.shutdown()
