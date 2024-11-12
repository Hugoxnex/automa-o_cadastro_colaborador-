import openpyxl
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 0.5

# Links para os portais
link_formulario = 'https://bracell.service-now.com/sys_user.do?sys_id=-1&sys_is_list=true&sys_target=sys_user&sysparm_checked_items=&sysparm_fixed_query=&sysparm_group_sort=&sysparm_list_css=&sysparm_query=&sysparm_referring_url=sys_user_list.do&sysparm_target=&sysparm_view='  
link_chamados_portalS = 'https://bracell.service-now.com/sp?id=sc_cat_item&sys_id=0ab3eb5c1bf1c21017fd404be54bcbd4&referrer=popular_items'
link_fecharRitm = 'https://bracell.service-now.com/nav_to.do?uri=%2F$pa_dashboard.do'

def abrir_navegador():
    pyautogui.press('winleft')
    pyautogui.write('edge')
    pyautogui.press('enter')
    time.sleep(2)

def acessar_formulario_chamado():
    pyperclip.copy(link_formulario)
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(5)                               
    pyautogui.click(561, 427, duration=1)

def preencher_campo(texto, usar_clipboard=False):
    if usar_clipboard:
        pyperclip.copy(texto)
        pyautogui.hotkey('ctrl', 'v')
    else:
        pyautogui.write(texto)

# Entrar na planilha
workbook = openpyxl.load_workbook('chamados_database.xlsx')
sheet_chamados = workbook['PrevRonda']

# Chamadas das funções
abrir_navegador()
acessar_formulario_chamado()

for linha in sheet_chamados.iter_rows(min_row=2):
    # Preenche o campo "Nome"
    pyautogui.click(369, 219, duration=1)
    id_user = linha[0].value
    preencher_campo(id_user)
    time.sleep(0.1)

    # Preenche o campo "Problema"
    nome_user = linha[1].value
    preencher_campo(nome_user, usar_clipboard=True)

    # Preenche o campo "Descrição"
    pyautogui.click(387, 418, duration=1)
    local_user = linha[2].value
    preencher_campo(local_user, usar_clipboard=True)

    # Preenche o campo "Gerente"
    pyautogui.click(1320, 217, duration=1)
    gerente = linha[3].value
    preencher_campo(gerente, usar_clipboard=True)

    # Salva o chamado
    # pyautogui.click(1839, 90, duration=1)

    # Pegando o número do chamado
    nome_re = str(linha[6].value)              
    pyperclip.copy(nome_re)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('tab')

    solicitacao = linha[8].value
    pyautogui.write(solicitacao)
    time.sleep(0.5)
    pyautogui.press('tab')

    descricao = linha[12].value
    descricao_2 = str(linha[13].value)
    texto_concatenado = descricao + descricao_2
    pyautogui.write(texto_concatenado)
    time.sleep(2)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(3)

    # Acessar o portal de chamados
    pyperclip.copy(link_chamados_portalS)                        
    pyautogui.hotkey('ctrl', 'l')                               
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(5)        

    pyautogui.click(418, 422, duration=0.5)                         
    time.sleep(3) 
    pyautogui.click(421, 380, duration=0.5)                         
    time.sleep(3) 
    pyautogui.doubleClick(1421, 281, duration=0.5)                   
    pyautogui.hotkey('ctrl', 'c')                                  
    numero_chamado = pyperclip.paste()                             
    print(numero_chamado)
    time.sleep(2) 

    # Indo para fechar o chamado
    pyperclip.copy(link_fecharRitm)
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(4)

    pyautogui.click(1738, 93, duration=1)                            
    pyperclip.copy(numero_chamado)                                  
    pyautogui.hotkey('ctrl', 'v')                                   
    pyautogui.press('enter')
    time.sleep(3)

    pyautogui.click(1453, 349, duration=1)                         
    estado = linha[8].value
    pyperclip.copy(estado)
    pyautogui.write(estado)
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    analista = linha[9].value
    pyperclip.copy(analista)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('tab')

    pyautogui.rightClick(1031, 145, duration=1)  
    pyautogui.press('enter')

    pyautogui.click(1437, 295, duration=1)                         
    resolvido = linha[10].value
    pyperclip.copy(resolvido)
    pyautogui.write(resolvido)
    time.sleep(1)

    pyautogui.click(1782, 143, duration=1)                         

    # Abre uma nova aba e acessa o site novamente
    pyautogui.hotkey('ctrl', 't')
    acessar_formulario_chamado()

# Alerta final
pyautogui.alert("O código acabou de rodar. Clique Ok para fechar a janela")
