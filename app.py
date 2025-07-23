import openpyxl
import pyperclip
import pyautogui
from time import sleep
#Entrar na planilha
workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook["Produtos"]
#Copiar infos de um campo e colar no correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1156,354,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1149,446,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1157,572,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1155,655,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1167,741,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1173,830,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1170,882,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1168,396,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1158,474,duration=1)
    pyautogui.hotkey('ctrl','v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(1156,563,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1153,661,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(1155,738,duration=1)
    if tamanho == "Pequeno":
        pyautogui.click(1164,744,duration=1)
    elif tamanho == "Médio":
        pyautogui.click(1150,800,duration=1)
    else:
        pyautogui.click(1158,834,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1154,818,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1134,887,duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1155,412,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1139,500,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1141,588,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1189,713,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1242,812,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #concluir
    pyautogui.click(1144,864,duration=1)
    #confirmar
    pyautogui.click(1624,192,duration=1)
    #confirmar 2
    pyautogui.click(1150,625,duration=1)
    #iniciar cadastro novamente
    pyautogui.click(1474,637,duration=1)
