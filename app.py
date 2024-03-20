import openpyxl
#pyperclip serve para pegar as informações com pontuação do excel
import pyperclip
import pyautogui
from time import sleep

#abrir a planilha 
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
#como no excel so tem uma pagina que seria Produto, tenho declarar essa pagina.   
sheet_produto = workbook['Produtos']
#copiar dados da planilha de um campo e colar no campo correspondente, esse 2 é a linha que inicia excel
for linha in sheet_produto.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto) 
    pyautogui.click(917,202, duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(883,288, duration=1)  #coordenada do mouse 
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(925,420, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(930,512, duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(937,592, duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(947,632, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(898,691,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(927,222,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(911,307,duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(911,398,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(914,477,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    tamanho = linha[10].value
    pyautogui.click(946,566,duration=1)

    if tamanho == 'pequeno':
        pyautogui.click(919,599, duration=1)
    elif tamanho == 'medio':
        pyautogui.click(915,619,duration=1)
    else:
        pyautogui.click(915,643,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(954,639,duration=1)
    pyautogui.hotkey('ctrl','v')
    
#botão proximo
    pyautogui.click(897,719,duration=1)
   
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(911,248,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(934,330,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(898,420,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(908,548,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(890,641,duration=1)
    pyautogui.hotkey('ctrl','v')
#botão concluir
    pyautogui.click(900,701,duration=1)
#botão confirmar 
    pyautogui.click(1267,186,duration=1)
#botão confirmar 2
    pyautogui.click(1095,467,duration=1)
#botão iniciar novamente
    pyautogui.click(1118,471,duration=1)
    
