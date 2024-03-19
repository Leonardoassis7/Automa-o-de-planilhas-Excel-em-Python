import openpyxl
#pyperclip serve para pega as informçao com pontuação do excel
import pyperclip
import pyautogui

#abrir a planilha 
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
#como no excel so tem uma pagina que seria Produto, tenho declara essa pagina.   
sheet_produto = workbook['Produtos']
#copiar dados da planilha de um campo e colar no campo correspondente, esse 2 é a linha que inicia excel
for linha in sheet_produto.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    #coordenada do mouse 
    pyautogui.click(85,162, duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip(descricao)
    pyautogui.click(84,244, duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip(categoria)
    pyautogui.click(94,386, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip(codigo_produto)
    pyautogui.click(104,475, duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip(peso)
    pyautogui.click(110,552, duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip(dimensoes)
    pyautogui.click(128,157, duration=1)
    pyautogui.hotkey('ctrl','v')

    preco = linha[6].value
    quantidade_em_estoque = linha[7].value
    data_de_validade = linha[8].value
    cor = linha[9].value
    tamanho = linha[10].value
    material = linha[11].value
    fabricante = linha[12].value
    pais_origem = linha[13].value
    observacoes = linha[14].value
    codigo_de_barras = linha[15].value
    localizacao_armazem = linha[16].value
    
