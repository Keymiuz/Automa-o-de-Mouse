import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(530, 290, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(350, 430, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(1)
    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(276, 562, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Codigo Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(290, 670, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1463, 691, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(255, 775, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão próximo
    pyautogui.click(215, 960, duration=1)
    sleep(2)

    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(330, 320, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(300, 430, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(300, 540, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(250, 645, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(280, 750, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(225, 790, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(252, 820, duration=1)
    else:
        pyautogui.click(240, 845, duration=1)

    # material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(280, 850, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    # Botão próximo
    pyautogui.click(220, 940, duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(200, 340, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(300, 455, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(320, 600, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(320, 720, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(310, 840, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão concluir
    pyautogui.click(220, 910, duration=1)
    # Botão confirmar inclusão
    pyautogui.click(1210, 190, duration=1)
    # iniciar cadastro novamente
    pyautogui.click(965, 630, duration=1)
