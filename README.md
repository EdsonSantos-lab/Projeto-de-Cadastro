# Projeto: Automação Completa de Cadastro de Produtos em Web App
Descrição:
Este projeto automatiza o cadastro de produtos em um sistema web, lendo informações de uma planilha Excel e preenchendo campos de formulários de maneira precisa e rápida. Ele combina técnicas de automação de navegador com Selenium, automação de interface gráfica com PyAutoGUI e manipulação de arquivos Excel com OpenPyXL.

Funcionalidades Implementadas:
Inicialização e Navegação Automática

Inicializa o navegador Chrome com Selenium WebDriver.

Acessa a URL do sistema web de cadastro de produtos.

python
Copiar
Editar
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.get('https://cadastro-produtos-devaprender.netlify.app/index.html')
Gerenciamento de Janelas no Windows

Maximiza a janela ativa do navegador.

Organiza a janela à direita da tela usando atalhos do Windows (Win + Right) e Enter.

python
Copiar
Editar
hwnd = win32gui.GetForegroundWindow()
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
pyautogui.hotkey('win','right')
pyautogui.press('enter')
Leitura de Dados do Excel

Carrega o arquivo produtos_ficticios.xlsx com OpenPyXL.

Itera sobre cada linha da planilha, lendo campos como Nome, Descrição, Categoria, Código do Produto, Peso, Dimensões, Preço, Quantidade, Data de Validade, Cor, Tamanho, Material, Fabricante, País de Origem, Observações, Código de Barras e Localização.

python
Copiar
Editar
abrir_planlha = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = abrir_planlha['Produtos']
linhas = sheet_produtos.iter_rows(min_row=2)
Preenchimento Automático com PyAutoGUI

Para cada campo do formulário, o valor é copiado para a área de transferência com Pyperclip e colado nos campos usando PyAutoGUI.

Simula cliques precisos em coordenadas da tela e atalhos de teclado (Ctrl+V) para garantir preenchimento correto.

Campos condicionais, como Tamanho, são selecionados dinamicamente com base no valor.

python
Copiar
Editar
pyperclip.copy(tamanho)
pyautogui.click(1369, 784, duration=1)
if tamanho == 'Pequeno':
    pyautogui.click(1391, 823,duration=1)
elif tamanho == 'Médio':
    pyautogui.click(1419, 860,duration=1)
else:
    pyautogui.click(1344,901,duration=1)
Processamento Multi-etapa

O cadastro é realizado em várias etapas, preenchendo informações básicas, dados comerciais e detalhes do fabricante.

Cada etapa é separada por cliques em botões e navegação entre seções.

Benefícios do Projeto:
Redução de tarefas manuais: Preenche centenas de produtos automaticamente.

Precisão: Evita erros humanos no cadastro de dados complexos.

Flexibilidade: Permite alterar planilha e coordenadas para se adaptar a diferentes sistemas.

Integração de tecnologias: Combina Selenium, PyAutoGUI, Pyperclip e OpenPyXL em um único fluxo
