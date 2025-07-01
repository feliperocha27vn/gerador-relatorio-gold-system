import configparser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# Read credentials from config.ini
config = configparser.ConfigParser()
config.read('config.ini')

URL = config['Credentials']['URL']
STORE = config['Credentials']['STORE']
USERNAME = config['Credentials']['USERNAME']
PASSWORD = config['Credentials']['PASSWORD']

# INPUTS DO USUÁRIO
print("=== CONFIGURAÇÃO DOS FILTROS ===")
secao_desejada = input("Digite a seção desejada (ex: JOIAS EM OURO): ").upper()
familia_desejada = input("Digite a família desejada (ex: ANEL) ou deixe vazio para pular: ").upper().strip()
loja_desejada = input("Digite a loja desejada (ex: LOJA 1): ").upper()
print("=====================================\n")

# Configurar Chrome para salvar PDFs automaticamente
chrome_options = Options()
download_dir = os.path.join(os.getcwd(), "downloads")
os.makedirs(download_dir, exist_ok=True)

prefs = {
    "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
    "savefile.default_directory": download_dir,
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--kiosk-printing")

driver = webdriver.Chrome(options=chrome_options)

try:
    # PASSO 1: LOGIN
    print("Iniciando login...")
    driver.get(URL)
    wait = WebDriverWait(driver, 15)
    
    store_field = wait.until(EC.presence_of_element_located((By.NAME, 'loja')))
    username_field = driver.find_element(By.NAME, 'nome')
    password_field = driver.find_element(By.NAME, 'senha')
    login_button = driver.find_element(By.NAME, 'submeteLogin')
    
    store_field.send_keys(STORE)
    username_field.send_keys(USERNAME)
    password_field.send_keys(PASSWORD)
    login_button.click()
    
    print("Login realizado com sucesso!")
    time.sleep(3)

    # PASSO 2: CLICAR EM RELATÓRIOS
    print("Clicando em Relatórios...")
    driver.switch_to.default_content()
    driver.switch_to.frame("a")
    
    relatorios_btn = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Relatorios")))
    relatorios_btn.click()
    
    print("Clicou no botão Relatórios com sucesso!")
    driver.switch_to.default_content()
    time.sleep(3)

    # PASSO 3: CLICAR EM "PRODUTO - ESTOQUE"
    print("Clicando em Produto - Estoque...")
    driver.switch_to.frame("b")
    
    produto_estoque_btn = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Produto - Estoque")))
    produto_estoque_btn.click()
    
    print("Clicou em 'Produto - Estoque' com sucesso!")
    driver.switch_to.default_content()
    time.sleep(5)

 # PASSO 4: CLICAR EM "Listar Produto/Estoque" - VERSÃO COM DEBUG
    print("Clicando em Listar Produto/Estoque...")
    
    # Tentar no frame "b" primeiro
    try:
        print("Tentando no frame 'b'...")
        driver.switch_to.frame("b")
        
        # Várias estratégias de busca
        strategies = [
            (By.XPATH, "//span[contains(text(), 'Listar Produto/Estoque')]/parent::a")
        ]
        
        link_found = False
        for strategy in strategies:
            try:
                print(f"Tentando estratégia: {strategy}")
                listar_produto_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable(strategy)
                )
                listar_produto_btn.click()
                print("Clicou em 'Listar Produto/Estoque' com sucesso!")
                link_found = True
                break
            except:
                continue
        
        if not link_found:
            raise Exception("Link não encontrado no frame 'b'")
            
    except Exception as e:
        print(f"Erro no frame 'b': {e}")
        print("Tentando no frame 'c'...")
        
        # Tentar no frame "c"
        driver.switch_to.default_content()
        driver.switch_to.frame("c")
        
        for strategy in strategies:
            try:
                print(f"Tentando estratégia no frame 'c': {strategy}")
                listar_produto_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable(strategy)
                )
                listar_produto_btn.click()
                print("Clicou em 'Listar Produto/Estoque' no frame 'c' com sucesso!")
                break
            except:
                continue
        else:
            raise Exception("Link não encontrado em nenhum frame")


    # PASSO 5: MARCAR RADIO "PRODUTO"
    print("Marcando o radio 'Produto'...")
    driver.switch_to.default_content()
    driver.switch_to.frame("c")
    
    radio_produto = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='radio' and @name='agrupar' and @value='produto']")))
    radio_produto.click()
    print("Radio 'Produto' marcado com sucesso!")

    # PASSO 6: SELECIONAR SEÇÃO
    print(f"Selecionando seção: {secao_desejada}...")
    select_secao = wait.until(EC.presence_of_element_located((By.NAME, "idsecao")))
    opcao_secao = select_secao.find_element(By.XPATH, f".//option[contains(text(), '{secao_desejada}')]")
    opcao_secao.click()
    print(f"Seção '{secao_desejada}' selecionada com sucesso!")
    
    # Aguardar a página recarregar após submeterSecao()
    time.sleep(3)
    
    # PASSO 7: SELECIONAR FAMÍLIA (OPCIONAL)
    if familia_desejada:
        print(f"Selecionando família: {familia_desejada}...")
        driver.switch_to.default_content()
        driver.switch_to.frame("c")
        select_familia = wait.until(EC.presence_of_element_located((By.NAME, "idfamilia")))
        opcao_familia = select_familia.find_element(By.XPATH, f".//option[normalize-space(text())='{familia_desejada}']")
        opcao_familia.click()
        print(f"Família '{familia_desejada}' selecionada com sucesso!")
        time.sleep(3)
    else:
        print("Família não informada, pulando...")

     # PASSO 8: SELECIONAR ESTOQUE POSITIVO
    print("Selecionando estoque positivo...")
    
    # Garantir que está no frame correto
    driver.switch_to.default_content()
    driver.switch_to.frame("c")
    
    # Aguardar mais tempo e relocalizar o elemento
    time.sleep(2)
    select_estoque = wait.until(EC.presence_of_element_located((By.NAME, "estoque")))
    opcao_positivo = select_estoque.find_element(By.XPATH, ".//option[@value='positivo']")
    opcao_positivo.click()
    print("Estoque 'POSITIVO' selecionado com sucesso!")

    # PASSO 9: SELECIONAR LOJA
    print(f"Selecionando loja: {loja_desejada}...")
    driver.switch_to.default_content()
    driver.switch_to.frame("c")
    select_loja = wait.until(EC.presence_of_element_located((By.NAME, "idloja")))
    opcao_loja = select_loja.find_element(By.XPATH, f".//option[contains(text(), '{loja_desejada}')]")
    opcao_loja.click()
    print(f"Loja '{loja_desejada}' selecionada com sucesso!")
    time.sleep(3)

    # PASSO 10: MARCAR CHECKBOXES
    print("Marcando checkboxes: Fornecedor, Qtde e Custo R$...")
    driver.switch_to.default_content()
    driver.switch_to.frame("c")
    
    # Marcar checkbox Fornecedor
    checkbox_fornecedor = wait.until(EC.presence_of_element_located((By.NAME, "fornecedor")))
    if not checkbox_fornecedor.is_selected():
        checkbox_fornecedor.click()
    print("Checkbox 'Fornecedor' marcado!")
    
    # Marcar checkbox Qtde
    checkbox_quantidade = wait.until(EC.presence_of_element_located((By.NAME, "quantidade")))
    if not checkbox_quantidade.is_selected():
        checkbox_quantidade.click()
    print("Checkbox 'Qtde' marcado!")
    
    # Marcar checkbox Custo R$
    checkbox_custo_reais = wait.until(EC.presence_of_element_located((By.NAME, "precocustoreais")))
    if not checkbox_custo_reais.is_selected():
        checkbox_custo_reais.click()
    print("Checkbox 'Custo R$' marcado!")
    
    print("Todos os checkboxes selecionados com sucesso!")
    time.sleep(2)

    # PASSO 11: CLICAR NO BOTÃO LISTAR
    print("Clicando no botão 'Listar'...")
    driver.switch_to.default_content()
    driver.switch_to.frame("c")
    
    botao_listar = wait.until(EC.element_to_be_clickable((By.NAME, "submeteListagem")))
    botao_listar.click()
    print("Botão 'Listar' clicado com sucesso!")
    
    # AGUARDAR A LISTAGEM CARREGAR (até 60 segundos)
    print("Aguardando listagem carregar...")
    try:
        # Esperar até 60 segundos pelo botão de imprimir aparecer
        wait_longo = WebDriverWait(driver, 60)
        botao_imprimir_check = wait_longo.until(EC.element_to_be_clickable((By.NAME, "submeteImpressaoListagem")))
        print("Listagem carregada com sucesso!")
    except Exception as e:
        print(f"Timeout ou erro ao aguardar listagem: {e}")
        print("Continuando mesmo assim...")

    # PASSO 13: RENOMEAR O ARQUIVO PDF COM O NOME DA FAMÍLIA
    # RENOMEAR PDF COM NOVA FAMÍLIA OU SEÇÃO
    print("Renomeando arquivo PDF...")
    import glob
    import shutil
    
    time.sleep(3)
    pdf_files = glob.glob(os.path.join(download_dir, "*.pdf"))
    if pdf_files:
        arquivo_mais_recente = max(pdf_files, key=os.path.getctime)
    
        # Se tem família, usa família; senão usa seção
        if familia_desejada:
            nome_pdf = familia_desejada
        else:
            nome_pdf = secao_desejada
    
        novo_nome = os.path.join(download_dir, f"{nome_pdf}.pdf")
    
        shutil.move(arquivo_mais_recente, novo_nome)
        print(f"Arquivo renomeado para: {nome_pdf}.pdf")
    else:
        print("Nenhum arquivo PDF encontrado para renomear")

    print("Automação concluída com sucesso!")

except Exception as e:
    print(f"Erro durante a automação: {e}")

# LOOP PARA GERAR MAIS RELATÓRIOS SEM FAZER LOGIN NOVAMENTE
while True:
    continuar = input("\nDeseja gerar relatório de outra seção/família? (s/n): ").lower().strip()
    
    if continuar in ['s', 'sim', 'y', 'yes']:
        # Se não tinha família no primeiro input, pedir nova seção
        if not familia_desejada:
            nova_secao = input("Digite a nova seção desejada (ex: JOIAS EM OURO): ").upper()
            
            print(f"\nGerando relatório para seção: {nova_secao}")
            print("="*50)
            
            # TROCAR SEÇÃO
            print(f"Selecionando nova seção: {nova_secao}...")
            driver.switch_to.default_content()
            driver.switch_to.frame("c")
            
            select_secao = wait.until(EC.presence_of_element_located((By.NAME, "idsecao")))
            opcao_secao = select_secao.find_element(By.XPATH, f".//option[contains(text(), '{nova_secao}')]")
            opcao_secao.click()
            print(f"Seção '{nova_secao}' selecionada com sucesso!")
            time.sleep(3)
            
            nome_pdf = nova_secao  # Para renomear o arquivo
            
        else:
            # Se tinha família, pedir nova família
            nova_familia = input("Digite a nova família desejada (ex: BRINCO): ").upper()
            
            print(f"\nGerando relatório para família: {nova_familia}")
            print("="*50)
            
            # TROCAR FAMÍLIA
            print(f"Selecionando nova família: {nova_familia}...")
            driver.switch_to.default_content()
            driver.switch_to.frame("c")
            
            select_familia = wait.until(EC.presence_of_element_located((By.NAME, "idfamilia")))
            opcao_familia = select_familia.find_element(By.XPATH, f".//option[normalize-space(text())='{nova_familia}']")
            opcao_familia.click()
            print(f"Família '{nova_familia}' selecionada com sucesso!")
            time.sleep(3)
            
            nome_pdf = nova_familia  # Para renomear o arquivo
        
        # CLICAR EM LISTAR NOVAMENTE
        print("Clicando no botão 'Listar'...")
        botao_listar = wait.until(EC.element_to_be_clickable((By.NAME, "submeteListagem")))
        botao_listar.click()
        print("Botão 'Listar' clicado com sucesso!")
        time.sleep(10)
        
        # IMPRIMIR LISTAGEM
        print("Clicando no botão 'Imprimir Listagem'...")
        botao_imprimir = wait.until(EC.element_to_be_clickable((By.NAME, "submeteImpressaoListagem")))
        botao_imprimir.click()
        print("Botão 'Imprimir Listagem' clicado - PDF será salvo automaticamente!")
        time.sleep(8)
        
        # RENOMEAR PDF
        print("Renomeando arquivo PDF...")
        import glob
        import shutil
        
        time.sleep(3)
        pdf_files = glob.glob(os.path.join(download_dir, "*.pdf"))
        if pdf_files:
            arquivo_mais_recente = max(pdf_files, key=os.path.getctime)
            novo_nome = os.path.join(download_dir, f"{nome_pdf}.pdf")
            
            shutil.move(arquivo_mais_recente, novo_nome)
            print(f"Arquivo renomeado para: {nome_pdf}.pdf")
        else:
            print("Nenhum arquivo PDF encontrado para renomear")
            
        print(f"Relatório concluído!")
        
    elif continuar in ['n', 'nao', 'não', 'no']:
        break
    else:
        print("Digite 's' para sim ou 'n' para não")

print("Encerrando...")
driver.quit()