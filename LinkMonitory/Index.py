import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# 1. Configurações Iniciais
load_dotenv()
CAMINHO_DRIVE = r"C:\Users\Desktop\Meu Drive\Cliente\Base_Monitoramento_Links_v1.xlsx"
NOME_ABA = "Monitoramento"

def validar_acesso_drive():
    """Verifica se o Google Drive está mapeado e acessível no PC."""
    print("🔍 Validando acesso ao Google Drive...")
    
    # Verifica se o arquivo existe no caminho especificado
    if os.path.exists(CAMINHO_DRIVE):
        print("✅ Drive detectado e arquivo encontrado!")
        return True
    else:
        print(f"❌ ERRO CRÍTICO: O caminho '{CAMINHO_DRIVE}' não está acessível.")
        print("Certifique-se de que o Google Drive Desktop está aberto e logado.")
        return False

def validar_estoque_no_terminal(url, id_item, nome_produto):
    """Acessa o link e imprime o status diretamente no terminal."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        try:
            # Aguarda o redirecionamento completo do link /sec/
            page.goto(url, wait_until="networkidle", timeout=60000)
            time.sleep(2) 

            # Busca pelo botão que confirma a área de estoque ativa
            area_estoque = page.query_selector("text='Ir para produto'")
            
            if area_estoque:
                print(f"✅ [ID {id_item}] {nome_produto}: Em estoque.")
                return True
            else:
                # MENSAGEM NO TERMINAL SE NÃO CONSEGUIR LER O ESTOQUE
                print(f"❌ [ID {id_item}] ALERTA: Não foi possível ler estoque para: {nome_produto}")
                return False
                

                #Incluir aqui a lógica de encontrar no estoque validando o endereço do site
        except Exception as e:
            print(f"⚠️ [ID {id_item}] ERRO TÉCNICO NO LINK: {e}")
            return False
        finally:
            browser.close()

def iniciar_monitoramento():
    # PASSO 1: Validar se o Drive está acessível antes de começar
    if not validar_acesso_drive():
        return

    # PASSO 2: Ler a aba específica 'monitoramento'
    try:
        print(f"📂 Abrindo a aba '{NOME_ABA}' da planilha...")
        df = pd.read_excel(CAMINHO_DRIVE, sheet_name=NOME_ABA)
    except Exception as e:
        print(f"❌ ERRO: Não foi possível carregar a aba '{NOME_ABA}'. Verifique o nome no Excel. ({e})")
        return
    
    print(f"🚀 Iniciando verificação de {len(df)} itens...\n")
    print("-" * 70)

    for index, linha in df.iterrows():
        # Pegando os dados baseados no cabeçalho da sua planilha
        id_item = linha.get('ID', index)
        nome = linha.get('Nome do Produto', 'Produto Sem Nome')
        link = linha.get('Link Principal (Link 1)')

        if pd.isna(link):
            print(f"⏩ [ID {id_item}] Linha vazia, pulando...")
            continue

        validar_estoque_no_terminal(link, id_item, nome)

    print("-" * 70)
    print("🏁 Monitoramento concluído.")

if __name__ == "__main__":
    iniciar_monitoramento()