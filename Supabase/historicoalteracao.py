import os
from datetime import datetime, timezone
from dotenv import load_dotenv
from supabase import create_client, Client

# =====================================================================
# 1. AUTENTICAÇÃO E CONEXÃO COM A API
# =====================================================================
load_dotenv()
supabase_url = os.getenv("API_URL_SUPABASE")
supabase_key = os.getenv("SERVICE_ROLE_SUPABASE_API_KEY")


if not supabase_url or not supabase_key:
    print("❌ ERRO CRÍTICO: Chaves de API do Supabase não encontradas no arquivo .env")
    exit(1)


supabase: Client = create_client(supabase_url, supabase_key)

# =====================================================================
# 2. FUNÇÃO DE INSERÇÃO NO HISTÓRICO
# =====================================================================
def registrar_historico_autocura(link_id_banco: str, url_velha: str, url_nova: str, motivo_troca: str):
    """
    Insere um novo registro na tabela 'historico_ajustes' do Supabase.
    Documenta a ação de auto-cura realizada pelo robô.
    """
    print("\n🚀 Iniciando gravação na tabela 'historico_ajustes'...")

    # Montagem do pacote de dados (Payload) mapeado exatamente para as colunas da tabela
    dados_historico = {
        "link_id": link_id_banco,                                  # Foreign Key: A Chave unica do link na tabela principal
        "data_alteracao": datetime.now(timezone.utc).isoformat(),  # Timestamp com fuso horário UTC para padronização global
        "url_antiga": url_velha,                                   # A URL que o Playwright detectou como esgotada
        "url_nova": url_nova,                                      # A URL reserva que assumiu o lugar
        "motivo": motivo_troca                                     # Descrição em texto do motivo da troca
    }

    try:
        # Comunicação com a API: Método de inserção (POST)
        resposta = supabase.table("historico_ajustes").insert(dados_historico).execute()
        
        # Extração do ID do histórico recém-criado para confirmação
        id_gerado = resposta.data[0]['id']
        print(f"✅ Histórico documentado com Sucesso!")
        print(f"   Identificador da Alteração (ID): {id_gerado}")
        return True

    except Exception as erro:
        # Captura de erro da API (Ex: UUID inválido, violação de chave estrangeira, etc.)
        print(f"❌ Falha de comunicação com a API ao registrar histórico:")
        print(f"   Detalhes do Erro: {erro}")
        return False

# =====================================================================
# 3. TESTE PRÁTICO DO SCRIPT
# =====================================================================
if __name__ == "__main__":
    #  Caso queira testar isoladamente, você precisa fornecer um UUID (link_id) 
    # que já exista previamente na sua tabela 'links_monitoramento'.
    # Caso envie um UUID inventado, a API do Supabase retornará um erro de Foreign Key.
    
    UUID_REAL_DO_SEU_BANCO = "b4120351-18c0-4e38-b067-2aa71bc7de83"
    
    registrar_historico_autocura(
        link_id_banco=UUID_REAL_DO_SEU_BANCO,
        url_velha="https://mercadolivre.com/sec/farol-civic-esgotado-exemplo",
        url_nova="https://mercadolivre.com/sec/farol-civic-reserva-exemplo",
        motivo_troca="Falta de estoque: O Playwright não encontrou o botão 'Ir para produto'."
    )