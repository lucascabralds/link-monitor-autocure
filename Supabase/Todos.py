import os
from datetime import datetime, timezone
from dotenv import load_dotenv
from supabase import create_client, Client

# 1. Inicialização: Carregando variáveis do .env
load_dotenv()
supabase_url = os.getenv("API_URL_SUPABASE")
supabase_key = os.getenv("SERVICE_ROLE_SUPABASE_API_KEY")

if not supabase_url or not supabase_key:
    print("❌ ERRO: Variáveis do Supabase não encontradas no .env")
    exit(1)

supabase: Client = create_client(supabase_url, supabase_key)

def popular_banco_dados_teste():
    """
    Função base para demonstrar o fluxo correto de inserção de dados
    respeitando os relacionamentos (Foreign Keys) do diagrama.
    """
    print("\n🚀 Iniciando Inserção de Dados no Supabase...")

    # =====================================================================
    # PASSO 1: INSERÇÃO NA TABELA 'clientes'
    # =====================================================================
    # Nota: Não enviamos 'id' nem 'created_at' porque o Supabase gera automaticamente.
    # Nota: Dados fictícios para teste
    dados_cliente = {
        "nome_cliente": "Agência Alpha",
        "responsavel": "João Silva",
        "email_contato": "joao@agenciaalpha.com",
        "ativo": True,
        "whatsapp": "+5511999999999",
        "youtube_url": "https://youtube.com/c/CanalAlpha"
    }

    try:
        print("📝 Inserindo Cliente...")
        res_cliente = supabase.table("clientes").insert(dados_cliente).execute()
        
        # Coletamos o UUID gerado pelo banco para usar no próximo passo
        cliente_uuid = res_cliente.data[0]['id']
        print(f"✅ Cliente criado com Sucesso! ID: {cliente_uuid}")

    except Exception as e:
        print(f"❌ Erro ao criar cliente: {e}")
        return

    # =====================================================================
    # PASSO 2: INSERÇÃO NA TABELA 'links_monitoramento'
    # =====================================================================
    # Aqui fazemos o VÍNCULO usando o 'cliente_id' (Foreign Key)
    # Incluimos também dados fictícios para teste, respeitando o formato esperado
    dados_link = {
        "cliente_id": cliente_uuid,  # Vínculo obrigatório com a tabela clientes
        "nome_produto": "Microfone de Lapela Sem Fio",
        "link_principal": "https://mercadolivre.com/sec/testemicrofone-lapela",
        "link_reserva": "https://mercadolivre.com/sec/microfone-lapela-reserva",
        "short_id": "lnk_abc12345",
        "short_url": "https://encurtador.io/mic",
        "videos_atrelados": "https://youtu.be/video1",
        "status_estoque": "em_estoque",
        "ultima_verificacao": datetime.now(timezone.utc).isoformat(),
        "plataforma": "Mercado Livre"
    }

    try:
        print("\n📝 Inserindo Link de Monitoramento...")
        res_link = supabase.table("links_monitoramento").insert(dados_link).execute()
        
        # Coletamos o UUID do link gerado para usar no histórico
        link_uuid = res_link.data[0]['id']
        print(f"✅ Link criado com Sucesso e vinculado ao cliente! ID: {link_uuid}")

    except Exception as e:
        print(f"❌ Erro ao criar link: {e}")
        return

    # =====================================================================
    # PASSO 3: INSERÇÃO NA TABELA 'historico_ajustes'
    # =====================================================================
    # Simulando que o robô (processo de auto-cura) precisou trocar o link
    dados_historico = {
        "link_id": link_uuid,  # Vínculo obrigatório com a tabela links_monitoramento
        "data_alteracao": datetime.now(timezone.utc).isoformat(),
        "url_antiga": dados_link["link_principal"],
        "url_nova": dados_link["link_reserva"],
        "motivo": "Simulação: Botão de estoque não encontrado. Auto-cura acionada."
    }

    try:
        print("\n📝 Inserindo Histórico de Alteração...")
        res_historico = supabase.table("historico_ajustes").insert(dados_historico).execute()
        print("✅ Histórico registrado com Sucesso!")

    except Exception as e:
        print(f"❌ Erro ao registrar histórico: {e}")
        return

    print("\n🎉 Fluxo de inserção relacional concluído com sucesso!")

# Ponto de entrada do script
if __name__ == "__main__":
    popular_banco_dados_teste()