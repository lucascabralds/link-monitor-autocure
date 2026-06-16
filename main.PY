"""
╔══════════════════════════════════════════════════════════════════════════╗
║     SISTEMA DE MONITORAMENTO, CADASTRO E AUTO-CURA DE LINKS             ║
║     Versão: 7.0.0                                                        ║
╠══════════════════════════════════════════════════════════════════════════╣
║  NOVIDADES v7.0.0 — Caso E: link encurtado sem principal → VALIDAR:     ║
║                                                                          ║
║  🟢 Caso E (NOVO): Sem principal, COM encurtado (com ou sem reserva)    ║
║     → ação = VALIDAR. O Short.io é consultado como fonte de verdade     ║
║       para recuperar o link_principal antes da verificação.             ║
║     → ID_Banco é buscado/preenchido normalmente (igual aos outros).     ║
║                                                                          ║
║  🟢 ID: auto-sequencial baseado no ÚLTIMO ID preenchido na planilha     ║
║     Não sobrescreve IDs já existentes                                    ║
║                                                                          ║
║  🟢 ID_Banco: NUNCA alterado pelo código. Preservado como está.         ║
║                                                                          ║
║  🟢 Nome do Produto: preenchido pelo cliente. Sem ação se vazio.        ║
║                                                                          ║
║  🟢 Plataforma: sempre "mercadolivre" (fixo)                            ║
║                                                                          ║
║  🟢 Link Principal (Link 1): link destino do encurtador.               ║
║     Se tem link mas sem encurtado → CRIAR encurtador                    ║
║                                                                          ║
║  🟢 Links Reservas (Link 2):                                            ║
║     Caso A: Não tem principal, tem reserva, sem encurtado → CRIAR       ║
║     Caso B: Tem principal, tem reserva, sem encurtado → CRIAR           ║
║     Caso C: Tem principal, SEM reserva, COM encurtado → VALIDAR        ║
║     Caso D: Tem principal, COM reserva, COM encurtado → VALIDAR        ║
║     Linha sem Nome do Produto → nenhuma ação                            ║
║                                                                          ║
║  🟢 Link Encurtado: padrão reidascapotas.s.gy/*                        ║
║     Todos testados via Playwright (botão "Ir para produto")             ║
║                                                                          ║
║  🟢 Vídeos Atrelados (YouTube): não alterado                            ║
║  🟢 Status / Ultima Atualização: sempre atualizados pelo código         ║
║                                                                          ║
║  🟦 FLUXO DE VERIFICAÇÃO v6.0:                                          ║
║     SINC  → Garante IDs e plataforma                                    ║
║     PASSO A → Testa encurtador (short_url) via Playwright               ║
║     PASSO B → Se A falhou: testa link_principal                         ║
║     PASSO C → Se B passou mas A falhou: PATCH Short.io (sincroniza)     ║
║     PASSO D → Se ambos falharam: auto-cura com fila de reservas         ║
╚══════════════════════════════════════════════════════════════════════════╝

SCHEMA (conforme imagens do banco):
  clientes            → id, created_at, nome_cliente, responsavel,
                        email_contato, ativo, whatsapp, youtube_url
  links_monitoramento → id, created_at, cliente_id, nome_produto,
                        link_principal, link_reserva, short_id, short_url,
                        videos_atrelados, status_estoque,
                        ultima_verificacao, plataforma
  historico_ajustes   → id, link_id, data_alteracao, url_antiga, url_nova,
                        motivo, tipo_operacao, produto_nome, plataforma

DEPENDÊNCIAS:
    pip install python-dotenv supabase playwright pandas openpyxl requests
    playwright install chromium
"""

import os
import re
import time
import requests
import pandas as pd
from datetime import datetime, timezone
from dotenv import load_dotenv
from supabase import create_client, Client
from playwright.sync_api import sync_playwright


# ══════════════════════════════════════════════════════════════════════════
# 1. CONFIGURAÇÕES
# ══════════════════════════════════════════════════════════════════════════

CAMINHO_DRIVE  = os.getenv("CAMINHO_DRIVE", r"C:\Users\Desktop\Meu Drive\Cliente\Base_Monitoramento_Links_v1.xlsx")
NOME_ABA       = os.getenv("NOME_ABA", "Monitoramento")
MODO_EXECUCAO  = os.getenv("MODO_EXECUCAO", "completo")

COL_ID              = "ID"
COL_ID_BANCO        = "ID_Banco"
COL_NOME_PRODUTO    = "Nome do Produto"
COL_PLATAFORMA      = "Plataforma"
COL_LINK_PRINCIPAL  = "Link Principal (Link 1)"
COL_LINK_RESERVA    = "Links Reservas (Link 2)"
COL_VIDEOS          = "Vídeos Atrelados (YouTube)"
COL_LINK_ENCURTADO  = "Link Encurtado"
COL_STATUS          = "Status"
COL_ULT_ATUALIZACAO = "Ultima Atualização"

# Padrão de domínio do encurtador da conta
SHORTIO_DOMINIO_ESPERADO = "reidascapotas.s.gy"

# Status normalizados
STATUS_EM_ESTOQUE    = "em_estoque"
STATUS_ESGOTADO      = "esgotado"
STATUS_LINK_QUEBRADO = "link_quebrado"
STATUS_SEM_RESERVA   = "sem_reserva"

# Tipos de operação para historico_ajustes
TIPO_CRIACAO       = "CRIACAO"
TIPO_AUTO_CURA     = "AUTO_CURA"
TIPO_ESGOTAMENTO   = "ESGOTAMENTO"
TIPO_SINCRONIZACAO = "SINCRONIZACAO"
TIPO_VERIFICACAO   = "VERIFICACAO"
TIPO_CORRECAO      = "CORRECAO_ENCURTADOR"

# ══════════════════════════════════════════════════════════════════════════
# 2. DELAY — 5 segundos entre toda chamada de API
# ══════════════════════════════════════════════════════════════════════════

DELAY_API        = 5.0
DELAY_PLAYWRIGHT = 3.0


def _aguardar(contexto: str = ""):
    print(f"   ⏳ [{contexto}] Aguardando {DELAY_API}s...")
    time.sleep(DELAY_API)


# ══════════════════════════════════════════════════════════════════════════
# 3. INICIALIZAÇÃO
# ══════════════════════════════════════════════════════════════════════════

load_dotenv()

supabase_url        = os.getenv("API_URL_SUPABASE")
supabase_key        = os.getenv("SERVICE_ROLE_SUPABASE_API_KEY")
shortio_key         = os.getenv("SHORT_IO_API_KEY")
shortio_domain      = os.getenv("SHORT_IO_DOMAIN")
ID_CLIENTE_SUPABASE = os.getenv("ID_CLIENTE_SUPABASE")

_erros_env = []
if not supabase_url:         _erros_env.append("API_URL_SUPABASE")
if not supabase_key:         _erros_env.append("SERVICE_ROLE_SUPABASE_API_KEY")
if not shortio_key:          _erros_env.append("SHORT_IO_API_KEY")
if not shortio_domain:       _erros_env.append("SHORT_IO_DOMAIN")
if not ID_CLIENTE_SUPABASE:  _erros_env.append("ID_CLIENTE_SUPABASE")

if _erros_env:
    print(f"❌ ERRO CRÍTICO: Variáveis ausentes no .env: {', '.join(_erros_env)}")
    print(f"   Adicione as variáveis faltantes e reinicie o robô.")
    exit(1)

supabase: Client = create_client(supabase_url, supabase_key)
print("✅ Conexão com Supabase estabelecida.")
print(f"   Cliente UUID     : {ID_CLIENTE_SUPABASE[:8]}...")
print(f"   Delay entre chamadas: {DELAY_API}s")

SHORTIO_HEADERS = {
    "accept":        "application/json",
    "content-type":  "application/json",
    "authorization": shortio_key,
}


# ══════════════════════════════════════════════════════════════════════════
# 4. UTILITÁRIOS
# ══════════════════════════════════════════════════════════════════════════

def _ts_utc() -> str:
    return datetime.now(timezone.utc).isoformat()

def _ts_br() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

def _str(valor) -> str:
    v = str(valor).strip() if valor is not None else ""
    return "" if v.lower() in ("nan", "none") else v

def _extrair_slug(url_curta) -> str:
    valor = _str(url_curta)
    return valor.rstrip("/").split("/")[-1] if valor else ""

def _sep(titulo: str = ""):
    linha = "─" * 70
    if titulo:
        print(f"\n{linha}\n  {titulo}\n{linha}")
    else:
        print(linha)


# ══════════════════════════════════════════════════════════════════════════
# 4b. VALIDAÇÃO DE PADRÃO DO LINK ENCURTADO
# ══════════════════════════════════════════════════════════════════════════

def validar_padrao_link_encurtado(url_encurtada: str) -> tuple[bool, str]:
    """
    Valida se o link encurtado está dentro do padrão esperado:
      - Deve ser uma URL válida
      - Deve pertencer ao domínio: reidascapotas.s.gy
      - Deve ter um slug (path) não vazio após o domínio
      - Não deve ter parâmetros de query desnecessários

    Retorna (valido: bool, motivo: str).
    """
    url = _str(url_encurtada)
    if not url:
        return False, "URL vazia"

    # Verificar se começa com http/https
    if not url.startswith(("http://", "https://")):
        return False, f"URL sem protocolo http/https: {url[:60]}"

    from urllib.parse import urlparse
    parsed = urlparse(url)

    # Verificar domínio
    dominio = parsed.netloc.lower()
    if SHORTIO_DOMINIO_ESPERADO not in dominio:
        return False, f"Domínio fora do padrão. Esperado '{SHORTIO_DOMINIO_ESPERADO}', encontrado '{dominio}'"

    # Verificar slug não vazio
    slug = parsed.path.strip("/")
    if not slug:
        return False, "Slug vazio — URL não possui path após o domínio"

    # Verificar tamanho mínimo do slug (Short.io gera slugs de 4+ chars)
    if len(slug) < 4:
        return False, f"Slug muito curto: '{slug}' ({len(slug)} chars)"

    return True, f"OK — {dominio}/{slug}"


def _listar_links_encurtados_invalidos(df: pd.DataFrame) -> list[dict]:
    """
    Percorre todas as linhas e retorna lista de dicionários com
    links encurtados que não passam na validação de padrão.
    Linhas sem Nome do Produto são ignoradas.
    """
    invalidos = []
    for index, linha in df.iterrows():
        nome = _str(linha.get(COL_NOME_PRODUTO, ""))
        if not nome:
            continue
        url_enc = _str(linha.get(COL_LINK_ENCURTADO, ""))
        if not url_enc:
            continue
        valido, motivo = validar_padrao_link_encurtado(url_enc)
        if not valido:
            invalidos.append({
                "index":    index,
                "id":       _str(linha.get(COL_ID, str(index + 1))),
                "produto":  nome,
                "url":      url_enc,
                "motivo":   motivo,
            })
    return invalidos


# ══════════════════════════════════════════════════════════════════════════
# 5. PLAYWRIGHT — verificação via encurtador e link direto
# ══════════════════════════════════════════════════════════════════════════

def _sanitizar_url(url: str) -> str:
    """
    Remove parâmetros de campanha da URL do Mercado Livre antes de acessar.
    """
    from urllib.parse import urlparse, urlunparse
    valor = _str(url)
    if not valor:
        return valor
    parsed = urlparse(valor)
    url_limpa = urlunparse((parsed.scheme, parsed.netloc, parsed.path, "", "", ""))
    if url_limpa != valor:
        print(f"   🧹 URL sanitizada (parâmetros removidos)")
        print(f"      Original : {valor[:80]}")
        print(f"      Limpa    : {url_limpa[:80]}")
    return url_limpa


def _abrir_e_verificar(url: str, descricao: str) -> tuple[bool, str]:
    """
    Abre URL com Playwright e verifica se o botão 'Ir para produto' está presente.
    Único critério de sucesso: presença do botão 'Ir para produto' (Gateway Social ML).
    """
    url_limpa = _sanitizar_url(url)
    print(f"   🌐 [{descricao}] Acessando: {url_limpa[:80]}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/119.0.0.0 Safari/537.36"
        ))
        page = context.new_page()
        try:
            page.goto(url_limpa, wait_until="networkidle", timeout=60000)
            time.sleep(DELAY_PLAYWRIGHT)
            url_final = page.url
            titulo    = page.title()
            print(f"   📄 Título    : {titulo[:70]}")
            print(f"   🔁 URL final : {url_final[:80]}")

            botao = page.query_selector("text='Ir para produto'")
            encontrado = botao is not None

            if encontrado:
                print(f"   ✅ Botão 'Ir para produto' ENCONTRADO — link válido.")
            else:
                if page.query_selector("text='Minhas listas'") or page.query_selector("text='Minhas recomendações'"):
                    print(f"   ❌ Botão NÃO encontrado — página é uma VITRINE/LISTA do vendedor.")
                    print(f"       → Substitua o link por um link direto de produto.")
                elif page.query_selector("text='Comprar agora'") or page.query_selector("text='Adicionar ao carrinho'"):
                    print(f"   ❌ Botão NÃO encontrado — página é um PRODUTO DIRETO (sem gateway).")
                    print(f"       → Use o link da página social/gateway que contém 'Ir para produto'.")
                else:
                    print(f"   ❌ Botão NÃO encontrado — estrutura de página não reconhecida.")
                    print(f"       → Pode ser CAPTCHA, erro 404 ou layout fora do padrão.")

            return encontrado, url_final

        except Exception as e:
            print(f"   ⚠️  Erro Playwright: {e}")
            return False, url
        finally:
            browser.close()


def verificar_produto_completo(
    link_principal: str,
    short_url: str,
    nome_produto: str,
) -> dict:
    """
    Verificação dupla: encurtador (primário) + link_principal (secundário).
    """
    print(f"\n   🔬 Verificação dupla para: {nome_produto[:50]}")

    resultado = {
        "encurtador_ok":    False,
        "principal_ok":     False,
        "precisa_patch":    False,
        "precisa_autocura": False,
        "url_final_enc":    "",
    }

    if _str(short_url):
        print(f"\n   ─── PASSO A: Teste via ENCURTADOR ───────────────────────────────")
        enc_ok, url_final_enc = _abrir_e_verificar(short_url, "ENCURTADOR")
        resultado["encurtador_ok"]  = enc_ok
        resultado["url_final_enc"]  = url_final_enc
    else:
        print(f"   ⚠️  PASSO A pulado — sem short_url disponível.")

    if not resultado["encurtador_ok"] and _str(link_principal):
        print(f"\n   ─── PASSO B: Teste via LINK PRINCIPAL ───────────────────────────")
        princ_ok, _ = _abrir_e_verificar(link_principal, "LINK PRINCIPAL")
        resultado["principal_ok"] = princ_ok
    elif resultado["encurtador_ok"]:
        print(f"\n   ✅ PASSO B pulado — encurtador já passou. Link principal OK por inferência.")
        resultado["principal_ok"] = True

    enc_ok   = resultado["encurtador_ok"]
    princ_ok = resultado["principal_ok"]

    if enc_ok:
        print(f"\n   ✅ RESULTADO: Encurtador FUNCIONA — produto acessível pelo link curto.")
    elif not enc_ok and princ_ok:
        resultado["precisa_patch"] = True
        print(f"\n   🔧 RESULTADO: Encurtador FALHOU, link principal OK.")
        print(f"       → Ação: PATCH Short.io para corrigir destino do encurtador.")
    else:
        resultado["precisa_autocura"] = True
        print(f"\n   ❌ RESULTADO: AMBOS falharam — produto fora de estoque.")
        print(f"       → Ação: auto-cura com fila de reservas.")

    return resultado


def testar_lista_de_links(links: list[str], nome: str) -> tuple[str, list[str]]:
    """Testa lista sequencialmente via Playwright. Retorna (primeiro_ativo, restantes)."""
    falhos = []
    for i, url in enumerate(links):
        if not url:
            continue
        posicao = "principal" if i == 0 else f"reserva [{i}]"
        print(f"\n   🧪 Testando {posicao}: {url[:70]}")
        ok, _ = _abrir_e_verificar(url, posicao.upper())
        if ok:
            restantes = [l for l in links[i+1:] if l]
            if falhos:
                print(f"   📋 Reprovados antes: {len(falhos)}")
                for fl in falhos:
                    print(f"      ✗ {fl[:70]}")
            return url, restantes
        falhos.append(url)
        print(f"   ❌ Reprovado — próximo...")
    print(f"   🚨 Todos os {len(links)} link(s) falharam.")
    return "", []


# ══════════════════════════════════════════════════════════════════════════
# 6. SUPABASE — com cascata de busca e logs completos
# ══════════════════════════════════════════════════════════════════════════

def buscar_link_no_supabase(
    nome_produto: str = "",
    short_id:     str = "",
    short_url:    str = "",
) -> tuple[str, str, str, str, str]:
    """
    Busca o produto em links_monitoramento por cascata de fallbacks:
      1º) nome_produto (correspondência exata)
      2º) short_id (idString lnk_xxx)
      3º) short_url (URL encurtada completa)
    """
    campos_select = "id, short_id, short_url, plataforma, nome_produto"

    if _str(nome_produto):
        _aguardar("Supabase SELECT por nome_produto")
        try:
            res = (supabase.table("links_monitoramento")
                   .select(campos_select)
                   .eq("nome_produto", nome_produto)
                   .limit(1).execute())
            if res.data:
                r = res.data[0]
                print(f"   ✅ [SUPABASE] Produto encontrado por nome_produto.")
                return r.get("id",""), r.get("short_id",""), r.get("short_url",""), r.get("plataforma",""), r.get("nome_produto","")
        except Exception as e:
            print(f"   ⚠️  [SUPABASE] Erro busca por nome: {e}")

    if _str(short_id) and short_id.startswith("lnk_"):
        print(f"   🔄 [SUPABASE] Fallback: buscando por short_id = {short_id[:20]}...")
        _aguardar("Supabase SELECT por short_id")
        try:
            res = (supabase.table("links_monitoramento")
                   .select(campos_select)
                   .eq("short_id", short_id)
                   .limit(1).execute())
            if res.data:
                r = res.data[0]
                print(f"   ✅ [SUPABASE] Produto encontrado por short_id!")
                return r.get("id",""), r.get("short_id",""), r.get("short_url",""), r.get("plataforma",""), r.get("nome_produto","")
        except Exception as e:
            print(f"   ⚠️  [SUPABASE] Erro busca por short_id: {e}")

    if _str(short_url):
        print(f"   🔄 [SUPABASE] Fallback: buscando por short_url...")
        _aguardar("Supabase SELECT por short_url")
        try:
            res = (supabase.table("links_monitoramento")
                   .select(campos_select)
                   .eq("short_url", short_url)
                   .limit(1).execute())
            if res.data:
                r = res.data[0]
                print(f"   ✅ [SUPABASE] Produto encontrado por short_url!")
                return r.get("id",""), r.get("short_id",""), r.get("short_url",""), r.get("plataforma",""), r.get("nome_produto","")
        except Exception as e:
            print(f"   ⚠️  [SUPABASE] Erro busca por short_url: {e}")

    print(f"   ❌ [SUPABASE] Produto NÃO encontrado em nenhum dos fallbacks.")
    return "", "", "", "", ""


def registrar_historico(
    link_uuid:     str,
    url_antiga:    str,
    url_nova:      str,
    motivo:        str,
    tipo_operacao: str,
    produto_nome:  str = "",
    plataforma:    str = "",
) -> bool:
    if not _str(link_uuid):
        print("   ⚠️  [HISTÓRICO] Não registrado — link_uuid ausente.")
        return False

    registro = {
        "link_id":        link_uuid,
        "data_alteracao": _ts_utc(),
        "url_antiga":     _str(url_antiga),
        "url_nova":       _str(url_nova),
        "motivo":         f"[{tipo_operacao}] {motivo}",
        "tipo_operacao":  _str(tipo_operacao),
        "produto_nome":   _str(produto_nome),
        "plataforma":     _str(plataforma),
    }

    print(f"\n   📝 [SUPABASE INSERT] → historico_ajustes")
    print(f"      link_id       : {link_uuid[:8]}...")
    print(f"      tipo_operacao : {tipo_operacao}")
    print(f"      produto_nome  : {produto_nome or '—'}")
    print(f"      url_antiga    : {_str(url_antiga)[:70] or '— (criação)'}")
    print(f"      url_nova      : {_str(url_nova)[:70] or '— (esgotamento)'}")
    _aguardar("Supabase INSERT historico_ajustes")

    try:
        supabase.table("historico_ajustes").insert(registro).execute()
        print(f"   ✅ [HISTÓRICO] Inserido com sucesso.")
        return True
    except Exception as e:
        if any(col in str(e) for col in ("tipo_operacao", "produto_nome", "plataforma")):
            print(f"   ⚠️  Colunas BI ausentes — inserindo sem elas.")
            try:
                registro_base = {k: v for k, v in registro.items()
                                 if k not in ("tipo_operacao","produto_nome","plataforma")}
                supabase.table("historico_ajustes").insert(registro_base).execute()
                print(f"   ✅ [HISTÓRICO] Inserido (sem colunas BI).")
                return True
            except Exception as e2:
                print(f"   ❌ [HISTÓRICO] FALHA: {e2}")
                return False
        print(f"   ❌ [HISTÓRICO] FALHA: {e}")
        return False


def atualizar_verificacao_ok(
    link_uuid:    str,
    produto_nome: str = "",
    plataforma:   str = "",
    short_url:    str = "",
) -> bool:
    if not _str(link_uuid):
        return False

    print(f"\n   ✏️  [SUPABASE UPDATE] → links_monitoramento (verificacao OK)")
    _aguardar("Supabase UPDATE verificacao_ok")

    try:
        supabase.table("links_monitoramento").update({
            "ultima_verificacao": _ts_utc(),
            "status_estoque":     STATUS_EM_ESTOQUE,
        }).eq("id", link_uuid).execute()
        print(f"   ✅ [SUPABASE] Timestamp atualizado.")
        registrar_historico(
            link_uuid=link_uuid, url_antiga="", url_nova=short_url or "",
            motivo="Verificação de rotina — produto em estoque pelo encurtador",
            tipo_operacao=TIPO_VERIFICACAO, produto_nome=produto_nome, plataforma=plataforma,
        )
        return True
    except Exception as e:
        print(f"   ❌ [SUPABASE] FALHA: {e}")
        return False


def corrigir_encurtador_supabase(
    link_uuid:    str,
    short_id:     str,
    nova_url:     str,
    produto_nome: str = "",
    plataforma:   str = "",
    url_antiga:   str = "",
) -> bool:
    if not _str(link_uuid):
        return False

    print(f"\n   ✏️  [SUPABASE UPDATE] → links_monitoramento (correcao encurtador)")
    _aguardar("Supabase UPDATE correcao_encurtador")

    try:
        supabase.table("links_monitoramento").update({
            "link_principal":     nova_url,
            "ultima_verificacao": _ts_utc(),
            "status_estoque":     STATUS_EM_ESTOQUE,
        }).eq("id", link_uuid).execute()
        print(f"   ✅ [SUPABASE] link_principal atualizado.")
        registrar_historico(
            link_uuid=link_uuid, url_antiga=url_antiga, url_nova=nova_url,
            motivo="Encurtador apontava para link desatualizado — corrigido automaticamente",
            tipo_operacao=TIPO_CORRECAO, produto_nome=produto_nome, plataforma=plataforma,
        )
        return True
    except Exception as e:
        print(f"   ❌ [SUPABASE] FALHA: {e}")
        return False


def atualizar_link_supabase(
    link_uuid:          str,
    nova_url_principal: str,
    status:             str,
    url_antiga:         str,
    motivo:             str,
    tipo_operacao:      str,
    produto_nome:       str = "",
    plataforma:         str = "",
) -> bool:
    if not _str(link_uuid):
        return False

    print(f"\n   ✏️  [SUPABASE UPDATE] → links_monitoramento")
    _aguardar("Supabase UPDATE link_principal")

    try:
        supabase.table("links_monitoramento").update({
            "link_principal":     nova_url_principal,
            "status_estoque":     status,
            "ultima_verificacao": _ts_utc(),
        }).eq("id", link_uuid).execute()
        print(f"   ✅ [SUPABASE] Atualizado.")
        registrar_historico(
            link_uuid=link_uuid, url_antiga=url_antiga, url_nova=nova_url_principal,
            motivo=motivo, tipo_operacao=tipo_operacao,
            produto_nome=produto_nome, plataforma=plataforma,
        )
        return True
    except Exception as e:
        print(f"   ❌ [SUPABASE] FALHA: {e}")
        return False


def atualizar_shortid_supabase(
    link_uuid:    str,
    short_id:     str,
    short_url:    str,
    produto_nome: str = "",
    plataforma:   str = "",
    motivo:       str = "Sincronização de idString",
) -> bool:
    if not _str(link_uuid):
        return False

    print(f"\n   ✏️  [SUPABASE UPDATE] → short_id/short_url")
    _aguardar("Supabase UPDATE short_id")

    try:
        supabase.table("links_monitoramento").update({
            "short_id":  _str(short_id),
            "short_url": _str(short_url),
        }).eq("id", link_uuid).execute()
        print(f"   ✅ [SUPABASE] short_id atualizado.")
        registrar_historico(
            link_uuid=link_uuid, url_antiga="", url_nova=short_url,
            motivo=motivo, tipo_operacao=TIPO_SINCRONIZACAO,
            produto_nome=produto_nome, plataforma=plataforma,
        )
        return True
    except Exception as e:
        print(f"   ❌ [SUPABASE] FALHA: {e}")
        return False


def cadastrar_link_monitoramento(
    cliente_id: str, nome_produto: str, link_principal: str,
    link_reserva: str = "", short_id: str = "", short_url: str = "",
    videos_atrelados: str = "", status_estoque: str = STATUS_EM_ESTOQUE,
    plataforma: str = ""
) -> str | None:
    dados = {
        "cliente_id":        cliente_id,
        "nome_produto":      _str(nome_produto),
        "link_principal":    _str(link_principal),
        "link_reserva":      _str(link_reserva),
        "short_id":          _str(short_id),
        "short_url":         _str(short_url),
        "videos_atrelados":  _str(videos_atrelados),
        "status_estoque":    status_estoque,
        "ultima_verificacao": _ts_utc(),
        "plataforma":        _str(plataforma),
    }
    print(f"\n   📝 [SUPABASE INSERT] → links_monitoramento")
    for k, v in dados.items():
        print(f"      {k:<22}: {str(v)[:70]}")
    _aguardar("Supabase INSERT links_monitoramento")
    try:
        res = supabase.table("links_monitoramento").insert(dados).execute()
        link_uuid = res.data[0]["id"]
        print(f"   ✅ [LINK] Criado! UUID: {link_uuid}")
        registrar_historico(
            link_uuid=link_uuid, url_antiga="", url_nova=link_principal,
            motivo=f"Produto '{nome_produto}' cadastrado no sistema",
            tipo_operacao=TIPO_CRIACAO, produto_nome=nome_produto, plataforma=plataforma,
        )
        return link_uuid
    except Exception as e:
        print(f"   ❌ [LINK] FALHA: {e}")
        return None


def auto_cadastrar_produto_do_excel(
    nome_produto:    str,
    link_principal:  str,
    link_reserva:    str = "",
    plataforma:      str = "",
    videos_atrelados: str = "",
    short_id:        str = "",
    short_url:       str = "",
) -> str:
    print(f"\n   🆕 [AUTO-CADASTRO] Produto novo detectado no Excel → cadastrando no Supabase...")
    print(f"      nome_produto  : {nome_produto[:60]}")
    print(f"      link_principal: {link_principal[:70]}")
    print(f"      cliente_id    : {ID_CLIENTE_SUPABASE[:8]}...")

    link_uuid = cadastrar_link_monitoramento(
        cliente_id       = ID_CLIENTE_SUPABASE,
        nome_produto     = nome_produto,
        link_principal   = link_principal,
        link_reserva     = link_reserva,
        short_id         = short_id,
        short_url        = short_url,
        videos_atrelados = videos_atrelados,
        status_estoque   = STATUS_EM_ESTOQUE,
        plataforma       = plataforma,
    )

    if link_uuid:
        print(f"   ✅ [AUTO-CADASTRO] Produto cadastrado! UUID: {link_uuid[:8]}...")
    else:
        print(f"   ❌ [AUTO-CADASTRO] Falha ao cadastrar. Verifique permissões do SERVICE_ROLE.")

    return link_uuid or ""


# ══════════════════════════════════════════════════════════════════════════
# 7. SHORT.IO
# ══════════════════════════════════════════════════════════════════════════

def criar_link_shortio(url_destino: str, titulo: str = None) -> dict | None:
    if not _str(url_destino):
        return None
    payload = {"domain": shortio_domain, "originalURL": url_destino}
    if titulo:
        payload["title"] = titulo
    print(f"\n   📡 [SHORT.IO POST] → /links")
    print(f"      destino : {url_destino[:80]}")
    _aguardar("Short.io POST /links")
    try:
        r = requests.post("https://api.short.io/links", json=payload,
                          headers=SHORTIO_HEADERS, timeout=15)
        if r.status_code in (200, 201):
            d = r.json()
            print(f"   ✅ [SHORT.IO] Criado! shortURL: {d.get('shortURL')} | idString: {d.get('idString')}")
            return d
        print(f"   ❌ [SHORT.IO] FALHA HTTP {r.status_code}: {r.text[:200]}")
        return None
    except Exception as e:
        print(f"   ❌ [SHORT.IO] EXCEÇÃO: {e}")
        return None


def buscar_destino_atual_shortio(slug: str) -> str:
    if not slug:
        return ""
    dados = buscar_link_por_slug(slug)
    if dados:
        destino = dados.get("originalURL", "") or dados.get("original_url", "")
        if destino:
            print(f"   🔗 [SHORT.IO] Destino atual do encurtador: {destino[:80]}")
        return destino
    return ""


def buscar_link_por_slug(slug: str) -> dict | None:
    if not slug:
        return None
    print(f"\n   🔎 [SHORT.IO GET] → /links/expand?path={slug}")
    _aguardar("Short.io GET /expand")
    try:
        r = requests.get("https://api.short.io/links/expand",
                         headers=SHORTIO_HEADERS,
                         params={"domain": shortio_domain, "path": slug},
                         timeout=15)
        if r.status_code == 200:
            d = r.json()
            print(f"   ✅ [SHORT.IO] idString: {d.get('idString','—')}")
            return d
        if r.status_code == 404:
            print(f"   ⚠️  Slug '{slug}' não existe.")
            return None
        print(f"   ⚠️  HTTP {r.status_code}")
        return None
    except Exception as e:
        print(f"   ⚠️  [SHORT.IO] Exceção: {e}")
        return None


def atualizar_destino_shortio(id_string: str, nova_url: str) -> bool:
    if not _str(id_string):
        print("   ⚠️  [SHORT.IO] Atualização abortada — id_string vazio.")
        return False

    if not id_string.startswith("lnk_"):
        print(f"   ❌ [SHORT.IO] '{id_string[:20]}' não é idString válido (deve começar com 'lnk_').")
        return False

    print(f"\n   🔄 [SHORT.IO POST] → /links/{id_string[:20]}...")
    print(f"      nova URL : {nova_url[:80]}")
    _aguardar("Short.io POST /links")

    try:
        r = requests.post(f"https://api.short.io/links/{id_string}",
                          json={"originalURL": nova_url},
                          headers=SHORTIO_HEADERS, timeout=15)
        if r.status_code in (200, 201):
            print(f"   ✅ [SHORT.IO] Destino ATUALIZADO com sucesso!")
            return True
        print(f"   ❌ [SHORT.IO] FALHA HTTP {r.status_code}: {r.text[:200]}")
        return False
    except Exception as e:
        print(f"   ❌ [SHORT.IO] EXCEÇÃO: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════
# 8. SINCRONIZAÇÃO DE IDENTIDADES
# ══════════════════════════════════════════════════════════════════════════

def _normalizar_lista_separada(valor: str) -> str:
    if not valor:
        return ""
    itens = [i.strip() for i in valor.split(";") if i.strip()]
    return ";".join(itens)


def _calcular_proximo_id(df: pd.DataFrame) -> str:
    """
    Encontra o último ID numérico preenchido na planilha e retorna o próximo.
    Suporta IDs numéricos puros (1, 2, 3...) e com prefixo (A-001, B-002...).
    """
    ultimo_id = ""
    for idx in reversed(df.index):
        val = _str(df.at[idx, COL_ID] if COL_ID in df.columns else "")
        if val:
            ultimo_id = val
            break

    if not ultimo_id:
        return "1"

    match = re.search(r'^(.*?)(\d+)$', ultimo_id)
    if match:
        prefixo  = match.group(1)
        numero   = int(match.group(2))
        largura  = len(match.group(2))
        return f"{prefixo}{str(numero + 1).zfill(largura)}"

    return str(len(df) + 1)


def _determinar_acao_linha(
    nome:          str,
    link_principal: str,
    link_reserva:  str,
    link_encurtado: str,
) -> str:
    """
    Determina a ação correta para a linha com base nas regras de colunas:

    CRIAR   → quando não há link encurtado mas há link(s) para encurtar
    VALIDAR → quando há link encurtado e link(s) — verificar se está atualizado
    IGNORAR → linha sem nome do produto
    SKIP    → sem links suficientes para qualquer ação
    """
    if not nome:
        return "IGNORAR"

    tem_principal  = bool(_str(link_principal))
    tem_reserva    = bool(_str(link_reserva))
    tem_encurtado  = bool(_str(link_encurtado))

    # Caso A: não tem principal, tem reserva, sem encurtado → CRIAR com reserva
    if not tem_principal and tem_reserva and not tem_encurtado:
        return "CRIAR_DE_RESERVA"

    # Caso B: tem principal, tem reserva, sem encurtado → CRIAR com principal
    if tem_principal and not tem_encurtado:
        return "CRIAR"

    # Caso C: tem principal, sem reserva, com encurtado → VALIDAR
    if tem_principal and not tem_reserva and tem_encurtado:
        return "VALIDAR"

    # Caso D: tem principal, com reserva, com encurtado → VALIDAR
    if tem_principal and tem_reserva and tem_encurtado:
        return "VALIDAR"

    # Caso E (NOVO): sem principal, com encurtado (com ou sem reserva) → VALIDAR
    # O encurtador é a fonte de verdade — buscar link principal via Short.io
    if not tem_principal and tem_encurtado:
        return "VALIDAR"

    # Sem links → nada a fazer
    if not tem_principal and not tem_reserva:
        return "SKIP"

    return "SKIP"


def sincronizar_identidades(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str], int]:
    _sep("🔄 SINCRONIZAÇÃO DE IDENTIDADES v6.0")

    # Garantir colunas obrigatórias
    if COL_ID not in df.columns:
        df.insert(0, COL_ID, "")
    if COL_ID_BANCO not in df.columns:
        df.insert(1, COL_ID_BANCO, "")
    if COL_PLATAFORMA not in df.columns:
        df[COL_PLATAFORMA] = ""
    if COL_STATUS not in df.columns:
        df[COL_STATUS] = ""
    if COL_ULT_ATUALIZACAO not in df.columns:
        df[COL_ULT_ATUALIZACAO] = ""

    nao_encontrados: list[str] = []
    auto_cadastros: int = 0

    # ── Relatório de padrão de links encurtados ───────────────────────────
    invalidos_padrao = _listar_links_encurtados_invalidos(df)
    if invalidos_padrao:
        print(f"\n   ⚠️  LINKS ENCURTADOS FORA DO PADRÃO ({len(invalidos_padrao)}):")
        for item in invalidos_padrao:
            print(f"      [ID={item['id']}] {item['produto'][:45]}")
            print(f"         URL    : {item['url'][:70]}")
            print(f"         Motivo : {item['motivo']}")
    else:
        print(f"\n   ✅ Todos os links encurtados existentes estão dentro do padrão.")

    for index, linha in df.iterrows():
        nome          = _str(linha.get(COL_NOME_PRODUTO, ""))
        id_banco      = _str(linha.get(COL_ID_BANCO, ""))       # NUNCA alterar
        id_item       = _str(linha.get(COL_ID, "") or "")
        url_curta     = _str(linha.get(COL_LINK_ENCURTADO, ""))
        link_princ    = _str(linha.get(COL_LINK_PRINCIPAL, ""))
        link_reserva  = _str(linha.get(COL_LINK_RESERVA, ""))
        plataforma    = _str(linha.get(COL_PLATAFORMA, ""))
        videos_raw    = _str(linha.get(COL_VIDEOS, ""))
        reservas_raw  = _str(linha.get(COL_LINK_RESERVA, ""))

        # REGRA: sem nome do produto → nenhuma ação
        if not nome:
            continue

        # ── REGRA ID: preencher apenas se vazio ──────────────────────────
        if not id_item:
            novo_id = _calcular_proximo_id(df)
            df.at[index, COL_ID] = novo_id
            id_item = novo_id
            print(f"\n  ▶ [{id_item}] {nome[:58]} (ID gerado)")
        else:
            print(f"\n  ▶ [{id_item}] {nome[:58]}")

        # ── REGRA PLATAFORMA: sempre "mercadolivre" ───────────────────────
        if plataforma.lower() != "mercadolivre":
            df.at[index, COL_PLATAFORMA] = "mercadolivre"
            plataforma = "mercadolivre"
            print(f"   🏪 Plataforma → 'mercadolivre'")

        # ── REGRA VÍDEOS: normalizar separador ; ─────────────────────────
        videos_norm = _normalizar_lista_separada(videos_raw)
        if videos_norm != videos_raw:
            df.at[index, COL_VIDEOS] = videos_norm

        # ── REGRA RESERVAS: normalizar separador ; ────────────────────────
        reservas_norm = _normalizar_lista_separada(reservas_raw)
        if reservas_norm != reservas_raw:
            df.at[index, COL_LINK_RESERVA] = reservas_norm

        # ── DETERMINAR AÇÃO ───────────────────────────────────────────────
        acao = _determinar_acao_linha(nome, link_princ, reservas_norm, url_curta)
        print(f"   🎯 Ação determinada: {acao}")

        slug = _extrair_slug(url_curta)

        # ── Validar padrão do encurtador existente ────────────────────────
        if url_curta:
            valido, motivo_val = validar_padrao_link_encurtado(url_curta)
            if valido:
                print(f"   ✅ Encurtado dentro do padrão: {motivo_val}")
            else:
                print(f"   ⚠️  Encurtado FORA do padrão: {motivo_val}")

        # ── AÇÃO: CRIAR encurtador (link principal disponível) ────────────
        if acao == "CRIAR":
            print(f"   🆕 Criando encurtador para link principal: {link_princ[:60]}")
            novo = criar_link_shortio(url_destino=link_princ, titulo=nome[:100])
            if novo:
                nova_url_curta = novo.get("shortURL", "")
                df.at[index, COL_LINK_ENCURTADO] = nova_url_curta
                url_curta = nova_url_curta
                slug = _extrair_slug(nova_url_curta)
                print(f"   ✅ Encurtador criado: {nova_url_curta}")
                # ID_Banco: não alterar — somente atualizar no Supabase se id_banco já existe
                if id_banco:
                    atualizar_shortid_supabase(
                        link_uuid=id_banco, short_id=novo.get("idString",""),
                        short_url=nova_url_curta, produto_nome=nome,
                        plataforma=plataforma, motivo="Encurtador criado na sincronização",
                    )

        # ── AÇÃO: CRIAR encurtador a partir da reserva ────────────────────
        elif acao == "CRIAR_DE_RESERVA":
            fila_reservas = [l.strip() for l in reservas_norm.split(";") if l.strip()]
            if fila_reservas:
                primeiro_reserva = fila_reservas[0]
                print(f"   🆕 Criando encurtador para link de reserva: {primeiro_reserva[:60]}")
                novo = criar_link_shortio(url_destino=primeiro_reserva, titulo=nome[:100])
                if novo:
                    nova_url_curta = novo.get("shortURL", "")
                    # Promover a reserva para link principal
                    df.at[index, COL_LINK_PRINCIPAL]  = primeiro_reserva
                    df.at[index, COL_LINK_ENCURTADO]  = nova_url_curta
                    df.at[index, COL_LINK_RESERVA]    = ";".join(fila_reservas[1:])
                    url_curta  = nova_url_curta
                    link_princ = primeiro_reserva
                    slug       = _extrair_slug(nova_url_curta)
                    print(f"   ✅ Encurtador criado: {nova_url_curta}")
                    print(f"   🔄 Reserva promovida para link principal.")
                    if id_banco:
                        atualizar_shortid_supabase(
                            link_uuid=id_banco, short_id=novo.get("idString",""),
                            short_url=nova_url_curta, produto_nome=nome,
                            plataforma=plataforma,
                            motivo="Encurtador criado a partir de reserva (sem link principal)",
                        )

        # ── AÇÃO: VALIDAR encurtador existente vs planilha ────────────────
        elif acao == "VALIDAR":
            if slug:
                destino_atual = buscar_destino_atual_shortio(slug)
                if destino_atual and destino_atual != link_princ:
                    print(f"   🔄 Link Principal desatualizado — sincronizando com Short.io:")
                    print(f"      Planilha : {link_princ[:70]}")
                    print(f"      Short.io : {destino_atual[:70]}")
                    df.at[index, COL_LINK_PRINCIPAL] = destino_atual
                    link_princ = destino_atual
                    print(f"   ✅ Link Principal atualizado na planilha.")
                elif destino_atual == link_princ:
                    print(f"   ✅ Link Principal e encurtador já estão sincronizados.")

        # ── ID_Banco: APENAS recuperar se vazio (NUNCA sobrescrever) ──────
        if not id_banco:
            id_string_atual = ""
            if slug:
                dados_api = buscar_link_por_slug(slug)
                if dados_api:
                    id_string_atual = dados_api.get("idString", "")
                    # Se não temos link_principal, recuperar do Short.io como fonte de verdade
                    if not link_princ:
                        destino_shortio = dados_api.get("originalURL", "") or dados_api.get("original_url", "")
                        if destino_shortio:
                            df.at[index, COL_LINK_PRINCIPAL] = destino_shortio
                            link_princ = destino_shortio
                            print(f"   🔗 Link Principal recuperado do Short.io: {destino_shortio[:70]}")

            uuid_enc, _, _, _, _ = buscar_link_no_supabase(
                nome_produto=nome,
                short_id=id_string_atual,
                short_url=url_curta,
            )
            if uuid_enc:
                # Encontrado no banco → preencher na planilha
                df.at[index, COL_ID_BANCO] = uuid_enc
                id_banco = uuid_enc
                print(f"   ✏️  ID_Banco recuperado do Supabase: {uuid_enc[:8]}...")
            elif link_princ:
                # Produto não encontrado → auto-cadastro
                print(f"   ⚠️  Produto não encontrado no Supabase → auto-cadastro...")
                novo_uuid = auto_cadastrar_produto_do_excel(
                    nome_produto     = nome,
                    link_principal   = link_princ,
                    link_reserva     = reservas_norm,
                    plataforma       = plataforma,
                    videos_atrelados = videos_norm,
                    short_id         = id_string_atual,
                    short_url        = url_curta,
                )
                if novo_uuid:
                    df.at[index, COL_ID_BANCO] = novo_uuid
                    id_banco = novo_uuid
                    auto_cadastros += 1
                    print(f"   ✅ ID_Banco gravado na planilha: {novo_uuid[:8]}...")
                else:
                    print(f"   ❌ Auto-cadastro falhou.")
            else:
                print(f"   ⚠️  Sem link_principal disponível — auto-cadastro ignorado (mesmo após busca no Short.io).")
        else:
            print(f"   🔒 ID_Banco preservado: {id_banco[:8]}... (não alterado)")

        # ── Sincronizar short_id/short_url no banco se necessário ─────────
        if slug and id_banco:
            dados_api = buscar_link_por_slug(slug)
            if dados_api:
                id_str = dados_api.get("idString", "")
                if id_str:
                    atualizar_shortid_supabase(
                        link_uuid=id_banco, short_id=id_str,
                        short_url=url_curta, produto_nome=nome, plataforma=plataforma,
                        motivo=f"Sincronização de idString pelo slug '{slug}'",
                    )
            elif acao != "CRIAR" and acao != "CRIAR_DE_RESERVA":
                # Slug existe na planilha mas não encontrado no Short.io
                nao_encontrados.append(f"[ID={id_item}] {nome} (slug '{slug}')")

    return df, nao_encontrados, auto_cadastros


# ══════════════════════════════════════════════════════════════════════════
# 9. AUTO-CURA
# ══════════════════════════════════════════════════════════════════════════

def executar_autocura(
    df: pd.DataFrame, index: int,
    url_falhou: str, link_uuid: str,
    short_id_banco: str, plataforma: str = "",
) -> tuple[bool, str]:
    nome         = _str(df.at[index, COL_NOME_PRODUTO])
    reservas_raw = _str(df.at[index, COL_LINK_RESERVA])

    if not reservas_raw:
        print(f"   🚨 SEM RESERVA — '{nome[:40]}' ficará esgotado.")
        if link_uuid:
            _aguardar("Supabase UPDATE esgotado")
            try:
                supabase.table("links_monitoramento").update({
                    "status_estoque":     STATUS_ESGOTADO,
                    "ultima_verificacao": _ts_utc(),
                }).eq("id", link_uuid).execute()
            except Exception as e:
                print(f"   ❌ [SUPABASE] Falha: {e}")
            registrar_historico(
                link_uuid=link_uuid, url_antiga=url_falhou, url_nova="",
                motivo="Nenhum link de reserva cadastrado",
                tipo_operacao=TIPO_ESGOTAMENTO, produto_nome=nome, plataforma=plataforma,
            )
        return False, ""

    fila = [l.strip() for l in reservas_raw.split(";") if l.strip()]
    if not fila:
        return False, ""

    print(f"   🧪 Testando {len(fila)} reserva(s)...")
    nova_url, fila_restante = testar_lista_de_links(fila, nome)

    if not nova_url:
        print(f"   🚨 Todas as reservas falharam.")
        if link_uuid:
            _aguardar("Supabase UPDATE sem_reserva")
            try:
                supabase.table("links_monitoramento").update({
                    "status_estoque":     STATUS_SEM_RESERVA,
                    "ultima_verificacao": _ts_utc(),
                }).eq("id", link_uuid).execute()
            except Exception:
                pass
            registrar_historico(
                link_uuid=link_uuid, url_antiga=url_falhou, url_nova="",
                motivo=f"Todos os {len(fila)} links de reserva reprovados",
                tipo_operacao=TIPO_ESGOTAMENTO, produto_nome=nome, plataforma=plataforma,
            )
        return False, ""

    df.at[index, COL_LINK_PRINCIPAL]  = nova_url
    df.at[index, COL_LINK_RESERVA]    = ";".join(fila_restante)
    df.at[index, COL_STATUS]          = STATUS_EM_ESTOQUE
    df.at[index, COL_ULT_ATUALIZACAO] = _ts_br()

    if short_id_banco and short_id_banco.startswith("lnk_"):
        atualizar_destino_shortio(short_id_banco, nova_url)
    else:
        print(f"   ⚠️  short_id ausente — Short.io não atualizado. Execute Sincronizar.")

    atualizar_link_supabase(
        link_uuid=link_uuid, nova_url_principal=nova_url,
        status=STATUS_EM_ESTOQUE, url_antiga=url_falhou,
        motivo="Link reprovado substituído por reserva aprovada no teste",
        tipo_operacao=TIPO_AUTO_CURA, produto_nome=nome, plataforma=plataforma,
    )
    return True, nova_url


# ══════════════════════════════════════════════════════════════════════════
# 10. FLUXO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════

def iniciar_monitoramento():
    print("\n" + "═" * 70)
    print("  🤖 SISTEMA DE MONITORAMENTO E AUTO-CURA v6.0.0")
    print(f"  Modo  : {MODO_EXECUCAO.upper()}")
    print(f"  Delay : {DELAY_API}s por chamada de API")
    print("═" * 70)

    if not os.path.exists(CAMINHO_DRIVE):
        print(f"❌ Arquivo não encontrado: {CAMINHO_DRIVE}")
        return

    try:
        print(f"\n📂 Lendo aba '{NOME_ABA}'...")
        df = pd.read_excel(CAMINHO_DRIVE, sheet_name=NOME_ABA, dtype=str)
        df = df.where(df.notna(), other=None)
        print(f"   {len(df)} linha(s) carregadas.")
    except Exception as e:
        print(f"❌ Erro ao ler planilha: {e}")
        return

    df, nao_encontrados, auto_cadastrados_sinc = sincronizar_identidades(df)

    if MODO_EXECUCAO == "sincronizar":
        try:
            df.to_excel(CAMINHO_DRIVE, sheet_name=NOME_ABA, index=False)
            print("\n💾 Planilha salva.")
        except Exception as e:
            print(f"❌ Erro: {e}")
        print(f"   🆕 Produtos auto-cadastrados: {auto_cadastrados_sinc}")
        print("\n✅ Sincronização concluída.\n")
        return

    relatorio = {
        "total": 0, "em_estoque": 0, "corrigidos": 0,
        "trocas": 0, "esgotados": 0, "sem_id_banco": [],
        "sem_nome": 0, "criados": 0,
        "auto_cadastrados": auto_cadastrados_sinc,
        "detalhes_trocas": [], "detalhes_correcoes": [],
        "links_invalidos_padrao": [],
    }

    _sep()

    for index, linha in df.iterrows():
        id_item    = _str(linha.get(COL_ID) or "") or str(index + 1)
        nome       = _str(linha.get(COL_NOME_PRODUTO) or "")
        link       = _str(linha.get(COL_LINK_PRINCIPAL) or "")
        url_curta  = _str(linha.get(COL_LINK_ENCURTADO) or "")
        id_banco   = _str(linha.get(COL_ID_BANCO) or "")
        plataforma = _str(linha.get(COL_PLATAFORMA) or "")
        reserva    = _str(linha.get(COL_LINK_RESERVA) or "")

        # REGRA: sem nome → pular
        if not nome:
            relatorio["sem_nome"] += 1
            continue

        # REGRA: plataforma sempre "mercadolivre"
        if plataforma.lower() != "mercadolivre":
            df.at[index, COL_PLATAFORMA] = "mercadolivre"
            plataforma = "mercadolivre"

        print(f"\n🔎 [{id_item}] {nome[:60]}")

        # ── Validar padrão do link encurtado ─────────────────────────────
        if url_curta:
            valido, motivo_val = validar_padrao_link_encurtado(url_curta)
            if not valido:
                print(f"   ⚠️  Link encurtado FORA do padrão: {motivo_val}")
                relatorio["links_invalidos_padrao"].append({
                    "id": id_item, "produto": nome,
                    "url": url_curta, "motivo": motivo_val,
                })

        # ── Determinar ação com base nas regras de colunas ───────────────
        acao = _determinar_acao_linha(nome, link, reserva, url_curta)

        if acao in ("SKIP",):
            print(f"   ⚠️  Sem links — pulando.")
            _sep()
            continue

        # ── CRIAR encurtador (sem encurtado, com link) ────────────────────
        if acao in ("CRIAR", "CRIAR_DE_RESERVA"):
            url_para_criar = link if acao == "CRIAR" else reserva.split(";")[0].strip()
            print(f"   🆕 Criando encurtador: {url_para_criar[:60]}")
            novo = criar_link_shortio(url_destino=url_para_criar, titulo=nome[:100])
            if novo:
                nova_url_curta = novo.get("shortURL", "")
                df.at[index, COL_LINK_ENCURTADO] = nova_url_curta
                url_curta = nova_url_curta
                relatorio["criados"] += 1
                print(f"   ✅ Encurtador criado: {nova_url_curta}")
                if id_banco:
                    atualizar_shortid_supabase(
                        link_uuid=id_banco, short_id=novo.get("idString",""),
                        short_url=nova_url_curta, produto_nome=nome, plataforma=plataforma,
                        motivo="Encurtador criado no monitoramento",
                    )
            _sep()
            continue

        # ── VALIDAR: verificar se encurtador e planilha estão alinhados ───
        short_url_planilha = url_curta
        slug               = _extrair_slug(url_curta)

        # Se não há link principal mas há encurtado, recuperar destino do Short.io
        if not link and slug:
            print(f"   🔗 Link principal ausente — consultando Short.io para recuperar destino...")
            dados_slug = buscar_link_por_slug(slug)
            if dados_slug:
                link = dados_slug.get("originalURL", "") or dados_slug.get("original_url", "")
                if link:
                    df.at[index, COL_LINK_PRINCIPAL] = link
                    print(f"   ✅ Link principal recuperado do Short.io: {link[:70]}")
                else:
                    print(f"   ⚠️  Short.io não retornou destino para o slug '{slug}' — validação continuará sem link principal.")
            else:
                print(f"   ⚠️  Slug '{slug}' não encontrado no Short.io — validação continuará apenas pelo encurtador.")

        print(f"   Link principal : {link[:75]}")
        print(f"   Encurtador     : {short_url_planilha[:60] or '—'}")
        print(f"   ID_Banco       : {id_banco[:8] + '...' if id_banco else '❌ NÃO ENCONTRADO'}")
        print(f"   Plataforma     : {plataforma or '—'}")

        if not id_banco:
            print(f"   ⚠️  ID_Banco ausente — atualizações Supabase não serão salvas nesta rodada.")
            relatorio["sem_id_banco"].append(f"[ID={id_item}] {nome}")

        short_id_banco = ""
        if id_banco:
            _, short_id_banco, _, _, _ = buscar_link_no_supabase(
                nome_produto=nome, short_id=slug, short_url=short_url_planilha
            )
            if not short_id_banco and slug:
                dados_api = buscar_link_por_slug(slug)
                if dados_api:
                    short_id_banco = dados_api.get("idString", "")

        relatorio["total"] += 1

        resultado = verificar_produto_completo(
            link_principal=link,
            short_url=short_url_planilha,
            nome_produto=nome,
        )

        if resultado["encurtador_ok"]:
            relatorio["em_estoque"] += 1
            df.at[index, COL_STATUS]          = STATUS_EM_ESTOQUE
            df.at[index, COL_ULT_ATUALIZACAO] = _ts_br()
            atualizar_verificacao_ok(
                link_uuid=id_banco, produto_nome=nome,
                plataforma=plataforma, short_url=short_url_planilha,
            )

        elif resultado["precisa_patch"]:
            print(f"\n   🔧 Corrigindo encurtador: PATCH Short.io → link principal")
            relatorio["em_estoque"] += 1
            relatorio["corrigidos"] += 1
            df.at[index, COL_STATUS]          = STATUS_EM_ESTOQUE
            df.at[index, COL_ULT_ATUALIZACAO] = _ts_br()

            if short_id_banco and short_id_banco.startswith("lnk_"):
                patch_ok = atualizar_destino_shortio(short_id_banco, link)
                print(f"   {'✅' if patch_ok else '❌'} PATCH Short.io {'concluído' if patch_ok else 'falhou'}")
            else:
                print(f"   ⚠️  Sem short_id válido — PATCH Short.io não executado.")

            corrigir_encurtador_supabase(
                link_uuid=id_banco, short_id=short_id_banco,
                nova_url=link, produto_nome=nome, plataforma=plataforma,
                url_antiga=resultado["url_final_enc"],
            )
            relatorio["detalhes_correcoes"].append({
                "id": id_item, "produto": nome,
                "url_encurtador_antiga": resultado["url_final_enc"],
                "link_principal": link
            })

        elif resultado["precisa_autocura"]:
            print(f"\n   🚨 Ambos falharam — iniciando auto-cura...")
            sucesso, nova_url = executar_autocura(
                df=df, index=index,
                url_falhou=link, link_uuid=id_banco,
                short_id_banco=short_id_banco, plataforma=plataforma,
            )
            if sucesso:
                relatorio["trocas"] += 1
                relatorio["detalhes_trocas"].append({
                    "id": id_item, "produto": nome,
                    "url_antiga": link, "url_nova": nova_url
                })
            else:
                relatorio["esgotados"] += 1

        _sep()

    # ── Salvar planilha ───────────────────────────────────────────────────
    try:
        print("\n💾 Salvando planilha...")
        df.to_excel(CAMINHO_DRIVE, sheet_name=NOME_ABA, index=False)
        print("✅ Salva!")
    except Exception as e:
        print(f"❌ Erro: {e}")

    # ── Relatório Final ───────────────────────────────────────────────────
    print("\n" + "═" * 70)
    print("  📊 RELATÓRIO FINAL — v6.0.0")
    print("═" * 70)
    print(f"  Delay aplicado           : {DELAY_API}s")
    print(f"  Linhas sem nome produto  : {relatorio['sem_nome']} (ignoradas)")
    print(f"  Total verificado         : {relatorio['total']}")
    print(f"  🆕 Encurtadores criados  : {relatorio['criados']}")
    print(f"  ✅ Em estoque            : {relatorio['em_estoque']}")
    print(f"  🆕 Auto-cadastrados      : {relatorio['auto_cadastrados']}")
    print(f"  🔧 Encurtadores corrigidos: {relatorio['corrigidos']}")
    print(f"  🔄 Trocas auto-cura      : {relatorio['trocas']}")
    print(f"  🚨 Esgotados             : {relatorio['esgotados']}")

    if relatorio["links_invalidos_padrao"]:
        print(f"\n  ⚠️  Links encurtados FORA do padrão ({len(relatorio['links_invalidos_padrao'])}):")
        for item in relatorio["links_invalidos_padrao"]:
            print(f"     [{item['id']}] {item['produto'][:45]}")
            print(f"          URL    : {item['url'][:70]}")
            print(f"          Motivo : {item['motivo']}")

    if relatorio["sem_id_banco"]:
        print(f"\n  ❌ Produtos SEM ID_Banco ({len(relatorio['sem_id_banco'])}):")
        for item in relatorio["sem_id_banco"]:
            print(f"     • {item}")

    if relatorio["detalhes_correcoes"]:
        print(f"\n  🔧 Encurtadores corrigidos:")
        for t in relatorio["detalhes_correcoes"]:
            print(f"    [{t['id']}] {t['produto'][:45]}")
            print(f"         Enc. apontava: {t['url_encurtador_antiga'][:70]}")
            print(f"         Corrigido p/ : {t['link_principal'][:70]}")

    if relatorio["detalhes_trocas"]:
        print(f"\n  🔄 Trocas realizadas:")
        for t in relatorio["detalhes_trocas"]:
            print(f"    [{t['id']}] {t['produto'][:45]}")
            print(f"         Antigo: {t['url_antiga'][:70]}")
            print(f"         Novo  : {t['url_nova'][:70]}")

    if nao_encontrados:
        print(f"\n  🔴 Slugs inativos no Short.io:")
        for item in nao_encontrados:
            print(f"     • {item}")

    print("═" * 70)
    print(f"  🏁 Concluído em {_ts_br()}")
    print("═" * 70 + "\n")


# ══════════════════════════════════════════════════════════════════════════
# 11. MODO CADASTRAR
# ══════════════════════════════════════════════════════════════════════════

def cadastrar_cliente(
    nome_cliente: str, responsavel: str = "", email_contato: str = "",
    ativo: bool = True, whatsapp: str = "", youtube_url: str = ""
) -> str | None:
    dados = {
        "nome_cliente":  _str(nome_cliente),
        "responsavel":   _str(responsavel),
        "email_contato": _str(email_contato),
        "ativo":         ativo,
        "whatsapp":      _str(whatsapp),
        "youtube_url":   _str(youtube_url),
    }
    print(f"\n   📝 [SUPABASE INSERT] → clientes")
    _aguardar("Supabase INSERT clientes")
    try:
        res = supabase.table("clientes").insert(dados).execute()
        uuid = res.data[0]["id"]
        print(f"   ✅ [CLIENTE] Criado! UUID: {uuid}")
        return uuid
    except Exception as e:
        print(f"   ❌ [CLIENTE] FALHA: {e}")
        return None


def cadastrar_novo_produto(
    nome_cliente: str, responsavel: str, whatsapp: str, youtube_url: str,
    nome_produto: str, link_principal: str, links_reserva: list[str] = None,
    videos_atrelados: str = "", plataforma: str = "",
    email_contato: str = "", ativo: bool = True,
):
    print("\n" + "═" * 70)
    print(f"  📋 CADASTRO — {nome_produto}")
    print("═" * 70)

    links_reserva = links_reserva or []
    todos = [link_principal] + [l for l in links_reserva if l]

    print(f"\n── PASSO 1: Teste sequencial de links ({len(todos)}) ─────────────────")
    link_aprovado, reservas_restantes = testar_lista_de_links(todos, nome_produto)
    if not link_aprovado:
        print("❌ Cadastro abortado — nenhum link passou no teste.")
        return None

    print(f"\n── PASSO 2: Cadastrar cliente ────────────────────────────────────")
    cliente_uuid = cadastrar_cliente(
        nome_cliente=nome_cliente, responsavel=responsavel,
        email_contato=email_contato, ativo=ativo,
        whatsapp=whatsapp, youtube_url=youtube_url
    )
    if not cliente_uuid:
        print("❌ Abortado — falha no cliente.")
        return None

    print(f"\n── PASSO 3: Criar encurtador (Short.io) ──────────────────────────")
    dados_shortio    = criar_link_shortio(link_aprovado, titulo=nome_produto)
    short_id_gerado  = dados_shortio.get("idString", "") if dados_shortio else ""
    short_url_gerada = dados_shortio.get("shortURL", "") if dados_shortio else ""

    print(f"\n── PASSO 4: Cadastrar no banco ───────────────────────────────────")
    link_uuid = cadastrar_link_monitoramento(
        cliente_id=cliente_uuid, nome_produto=nome_produto,
        link_principal=link_aprovado,
        link_reserva=";".join(reservas_restantes),
        short_id=short_id_gerado, short_url=short_url_gerada,
        videos_atrelados=videos_atrelados, plataforma=plataforma,
    )

    print("\n" + "═" * 70)
    print("  🎉 CADASTRO CONCLUÍDO")
    print(f"     Link UUID     : {link_uuid or '—'}")
    print(f"     Link aprovado : {link_aprovado}")
    print(f"     Short URL     : {short_url_gerada or '—'}")
    print(f"     idString      : {short_id_gerado or '—'}")
    print("═" * 70)
    return link_uuid


# ══════════════════════════════════════════════════════════════════════════
# 12. ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    iniciar_monitoramento()