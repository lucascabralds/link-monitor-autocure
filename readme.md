# 🔗 LinkMonitory — Sistema de Monitoramento, Cadastro e Auto-Cura de Links

> Robô de monitoramento automatizado para links de produtos no Mercado Livre, com encurtamento via Short.io, sincronização em banco de dados Supabase e mecanismo de auto-cura com links de reserva.

---

## 📌 Índice

- [Sobre o Projeto](#sobre-o-projeto)
- [O Problema](#o-problema)
- [A Solução](#a-solução)
- [Funcionalidades](#funcionalidades)
- [Arquitetura e Tecnologias](#arquitetura-e-tecnologias)
- [Estrutura do Banco de Dados](#estrutura-do-banco-de-dados)
- [Estrutura da Planilha](#estrutura-da-planilha)
- [Fluxo de Funcionamento](#fluxo-de-funcionamento)
- [Casos de Uso (A → E)](#casos-de-uso-a--e)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Como Usar](#como-usar)
- [Modos de Execução](#modos-de-execução)
- [Relatório Final](#relatório-final)
- [Autor](#autor)

---

## Sobre o Projeto

O **LinkMonitory** é um sistema Python automatizado desenvolvido para lojistas do Mercado Livre que utilizam links encurtados (via Short.io) em campanhas e divulgações. O robô monitora continuamente uma planilha Excel de controle, verifica se os links dos produtos estão ativos, cria encurtadores quando necessário e, caso um produto fique fora de estoque, realiza a **auto-cura automática** trocando o link pelo próximo da fila de reservas — sem nenhuma intervenção manual.

**Versão atual:** `7.0.0`

---

## O Problema

Lojistas que vendem no Mercado Livre e divulgam seus produtos em redes sociais, WhatsApp e YouTube enfrentam um problema recorrente: **produtos saem do ar**. Seja por esgotamento de estoque, remoção do anúncio ou mudança de URL, os links compartilhados deixam de funcionar, gerando perda de vendas e experiência negativa para o cliente.

Além disso, gerenciar dezenas ou centenas de links manualmente — criando encurtadores, monitorando o status e atualizando as campanhas — é um processo lento, sujeito a erros e que demanda tempo valioso do vendedor.

**Os principais pontos de dor:**

- Links de produtos expiram ou mudam sem aviso.
- Encurtadores desatualizados apontam para produtos fora de estoque.
- Sem automação, o vendedor precisa revisar cada link individualmente.
- O histórico de alterações não é registrado, dificultando auditorias.

---

## A Solução

O **LinkMonitory** resolve esses problemas com uma abordagem de três camadas:

1. **Monitoramento automatizado** — verifica cada link via Playwright (browser headless), simulando um usuário real e checando a presença do botão *"Ir para produto"* do gateway social do Mercado Livre.
2. **Auto-cura inteligente** — quando um link falha, o sistema testa automaticamente os links de reserva cadastrados e substitui o link principal pelo primeiro que passar no teste, atualizando o encurtador e o banco de dados em tempo real.
3. **Sincronização centralizada** — todas as alterações são refletidas simultaneamente na planilha Excel, no Short.io e no banco de dados Supabase, garantindo consistência em todas as fontes de verdade.

---

## Funcionalidades

- ✅ **Verificação dupla de links** — testa encurtador e link principal separadamente (Passos A e B)
- 🔧 **Correção automática de encurtador** — se o link principal funciona mas o encurtador aponta para URL desatualizada, realiza PATCH via API do Short.io
- 🔄 **Auto-cura com fila de reservas** — substitui links quebrados automaticamente, consumindo a fila de reservas em ordem de prioridade
- 🆕 **Criação automática de encurtadores** — gera links curtos via Short.io para produtos que ainda não possuem encurtador
- 🗄️ **Auto-cadastro no Supabase** — produtos novos detectados na planilha são automaticamente registrados no banco de dados
- 📋 **Histórico completo de alterações** — toda modificação é registrada na tabela `historico_ajustes` com tipo de operação, URL antiga, URL nova e timestamp
- 🔒 **Preservação de IDs** — `ID_Banco` nunca é sobrescrito; IDs sequenciais são gerados apenas para linhas sem identificação
- 🧹 **Validação de padrão de links** — verifica se os encurtadores seguem o domínio configurado 
- 📊 **Relatório detalhado** — ao final de cada execução, exibe resumo com contagem de verificados, corrigidos, trocas e esgotados
- ⏱️ **Rate limiting respeitoso** — delay configurável entre chamadas de API (padrão: 5 segundos) para evitar bloqueios

---

## Arquitetura e Tecnologias

```
┌─────────────────────────────────────────────────────────┐
│                     LinkMonitory                        │
│                                                         │
│  ┌──────────────┐   ┌──────────────┐   ┌────────────┐  │
│  │  Excel/Drive │   │   Short.io   │   │  Supabase  │  │
│  │  (Planilha)  │◄──►│ (Encurtador) │◄──►│  (Banco)   │  │
│  └──────┬───────┘   └──────────────┘   └────────────┘  │
│         │                                               │
│  ┌──────▼───────────────────────────────────────────┐   │
│  │          Engine de Monitoramento (Python)         │   │
│  │   Playwright ► Verificação ► Auto-cura ► Sync     │   │
│  └──────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────┘
```

| Tecnologia | Uso |
|---|---|
| **Python 3.10+** | Linguagem principal |
| **Playwright** | Verificação de links via browser headless (Chromium) |
| **Supabase** | Banco de dados PostgreSQL gerenciado (persistência e histórico) |
| **Short.io** | Criação e gerenciamento de links encurtados |
| **Pandas + OpenPyXL** | Leitura e escrita da planilha Excel de controle |
| **python-dotenv** | Gerenciamento de variáveis de ambiente |
| **Requests** | Chamadas à API REST do Short.io |

---

## Estrutura do Banco de Dados

### `clientes`
| Coluna | Tipo | Descrição |
|---|---|---|
| `id` | UUID | Identificador único |
| `created_at` | timestamp | Data de criação |
| `nome_cliente` | text | Nome do cliente/loja |
| `responsavel` | text | Nome do responsável |
| `email_contato` | text | E-mail |
| `ativo` | boolean | Status ativo/inativo |
| `whatsapp` | text | WhatsApp de contato |
| `youtube_url` | text | Canal do YouTube |

### `links_monitoramento`
| Coluna | Tipo | Descrição |
|---|---|---|
| `id` | UUID | Identificador único |
| `created_at` | timestamp | Data de criação |
| `cliente_id` | UUID | FK → clientes |
| `nome_produto` | text | Nome do produto |
| `link_principal` | text | URL do produto ativo |
| `link_reserva` | text | Fila de reservas separadas por `;` |
| `short_id` | text | idString do Short.io (`lnk_xxx`) |
| `short_url` | text | URL encurtada completa |
| `videos_atrelados` | text | URLs de vídeos YouTube |
| `status_estoque` | text | `em_estoque`, `esgotado`, `sem_reserva`, `link_quebrado` |
| `ultima_verificacao` | timestamp | Último check realizado |
| `plataforma` | text | Sempre `mercadolivre` |

### `historico_ajustes`
| Coluna | Tipo | Descrição |
|---|---|---|
| `id` | UUID | Identificador único |
| `link_id` | UUID | FK → links_monitoramento |
| `data_alteracao` | timestamp | Data da alteração |
| `url_antiga` | text | URL antes da alteração |
| `url_nova` | text | URL após a alteração |
| `motivo` | text | Descrição do motivo |
| `tipo_operacao` | text | `CRIACAO`, `AUTO_CURA`, `ESGOTAMENTO`, `SINCRONIZACAO`, `VERIFICACAO`, `CORRECAO_ENCURTADOR` |
| `produto_nome` | text | Nome do produto |
| `plataforma` | text | Plataforma do produto |

---

## Estrutura da Planilha

A planilha `Base_Monitoramento_Links_v1.xlsx` (aba `Monitoramento`) é a fonte de controle central:

| Coluna | Descrição |
|---|---|
| `ID` | Identificador sequencial (auto-gerado se vazio) |
| `ID_Banco` | UUID do Supabase (nunca sobrescrito pelo robô) |
| `Nome do Produto` | Nome do produto (obrigatório para processar a linha) |
| `Plataforma` | Sempre `mercadolivre` (preenchido automaticamente) |
| `Link Principal (Link 1)` | URL ativa do produto |
| `Links Reservas (Link 2)` | URLs de backup separadas por `;` |
| `Vídeos Atrelados (YouTube)` | Links de vídeos, separados por `;` |
| `Link Encurtado` | URL do Short.io (ex: `https://reidascapotas.s.gy/xxxxx`) |
| `Status` | Status atual do produto |
| `Ultima Atualização` | Data/hora da última verificação |

---

## Fluxo de Funcionamento

```
INÍCIO
  │
  ├─► SINC — Garante IDs, plataforma e cria encurtadores ausentes
  │
  ├─► Para cada produto com nome preenchido:
  │     │
  │     ├─► PASSO A — Testa encurtador via Playwright
  │     │     ├─► ✅ OK → atualiza timestamp, status = em_estoque
  │     │     └─► ❌ Falhou → vai para Passo B
  │     │
  │     ├─► PASSO B — Testa link principal via Playwright
  │     │     ├─► ✅ OK → PASSO C (corrigir encurtador)
  │     │     └─► ❌ Falhou → PASSO D (auto-cura)
  │     │
  │     ├─► PASSO C — PATCH Short.io + atualiza Supabase
  │     │
  │     └─► PASSO D — Testa fila de reservas → substitui principal + PATCH Short.io
  │
  └─► Salva planilha + exibe relatório final
```

---

## Casos de Uso (A → E)

| Caso | Principal | Reserva | Encurtado | Ação |
|---|---|---|---|---|
| **A** | ❌ | ✅ | ❌ | `CRIAR` encurtador a partir da reserva, promovendo-a para principal |
| **B** | ✅ | qualquer | ❌ | `CRIAR` encurtador para o link principal |
| **C** | ✅ | ❌ | ✅ | `VALIDAR` se encurtador e principal estão sincronizados |
| **D** | ✅ | ✅ | ✅ | `VALIDAR` encurtador + testar reservas se necessário |
| **E** *(novo v7)* | ❌ | qualquer | ✅ | `VALIDAR` — Short.io é consultado como fonte de verdade para recuperar o link principal |

---

## Instalação

### Pré-requisitos

- Python 3.10 ou superior
- Conta no [Supabase](https://supabase.com)
- Conta no [Short.io](https://short.io)
- Planilha Excel no caminho configurado (Google Drive ou local)

### Passo a passo

```bash
# 1. Clone o repositório
git clone https://github.com/seu-usuario/LinkMonitory.git
cd LinkMonitory

# 2. Crie e ative um ambiente virtual (recomendado)
python -m venv venv
source venv/bin/activate      # Linux/macOS
venv\Scripts\activate         # Windows

# 3. Instale as dependências
pip install python-dotenv supabase playwright pandas openpyxl requests

# 4. Instale o browser Chromium para o Playwright
playwright install chromium
```

---

## Configuração

Crie um arquivo `.env` na raiz do projeto:

```env
# Supabase
API_URL_SUPABASE=https://xxxxxxxxxxx.supabase.co
SERVICE_ROLE_SUPABASE_API_KEY=seu_service_role_key_aqui
ID_CLIENTE_SUPABASE=uuid-do-cliente-no-banco

# Short.io
SHORT_IO_API_KEY=sua_chave_api_shortio
SHORT_IO_DOMAIN=seudominio.s.gy

# Planilha
CAMINHO_DRIVE=C:\Users\SeuUsuario\Google Drive\Cliente\Base_Monitoramento_Links_v1.xlsx
NOME_ABA=Monitoramento

# Modo de execução: completo | sincronizar
MODO_EXECUCAO=completo
```

> ⚠️ **Importante:** Nunca versione seu arquivo `.env`. Adicione-o ao `.gitignore`.

---

## Como Usar

```bash
# Executar o monitoramento completo
python TESTE_V7.py

# Definindo o modo via variável de ambiente
MODO_EXECUCAO=sincronizar python TESTE_V7.py
```

---

## Modos de Execução

| Modo | Descrição |
|---|---|
| `completo` | Sincroniza identidades + verifica todos os links + auto-cura + salva planilha |
| `sincronizar` | Apenas sincroniza IDs, plataforma e cria encurtadores ausentes, sem verificação Playwright |

---

## Relatório Final

Ao final de cada execução no modo `completo`, o sistema exibe um relatório como:

```
══════════════════════════════════════════════════════════════════════
  📊 RELATÓRIO FINAL — v7.0.0
══════════════════════════════════════════════════════════════════════
  Delay aplicado            : 5.0s
  Linhas sem nome produto   : 2 (ignoradas)
  Total verificado          : 47
  🆕 Encurtadores criados   : 0
  ✅ Em estoque             : 45
  🆕 Auto-cadastrados       : 0
  🔧 Encurtadores corrigidos : 1
  🔄 Trocas auto-cura       : 1
  🚨 Esgotados              : 0
══════════════════════════════════════════════════════════════════════
```

---

## Autor

Desenvolvido por **[Lucas Cabral]**

- 🐙 GitHub: [github.com/lucascabralds](https://github.com/lucascabralds)
- 💼 LinkedIn: [linkedin.com/in/lucascasantos](https://www.linkedin.com/in/lucascasantos/)

---

> Projeto desenvolvido para automação de operações de e-commerce no Mercado Livre, integrando Short.io e Supabase para monitoramento contínuo e recuperação automática de links de produtos.