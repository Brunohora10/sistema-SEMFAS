# Sistema SEMFAS

Sistema web completo em **Google Apps Script** para gestão de benefícios sociais, RMA e chamados de TI da SEMFAS.

## Visão geral

| Módulo | Descrição |
|--------|-----------|
| **Hub** | Painel central com acesso a todos os módulos |
| **Benefício Eventual** | Cadastro e gestão de benefícios eventuais |
| **RMA Mensal** | Formulários RMA por unidade, com geração automática de PDF |
| **Central** | Painel de controle e relatórios |
| **TI — Chamados** | Sistema de chamados técnicos com chat em tempo real e anexos |
| **TI — Formulário Público** | Portal externo para abertura de chamados (subprojeto separado) |
| **Admin** | Gerenciamento de usuários e configurações |
| **Vigilância** | Módulo de vigilância sanitária |

## Estrutura do projeto

```
sistema-SEMFAS/
├── code.js                  # Backend principal (API + lógica de negócio)
├── index.html               # Página inicial
├── login.html               # Tela de login
├── hub.html                 # Hub de acesso aos módulos
├── hub_beneficio.html       # Hub de benefícios
├── hub_beneficio_eventual.html
├── hub_beneficio_rma.html
├── formulario.html          # Formulário de benefício
├── baixa.html               # Tela de baixa de benefícios
├── rma.html                 # Formulário RMA
├── rma-templates/           # Templates para geração de PDF
├── central.html             # Central de controle
├── admin.html               # Painel administrativo
├── analista.html            # Painel do analista
├── ti.html                  # Painel de chamados TI (técnicos)
├── vigilancia.html          # Módulo de vigilância
├── ti-formulario-publico/   # Subprojeto — formulário público de TI
│   ├── Code.gs              #   Backend do formulário público
│   ├── index.html           #   Frontend do formulário público
│   └── appsscript.json      #   Manifest do subprojeto
└── appsscript.json          # Manifest do projeto principal
```

## Requisitos

- Conta Google com acesso ao Apps Script e Drive
- Planilha Google configurada como banco de dados
- Node.js + [clasp](https://github.com/nicromancero/clasp) (para publicar)

## Configuração rápida

1. Abra o projeto no Apps Script e autorize as permissões.
2. Edite as constantes em `code.js`:
   - `PLANILHA_ID` — ID da planilha banco de dados
   - `FOLDER_PDFS_ID` — pasta do Drive para mídia TI
   - `RMA_TEMPLATES_FOLDER_ID` — pasta de templates RMA
   - `RMA_OUTPUT_FOLDER_ID` — pasta de saída dos PDFs RMA
3. Para o formulário público (`ti-formulario-publico/Code.gs`):
   - `TARGET_BASE_URL` — URL do deploy do projeto principal
   - `PLANILHA_ID` — mesma planilha do projeto principal
   - `FOLDER_PDFS_ID` — mesma pasta de mídia TI
4. Publique ambos os projetos como Web App.

## Módulo TI — Chamados

### Funcionalidades

- Abertura de chamados pelo **portal público** (sem acesso à planilha)
- Upload de **fotos e vídeos** do problema (até 50 MB via form submission)
- **Chat em tempo real** entre solicitante e técnico
- Anexos organizados em subpastas por protocolo no Google Drive
- Visualização de anexos diretamente no painel do técnico (preview + tela cheia)
- Histórico completo de ações por chamado

### Fluxo de anexos

1. Solicitante seleciona foto/vídeo no formulário público
2. Arquivo é enviado como **Blob nativo** via `google.script.run` (form submission)
3. Backend cria subpasta `TI-AAAA-NNNNNN/` no Drive e salva o arquivo
4. ID da pasta é gravado na coluna `Anexo (Pasta ID)` da planilha
5. Painel TI busca arquivos direto pela pasta, sem varredura

### API pública (formulário externo)

| Rota | Descrição |
|------|-----------|
| `?api=ti_public_meta` | Metadados (prioridades, categorias, setores) |
| `?api=ti_public_create` | Abre chamado público (retorna protocolo) |
| `?api=ti_public_chat_send` | Envia mensagem no chat |
| `?api=ti_public_chat_list` | Lista mensagens por protocolo |

### API interna (painel TI)

| Rota | Descrição |
|------|-----------|
| `?api=ti_chat_send` | Envia mensagem (técnico) |
| `?api=ti_chat_list` | Lista mensagens |
| `?api=ti_drive_anexos` | Lista arquivos do Drive por protocolo |

### Token opcional

Para proteger as rotas públicas, configure em **Script Properties**:

- Chave: `TI_PUBLIC_TOKEN`
- Valor: token secreto definido pelo TI

Se configurado, as rotas `ti_public_*` exigem `token` no payload.

### Persistência (abas da planilha)

| Aba | Conteúdo |
|-----|----------|
| `Chamados` | Dados dos chamados (protocolo, status, anexos, etc.) |
| `Chamados_Historico` | Histórico de ações por chamado |
| `Chamados_Chat` | Mensagens do chat em tempo real |

## Rotas (views)

| View | Página |
|------|--------|
| `?view=login` | Login |
| `?view=hub` | Hub principal |
| `?view=hub_beneficio_eventual` | Hub benefício eventual |
| `?view=rma` | Formulário RMA |
| `?view=central` | Central de controle |
| `?view=admin` | Painel administrativo |
| `?view=ti` | Painel de chamados TI |

## Como usar o RMA

1. Abra o hub e selecione **RMA (Mensal)**.
2. Clique na unidade desejada.
3. Preencha o formulário e clique em **Salvar**.
4. O sistema grava na planilha e gera o PDF na pasta configurada.

## Desenvolvimento

### Publicar projeto principal

```bash
cd sistema-SEMFAS
clasp push --force
```

### Publicar formulário público TI

```bash
cd ti-formulario-publico
clasp push --force
```

> Cada subprojeto tem seu próprio `.clasp.json` e deve ser publicado separadamente.

## Observações

- PDFs são gerados via Google Docs temporário e salvos no Drive.
- A estrutura de pastas RMA é criada automaticamente em `RMA/AAAA-MM/UNIDADE`.
- Anexos TI ficam em `Mídia TI/<PROTOCOLO>/` no Drive.
- O formulário público usa fallback local quando a API principal está indisponível.
