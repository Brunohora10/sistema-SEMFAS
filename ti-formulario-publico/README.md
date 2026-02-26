# Formulario Publico TI (Apps Script separado)

Este projeto e separado do sistema principal e gera um segundo link (novo Web App) para abrir chamados e conversar no chat.

## 1) Configurar conexao com o sistema principal

Edite `Code.gs`:

- `TARGET_BASE_URL`: URL do Web App principal do sistema SEMFAS
- `PUBLIC_TOKEN`: mesmo token configurado em `TI_PUBLIC_TOKEN` no projeto principal (se voce ativou token)

## 2) Criar repositorio separado

Copie esta pasta `ti-formulario-publico` para um novo repositorio.

Arquivos minimos:

- `Code.gs`
- `index.html`
- `appsscript.json`

## 3) Subir no Apps Script com clasp

No novo repositorio:

1. `clasp login`
2. `clasp create --type webapp --title "SEMFAS - Formulario Publico TI"`
3. `clasp push`
4. `clasp deploy --description "v1 formulario publico ti"`

Depois, no Apps Script, configure o deploy como:

- Executar como: voce (dono do script)
- Quem tem acesso: Qualquer pessoa com o link

## 4) Resultado esperado

- Link 1: sistema principal (ja existente)
- Link 2: formulario/chat publico (este projeto)

Ambos gravam no mesmo backend de chamados via API.