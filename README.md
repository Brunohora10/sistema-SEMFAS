# Sistema SEMFAS

Sistema web em Google Apps Script para gestao de beneficios e RMA.

## Visao geral

- Hub de acesso com modulos de Beneficio Eventual e RMA Mensal
- Formularios RMA por unidade, com salvamento em planilha
- Geracao de PDF ao salvar o RMA

## Requisitos

- Conta Google com acesso ao Apps Script e Drive
- Planilha configurada (banco de dados)
- Node.js + clasp (para publicar o projeto)

## Configuracao rapida

1. Abra o projeto no Apps Script e autorize as permissoes.
2. Edite as constantes em code.js:
   - PLANILHA_ID
   - RMA_TEMPLATES_FOLDER_ID
   - RMA_OUTPUT_FOLDER_ID
3. Salve e publique o web app.

## Rotas (views)

- ?view=login
- ?view=hub
- ?view=hub_beneficio_eventual
- ?view=rma
- ?view=central
- ?view=admin
- ?view=ti

## Como usar o RMA

1. Abra o hub e selecione RMA (Mensal).
2. Clique na unidade desejada.
3. Preencha o formulario e clique em Salvar.
4. O sistema grava na planilha e gera o PDF na pasta configurada.

## Desenvolvimento

Publicar com clasp:

- clasp push --force

## Observacoes

- O PDF e gerado via Google Docs temporario e salvo no Drive.
- A estrutura de pastas e criada automaticamente em RMA/AAAA-MM/UNIDADE.
