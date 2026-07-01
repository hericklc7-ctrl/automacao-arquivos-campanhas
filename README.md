# Automação de Arquivos de Campanhas

Projeto em Python para automatizar o tratamento de arquivos `.txt` de campanhas, gerar planilhas padronizadas, organizar arquivos por data e, opcionalmente, enviar contatos para uma API externa.

## Por que este projeto existe

A ideia nasceu de uma necessidade operacional: reduzir tarefas manuais de copiar, colar, validar telefones, padronizar arquivos e organizar importações. O objetivo é transformar um fluxo repetitivo em um processo mais confiável, rastreável e rápido.

## O que a automação faz

- Monitora pastas de entrada de WhatsApp e SMS.
- Valida e normaliza telefones brasileiros.
- Remove duplicidades por telefone no fluxo de WhatsApp.
- Gera planilha `.xlsx` a partir de um modelo.
- Padroniza arquivos de SMS no formato `TELEFONE;NOME;IDCLIENTE`.
- Cria logs de execução.
- Controla arquivos já copiados para evitar duplicidade.
- Permite integração opcional com API externa usando variáveis de ambiente.

## Segurança e privacidade

Este repositório foi preparado para publicação pública. Foram removidos:

- caminhos reais de rede;
- nomes reais de empresas, carteiras e contratantes;
- URLs reais de APIs privadas;
- tokens, credenciais e dados sensíveis;
- arquivos reais de entrada/saída.

Use `.env.example` como base e mantenha seu `.env` fora do GitHub.

## Estrutura

```bash
.
├── src/
│   └── automacao_importacao.py
├── examples/
│   ├── entrada_whats/
│   └── entrada_sms/
├── modelo_importacao.xlsx
├── .env.example
├── .gitignore
├── requirements.txt
└── rodar_automacao.bat
```

## Como usar

1. Instale as dependências:

```bash
pip install -r requirements.txt
```

2. Copie o arquivo de exemplo de ambiente:

```bash
copy .env.example .env
```

3. Ajuste os caminhos no `.env` conforme sua máquina.

4. Rode a automação:

```bash
python src/automacao_importacao.py
```

No Windows, você também pode executar:

```bash
rodar_automacao.bat
```

## Formato dos arquivos

### WhatsApp

```txt
telefone;nome;id_cliente;valor;data_vencimento
51999999999;Maria;1001;150,00;10/07/2026
```

### SMS

```txt
TELEFONE;NOME;IDCLIENTE
51977777777;ANA;2001
```

## Integração com API

A integração vem desativada por padrão:

```env
API_ENABLED=false
API_DRY_RUN=true
```

Para usar em ambiente real, configure `API_BASE_URL` e `API_TOKEN` apenas no seu `.env` local.

## Aprendizados envolvidos

Este projeto trabalha conceitos importantes para automação e dados:

- leitura e escrita de arquivos;
- tratamento de texto;
- validação de dados;
- manipulação de Excel com `openpyxl`;
- variáveis de ambiente;
- logs e auditoria;
- integração HTTP com API;
- organização de projeto para GitHub.

