# Analisador Contábil

**Conferência automática de saldos no razão SCI.**

O Analisador Contábil é uma aplicação em Streamlit criada para apoiar a revisão de razões contábeis exportados do SCI. O app cruza o razão diário com o plano de contas, identifica saldos que precisam de verificação e apresenta um resumo limpo para facilitar a conferência.

## Principais recursos

- Leitura de plano de contas em CSV, XLSX ou XLS.
- Leitura de razão diário exportado do SCI em CSV, XLSX ou XLS.
- Identificação de contas e participantes dentro do razão.
- Conversão automática de valores no padrão brasileiro.
- Análise de saldos por conta, participante e data.
- Resumo de ocorrências repetidas para evitar poluição visual.
- Dashboard com indicadores principais.
- Tabela com apenas os casos que precisam de revisão.
- Exportação do resultado em Excel.

## Objetivo

O app foi pensado para reduzir o trabalho manual de abrir conta por conta no sistema. Em vez disso, a análise é feita de uma vez sobre o arquivo completo do razão, destacando apenas os pontos que merecem conferência.

## Arquivos de entrada

O sistema utiliza dois arquivos:

```text
Plano de contas CSV
Razão SCI CSV
```

Os arquivos podem estar em CSV separado por ponto e vírgula, XLSX ou XLS.

## Hospedagem no Streamlit Cloud

Para publicar o app no Streamlit Cloud:

1. Envie os arquivos do projeto para um repositório no GitHub.
2. Acesse o Streamlit Cloud.
3. Crie um novo app usando esse repositório.
4. Em **Main file path**, informe:

```text
streamlit_app.py
```

5. O Streamlit instalará automaticamente as dependências do arquivo `requirements.txt`.

## Arquivos do projeto

```text
streamlit_app.py              Aplicação principal para Streamlit Cloud
core.py                       Motor de leitura, análise e exportação
requirements.txt              Dependências do projeto
logo_analisador_contabil.svg  Logo do aplicativo
app.py                        Versão local para testes
```

## Segurança dos dados

Arquivos reais de empresas não devem ser enviados ao GitHub.

O projeto já inclui um `.gitignore` para evitar o envio de arquivos como:

```text
*.csv
*.xlsx
*.xls
*.txt
*.log
```

Assim, planos de contas, razões contábeis e relatórios exportados ficam somente no ambiente de uso.

## Execução local

Para testar a versão local:

```text
abrir_analisador.bat
```

Para testar a versão Streamlit localmente:

```text
abrir_streamlit.bat
```

## Autoria

Desenvolvido por **Rayssa Mayara**.

GitHub: [rayssamayarax](https://github.com/rayssamayarax)
