# Automação - Download das bases do Cognos (IBM Planning Analytics)

Automatiza a etapa de **atualização das bases** do relatório Power BI
"Orçamento e Forecast Gerencial": baixa as 6 exportações do Planning Analytics
e salva cada arquivo na pasta de rede correspondente
(`\\CBL1CM68919\Planejamento\...`), fazendo backup local do arquivo anterior.

## Exportações configuradas (config.json)

| Nome | Servidor | View pesquisada | Pasta destino |
|---|---|---|---|
| Receitas | Sdata ACOMPANHAMENTO | Receita DRE PowerBI V2 (irat950) | ...\IRAT.950 |
| Físicos | Sdata ACOMPANHAMENTO | Fisico Receita - FIS 900 (Power BI) | ...\FIS.900 |
| Custos | Sdata ACOMPANHAMENTO | Receita DRE PowerBI V2 (irat950) CUSTO | ...\IRAT.950_Custo |
| Abertura Receita (Waterfall) | NET_PLAN_UF | REV.900.Receita_Consolidada_DRE Power BI | ...\REC.900 - Receita Abertura (PROVISORIO) |
| Pré-Pago Parte 1 | NET_PLAN_UF | CTS.100 (Power BI) | ...\CTS.100 |
| Pré-Pago Parte 2 | NET_PLAN_UF | REV.420 Power BI | ...\REV.420 |

## Pré-requisitos

- Conectado à **rede corporativa / VPN** (para acessar o Cognos e a pasta `\\CBL1CM68919`).
- Python 3.12 (já instalado em `%LOCALAPPDATA%\Programs\Python\Python312`).
- Microsoft Edge instalado (o Selenium baixa o driver automaticamente).
- Dependências: `python -m pip install -r requirements.txt`

## Como executar

Duplo clique (recomendado):

```text
rodar_exportacoes.bat
```

Abre um **painel** em outra janela com status de cada exportação, tempo gasto e ETA.

Ou via PowerShell:

```powershell
cd "C:\Users\n5919189\OneDrive - Claro SA\FINANCEIRO\Github\att_cognos_pbi"
.\rodar_exportacoes.bat
```

Opções úteis direto no Python:

```powershell
# Executa só uma exportação (filtra pelo nome)
python baixar_cognos.py --somente "FIS.900"

# Baixa os arquivos mas NÃO copia para a rede (bom para testar)
python baixar_cognos.py --sem-mover
```

## Pontos de atenção (do documento original)

- O **nome do arquivo** e o **nome da aba** do Excel devem ser idênticos aos dos
  arquivos anteriores — o Power BI busca por eles. Se o nome baixado do Cognos
  for diferente do esperado, preencha o campo `nome_arquivo_destino` no
  `config.json` com o nome exato do arquivo anterior.
- Antes de sobrescrever, o script salva uma cópia do arquivo anterior em
  `backup\<nome da exportação>\`.

## Ajustes pendentes (primeira execução)

Este script foi construído **sem acesso à rede corporativa**, então os
seletores da interface do Planning Analytics (dicionário `SELETORES` no início
de `baixar_cognos.py`) são baseados no padrão do produto e **podem precisar de
ajuste na primeira execução** — em especial os passos de pesquisa e do menu
"Exportar". Se algum passo falhar, o log indica qual elemento não foi
encontrado; basta rodar de novo com a VPN conectada junto do assistente para
corrigir os seletores.
