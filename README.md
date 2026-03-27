# Análise de Pesquisa de Avaliação Institucional

## 1. Descrição
Sistema em Python para análise de pesquisas de avaliação institucional de Institutos Federais. Processa dados de Discentes, Docentes e TAEs, incluindo dados demográficos, e gera estatísticas, gráficos e relatórios completos (PDF, Excel e CSV) formatados automaticamente.

## 2. Funcionalidades
* Processamento automático de planilhas Excel da pesquisa.
* Análise de representatividade com dados demográficos da instituição.
* Cálculo de estatísticas e médias por perfil e por eixo do SINAES.
* Geração de gráficos, infográficos e relatório PDF consolidado.

## 3. Como Usar

### Google Colab (Recomendado)
A forma mais ágil para não precisar instalar nada.
1. Acesse o Colab e crie um novo notebook ou abra o seu existente.
2. Faca o upload via aba de arquivos ou importe as planilhas da pesquisa e de demografia.
3. Cole todo o código do `analise_pesquisa_2025.py` e execute a célula.
4. Baixe os resultados emitidos dentro da pasta subjacente `resultados_analise/`.

### Python Local
1. Instale os requisitos fundamentais em seu ambiente:
```bash
pip install -r requirements.txt
```
2. Edite os caminhos das planilhas no bloco principal final do script `analise_pesquisa_2025.py`:
```python
ARQUIVO_PESQUISA    = 'dados/Gyn.xlsx'
ARQUIVO_DEMOGRAFICO = 'dados/Numero_Matriculas_alunos-Gyn_Servidores_2025.xlsx'
```
3. Execute o script no seu terminal ou editor favorito:
```bash
python analise_pesquisa_2025.py
```
4. Verifique os relatórios originados na pasta `resultados_analise/`.

## 4. Estrutura dos Dados Cativados

### Arquivo de Pesquisa (Obrigatório)
Planilha com três abas exatas: DISCENTES, DOCENTES e TAES.
* Coluna 1: Vínculo com o IF
* Coluna 2: Campus
* Colunas 3 em diante: Questões dispostas na pesquisa.
As respostas obrigatoriamente aceitas são: Ótimo, Bom, Regular, Ruim, Péssimo, Não sei.

### Arquivo Demográfico (Opcional)
Planilha extra com total de matrículas de alunos em cursos, e totais de servidores ativos para o devido cálculo da taxa de resposta ser verídico.

## 5. Resultados Gerados na Saída (`resultados_analise/`)
* **Relatório em PDF**: Dossiê completo unificado unindo todos os gráficos (`relatorio_graficos_*.pdf`).
* **Imagens soltas (PNG)**: Gráficos demográficos de participação, infográfico executivo condensado e resultados individuais por cada eixo e dimensão.
* **Tabelas de Dados**: Planilha Excel máster completa (`relatorio_completo.xlsx`) contendo as médias tratadas, e arquivos brutos CSV para outros softwares.

## 6. Personalização do Motor Histórico
* **Eixos SINAES**: Edite os vetores strings no dicionário `estrutura_eixos` logo no início da classe se o instrumento de teste da CPA diferir perguntas da base de 2025/2026.
* **Cores Oficiais**: Edite as constantes literais hexadecimais (ex: `COR_OTIMO`, `COR_RUIM` ou `COR_TITULO`) declaradas no topo do arquivo.


