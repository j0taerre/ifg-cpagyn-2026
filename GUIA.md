# Guia Rápido de Uso

Este projeto automatiza a contagem dos dados das planilhas respostas da CPA, e exporta em visualizações gráficas avançadas num único pacote útil.

## 1. Preparando suas Planilhas

Você precisa de até dois arquivos do Excel em disco:
* **Planilha de Pesquisa (Obrigatória)**: Deve ter 3 abas essenciais prontas (DISCENTES, DOCENTES, TAES). As duas primeiras colunas de cada aba representam o Vínculo e o Campus do eleitorado. As colunas seguintes são as respostas fechadas das escalas (Ótimo, Bom, Regular, Ruim, Péssimo, Não sei).
* **Planilha Demográfica (Opcional)**: Tabela contendo num somatório os totais corretos de alunos matriculados e servidores. Serve somente para calcular a fatia da sua amostra (Representatividade percentual real).

## 2. Rodando no Google Colab (Para qualquer Computador)

Opção veloz pois o computador do Google processa e já tem o ambiente empacotador.

1. Navegue e Acesse: https://colab.research.google.com
2. Abra um Novo Notebook limpo.
3. Copie o nosso código inteiro de dentro de `analise_pesquisa_2025.py` e cole em um quadro (célula).
4. Rode a Célula apertando Play.
5. Ele poderá interromper momentaneamente para lhe solicitar (via upload) seus os arquivos gerados no passo 1.
6. Uma pasta chamada `resultados_analise/` aparecerá no canto esquerdo. Todos os arquivos e o PDF estarão agrupados e unificados para baixar.

## 3. Rodando no Próprio Computador Local (Python)

Destinado aos avaliadores que possuem desenvolvimento pré estabelecido ou já mantêm o interpretador Python na máquina com as bibliotecas Pandas e afins.

1. Instale as extensões vitais para o roteador de estatística funcionar:
```bash
pip install -r requirements.txt
```
2. Abra o arquivo matriz `analise_pesquisa_2025.py` num leitor de texto puro. Vá literalmente às últimas linhas finais no bloco `if __name__ == "__main__":`. Modifique para as referências internas aos seus próprios arquivos locais os caminhos e variáveis `ARQUIVO_PESQUISA` e `ARQUIVO_DEMOGRAFICO`.
3. Rode seu comando preferido para o gatilho da linguagem:
```bash
python analise_pesquisa_2025.py
```
4. A tela piscará informando as tabelas achadas em suas abas. Uma vez concluído, as saídas serão alocadas silenciosamente na mesma subpasta denominada `resultados_analise/`.

## 4. O Que Devo Aguardar de Resultado?

* **O PDF Oficial**: Um documento multi folha (`relatorio_graficos_*.pdf`) paginado e pronto indicando os acertos. Pode ser remetido anexamente.
* **Imagens PNG Individuais**: Os retângulos e diagramas que seriam formados dentro do PDF ficam expostos ali avulsamente. Estudo de público em gráficos de radar, e barrinhas.
* **Acervos de Tabelas**: Documento gigante em Excel computando a matemática bruta inteira das avaliações e subdivisões que geraram as parcelas das porcentagens. Assim como as tabelas CSV isoladas de Eixo a Eixo prontas para planilhas da coordenação institucional.

