# 🚀 GUIA RÁPIDO DE USO

## Para Usar no Google Colab 

### Passo 1: Abrir o Notebook
1. Acesse https://colab.research.google.com
2. Clique em "Arquivo" → "Fazer upload de notebook"
3. Selecione o arquivo `Analise_Pesquisa_IF_Colab.ipynb`

### Passo 2: Executar
1. Execute todas as células em ordem (Ctrl + F9)
2. Quando solicitado, faça upload dos arquivos Excel com as respostas da pesquisa, bem como do script python de análise (analise_pesquisa_2025.py)
3. Aguarde a geração dos gráficos e relatórios
4. Baixe os arquivos gerados na última célula

---

## Para Usar Localmente (Python)

### Passo 1: Instalar Dependências
```bash
pip install -r requirements.txt
```

### Passo 2: Ajustar Caminho do Arquivo
Edite a linha 657 do arquivo `analise_pesquisa_if_completo.py`:
```python
CAMINHO_ARQUIVO = 'seu_arquivo.xlsx'  # Coloque o caminho correto
```

### Passo 3: Executar
```bash
python analise_pesquisa_if_completo.py
```

### Passo 4: Ver Resultados
Os arquivos serão gerados no mesmo diretório:
- Gráficos (*.png)
- Relatório Excel (relatorio_completo.xlsx)
- Tabelas CSV (tabela_*.csv)

---

## Estrutura do Arquivo Excel Necessária

Seu arquivo Excel deve ter 3 abas:
- **DISCENTES**: Respostas dos discentes
- **DOCENTES**: Respostas dos docentes
- **TAES**: Respostas dos TAEs

Cada aba deve ter:
- Coluna 1: Vínculo com o IF
- Coluna 2: Campus
- Demais colunas: Questões da pesquisa

As respostas devem usar a escala: Ótimo, Bom, Regular, Ruim, Péssimo, Não sei

---

## O Que Será Gerado

### 📊 4 Gráficos PNG
- Gráfico do Eixo 1 - Dimensão 8 (Planejamento e Avaliação)
- Gráfico do Eixo 2 - Dimensão 1 (Missão e PDI)
- Gráfico do Eixo 2 - Dimensão 3 (Responsabilidade Social)
- Infográfico Executivo (resumo completo)

### 📑 4 Arquivos de Dados
- Relatório Excel completo (todas as análises em abas)
- 3 Tabelas CSV (uma para cada Eixo-Dimensão)

---

## Solução Rápida de Problemas

### "No such file or directory"
→ Verifique o caminho do arquivo Excel

### "ModuleNotFoundError"
→ Execute: `pip install -r requirements.txt`

### Gráficos não aparecem
→ No Colab funciona automaticamente
→ Localmente, os arquivos são salvos como PNG

### Caracteres estranhos nos gráficos
→ Certifique-se que o Excel está em UTF-8

---

## Personalização Rápida

### Mudar Cores
Edite o dicionário `cores_escala` (linha 64)

### Adicionar Eixo-Dimensão
Edite o dicionário `estrutura_eixos` (linha 72)

### Ajustar Tamanho dos Gráficos
Altere `plt.rcParams['figure.figsize']` (linha 54)

---

## Precisa de Ajuda?
📧 Email: regina.fonseca@ifg.edu.br
          jrs.joseroberto@gmail.com
🌐 Site: www.ifg.edu.br

---
