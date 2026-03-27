"""
====================================================================================================
ANÁLISE DE PESQUISA DE AVALIAÇÃO INSTITUCIONAL - INSTITUTO FEDERAL
====================================================================================================
Autor: Equipe de Avaliação Institucional
Data: Fevereiro de 2026
Versão: 1.3

Este script realiza análise completa dos dados da pesquisa de avaliação institucional,
contemplando respostas de Discentes, Docentes e Técnicos-Administrativos (TAEs).

FUNCIONALIDADES:
- Carregamento e limpeza de dados
- Análise estatística descritiva por perfil
- Geração de tabelas resumo por Eixo-Dimensão
- Visualizações gráficas detalhadas
- Infográfico executivo
- Exportação de relatórios

USO NO GOOGLE COLAB:
1. Faça upload dos arquivos Excel
2. Execute todas as células
3. Os resultados serão salvos automaticamente
====================================================================================================
"""

# ===================================================================================================
# IMPORTAÇÃO DE BIBLIOTECAS
# ===================================================================================================
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.gridspec import GridSpec
import warnings
from typing import Dict, List, Tuple, Optional
from datetime import datetime

warnings.filterwarnings('ignore')

# ===================================================================================================
# PALETA DE CORES CENTRALIZADA
# Altere aqui para modificar todas as cores da aplicação de uma só vez.
# ===================================================================================================

# --- Escala de avaliação (Ótimo → Péssimo) ---
# Gradiente deliberadamente contrastante: verde-floresta → amarelo-mostarda → laranja-ferrugem
COR_OTIMO    = '#1a5c38'   # Verde-floresta intenso
COR_BOM      = '#57a55a'   # Verde-médio vivo
COR_REGULAR  = '#c8b400'   # Amarelo-mostarda (neutro visível)
COR_RUIM     = '#d46b2f'   # Laranja-queimado
COR_PESSIMO  = '#a63215'   # Vermelho-tijolo escuro
COR_NAO_SEI  = '#8faa85'   # Cinza-verde apagado

# --- Perfis (Discentes, Docentes, TAEs) ---
# Três tons bem separados: escuro / médio / claro-acinzentado
COR_DISCENTES = '#1a5c38'   # Verde-floresta (igual ao Ótimo - âncora escura)
COR_DOCENTES  = '#57a55a'   # Verde-médio vivo
COR_TAES      = '#a8c5a0'   # Verde-acinzentado claro (contraste suficiente com os dois acima)

# --- Linhas de referência nos gráficos ---
COR_LINHA_MINIMO  = '#a63215'   # Vermelho-tijolo (alerta)
COR_LINHA_BOM     = '#1a5c38'   # Verde-floresta (meta positiva)
COR_LINHA_REGULAR = '#c8b400'   # Amarelo-mostarda (faixa neutra)

# --- Elementos de destaque e fundo ---
COR_TITULO       = '#0d2e1a'   # Verde quase-preto (máximo contraste em títulos)
COR_TEXTO        = '#1e2e20'   # Verde muito escuro para corpo de texto
COR_FUNDO_CAIXA  = '#eaf4e8'   # Verde-menta muito suave (fundo de caixas)
COR_BORDA_CAIXA  = '#57a55a'   # Verde-médio vivo (borda visível)
COR_GRADE        = '#c4d9bf'   # Verde-acinzentado para grades (discreto)
COR_SUBTITULO    = '#4682B4'   # azul-claro
COR_TEXTO_INTERNO_BARRA = 'white'

# --- Sequência para gráficos genéricos (pie, barras por turno/modalidade) ---
# Saltos deliberados de luminosidade para distinguir fatias/barras facilmente
PALETA_GERAL = [
    '#0d2e1a',   # Verde quase-preto
    '#1a5c38',   # Verde-floresta
    '#2e8b50',   # Verde-esmeralda
    '#57a55a',   # Verde-médio vivo
    '#82c077',   # Verde-maçã
    '#a8c5a0',   # Verde-acinzentado claro
    '#c8b400',   # Amarelo-mostarda (quebra de tom intencional)
    '#d4855a',   # Pêssego-terra (contraste quente no final)
]

# --- Avaliação consolidada positivo / neutro / negativo ---
COR_POSITIVO = '#2e8b50'   # Verde-esmeralda (leitura imediata: bom)
COR_NEUTRO   = '#c8b400'   # Amarelo-mostarda (atenção, não alerta)
COR_NEGATIVO = '#a63215'   # Vermelho-tijolo (alerta claro, sem ser gritante)
# --perfis--
PERFIS = ['DISCENTES', 'DOCENTES', 'TAES']
PERFIS_SIGLA = [('DISCENTES', 'DIS'), ('DOCENTES', 'DOC'), ('TAES', 'TAE')]

PASTA_RESULTADOS = "resultados_analise/"

# ===================================================================================================
# Configurações globais de visualização
# ===================================================================================================
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette(PALETA_GERAL)
plt.rcParams['figure.figsize']    = (14, 8)
plt.rcParams['font.size']         = 11
plt.rcParams['axes.titlesize']    = 13
plt.rcParams['axes.labelsize']    = 11
plt.rcParams['xtick.labelsize']   = 10
plt.rcParams['ytick.labelsize']   = 10
plt.rcParams['legend.fontsize']   = 10
plt.rcParams['figure.titlesize']  = 15
plt.rcParams['font.family']       = 'DejaVu Sans'
plt.rcParams['axes.titlecolor']   = COR_TITULO
plt.rcParams['axes.labelcolor']   = COR_TEXTO
plt.rcParams['xtick.color']       = COR_TEXTO
plt.rcParams['ytick.color']       = COR_TEXTO


# ===================================================================================================
# CLASSE PRINCIPAL
# ===================================================================================================
class AnalisadorPesquisa:
    """
    Classe principal para análise da pesquisa de avaliação institucional.
    """
    def __init__(self, ARQUIVO_PESQUISA: str, caminho_demografico: Optional[str] = None):
        """
        Inicializa o analisador com o arquivo Excel.

        Args:
            ARQUIVO_PESQUISA: Caminho para arquivo Excel com respostas da pesquisa
            caminho_demografico: Caminho para arquivo Excel com dados de matrículas e servidores
        """
        self.ARQUIVO_PESQUISA = ARQUIVO_PESQUISA
        self.caminho_demografico = caminho_demografico
        # local onde serão salvos os resultados
        #self.PASTA_RESULTADOS = "resultados_analise/"

        self.dados = {}
        self.dados_demograficos = {
            'total_alunos': 0,
            'total_docentes': 0,
            'total_taes': 0,
            'cursos': None,
            'taxa_resposta': {}
        }

        self.escalas = ['Ótimo', 'Bom', 'Regular', 'Ruim', 'Péssimo', 'Não sei']

        # Mapeamento escala → cor (usa as variáveis centralizadas)
        self.cores_escala = {
            'Ótimo':   COR_OTIMO,
            'Bom':     COR_BOM,
            'Regular': COR_REGULAR,
            'Ruim':    COR_RUIM,
            'Péssimo': COR_PESSIMO,
            'Não sei': COR_NAO_SEI,
        }

        # Mapeamento perfil → cor
        self.cores_perfil = {
            'Discentes': COR_DISCENTES,
            'Docentes':  COR_DOCENTES,
            'TAEs':      COR_TAES,
        }

        # Cores das barras de perfil na ordem DIS / DOC / TAE
        self.cores_barras_perfil = [COR_DISCENTES, COR_DOCENTES, COR_TAES]

        # Mapeamento de Eixos e Dimensões
        self.estrutura_eixos = {
            'Eixo 1 - Dimensão 8': {
                'nome': 'Planejamento e Avaliação Institucional',
                'questoes': [
                    '1 - O seu conhecimento acerca dos resultados do último processo de autoavaliação institucional realizado pela Comissão Própria de Avaliação (CPA) é:',
                    '2 - A resposta e o encaminhamento institucionais dados às demandas indicadas nos relatórios de avaliação produzidos pela Comissão Própria de Avaliação (CPA) são:',
                    '3 - A divulgação do planejamento anual do seu Câmpus é:',
                    '4 - A sua participação na elaboração do planejamento anual do seu Câmpus é:'
                ]
            },
            'Eixo 2 - Dimensão 1': {
                'nome': 'Missão e Plano de Desenvolvimento Institucional',
                'questoes': [
                    '1 - O seu conhecimento sobre a função social do IFG é:',
                    '2 - O seu conhecimento sobre o Plano de Desenvolvimento Institucional (PDI 2019 prorrogado até 2025) é:'
                ]
            },
            'Eixo 2 - Dimensão 3': {
                'nome': 'Responsabilidade Social da Instituição',
                'questoes': [
                    '1 - O respeito pelas diferenças étnicas, religiosas, políticas, de gênero e sexualidade na comunidade do IFG é:',
                    '2 - O incentivo à promoção de ações de desenvolvimento científico e tecnológico no IFG é:',
                    '3 - O desenvolvimento de ações de ensino, pesquisa, extensão, e administração, no IFG, voltadas para a preservação do meio ambiente é:',
                    '4 - A política de inclusão e acompanhamento de pessoas com necessidades educacionais específicas no IFG é:'
                ]
            }
        }

    def carregar_dados(self):
        """Carrega os dados das três abas do Excel."""
        print(" Carregando dados da pesquisa...")

        try:
            self.dados['DISCENTES'] = pd.read_excel(self.ARQUIVO_PESQUISA, sheet_name='DISCENTES')
            self.dados['DOCENTES']  = pd.read_excel(self.ARQUIVO_PESQUISA, sheet_name='DOCENTES')
            self.dados['TAES']      = pd.read_excel(self.ARQUIVO_PESQUISA, sheet_name='TAES')

            print(f" Dados da pesquisa carregados!")
            print(f"   - Discentes: {len(self.dados['DISCENTES'])} respostas")
            print(f"   - Docentes:  {len(self.dados['DOCENTES'])} respostas")
            print(f"   - TAEs:      {len(self.dados['TAES'])} respostas")
            print(f"   - Total:     {sum([len(df) for df in self.dados.values()])} respostas")

            #renomeia a coluna H da aba discentes, caso seja diferente da coluna I da aba docentes, igualando-as
            col_discente_7_raw = self.dados['DISCENTES'].columns[7]
            col_docente_8_raw = self.dados['DOCENTES'].columns[8]
            questao_discente = col_discente_7_raw.strip()
            questao_docente = col_docente_8_raw.strip()

            if questao_discente != questao_docente:
                self.dados['DISCENTES'].rename(columns={col_discente_7_raw: questao_docente}, inplace=True)  
                print(f"   - Coluna discente '{questao_discente[:30]}...' renomeada para '{questao_docente}'")
                print(f"")

            #renomeia a coluna J da aba discentes, e a coluna K das abas docentes e TAEs, deixando com o mesmo conteúdo definido na planilha de questões, igualando-as
            enunciado_questao = (self.estrutura_eixos['Eixo 2 - Dimensão 3']['questoes'][2]).strip() #pega o enunciado da questão 3 da dimensão 3 do eixo 2
            
            col_discente_9_raw = self.dados['DISCENTES'].columns[9]
            col_docente_10_raw = self.dados['DOCENTES'].columns[10]
            col_tae_10_raw = self.dados['TAES'].columns[10]
            
            questao_discente = col_discente_9_raw.strip()
            questao_docente = col_docente_10_raw.strip()
            questao_tae = col_tae_10_raw.strip()

            if questao_discente != enunciado_questao:
                self.dados['DISCENTES'].rename(columns={col_discente_9_raw: enunciado_questao}, inplace=True)  
                print(f"   - Coluna discente '{questao_discente[:30]}...' renomeada para '{enunciado_questao[:30]}...'")

            if questao_docente != enunciado_questao:
                self.dados['DOCENTES'].rename(columns={col_docente_10_raw: enunciado_questao}, inplace=True)  
                print(f"   - Coluna docente '{questao_docente[:30]}...' renomeada para '{enunciado_questao[:30]}...'")

            if questao_tae != enunciado_questao:
                self.dados['TAES'].rename(columns={col_tae_10_raw: enunciado_questao}, inplace=True)  
                print(f"   - Coluna TAE '{questao_tae[:30]}...' renomeada para '{enunciado_questao[:30]}...'")

        except Exception as e:
            print(f" Erro ao carregar dados: {e}")
            raise

    def carregar_dados_demograficos(self):
        """Carrega dados de matrículas e servidores."""
        if not self.caminho_demografico:
            print("Arquivo demográfico não fornecido. Pulando esta etapa.")
            return

        print("Carregando dados demográficos...")

        try:
            df_disc_raw = pd.read_excel(self.caminho_demografico, sheet_name='DISCENTES', header=None)

            cursos_data = []
            for idx, row in df_disc_raw.iterrows():
                if idx < 3:
                    continue
                if pd.notna(row[0]) and 'TOTAL' not in str(row[0]).upper() and 'SUBTOTAL' not in str(row[0]).upper():
                    if pd.notna(row[2]):
                        cursos_data.append({
                            'Curso':        row[0],
                            'Modalidade':   row[1] if pd.notna(row[1]) else 'N/A',
                            'Total_Alunos': row[2],
                            'Vagas':        row[3] if pd.notna(row[3]) else 'N/A',
                            'Turno':        row[4] if pd.notna(row[4]) else 'N/A'
                        })

            self.dados_demograficos['cursos']       = pd.DataFrame(cursos_data)
            self.dados_demograficos['total_alunos'] = int(self.dados_demograficos['cursos']['Total_Alunos'].sum())

            df_serv = pd.read_excel(self.caminho_demografico, sheet_name='SERVIDORES')
            self.dados_demograficos['total_taes']     = int(df_serv.iloc[0, 1])
            self.dados_demograficos['total_docentes'] = int(df_serv.iloc[1, 1])

            print(f" Dados demográficos carregados!")
            print(f"   - Total de alunos matriculados: {self.dados_demograficos['total_alunos']:,}")
            print(f"   - Total de cursos:   {len(self.dados_demograficos['cursos'])}")
            print(f"   - Total de docentes: {self.dados_demograficos['total_docentes']}")
            print(f"   - Total de TAEs:     {self.dados_demograficos['total_taes']}")

        except Exception as e:
            print(f" Erro ao carregar dados demográficos: {e}")
            print("   Continuando sem dados demográficos...")

    def calcular_taxas_resposta(self):
        """Calcula taxas de resposta baseado nos dados demográficos."""
        if self.dados_demograficos['total_alunos'] == 0:
            print("Dados demográficos não disponíveis para calcular taxas de resposta.")
            return

        print("\n Calculando taxas de resposta...")

        n_disc = len(self.dados.get('DISCENTES', []))
        n_doc  = len(self.dados.get('DOCENTES',  []))
        n_taes = len(self.dados.get('TAES',      []))

        self.dados_demograficos['taxa_resposta'] = {
            'DISCENTES': {
                'respostas':  n_disc,
                'populacao':  self.dados_demograficos['total_alunos'],
                'taxa': (n_disc / self.dados_demograficos['total_alunos']) * 100
                        if self.dados_demograficos['total_alunos'] > 0 else 0
            },
            'DOCENTES': {
                'respostas':  n_doc,
                'populacao':  self.dados_demograficos['total_docentes'],
                'taxa': (n_doc / self.dados_demograficos['total_docentes']) * 100
                        if self.dados_demograficos['total_docentes'] > 0 else 0
            },
            'TAES': {
                'respostas':  n_taes,
                'populacao':  self.dados_demograficos['total_taes'],
                'taxa': (n_taes / self.dados_demograficos['total_taes']) * 100
                        if self.dados_demograficos['total_taes'] > 0 else 0
            }
        }

        print(" Taxas de resposta calculadas:")
        for perfil, dados in self.dados_demograficos['taxa_resposta'].items():
            print(f"   - {perfil}: {dados['respostas']}/{dados['populacao']} = {dados['taxa']:.1f}%")

    def visualizar_dados_demograficos(self, salvar: bool = True):
        """Cria visualizações dos dados demográficos e taxas de resposta.
        Returns:
            fig: figura matplotlib gerada (usada por exportar_graficos_pdf)
        """
        if self.dados_demograficos['total_alunos'] == 0:
            print("  Dados demográficos não disponíveis para visualização.")
            return None

        print("\n Criando visualizações demográficas...")

        fig = plt.figure(figsize=(18, 12))
        gs  = GridSpec(3, 3, figure=fig, hspace=0.4, wspace=0.3)

        # Título em três linhas com cores independentes: as duas primeiras em COR_TITULO, a última em destaque (COR_BOM)
        fig.text(0.5, 0.984, 'INSTITUTO FEDERAL DE GOIÁS - CAMPUS GOIÂNIA',
                 ha='center', va='top', fontsize=20, fontweight='bold', color=COR_TITULO)
        fig.text(0.5, 0.968, 'PESQUISA DE AVALIAÇÃO INSTITUCIONAL',
                 ha='center', va='top', fontsize=20, fontweight='bold', color=COR_TITULO)
        fig.text(0.5, 0.952, '\nDados Demográficos e Representatividade da Pesquisa\n',
                 ha='center', va='top', fontsize=20, fontweight='bold', color=COR_SUBTITULO)

        # 1. População vs Respostas
        ax1 = fig.add_subplot(gs[0, 0])
        populacao = [self.dados_demograficos['total_alunos'],
                     self.dados_demograficos['total_docentes'],
                     self.dados_demograficos['total_taes']]
        respostas = [self.dados_demograficos['taxa_resposta'][p]['respostas'] for p in PERFIS]

        x     = np.arange(len(PERFIS))
        width = 0.35

        bars1 = ax1.bar(x - width/2, populacao, width, label='População Total',
                        color=COR_DISCENTES, alpha=0.6)
        bars2 = ax1.bar(x + width/2, respostas, width, label='Respostas',
                        color=COR_OTIMO, alpha=0.85)

        ax1.set_ylabel('Quantidade', fontweight='bold')
        ax1.set_title('População vs Respostas da Pesquisa', fontweight='bold')
        ax1.set_xticks(x)
        ax1.set_xticklabels(PERFIS) #perfis)
        ax1.legend()
        ax1.grid(axis='y', alpha=0.3, color=COR_GRADE)

        for bar in bars1:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                     f'{int(height):,}', ha='center', va='bottom', fontsize=9, color=COR_TEXTO)
        for bar in bars2:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                     f'{int(height)}', ha='center', va='bottom', fontsize=9,
                     fontweight='bold', color=COR_TEXTO)

        # 2. Taxa de Resposta por perfil
        ax2   = fig.add_subplot(gs[0, 1])
        taxas = [self.dados_demograficos['taxa_resposta'][p]['taxa'] for p in PERFIS] #perfis]

        # Cores dinâmicas por faixa - mantidas na paleta verde/terra
        cores_taxa = [
            COR_NEGATIVO if t < 10 else COR_REGULAR if t < 30 else COR_OTIMO
            for t in taxas
        ]

        bars = ax2.bar(PERFIS, taxas, color=cores_taxa, alpha=0.85, edgecolor=COR_TITULO, linewidth=1.2)
        ax2.set_ylabel('Taxa de Resposta (%)', fontweight='bold')
        ax2.set_title('Taxa de Resposta por Perfil', fontweight='bold')
        ax2.axhline(y=10, color=COR_LINHA_MINIMO, linestyle='--', alpha=0.7,
                    linewidth=2, label='Mínimo (10%)')
        ax2.axhline(y=30, color=COR_LINHA_BOM,    linestyle='--', alpha=0.7,
                    linewidth=2, label='Bom (30%)')
        ax2.set_ylim(0, max(taxas) * 1.3)
        ax2.legend(fontsize=8)
        ax2.grid(axis='y', alpha=0.3, color=COR_GRADE)

        for bar, taxa in zip(bars, taxas):
            ax2.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 1,
                     f'{taxa:.1f}%', ha='center', va='bottom',
                     fontweight='bold', fontsize=11, color=COR_TITULO)

        # 3. Distribuição de cursos por modalidade
        ax3 = fig.add_subplot(gs[0, 2])
        if self.dados_demograficos['cursos'] is not None:
            cursos_por_mod = self.dados_demograficos['cursos']['Modalidade'].value_counts()
            ax3.pie(cursos_por_mod.values,
                    labels=cursos_por_mod.index,
                    autopct='%1.1f%%',
                    startangle=90,
                    colors=PALETA_GERAL[:len(cursos_por_mod)],
                    textprops={'fontsize': 9, 'color': COR_TEXTO})
            ax3.set_title('Distribuição de Cursos por Modalidade', fontweight='bold')

        # 4. Top 10 cursos com mais alunos
        ax4 = fig.add_subplot(gs[1, :])
        if self.dados_demograficos['cursos'] is not None:
            top_cursos = self.dados_demograficos['cursos'].nlargest(10, 'Total_Alunos')

            y_pos = np.arange(len(top_cursos))
            bars  = ax4.barh(y_pos, top_cursos['Total_Alunos'],
                             color=COR_BOM, alpha=0.85)
            ax4.set_yticks(y_pos)
            ax4.set_yticklabels(
                [c[:50] + '...' if len(c) > 50 else c for c in top_cursos['Curso']],
                fontsize=9
            )
            ax4.set_xlabel('Número de Alunos', fontweight='bold')
            ax4.set_title('Top 10 Cursos com Mais Alunos Matriculados',
                          fontweight='bold', pad=15)
            ax4.grid(axis='x', alpha=0.3, color=COR_GRADE)

            for i, (bar, valor) in enumerate(zip(bars, top_cursos['Total_Alunos'])):
                ax4.text(valor + 5, i, f'{int(valor)}',
                         va='center', fontweight='bold', color=COR_TEXTO)
        # 5. Alunos por turno
        ax5 = fig.add_subplot(gs[2, 0])
        if self.dados_demograficos['cursos'] is not None:
            alunos_turno = (self.dados_demograficos['cursos']
                            .groupby('Turno')['Total_Alunos']
                            .sum()
                            .sort_values(ascending=False))

            ax5.bar(range(len(alunos_turno)), alunos_turno.values,
                    color=PALETA_GERAL[:len(alunos_turno)],
                    alpha=0.85, edgecolor=COR_TITULO, linewidth=0.8)
            ax5.set_xticks(range(len(alunos_turno)))
            ax5.set_xticklabels(alunos_turno.index, rotation=45, ha='right')
            ax5.set_ylabel('Número de Alunos', fontweight='bold')
            ax5.set_title('Distribuição de Alunos por Turno', fontweight='bold')
            ax5.grid(axis='y', alpha=0.3, color=COR_GRADE)

            for i, valor in enumerate(alunos_turno.values):
                ax5.text(i, valor + 20, f'{int(valor)}',
                         ha='center', fontweight='bold', color=COR_TEXTO)
        # 6. Representatividade da amostra
        ax6 = fig.add_subplot(gs[2, 1:])
        ax6.axis('off')

        obs_linhas = []
        for p in PERFIS: 
            taxa_resp = self.dados_demograficos['taxa_resposta'][p]['taxa']
            if taxa_resp < 10:
                classificacao, acao = "CRÍTICA", "indispensável ação para ampliar divulgação"
            elif taxa_resp < 30:
                classificacao, acao = "BAIXA", "ampliar divulgação"
            elif taxa_resp < 50:
                classificacao, acao = "BOA", "amostra representativa"
            else:
                classificacao, acao = "EXCELENTE", "alta representatividade"
            obs_linhas.append(f"• Taxa de resposta de {p}: {classificacao} ({taxa_resp:.1f}%)") # - {acao}")

        obs_texto  = "\n        ".join(obs_linhas)
        texto_repr = f"""
        ANÁLISE DE REPRESENTATIVIDADE DA AMOSTRA

        População Total do Campus:
        • Discentes: {self.dados_demograficos['total_alunos']:,} alunos em {len(self.dados_demograficos['cursos'])} cursos
        • Docentes: {self.dados_demograficos['total_docentes']} docentes
        • TAEs: {self.dados_demograficos['total_taes']} técnicos-administrativos
        • Total: {self.dados_demograficos['total_alunos'] + self.dados_demograficos['total_docentes'] + self.dados_demograficos['total_taes']:,} pessoas

        Respostas Obtidas:
        • Discentes: {self.dados_demograficos['taxa_resposta']['DISCENTES']['respostas']} ({self.dados_demograficos['taxa_resposta']['DISCENTES']['taxa']:.1f}%)
        • Docentes:  {self.dados_demograficos['taxa_resposta']['DOCENTES']['respostas']} ({self.dados_demograficos['taxa_resposta']['DOCENTES']['taxa']:.1f}%)
        • TAEs:      {self.dados_demograficos['taxa_resposta']['TAES']['respostas']} ({self.dados_demograficos['taxa_resposta']['TAES']['taxa']:.1f}%)

        Observações:
        {obs_texto}
        • Total de {sum([self.dados_demograficos['taxa_resposta'][p]['respostas'] for p in PERFIS])} respondentes
        """

        ax6.text(0.05, 0.95, texto_repr, transform=ax6.transAxes,
                 fontsize=11, verticalalignment='top', family='monospace',
                 color=COR_TEXTO,
                 bbox=dict(boxstyle='round', facecolor=COR_FUNDO_CAIXA, edgecolor=COR_BORDA_CAIXA, alpha=0.6))

        if salvar:
            nome_arquivo = f"{PASTA_RESULTADOS}analise_demografica.png"
            plt.savefig(nome_arquivo, dpi=300, bbox_inches='tight')
            print(f"    Gráfico salvo: {nome_arquivo}")

        plt.show()
        return fig

    def limpar_nome_coluna(self, nome: str) -> str:
        """Remove quebras de linha e espaços extras dos nomes de colunas."""
        return ' '.join(nome.replace('\n', ' ').split())

    def calcular_quantidades(self, serie: pd.Series) -> Dict[str, float]:
        """Calcula totais para cada categoria da escala."""
        total = len(serie)
        if total == 0:
            return {escala: 0.0 for escala in self.escalas}

        contagem = serie.value_counts()
        return {escala: contagem.get(escala, 0) for escala in self.escalas}

    def calcular_percentuais(self, serie: pd.Series) -> Dict[str, float]:
        """Calcula percentuais para cada categoria da escala."""
        total = len(serie)
        if total == 0:
            return {escala: 0.0 for escala in self.escalas}

        contagem = serie.value_counts()
        return {escala: (contagem.get(escala, 0) / total) * 100 for escala in self.escalas}

    def gerar_tabelas_resumo(self) -> Dict[str, pd.DataFrame]:
        """Gera tabelas resumo por Eixo-Dimensão."""
        print("\n Gerando tabelas resumo por Eixo-Dimensão...")
        tabelas = {}

        for eixo, info in self.estrutura_eixos.items():
            print(f"   - Processando {eixo}: {info['nome']}")
            linhas = []
            for questao_parcial in info['questoes']:
                for perfil in PERFIS:
                    df = self.dados[perfil]

                    q = " ".join(questao_parcial.split())
                    coluna_encontrada = None

                    # #debug
                    # if q.lower().startswith("3 - o desenvolvimento de ações de ensino, pesquisa, extensão"):
                    #     print(f"     - Procurando coluna para questão: '{q[:30]}...' no perfil {perfil}")
                    #     print(f"     - Colunas disponíveis: {[c for c in df.columns[2:]]}")

                    for col in df.columns[2:]:
                        c = " ".join(col.split())
                        if q.lower() in c.lower():
                            coluna_encontrada = col
                            break

                    if coluna_encontrada:
                        percentuais = self.calcular_percentuais(df[coluna_encontrada])
                        quantidades_abs = self.calcular_quantidades(df[coluna_encontrada])
                        for escala in self.escalas:
                            linhas.append({
                                'Eixo-Dimensão': eixo,
                                'Questão':       questao_parcial[:50] + '...'
                                                 if len(questao_parcial) > 50 else questao_parcial,
                                'Perfil':        perfil,
                                'Escala':        escala,
                                'Percentual':    round(percentuais[escala], 1),
                                'Qtd_Resposta':  int(quantidades_abs[escala]),
                                'N_Respostas':   len(df)
                            })
                    else:
                        print(f"     - Coluna correspondente à questão '{q[:50]}...' não encontrada para perfil {perfil}.")

            if linhas:
                tabelas[eixo] = pd.DataFrame(linhas)

        print(f" {len(tabelas)} tabelas geradas!")
        return tabelas

    def criar_tabela_formatada(self, eixo: str, dados_tabela: pd.DataFrame) -> pd.DataFrame:
        """Cria tabela formatada no estilo especificado no documento Word."""
        tabela_pivot = dados_tabela.pivot_table(
            index='Escala',
            columns=['Questão', 'Perfil'],
            values='Percentual',
            aggfunc='first'
        )

        tabela_pivot = tabela_pivot.reindex([e for e in self.escalas if e in tabela_pivot.index])

        valores_escala = {'Ótimo': 5, 'Bom': 4, 'Regular': 3, 'Ruim': 2, 'Péssimo': 1, 'Não sei': np.nan}
        medias = {}

        for col in tabela_pivot.columns:
            valores_ponderados = []
            for escala in tabela_pivot.index:
                if pd.notna(tabela_pivot.loc[escala, col]) and escala != 'Não sei':
                    valores_ponderados.append(valores_escala[escala] * tabela_pivot.loc[escala, col])

            if valores_ponderados:
                soma_valores     = sum(valores_ponderados)
                soma_percentuais = tabela_pivot[col].sum()
                medias[col]      = round(soma_valores / soma_percentuais, 2) if soma_percentuais > 0 else np.nan
            else:
                medias[col] = np.nan

        tabela_pivot.loc['Média'] = pd.Series(medias)
        return tabela_pivot

    def visualizar_por_eixo(self, eixo: str, dados_tabela: pd.DataFrame, salvar: bool = True):
        """Cria visualização gráfica para um Eixo-Dimensão específico.
        Returns:
            fig: figura matplotlib gerada (usada por exportar_graficos_pdf)
        """
        info     = self.estrutura_eixos[eixo]
        questoes = dados_tabela['Questão'].unique()

        n_questoes = len(questoes)
        fig, axes  = plt.subplots(n_questoes, 1, figsize=(16, 8 * n_questoes))

        if n_questoes == 1:
            axes = [axes]

        fig.suptitle(f"{eixo}: {info['nome']}",
                     fontsize=18, fontweight='bold', y=0.995, color=COR_TITULO)

        for idx, questao in enumerate(questoes):
            ax = axes[idx]
            dados_questao = dados_tabela[dados_tabela['Questão'] == questao]

            x      = np.arange(len(PERFIS))
            width  = 0.12

            for i, escala in enumerate(self.escalas):
                valores = []
                quantidades = []
                for perfil in PERFIS:
                    dados_perfil = dados_questao[
                        (dados_questao['Perfil'] == perfil) &
                        (dados_questao['Escala'] == escala)
                    ]
                    #adiciona o valor percentual
                    perc = dados_perfil['Percentual'].values[0] if not dados_perfil.empty else 0
                    valores.append(perc)
                    #adiciona a quantidade real da resposta
                    qtd = dados_perfil['Qtd_Resposta'].values[0] if not dados_perfil.empty else 0
                    quantidades.append(qtd)

                offset = (i - 2.5) * width #centraliza as barras
                bars   = ax.bar(x + offset, valores, width, label=escala,
                                color=self.cores_escala[escala], alpha=0.88,
                                edgecolor='white', linewidth=0.5)

                for idx,bar in enumerate(bars):
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height,
                            f'{height:.0f}%',
                            ha='center', va='bottom', fontweight='bold', fontsize=10, color=COR_TEXTO)
                    if height > 2:
                        #coloca a quantidade dentro da barra
                        ax.text(bar.get_x() + bar.get_width()/2., height -2,
                                f'{quantidades[idx]:.0f}',
                                ha='center', va='top', fontsize=8, color=COR_TEXTO_INTERNO_BARRA)

            #ax.set_xlabel('Perfil', fontweight='bold')
            ax.set_ylabel('Percentual (%)', fontweight='bold')
            ax.set_title(f"Questão {idx+1}: {questao[:100]}...", fontsize=12, pad=10)
            ax.set_xticks(x)
            ax.set_xticklabels(PERFIS)
            ax.legend(loc='upper right', ncol=3, framealpha=0.9)
            ax.set_ylim(0, 100)
            ax.grid(axis='y', alpha=0.3, color=COR_GRADE)

        plt.tight_layout(h_pad=4.0)

        if salvar:
            nome_arquivo = f"{PASTA_RESULTADOS}grafico_{eixo.replace(' ', '_').replace(',', '')}.png"
            plt.savefig(nome_arquivo, dpi=300, bbox_inches='tight')
            print(f"   Gráfico salvo: {nome_arquivo}")

        plt.show()
        return fig

    def criar_infografico_executivo(self, tabelas: Dict[str, pd.DataFrame], salvar: bool = True):
        """Cria infográfico executivo com resumo geral dos resultados.
        Returns:
            fig: figura matplotlib gerada (usada por exportar_graficos_pdf)
        """
        print("\n Criando infográfico executivo...")

        fig = plt.figure(figsize=(20, 14))
        gs  = GridSpec(4, 3, figure=fig, hspace=0.4, wspace=0.3)

        # Título em três linhas com cores independentes: as duas primeiras em COR_TITULO, a última em destaque (COR_BOM)
        fig.text(0.5, 0.984, 'INSTITUTO FEDERAL DE GOIÁS - CAMPUS GOIÂNIA',
                 ha='center', va='top', fontsize=20, fontweight='bold', color=COR_TITULO)
        fig.text(0.5, 0.968, 'PESQUISA DE AVALIAÇÃO INSTITUCIONAL',
                 ha='center', va='top', fontsize=20, fontweight='bold', color=COR_TITULO)
        fig.text(0.5, 0.952, '\nResumo Executivo',
                 ha='center', va='top', fontsize=20, fontweight='bold', color=COR_SUBTITULO)

        # Linha 0: subgridspec de 2 colunas para dividir igualmente entre os 2 gráficos remanescentes
        gs_row0 = gs[0, :].subgridspec(1, 2, wspace=0.35)

        # 1. Participação por perfil
        ax1 = fig.add_subplot(gs_row0[0, 0])
        participacao = {
            'Discentes': len(self.dados['DISCENTES']),
            'Docentes':  len(self.dados['DOCENTES']),
            'TAEs':      len(self.dados['TAES'])
        }
        wedges, texts, autotexts = ax1.pie(
            participacao.values(),
            labels=participacao.keys(),
            autopct='%1.1f%%',
            colors=[self.cores_perfil[p] for p in participacao],
            startangle=90,
            textprops={'fontsize': 11, 'weight': 'bold', 'color': COR_TITULO}
        )
        ax1.set_title('Participação por Perfil', fontsize=14, fontweight='bold', pad=15)

        total_respondentes = sum(participacao.values())
        """
        # 2. Total de respondentes
        ax2 = fig.add_subplot(gs[0, 1])
        ax2.axis('off')
        total_respondentes = sum(participacao.values())
        ax2.text(0.5, 0.6, f'{total_respondentes}', ha='center', va='center',
                 fontsize=60, fontweight='bold', color=COR_OTIMO)
        ax2.text(0.5, 0.25, 'Total de Respondentes', ha='center', va='center',
                 fontsize=14, fontweight='bold', color=COR_TITULO)
        """

        # 3. Distribuição geral de avaliações
        ax3 = fig.add_subplot(gs_row0[0, 1])
        todas_respostas = []
        for df in self.dados.values():
            for col in df.columns[2:]:
                if any(escala in str(df[col].values) for escala in self.escalas):
                    todas_respostas.extend(df[col].dropna().tolist())

        if todas_respostas:
            serie_total        = pd.Series(todas_respostas)
            percentuais_gerais = self.calcular_percentuais(serie_total)
            quantidades_gerais = self.calcular_quantidades(serie_total)

            y_pos  = np.arange(len(self.escalas))
            valores = [percentuais_gerais[e] for e in self.escalas]
            quantidades = [int(quantidades_gerais[e]) for e in self.escalas]
            cores   = [self.cores_escala[e]  for e in self.escalas]

            bars = ax3.barh(y_pos, valores, color=cores, alpha=0.88,
                            edgecolor='white', linewidth=0.5)
            ax3.set_yticks(y_pos)
            ax3.set_yticklabels(self.escalas, fontweight='bold')
            #ax3.set_xlabel('Percentual (%)', fontweight='bold')
            ax3.set_title('Distribuição Geral de Avaliações', fontsize=14, fontweight='bold', pad=15)
            ax3.grid(axis='x', alpha=0.3, color=COR_GRADE)

            for i, (bar, valor, qtd) in enumerate(zip(bars, valores, quantidades)):
                # Plota a porcentagem no final da barra
                ax3.text(valor + 1, i, f'{valor:.1f}%',
                         va='center', fontweight='bold', color=COR_TEXTO)
                # Plota as quantidades respectivas dentro da barra, centralizadas
                if valor > 0:
                    ax3.text(valor / 2, i, f'{qtd}', ha='center', va='center', 
                             fontsize=8, color=COR_TEXTO_INTERNO_BARRA)

        # 4-6. Resumo por Eixo-Dimensão
        eixos_lista = list(self.estrutura_eixos.keys())
        for idx, eixo in enumerate(eixos_lista):
            ax = fig.add_subplot(gs[1, idx])

            if eixo in tabelas:
                dados_eixo     = tabelas[eixo]
                valores_escala = {'Ótimo': 5, 'Bom': 4, 'Regular': 3, 'Ruim': 2, 'Péssimo': 1}
                medias_perfil  = {}

                #for perfil in ['DIS', 'DOC', 'TAE']:
                for perfil in PERFIS:
                    dados_perfil  = dados_eixo[dados_eixo['Perfil'] == perfil]
                    soma_pond     = 0
                    soma_perc     = 0
                    for _, row in dados_perfil.iterrows():
                        if row['Escala'] in valores_escala:
                            soma_pond += valores_escala[row['Escala']] * row['Percentual']
                            soma_perc += row['Percentual']
                    medias_perfil[perfil] = soma_pond / soma_perc if soma_perc > 0 else 0

                perfis  = list(medias_perfil.keys())
                valores = list(medias_perfil.values())

                bars = ax.bar(perfis, valores,
                              color=self.cores_barras_perfil,
                              alpha=0.88, edgecolor=COR_TITULO, linewidth=1.2)
                ax.set_ylim(0, 5)
                ax.set_ylabel('Média (1-5)', fontweight='bold')
                ax.set_title(
                    f"{eixo}\n{self.estrutura_eixos[eixo]['nome'][:40]}...",
                    fontsize=11, fontweight='bold', pad=10
                )
                ax.axhline(y=3, color=COR_LINHA_REGULAR, linestyle='--',
                           alpha=0.7, linewidth=2, label='Regular')
                ax.axhline(y=4, color=COR_LINHA_BOM,     linestyle='--',
                           alpha=0.7, linewidth=2, label='Bom')
                ax.grid(axis='y', alpha=0.3, color=COR_GRADE)
                ax.legend(fontsize=8)

                for bar, valor in zip(bars, valores):
                    ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.1,
                            f'{valor:.2f}',
                            ha='center', va='bottom', fontweight='bold',
                            fontsize=11, color=COR_TITULO)

        # 7-9. Comparação por perfil – barras empilhadas
        handles_legenda = []  # coletado do último eixo para usar na legenda compartilhada
        for idx, eixo in enumerate(eixos_lista):
            ax = fig.add_subplot(gs[2, idx])

            if eixo in tabelas:
                dados_eixo  = tabelas[eixo]
                categorias  = {
                    'Positivo': ['Ótimo', 'Bom'],
                    'Neutro':   ['Regular'],
                    'Negativo': ['Ruim', 'Péssimo']
                }
                resultados = {perfil: {'Positivo': 0, 'Neutro': 0, 'Negativo': 0}
                              for perfil in PERFIS}
                qtds = {
                    perfil: {'Positivo': 0, 'Neutro': 0, 'Negativo': 0}
                    for perfil in PERFIS
                }
                
                for perfil in PERFIS:
                    dados_perfil = dados_eixo[dados_eixo['Perfil'] == perfil]
                    total_perc   = dados_perfil[dados_perfil['Escala'] != 'Não sei']['Percentual'].sum()
                    for cat, escalas in categorias.items():
                        mask = dados_perfil['Escala'].isin(escalas)
                        soma_perc = dados_perfil[mask]['Percentual'].sum()
                        soma_qtd  = dados_perfil[mask]['Qtd_Resposta'].sum()

                        resultados[perfil][cat] = (soma_perc / total_perc * 100) if total_perc > 0 else 0
                        qtds[perfil][cat]       = int(soma_qtd)

                perfis    = list(resultados.keys())
                positivos = [resultados[p]['Positivo'] for p in perfis]
                neutros   = [resultados[p]['Neutro']   for p in perfis]
                negativos = [resultados[p]['Negativo'] for p in perfis]
                qtd_positivos = [qtds[p]['Positivo'] for p in perfis]
                qtd_neutros   = [qtds[p]['Neutro']   for p in perfis]
                qtd_negativos = [qtds[p]['Negativo'] for p in perfis]

                x     = np.arange(len(perfis))
                width = 0.6

                b1 = ax.bar(x, positivos, width,
                       label='Positivo (Ótimo+Bom)',       color=COR_POSITIVO, alpha=0.88)
                b2 = ax.bar(x, neutros,   width, bottom=positivos,
                       label='Neutro (Regular)',            color=COR_NEUTRO,   alpha=0.88)
                b3 = ax.bar(x, negativos, width,
                       bottom=[p+n for p, n in zip(positivos, neutros)],
                       label='Negativo (Ruim+Péssimo)',     color=COR_NEGATIVO, alpha=0.88)
                handles_legenda = [b1, b2, b3]  # atualiza a cada eixo; o último é suficiente

                # Rótulos dentro de cada segmento: percentual (negrito) + quantidade
                LIMIAR_EXIBIR = 6   # altura mínima do segmento (%) para exibir rótulo
                for i in range(len(perfis)):
                    # Segmento Positivo - centro em positivos[i] / 2
                    if positivos[i] >= LIMIAR_EXIBIR:
                        cy = positivos[i] / 2
                        rotulo = f'{positivos[i]:.0f}%' + '\n' + f'n={qtd_positivos[i]}'
                        ax.text(x[i], cy, rotulo,
                                ha='center', va='center',
                                fontsize=7, color='white',
                                linespacing=1.4)
 
                    # Segmento Neutro - centro em positivos[i] + neutros[i] / 2
                    if neutros[i] >= LIMIAR_EXIBIR:
                        cy = positivos[i] + neutros[i] / 2
                        rotulo = f'{neutros[i]:.0f}%' + '\n' + f'n={qtd_neutros[i]}'
                        ax.text(x[i], cy, rotulo,
                                ha='center', va='center',
                                fontsize=7, color='white',
                                linespacing=1.4)
 
                    # Segmento Negativo - centro em positivos[i] + neutros[i] + negativos[i] / 2
                    if negativos[i] >= LIMIAR_EXIBIR:
                        cy = positivos[i] + neutros[i] + negativos[i] / 2
                        rotulo = f'{negativos[i]:.0f}%' + '\n' + f'n={qtd_negativos[i]}'
                        ax.text(x[i], cy, rotulo,
                                ha='center', va='center',
                                fontsize=7, color='white',
                                linespacing=1.4)

                ax.set_ylabel('Percentual (%)', fontweight='bold')
                ax.set_title(f'Avaliação por Perfil - {eixo}',
                             fontsize=11, fontweight='bold', pad=10)
                ax.set_xticks(x)
                ax.set_xticklabels(perfis)
                ax.set_ylim(0, 100)
                ax.grid(axis='y', alpha=0.3, color=COR_GRADE)

        # Legenda compartilhada abaixo dos gráficos 7-9
        if handles_legenda:
            fig.legend(handles=handles_legenda,
                       labels=['Positivo (Ótimo+Bom)', 'Neutro (Regular)', 'Negativo (Ruim+Péssimo)'],
                       loc='upper center',
                       bbox_to_anchor=(0.5, 0.30),  # base da linha 2 do GridSpec(4,3,hspace=0.4)
                       ncol=3,
                       fontsize=10,
                       framealpha=0.9,
                       edgecolor=COR_BORDA_CAIXA)

        # 10. Quadro de observações
        ax10 = fig.add_subplot(gs[3, :])
        ax10.axis('off')

        melhor_eixo, melhor_media = None, 0
        pior_eixo,   pior_media   = None, 6

        for eixo in eixos_lista:
            if eixo in tabelas:
                dados_eixo     = tabelas[eixo]
                valores_escala = {'Ótimo': 5, 'Bom': 4, 'Regular': 3, 'Ruim': 2, 'Péssimo': 1}
                soma_pond, soma_perc = 0, 0
                for _, row in dados_eixo.iterrows():
                    if row['Escala'] in valores_escala:
                        soma_pond += valores_escala[row['Escala']] * row['Percentual']
                        soma_perc += row['Percentual']
                if soma_perc > 0:
                    media = soma_pond / soma_perc
                    if media > melhor_media:
                        melhor_media, melhor_eixo = media, eixo
                    if media < pior_media:
                        pior_media, pior_eixo = media, eixo


        observacoes = f"""
         Participação: Total de {total_respondentes} respondentes ({participacao['Discentes']} discentes, {participacao['Docentes']} docentes e {participacao['TAEs']} TAEs)
         Melhor Avaliado: {melhor_eixo} (Média: {melhor_media:.2f})
         Necessita Atenção: {pior_eixo} (Média: {pior_media:.2f})

             ** escala de notas utilizada: 5 = Ótimo, 4 = Bom, 3 = Regular, 2 = Ruim, 1 = Péssimo
             ** escala de classificação: Positivo  = (Ótimo+Bom), Neutro = (Regular), Negativo = (Ruim+Péssimo)
         
        """
        #Recomendações: Aprofundar análise dos eixos com média abaixo de 3.5 e desenvolver
        #  planos de ação para melhoria contínua.
        #"""

        ax10.text(0.05, 0.95, observacoes, transform=ax10.transAxes,
                  fontsize=12, verticalalignment='top',
                  color=COR_TEXTO,
                  bbox=dict(boxstyle='round', facecolor=COR_FUNDO_CAIXA,
                            edgecolor=COR_BORDA_CAIXA, alpha=0.6),
                  family='monospace')

        data_atual = datetime.now().strftime("%d/%m/%Y")
        fig.text(0.99, 0.01, f'Gerado em: {data_atual}',
                 ha='right', fontsize=10, style='italic', color=COR_NAO_SEI)

        if salvar:
            nome_arquivo = f"{PASTA_RESULTADOS}infografico_executivo.png"
            plt.savefig(nome_arquivo, dpi=300, bbox_inches='tight')
            print(f"    Infográfico salvo: {nome_arquivo}")

        plt.show()
        return fig

    def exportar_graficos_pdf(self, figuras: List):
        """
        Exporta as figuras já geradas em um único arquivo PDF multi-página.
        Não recria nenhum gráfico - recebe a lista de figuras produzidas por
        visualizar_dados_demograficos, visualizar_por_eixo e criar_infografico_executivo.

        Args:
            figuras: lista de objetos matplotlib.figure.Figure, na ordem desejada
                     de paginação. Figuras None (ex: demográfico indisponível) são ignoradas.
        """
        from matplotlib.backends.backend_pdf import PdfPages

        nome_pdf = f"{PASTA_RESULTADOS}relatorio_graficos_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        print(f"\n📄 Exportando gráficos para PDF: {nome_pdf} ...")

        figuras_validas = [f for f in figuras if f is not None]

        with PdfPages(nome_pdf) as pdf:
            for i, fig in enumerate(figuras_validas, start=1):
                pdf.savefig(fig, bbox_inches='tight')
                print(f" Página {i} exportada")

            # Metadados do PDF
            info_pdf = pdf.infodict()
            info_pdf['Title']        = 'Relatório de Avaliação Institucional - Instituto Federal'
            info_pdf['Author']       = 'Equipe de Avaliação Institucional'
            info_pdf['Subject']      = 'Pesquisa de Avaliação Institucional'
            info_pdf['CreationDate'] = datetime.now()

        print(f" PDF exportado com sucesso: {nome_pdf} ({len(figuras_validas)} páginas)")
        return nome_pdf

    def exportar_relatorios(self, tabelas: Dict[str, pd.DataFrame]):
        """Exporta relatórios em Excel e CSV."""
        print("\n Exportando relatórios...")
        nome_xls = f"{PASTA_RESULTADOS}relatorio_completo.xlsx"

        with pd.ExcelWriter(nome_xls, engine='openpyxl') as writer:
            resumo_geral = []
            for perfil, sigla in PERFIS_SIGLA: 
                df = self.dados[perfil]
                resumo_geral.append({
                    'Perfil':        perfil, #sigla,
                    'N_Respondentes': len(df),
                    'Campus':        df.iloc[:, 1].unique()[0] if len(df) > 0 else 'N/A'
                })
            pd.DataFrame(resumo_geral).to_excel(writer, sheet_name='Resumo_Geral', index=False)

            for eixo, dados_tabela in tabelas.items():
                nome_aba = eixo.replace(' ', '_').replace(',', '')[:31]
                tabela_formatada = self.criar_tabela_formatada(eixo, dados_tabela)
                tabela_formatada.to_excel(writer, sheet_name=nome_aba)

            dados_consolidados = []
            for perfil in PERFIS: 
                df_temp         = self.dados[perfil].copy()
                df_temp['Perfil'] = perfil
                dados_consolidados.append(df_temp)

            pd.concat(dados_consolidados, ignore_index=True).to_excel(
                writer, sheet_name='Dados_Brutos', index=False
            )

        print("    Relatório Excel salvo: relatorio_completo.xlsx")

        for eixo, dados_tabela in tabelas.items():
            nome_arquivo = f"{PASTA_RESULTADOS}tabela_{eixo.replace(' ', '_').replace(',', '')}.csv"
            dados_tabela.to_csv(nome_arquivo, index=False, encoding='utf-8-sig')
            print(f"    Tabela CSV salva: {nome_arquivo}")

    def executar_analise_completa(self):
        """Executa o pipeline completo de análise com dados demográficos."""
        print("\n" + "="*80)
        print("INICIANDO ANÁLISE COMPLETA DA PESQUISA DE AVALIAÇÃO INSTITUCIONAL")
        print("="*80)

        self.carregar_dados()
        self.carregar_dados_demograficos()
        self.calcular_taxas_resposta()

        # Coleta das figuras na ordem de paginação do PDF
        figuras = []

        fig_demografico = self.visualizar_dados_demograficos()
        figuras.append(fig_demografico)  # None se dados indisponíveis - ignorado no PDF

        tabelas = self.gerar_tabelas_resumo()

        print("\n Criando visualizações gráficas...")
        for eixo, dados_tabela in tabelas.items():
            fig_eixo = self.visualizar_por_eixo(eixo, dados_tabela)
            figuras.append(fig_eixo)

        fig_info = self.criar_infografico_executivo(tabelas)
        figuras.append(fig_info)

        # Exporta todas as figuras já renderizadas - sem recriar nada
        nome_pdf = self.exportar_graficos_pdf(figuras)

        self.exportar_relatorios(tabelas)

        print("\n" + "="*80)
        print(" ANÁLISE COMPLETA FINALIZADA COM SUCESSO!")
        print("="*80)
        print(f"\nArquivos gerados (em {PASTA_RESULTADOS}):")
        print("   Gráficos por eixo-dimensão (grafico_*.png)")
        print("   Infográfico executivo (infografico_executivo.png)")
        print(f"   Relatório PDF completo ({nome_pdf})")
        print("   Relatório Excel completo (relatorio_completo.xlsx)")
        print("   Tabelas CSV individuais (tabela_*.csv)")
        print("\n")


# ===================================================================================================
# EXECUÇÃO PRINCIPAL
# ===================================================================================================
if __name__ == "__main__":
    ARQUIVO_PESQUISA    = 'dados/Gyn.xlsx'
    ARQUIVO_DEMOGRAFICO = 'dados/Numero_Matriculas_alunos-Gyn_Servidores_2025.xlsx'

    analisador = AnalisadorPesquisa(ARQUIVO_PESQUISA, ARQUIVO_DEMOGRAFICO)
    analisador.executar_analise_completa()

    print("\n Para análises exploratórias adicionais, você pode acessar:")
    print("   - analisador.dados: dicionário com DataFrames de cada perfil")
    print("   - analisador.estrutura_eixos: informações sobre eixos e dimensões")
    print("   - analisador.gerar_tabelas_resumo(): regenerar tabelas")
    print("\n")
