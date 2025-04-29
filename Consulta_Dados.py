# Importação das Bibliotecas Necessárias
import sys
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry  # Para seleção de datas
from bcb import sgs
import yfinance as yf
import numpy as np

class DadosConsulta:
    def __init__(self):
        self.codigos_sgs = {
            'CDI': 12,
            'SELIC': 432,
            'IPCA': 433,
            'IGPM': 189,
            'Dólar': 1,
            'IBC-Br': 24363,
            'CDI Acumulado no Mês': 4391,
            'CDI Acumulado no Mês Anualizado': 4392,
            'SELIC Anualizado': 1178,
        }

        self.tickers_yfinance = {
            'IBOV': '^BVSP',
            'S&P 500': '^GSPC',
            'NASDAQ': '^IXIC',
        }

    def consulta_sgs(self, nome, codigo, data_inicio, data_fim):
        try:
            tabela = sgs.get(codigo, start=data_inicio, end=data_fim)
            tabela = tabela.reset_index()
            tabela = tabela.dropna()
            tabela.columns = ['Data', nome]
            return tabela
        except Exception as e:
            print(f"Erro ao coletar {nome}: {e}")
            return pd.DataFrame()

    def consulta_yfinance(self, ticker, nome, data_inicio, data_fim):
        try:
            tabela = yf.download(ticker, start=data_inicio, end=data_fim)[['Close']]
            tabela = tabela.reset_index()
            tabela = tabela.dropna()
            tabela.columns = ['Data', nome]
            return tabela
        except Exception as e:
            print(f"Erro ao coletar {nome}: {e}")
            return pd.DataFrame()

    def indicadores(self):
        return list(self.codigos_sgs.keys()) + list(self.tickers_yfinance.keys())

class MatplotlibCanvas(FigureCanvasTkAgg):
    def __init__(self, figure, master=None):
        self.fig = figure
        self.ax = self.fig.add_subplot(111)
        super().__init__(self.fig, master=master)
        plt.style.use('ggplot')
        self.fig.tight_layout()

        # Variáveis para tooltip
        self.lines = []
        self.annot = self.ax.annotate("", xy=(0, 0), xytext=(20, 20),
                                      textcoords="offset points",
                                      bbox=dict(boxstyle="round", fc="w"),
                                      arrowprops=dict(arrowstyle="->"))
        self.annot.set_visible(False)

        # Conectar eventos
        self.mpl_connect("motion_notify_event", self.hover)

    def hover(self, event):
        if event.inaxes == self.ax:
            closest_dist = float('inf')
            closest_line = None
            closest_x = None
            closest_y = None
            closest_index = None

            for line in self.lines:
                x_data, y_data = line.get_data()
                # Convert x_data (dates) to numerical values for distance calculation
                x_num = np.array([d.toordinal() for d in x_data])
                y_num = np.array(y_data)

                # Transform mouse position to data coordinates
                mouse_x, mouse_y = self.ax.transData.inverted().transform((event.x, event.y))
                mouse_x_num = mouse_x.toordinal()

                # Calculate distances in data coordinates
                distances = np.sqrt((x_num - mouse_x_num)**2 + (y_num - mouse_y)**2)
                min_dist_idx = np.argmin(distances)
                min_dist = distances[min_dist_idx]

                if min_dist < closest_dist:
                    closest_dist = min_dist
                    closest_line = line
                    closest_x = x_data[min_dist_idx]
                    closest_y = y_data[min_dist_idx]
                    closest_index = min_dist_idx

            # Define a threshold for distance to avoid showing tooltip for distant points
            # Adjust threshold based on plot scale (in days and y-value units)
            threshold = 50  # Arbitrary threshold in data units (days squared + value squared)
            if closest_dist < threshold and closest_line is not None:
                self.annot.xy = (closest_x, closest_y)
                self.annot.set_text(f"{closest_line.get_label()}\nX: {closest_x.strftime('%Y-%m-%d')}\nY: {closest_y:.2f}")
                self.annot.set_visible(True)
                self.fig.canvas.draw_idle()
            else:
                self.annot.set_visible(False)
                self.fig.canvas.draw_idle()
        else:
            self.annot.set_visible(False)
            self.fig.canvas.draw_idle()

class ConsultaFinanceiraApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Consulta de Dados Financeiros")
        self.geometry("1200x800")
        
        self.dados = DadosConsulta()
        self.df_resultados = {}  # Dicionário para armazenar os resultados das consultas
        
        self.initUI()
        
    def initUI(self):
        # Container principal
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Painel de controles (esquerda)
        controls_frame = ttk.Frame(main_frame, width=400)
        controls_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        controls_frame.pack_propagate(False)
        
        # Grupo para seleção de datas
        date_frame = ttk.LabelFrame(controls_frame, text="Período de Consulta", padding=10)
        date_frame.pack(fill=tk.X, pady=5)
        
        # Data de início
        tk.Label(date_frame, text="Data Inicial:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.date_inicio = DateEntry(date_frame, date_pattern='yyyy-mm-dd')
        self.date_inicio.set_date(datetime.now() - timedelta(days=365))
        self.date_inicio.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Data de fim
        tk.Label(date_frame, text="Data Final:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.date_fim = DateEntry(date_frame, date_pattern='yyyy-mm-dd')
        self.date_fim.set_date(datetime.now())
        self.date_fim.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Grupo para seleção de indicadores
        indicadores_frame = ttk.LabelFrame(controls_frame, text="Indicadores Disponíveis", padding=10)
        indicadores_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.indicadores_vars = {}
        self.indicadores_frame = ttk.Frame(indicadores_frame)
        self.indicadores_frame.pack(fill=tk.BOTH, expand=True)
        
        # Criar checkboxes para indicadores
        for i, indicador in enumerate(self.dados.indicadores()):
            var = tk.BooleanVar(value=False)
            self.indicadores_vars[indicador] = var
            chk = ttk.Checkbutton(self.indicadores_frame, text=indicador, variable=var)
            chk.grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
        
        # Botões de ação
        buttons_frame = ttk.Frame(controls_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        self.btn_consultar = ttk.Button(buttons_frame, text="Consultar Dados", command=self.consultar_dados)
        self.btn_consultar.pack(fill=tk.X, pady=5)
        
        self.btn_exportar_excel = ttk.Button(buttons_frame, text="Exportar para Excel", command=self.exportar_excel)
        self.btn_exportar_excel.pack(fill=tk.X, pady=5)
        self.btn_exportar_excel.config(state='disabled')
        
        self.btn_exportar_grafico = ttk.Button(buttons_frame, text="Exportar Gráfico (PNG)", command=self.exportar_grafico)
        self.btn_exportar_grafico.pack(fill=tk.X, pady=5)
        self.btn_exportar_grafico.config(state='disabled')
        
        # Painel do gráfico (direita)
        graph_frame = ttk.Frame(main_frame)
        graph_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # Canvas do matplotlib para o gráfico
        self.fig = plt.Figure(figsize=(8, 6), dpi=100)
        self.canvas = MatplotlibCanvas(self.fig, master=graph_frame)
        self.toolbar = NavigationToolbar2Tk(self.canvas, graph_frame, pack_toolbar=False)
        
        self.toolbar.pack(side=tk.TOP, fill=tk.X)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    def consultar_dados(self):
        # Limpar dados anteriores
        self.df_resultados.clear()
        self.canvas.ax.clear()
        
        # Obter datas selecionadas
        data_inicio = self.date_inicio.get()
        data_fim = self.date_fim.get()
        
        # Verificar se as datas são válidas
        try:
            inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
            fim = datetime.strptime(data_fim, '%Y-%m-%d')
            if inicio > fim:
                messagebox.showwarning("Erro", "Data inicial deve ser anterior à data final.")
                return
        except ValueError:
            messagebox.showwarning("Erro", "Formato de data inválido.")
            return
        
        # Obter indicadores selecionados
        indicadores_selecionados = [indicador for indicador, var in self.indicadores_vars.items() if var.get()]
        
        if not indicadores_selecionados:
            messagebox.showwarning("Erro", "Selecione pelo menos um indicador.")
            return
        
        # Consultar dados para cada indicador selecionado
        for indicador in indicadores_selecionados:
            if indicador in self.dados.codigos_sgs:
                codigo = self.dados.codigos_sgs[indicador]
                df = self.dados.consulta_sgs(indicador, codigo, data_inicio, data_fim)
            elif indicador in self.dados.tickers_yfinance:
                ticker = self.dados.tickers_yfinance[indicador]
                df = self.dados.consulta_yfinance(ticker, indicador, data_inicio, data_fim)
            else:
                continue
            
            if not df.empty:
                self.df_resultados[indicador] = df
        
        if not self.df_resultados:
            messagebox.showwarning("Erro", "Não foi possível obter dados para os indicadores selecionados.")
            return
        
        # Plotar os dados
        self.plotar_dados()
        
        # Habilitar botões de exportação
        self.btn_exportar_excel.config(state='normal')
        self.btn_exportar_grafico.config(state='normal')
    
    def plotar_dados(self):
        self.canvas.ax.clear()
        self.canvas.lines.clear()

        # Configurar estilo escuro
        plt.style.use('dark_background')
        self.canvas.fig.set_facecolor('#121212')
        self.canvas.ax.set_facecolor('#121212')

        # Plotar cada indicador com linha azul
        for indicador, df in self.df_resultados.items():
            line, = self.canvas.ax.plot(df['Data'], df[indicador], 
                                        label=indicador, 
                                        color='#1f77b4',
                                        linewidth=2)
            self.canvas.lines.append(line)

        # Configurar o gráfico
        self.canvas.ax.set_title('Dados Financeiros', color='white')
        self.canvas.ax.set_xlabel('Data', color='white')
        self.canvas.ax.set_ylabel('Valor', color='white')
        self.canvas.ax.legend(facecolor='#121212', edgecolor='white', labelcolor='white')
        self.canvas.ax.grid(True, color='gray', linestyle='--', alpha=0.5)

        # Configurar cores dos eixos
        self.canvas.ax.tick_params(colors='white', which='both')
        self.canvas.ax.spines['bottom'].set_color('white')
        self.canvas.ax.spines['top'].set_color('white')
        self.canvas.ax.spines['right'].set_color('white')
        self.canvas.ax.spines['left'].set_color('white')

        # Ajustar layout e exibir
        self.canvas.fig.tight_layout()
        self.canvas.draw()

    def exportar_excel(self):
        if not self.df_resultados:
            messagebox.showwarning("Erro", "Não há dados para exportar.")
            return
        
        # Criar um dataframe combinado para exportação
        df_combinado = None
        
        for indicador, df in self.df_resultados.items():
            if df_combinado is None:
                df_combinado = df.copy()
            else:
                df_combinado = pd.merge(df_combinado, df, on='Data', how='outer')
        
        # Solicitar ao usuário o local para salvar o arquivo
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        
        if filename:
            try:
                df_combinado.to_excel(filename, index=False)
                messagebox.showinfo("Sucesso", f"Dados exportados para {filename}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar dados: {str(e)}")
    
    def exportar_grafico(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Files", "*.png"), ("All Files", "*.*")]
        )
        
        if filename:
            try:
                self.canvas.fig.savefig(filename, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Sucesso", f"Gráfico salvo em {filename}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar gráfico: {str(e)}")

if __name__ == "__main__":
    app = ConsultaFinanceiraApp()
    app.mainloop()
