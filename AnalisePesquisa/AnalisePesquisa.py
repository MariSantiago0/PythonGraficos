import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
from collections import Counter

# Carregar os dados
data = pd.read_excel('PesquisaForms.xlsx')

# Função para processar respostas múltiplas
def process_multiple_answers(series, separator=';'):
    all_answers = []
    for answer in series:
        if pd.isna(answer):
            continue
        parts = [part.strip() for part in str(answer).split(separator) if part.strip()]
        all_answers.extend(parts)
    return Counter(all_answers)

# Função para criar gráfico de pizza com contagem de respondentes
def create_pie_chart(frame, title, data_dict, total_respondents):
    fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
    fig.patch.set_facecolor('#f0f0f0')
    ax.set_facecolor('#f0f0f0')
    
    labels = list(data_dict.keys())
    sizes = list(data_dict.values())
    
    # Cores modernas
    colors = plt.cm.Pastel1(range(len(labels)))
    
    # Criar o gráfico de pizza
    wedges, texts, autotexts = ax.pie(
        sizes, 
        labels=labels, 
        colors=colors,
        autopct=lambda p: f'{p:.1f}%\n({int(round(p*sum(sizes)/100))})',
        startangle=90,
        wedgeprops={'edgecolor': 'white', 'linewidth': 1},
        textprops={'fontsize': 8}
    )
    
    # Adicionar título com número total de respondentes
    ax.set_title(f"{title}\nTotal de respondentes: {total_respondents}", 
                fontsize=10, pad=20, fontweight='bold', color='#333333')
    
    # Ajustar layout
    plt.tight_layout()
    
    # Adicionar ao Tkinter
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    return canvas

# Função principal para criar a interface
def main():
    root = tk.Tk()
    root.title("Análise de Pesquisa sobre Animais em Situação de Rua")
    root.geometry("1200x800")
    root.configure(bg='#f0f0f0')
    
    # Frame principal para centralização
    main_frame = ttk.Frame(root)
    main_frame.pack(fill='both', expand=True, padx=20, pady=20)
    
    # Estilo
    style = ttk.Style()
    style.configure('TFrame', background='#f0f0f0')
    style.configure('TNotebook', background='#f0f0f0')
    style.configure('TNotebook.Tab', font=('Helvetica', 10, 'bold'), padding=[10, 5])
    
    # Notebook (abas)
    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill='both', expand=True)
    
    # Função auxiliar para criar abas
    def create_tab(notebook, tab_title, column_name, chart_title, is_multiple=False):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=tab_title)
        
        # Encontrar a coluna correta (case insensitive)
        col_name = None
        for col in data.columns:
            if column_name.lower() in col.lower():
                col_name = col
                break
                
        if col_name is None:
            print(f"Coluna '{column_name}' não encontrada")
            return tab
            
        if is_multiple:
            chart_data = process_multiple_answers(data[col_name])
            total = len(data[col_name].dropna())
        else:
            chart_data = data[col_name].value_counts().to_dict()
            total = len(data[col_name].dropna())
            
        if chart_data:
            canvas = create_pie_chart(tab, chart_title, chart_data, total)
            canvas.get_tk_widget().pack(fill='both', expand=True, padx=10, pady=10)
            
        return tab
    
    # Criar todas as abas
    create_tab(notebook, "Frequência", 
              "Com que frequência você vê animais em situação de rua na sua vizinhança?", 
              "Frequência de Visualização")
    
    create_tab(notebook, "Envolvimento", 
              "Você já se envolveu em alguma ação para ajudar animais em situação de rua?", 
              "Envolvimento em Ações")
    
    create_tab(notebook, "Problemas", 
              "Quais problemas você percebe com os animais em situação de rua", 
              "Problemas Percebidos", True)
    
    create_tab(notebook, "Barreiras", 
              "Qual a principal barreira para você ajudar mais os animais em situação de rua", 
              "Principais Barreiras", True)
    
    create_tab(notebook, "ONGs", 
              "Você conhece ONGs ou instituições próximas que resgatam animais em situação de rua", 
              "Conhecimento de ONGs")
    
    create_tab(notebook, "Interesse", 
              "Se existisse uma plataforma que conecta animais em situação de rua a ONGs locais", 
              "Interesse na Plataforma")
    
    create_tab(notebook, "Acesso", 
              "Qual seria a maneira mais prática para você acessar essa plataforma", 
              "Forma de Acesso Preferida")
    
    create_tab(notebook, "Confiança", 
              "Quais características fariam você confiar mais na plataforma", 
              "Características de Confiança", True)
    
    create_tab(notebook, "Informações", 
              "Quais informações você estaria disposto(a) a compartilhar", 
              "Informações Compartilhadas", True)
    
    root.mainloop()

if __name__ == "__main__":
    main()