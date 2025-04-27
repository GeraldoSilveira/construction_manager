import pandas as pd
import os
from tkinter import messagebox

# Função para importar cronograma do MS Project (CSV)
def import_ms_project_schedule(filename):
    try:
        if not os.path.exists(filename):
            raise FileNotFoundError("Arquivo de cronograma não encontrado.")
        
        df = pd.read_csv(filename)
        required_columns = ["Task Name", "Start", "Finish", "Duration"]
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Arquivo deve conter as colunas: Task Name, Start, Finish, Duration")
        
        return df.to_dict('records')
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao importar cronograma: {str(e)}")
        return []

# Função para exportar atividades para MS Project (CSV)
def export_to_ms_project(activities, filename="ms_project_export.csv"):
    if not activities:
        messagebox.showinfo("Informação", "Nenhuma atividade para exportar.")
        return

    df = pd.DataFrame(activities)
    df = df[["Descrição", "Data", "Responsável", "Status", "Custo"]]
    df.rename(columns={
        "Descrição": "Task Name",
        "Data": "Start",
        "Responsável": "Resource Names",
        "Status": "Notes",
        "Custo": "Cost"
    }, inplace=True)
    
    df.to_csv(filename, index=False)
    messagebox.showinfo("Sucesso", f"Dados exportados para {filename}")
