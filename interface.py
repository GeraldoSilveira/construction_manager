import json
import os
from tkinter import Tk, ttk, messagebox, filedialog, Toplevel
from tkinter import Label, Entry, Button, Listbox, Scrollbar, Text, PhotoImage
from datetime import datetime
from PIL import Image, ImageTk
from schedule import import_ms_project_schedule, export_to_ms_project
from reports import generate_excel_report, generate_pdf_report
from utils import optimize_photo, validate_inputs

class ConstructionManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador de Obras")
        self.activities = []
        self.schedule = []
        self.current_photo = None
        self.load_activities()  # Carregar atividades do JSON ao iniciar

        # Estilo ttk para um visual mais moderno
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0", font=("Helvetica", 10))
        style.configure("TEntry", font=("Helvetica", 10))
        style.configure("TButton", font=("Helvetica", 10))
        style.configure("TCombobox", font=("Helvetica", 10))

        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Importar cronograma MS Project
        ttk.Label(main_frame, text="Importar Cronograma MS Project (CSV):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Button(main_frame, text="Selecionar Arquivo", command=self.import_schedule).grid(row=0, column=1, pady=2)

        # Descrição da Atividade
        ttk.Label(main_frame, text="Descrição da Atividade:").grid(row=1, column=0, sticky="w", pady=2)
        self.description_var = ttk.Entry(main_frame, width=30)
        self.description_var.grid(row=1, column=1, pady=2)

        # Custo
        ttk.Label(main_frame, text="Custo (R$):").grid(row=2, column=0, sticky="w", pady=2)
        self.cost_var = ttk.Entry(main_frame, width=10)
        self.cost_var.grid(row=2, column=1, pady=2)

        # Data
        ttk.Label(main_frame, text="Data (dd/mm/aaaa):").grid(row=3, column=0, sticky="w", pady=2)
        self.date_var = ttk.Entry(main_frame, width=12)
        self.date_var.insert(0, datetime.now().strftime('%d/%m/%Y'))
        self.date_var.grid(row=3, column=1, pady=2)

        # Responsável
        ttk.Label(main_frame, text="Responsável:").grid(row=4, column=0, sticky="w", pady=2)
        self.responsible_var = ttk.Entry(main_frame, width=20)
        self.responsible_var.grid(row=4, column=1, pady=2)

        # Status
        ttk.Label(main_frame, text="Status:").grid(row=5, column=0, sticky="w", pady=2)
        self.status_combobox = ttk.Combobox(main_frame, state="readonly", width=15, values=["Em Andamento", "Concluído", "Atrasado"])
        self.status_combobox.set("Em Andamento")  # Valor padrão
        self.status_combobox.grid(row=5, column=1, pady=2)

        # Observações
        ttk.Label(main_frame, text="Observações:").grid(row=6, column=0, sticky="w", pady=2)
        self.notes_text = Text(main_frame, height=3, width=30, font=("Helvetica", 10))
        self.notes_text.grid(row=6, column=1, pady=2)

        # Adicionar Foto
        ttk.Label(main_frame, text="Adicionar Foto:").grid(row=7, column=0, sticky="w", pady=2)
        self.photo_path_var = ""
        ttk.Button(main_frame, text="Selecionar Foto", command=self.add_photo).grid(row=7, column=1, pady=2)

        # Botões para adicionar/editar/excluir
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text="Adicionar Atividade", command=self.add_activity).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Editar Atividade", command=self.edit_activity).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Excluir Atividade", command=self.delete_activity).grid(row=0, column=2, padx=5)

        # Lista de atividades
        ttk.Label(main_frame, text="Atividades Registradas:").grid(row=9, column=0, sticky="w", pady=2)
        self.activities_listbox = Listbox(main_frame, width=50, height=10, font=("Helvetica", 10))
        self.activities_listbox.grid(row=10, column=0, columnspan=2, pady=2)
        scrollbar = Scrollbar(main_frame, orient="vertical")
        scrollbar.grid(row=10, column=2, sticky="ns")
        self.activities_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.activities_listbox.yview)

        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, length=200, mode='determinate')
        self.progress.grid(row=11, column=0, columnspan=2, pady=5)
        self.progress.grid_remove()  # Esconder inicialmente

        # Botões para gerar relatórios e exportar
        report_frame = ttk.Frame(main_frame)
        report_frame.grid(row=12, column=0, columnspan=2, pady=10)
        ttk.Button(report_frame, text="Gerar Relatório Excel", command=self.generate_excel).grid(row=0, column=0, padx=5)
        ttk.Button(report_frame, text="Gerar Relatório PDF", command=self.generate_pdf).grid(row=0, column=1, padx=5)
        ttk.Button(report_frame, text="Exportar para MS Project", command=self.export_to_ms_project).grid(row=0, column=2, padx=5)

        # Carregar atividades na Listbox
        self.update_listbox()

    def load_activities(self):
        # Carregar atividades do arquivo JSON
        try:
            if os.path.exists("activities.json"):
                with open("activities.json", "r") as f:
                    self.activities = json.load(f)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar atividades: {str(e)}")

    def save_activities(self):
        # Salvar atividades no arquivo JSON
        try:
            with open("activities.json", "w") as f:
                json.dump(self.activities, f, indent=4)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar atividades: {str(e)}")

    def update_listbox(self):
        # Atualizar a Listbox com as atividades
        self.activities_listbox.delete(0, "end")
        for activity in self.activities:
            self.activities_listbox.insert("end", f"{activity['Data']} - {activity['Descrição']} - {activity['Responsável']} - {activity['Status']} - R${activity['Custo']:.2f}")

    def import_schedule(self):
        filename = filedialog.askopenfilename(
            title="Selecionar cronograma MS Project",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.schedule = import_ms_project_schedule(filename)
            if self.schedule:
                messagebox.showinfo("Sucesso", "Cronograma importado com sucesso!")

    def add_photo(self):
        filename = filedialog.askopenfilename(
            title="Selecionar foto",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            try:
                # Otimizar e salvar a foto
                self.photo_path_var = optimize_photo(filename)

                # Mostrar preview da foto
                img = Image.open(self.photo_path_var)
                img.thumbnail((200, 200))
                photo = ImageTk.PhotoImage(img)
                preview_window = Toplevel(self.root)
                preview_window.title("Preview da Foto")
                Label(preview_window, image=photo).pack(pady=10)
                self.current_photo = photo  # Manter referência para evitar garbage collection
                messagebox.showinfo("Sucesso", "Foto otimizada e adicionada com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

    def add_activity(self):
        description = self.description_var.get()
        cost = self.cost_var.get()
        date = self.date_var.get()
        responsible = self.responsible_var.get()
        status = self.status_combobox.get()
        notes = self.notes_text.get("1.0", "end-1c")
        if not notes:
            notes = "Sem observações"
        photo_path = self.photo_path_var if self.photo_path_var else ""

        # Validar entradas
        errors = validate_inputs(description, responsible, cost, date)
        if errors:
            messagebox.showwarning("Aviso", "\n".join(errors))
            return

        if not status:
            messagebox.showwarning("Aviso", "Selecione o status da atividade.")
            return

        activity = {
            "Data": date,
            "Descrição": description,
            "Responsável": responsible,
            "Status": status,
            "Observações": notes,
            "Custo": float(cost),
            "Foto": photo_path
        }
        self.activities.append(activity)
        self.save_activities()
        self.update_listbox()
        self.clear_fields()

    def edit_activity(self):
        selected = self.activities_listbox.curselection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma atividade para editar.")
            return

        index = selected[0]
        activity = self.activities[index]

        # Preencher os campos com os dados da atividade
        self.description_var.delete(0, "end")
        self.description_var.insert(0, activity["Descrição"])
        self.cost_var.delete(0, "end")
        self.cost_var.insert(0, str(activity["Custo"]))
        self.date_var.delete(0, "end")
        self.date_var.insert(0, activity["Data"])
        self.responsible_var.delete(0, "end")
        self.responsible_var.insert(0, activity["Responsável"])
        self.status_combobox.set(activity["Status"])
        self.notes_text.delete("1.0", "end")
        self.notes_text.insert("1.0", activity["Observações"])
        self.photo_path_var = activity["Foto"]

        # Remover a atividade antiga
        self.activities.pop(index)
        self.update_listbox()

    def delete_activity(self):
        selected = self.activities_listbox.curselection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma atividade para excluir.")
            return

        index = selected[0]
        if messagebox.askyesno("Confirmação", "Deseja excluir esta atividade?"):
            self.activities.pop(index)
            self.save_activities()
            self.update_listbox()

    def clear_fields(self):
        self.description_var.delete(0, "end")
        self.cost_var.delete(0, "end")
        self.date_var.delete(0, "end")
        self.date_var.insert(0, datetime.now().strftime('%d/%m/%Y'))
        self.responsible_var.delete(0, "end")
        self.status_combobox.set("Em Andamento")
        self.notes_text.delete("1.0", "end")
        self.photo_path_var = ""
        self.current_photo = None

    def generate_excel(self):
        self.progress.grid()
        self.progress.start(10)
        generate_excel_report(self.activities)
        self.progress.stop()
        self.progress.grid_remove()

    def generate_pdf(self):
        self.progress.grid()
        self.progress.start(10)
        generate_pdf_report(self.activities)
        self.progress.stop()
        self.progress.grid_remove()

    def export_to_ms_project(self):
        self.progress.grid()
        self.progress.start(10)
        export_to_ms_project(self.activities)
        self.progress.stop()
        self.progress.grid_remove()
