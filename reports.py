import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
import matplotlib.pyplot as plt
import io
import os
from PIL import Image as PilImage
from tkinter import messagebox

# Função para gerar gráfico de pizza
def generate_pie_chart(activities):
    if not activities:
        return None
    
    labels = [activity["Descrição"] for activity in activities]
    sizes = [activity["Custo"] for activity in activities]
    
    plt.figure(figsize=(6, 4))
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
    plt.title("Distribuição de Custos por Atividade")
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close()
    buf.seek(0)
    return buf

# Função para gerar relatório em Excel
def generate_excel_report(activities, filename="daily_report.xlsx"):
    if not activities:
        messagebox.showinfo("Informação", "Nenhuma atividade para gerar o relatório.")
        return

    df = pd.DataFrame(activities)
    df = df[["Data", "Descrição", "Responsável", "Status", "Observações", "Custo", "Foto"]]

    total_row = pd.DataFrame([{
        "Data": "",
        "Descrição": "TOTAL",
        "Responsável": "",
        "Status": "",
        "Observações": "",
        "Custo": df["Custo"].sum(),
        "Foto": ""
    }])
    df = pd.concat([df, total_row], ignore_index=True)

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Relatório Diário", index=False)
        worksheet = writer.sheets["Relatório Diário"]
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    messagebox.showinfo("Sucesso", f"Relatório Excel gerado: {filename}")

# Função para gerar relatório em PDF com fotos
def generate_pdf_report(activities, filename="daily_report.pdf"):
    if not activities:
        messagebox.showinfo("Informação", "Nenhuma atividade para gerar o relatório.")
        return

    doc = SimpleDocTemplate(filename, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    title = Paragraph(f"Relatório Diário de Obras - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Title'])
    elements.append(title)
    elements.append(Paragraph("<br/><br/>", styles['Normal']))

    data = [["Data", "Descrição", "Responsável", "Status", "Observações", "Custo (R$)"]]
    for activity in activities:
        data.append([
            activity["Data"],
            activity["Descrição"],
            activity["Responsável"],
            activity["Status"],
            activity["Observações"],
            f"{activity['Custo']:.2f}"
        ])

    total_cost = sum(activity["Custo"] for activity in activities)
    data.append(["", "TOTAL", "", "", "", f"{total_cost:.2f}"])

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
    ]))
    elements.append(table)
    elements.append(Paragraph("<br/><br/>", styles['Normal']))

    chart_buf = generate_pie_chart(activities)
    if chart_buf:
        chart_image = Image(chart_buf, width=300, height=200)
        elements.append(chart_image)
        elements.append(Paragraph("<br/><br/>", styles['Normal']))

    # Adicionar fotos
    for activity in activities:
        if activity["Foto"]:
            elements.append(Paragraph(f"Foto - {activity['Descrição']} ({activity['Data']})", styles['Heading2']))
            try:
                img = PilImage.open(activity["Foto"])
                img.thumbnail((200, 200))  # Redimensionar para caber no PDF
                temp_img_path = "temp_photo.png"
                img.save(temp_img_path, format="PNG")
                photo = Image(temp_img_path, width=200, height=200)
                elements.append(photo)
                os.remove(temp_img_path)
            except Exception as e:
                elements.append(Paragraph(f"Erro ao carregar foto: {str(e)}", styles['Normal']))
            elements.append(Paragraph("<br/><br/>", styles['Normal']))

    doc.build(elements)
    messagebox.showinfo("Sucesso", f"Relatório PDF gerado: {filename}")

