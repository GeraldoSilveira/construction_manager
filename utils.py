import os
import io
from PIL import Image as PilImage
from datetime import datetime

# Diretório para armazenar fotos
PHOTO_DIR = "photos"
if not os.path.exists(PHOTO_DIR):
    os.makedirs(PHOTO_DIR)

# Função para otimizar e salvar foto (máximo 500 KB)
def optimize_photo(filename):
    try:
        img = PilImage.open(filename)
        # Redimensionar para um tamanho máximo (ex.: 1024x1024)
        max_size = (1024, 1024)
        img.thumbnail(max_size, PilImage.Resampling.LANCZOS)

        # Salvar com qualidade ajustada até que o tamanho seja <= 500 KB
        output = io.BytesIO()
        quality = 95
        while True:
            output.seek(0)
            output.truncate(0)
            img.save(output, format="JPEG", quality=quality)
            size = output.tell()
            if size <= 500 * 1024 or quality <= 10:  # 500 KB ou qualidade mínima atingida
                break
            quality -= 5

        # Salvar a foto otimizada no diretório photos
        photo_name = os.path.basename(filename)
        new_photo_path = os.path.join(PHOTO_DIR, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{photo_name}")
        with open(new_photo_path, "wb") as f:
            f.write(output.getvalue())
        
        return new_photo_path
    except Exception as e:
        raise Exception(f"Erro ao otimizar foto: {str(e)}")

# Função para validar entradas
def validate_inputs(description, responsible, cost, date):
    errors = []
    if len(description.strip()) < 5:
        errors.append("Descrição deve ter no mínimo 5 caracteres.")
    if len(responsible.strip()) < 3:
        errors.append("Responsável deve ter no mínimo 3 caracteres.")
    try:
        float(cost)
        if float(cost) < 0:
            errors.append("Custo não pode ser negativo.")
    except ValueError:
        errors.append("Custo deve ser um número válido.")
    try:
        datetime.strptime(date, '%d/%m/%Y')
    except ValueError:
        errors.append("Data deve estar no formato dd/mm/aaaa.")
    
    return errors
