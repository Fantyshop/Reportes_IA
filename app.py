import os
import time
import base64
import requests
from urllib.parse import unquote
from io import BytesIO
from supabase import create_client, Client
from openai import OpenAI

# Bibliotecas para procesamiento de documentos
import PyPDF2
import pdfplumber
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

# ----------------------------------------------------
# 1. CONFIGURACIÓN E INICIALIZACIÓN
# ----------------------------------------------------

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
BUCKET_NAME = os.environ.get("SUPABASE_BUCKET", "whatsapp-media")

# Validar variables de entorno
if not all([SUPABASE_URL, SUPABASE_SERVICE_KEY, OPENAI_API_KEY]):
    raise ValueError("Faltan variables de entorno necesarias. Verifica SUPABASE_URL, SUPABASE_SERVICE_KEY y OPENAI_API_KEY")

# Inicializar clientes
supabase: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
openai_client = OpenAI(api_key=OPENAI_API_KEY)

# Tipos de archivos soportados
SUPPORTED_IMAGE_FORMATS = ['png', 'jpeg', 'jpg', 'gif', 'webp']
SUPPORTED_DOCUMENT_FORMATS = ['pdf', 'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt']

# ----------------------------------------------------
# 2. FUNCIONES DE UTILIDAD
# ----------------------------------------------------

def clean_url(url: str) -> str:
    """Limpia y decodifica la URL para evitar problemas de encoding."""
    try:
        return unquote(url)
    except Exception as e:
        print(f"⚠️ Error al limpiar URL: {e}")
        return url

def get_file_extension(url: str) -> str:
    """Extrae la extensión del archivo desde la URL."""
    url_lower = url.lower()
    for ext in SUPPORTED_IMAGE_FORMATS + SUPPORTED_DOCUMENT_FORMATS:
        if f".{ext}" in url_lower:
            return ext
    return None

def download_file(url: str) -> bytes:
    """Descarga un archivo desde una URL y devuelve su contenido en bytes."""
    try:
        clean_file_url = clean_url(url)
        response = requests.get(clean_file_url, timeout=30)
        response.raise_for_status()
        return response.content
    except Exception as e:
        print(f"❌ Error al descargar archivo {url}: {e}")
        return None

# ----------------------------------------------------
# 3. PROCESAMIENTO DE IMÁGENES
# ----------------------------------------------------

def get_image_mime_type(url: str) -> str:
    """Determina el tipo MIME de la imagen basado en la extensión."""
    url_lower = url.lower()
    
    if '.png' in url_lower:
        return 'image/png'
    elif '.jpg' in url_lower or '.jpeg' in url_lower:
        return 'image/jpeg'
    elif '.gif' in url_lower:
        return 'image/gif'
    elif '.webp' in url_lower:
        return 'image/webp'
    else:
        return 'image/jpeg'

def analyze_image_with_ai(image_base64: str, file_type: str) -> str:
    """Usa GPT-4o para obtener una descripción textual de la imagen."""
    
    prompt = (
        "Actúa como un analista experto de inteligencia de negocios en sector minero. "
        "Describe concisamente la imagen. Identifica cualquier texto relevante, "
        "avance de proyecto (si aplica), o problema visible. "
        "El objetivo es convertir la imagen en contexto textual para un reporte ejecutivo. "
        "Máximo 100 palabras."
    )

    try:
        response = openai_client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{file_type};base64,{image_base64}",
                            },
                        },
                    ],
                }
            ],
            max_tokens=200,
        )
        return response.choices[0].message.content.strip()

    except Exception as e:
        print(f"❌ Error en la API de OpenAI para la imagen: {e}")
        return None

def process_image(url: str) -> str:
    """Procesa una imagen y retorna su análisis textual."""
    try:
        file_content = download_file(url)
        if not file_content:
            return None
        
        # Codificar a Base64
        image_base64 = base64.b64encode(file_content).decode('utf-8')
        
        # Analizar con IA
        file_type = get_image_mime_type(url)
        description = analyze_image_with_ai(image_base64, file_type)
        
        return description
        
    except Exception as e:
        print(f"❌ Error procesando imagen: {e}")
        return None

# ----------------------------------------------------
# 4. PROCESAMIENTO DE PDFs
# ----------------------------------------------------

def extract_text_from_pdf(file_content: bytes) -> str:
    """Extrae texto de un archivo PDF."""
    text_parts = []
    
    try:
        # Método 1: PyPDF2 (más rápido)
        pdf_reader = PyPDF2.PdfReader(BytesIO(file_content))
        for page in pdf_reader.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
        
        # Si PyPDF2 no extrajo texto, intentar con pdfplumber (mejor para PDFs complejos)
        if not text_parts:
            with pdfplumber.open(BytesIO(file_content)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)
        
        return "\n\n".join(text_parts)
        
    except Exception as e:
        print(f"❌ Error extrayendo texto de PDF: {e}")
        return None

def process_pdf(url: str) -> str:
    """Procesa un PDF y retorna su contenido textual."""
    try:
        file_content = download_file(url)
        if not file_content:
            return None
        
        text = extract_text_from_pdf(file_content)
        
        if text and len(text.strip()) > 0:
            return f"[CONTENIDO PDF]: {text[:3000]}"  # Limitar a 3000 caracteres
        else:
            return "[PDF sin texto extraíble - posiblemente escaneado]"
            
    except Exception as e:
        print(f"❌ Error procesando PDF: {e}")
        return None

# ----------------------------------------------------
# 5. PROCESAMIENTO DE ARCHIVOS WORD
# ----------------------------------------------------

def extract_text_from_docx(file_content: bytes) -> str:
    """Extrae texto de un archivo Word (.docx)."""
    try:
        doc = Document(BytesIO(file_content))
        text_parts = []
        
        # Extraer párrafos
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                text_parts.append(paragraph.text)
        
        # Extraer tablas
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join([cell.text.strip() for cell in row.cells])
                if row_text.strip():
                    text_parts.append(row_text)
        
        return "\n".join(text_parts)
        
    except Exception as e:
        print(f"❌ Error extrayendo texto de DOCX: {e}")
        return None

def process_word(url: str) -> str:
    """Procesa un archivo Word y retorna su contenido textual."""
    try:
        file_content = download_file(url)
        if not file_content:
            return None
        
        text = extract_text_from_docx(file_content)
        
        if text and len(text.strip()) > 0:
            return f"[CONTENIDO WORD]: {text[:3000]}"
        else:
            return "[Documento Word vacío o sin contenido]"
            
    except Exception as e:
        print(f"❌ Error procesando Word: {e}")
        return None

# ----------------------------------------------------
# 6. PROCESAMIENTO DE ARCHIVOS EXCEL
# ----------------------------------------------------

def extract_text_from_xlsx(file_content: bytes) -> str:
    """Extrae texto de un archivo Excel (.xlsx)."""
    try:
        workbook = load_workbook(BytesIO(file_content), data_only=True)
        text_parts = []
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            text_parts.append(f"\n=== HOJA: {sheet_name} ===")
            
            # Extraer hasta 100 filas por hoja
            for row_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
                if row_idx > 100:  # Limitar filas
                    text_parts.append("[... contenido truncado ...]")
                    break
                
                row_text = " | ".join([str(cell) if cell is not None else "" for cell in row])
                if row_text.strip():
                    text_parts.append(row_text)
        
        return "\n".join(text_parts)
        
    except Exception as e:
        print(f"❌ Error extrayendo texto de XLSX: {e}")
        return None

def process_excel(url: str) -> str:
    """Procesa un archivo Excel y retorna su contenido textual."""
    try:
        file_content = download_file(url)
        if not file_content:
            return None
        
        text = extract_text_from_xlsx(file_content)
        
        if text and len(text.strip()) > 0:
            return f"[CONTENIDO EXCEL]: {text[:3000]}"
        else:
            return "[Archivo Excel vacío o sin contenido]"
            
    except Exception as e:
        print(f"❌ Error procesando Excel: {e}")
        return None

# ----------------------------------------------------
# 7. PROCESAMIENTO DE ARCHIVOS POWERPOINT
# ----------------------------------------------------

def extract_text_from_pptx(file_content: bytes) -> str:
    """Extrae texto de un archivo PowerPoint (.pptx)."""
    try:
        presentation = Presentation(BytesIO(file_content))
        text_parts = []
        
        for slide_idx, slide in enumerate(presentation.slides, 1):
            text_parts.append(f"\n=== DIAPOSITIVA {slide_idx} ===")
            
            # Extraer texto de todas las formas
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    text_parts.append(shape.text)
        
        return "\n".join(text_parts)
        
    except Exception as e:
        print(f"❌ Error extrayendo texto de PPTX: {e}")
        return None

def process_powerpoint(url: str) -> str:
    """Procesa un archivo PowerPoint y retorna su contenido textual."""
    try:
        file_content = download_file(url)
        if not file_content:
            return None
        
        text = extract_text_from_pptx(file_content)
        
        if text and len(text.strip()) > 0:
            return f"[CONTENIDO POWERPOINT]: {text[:3000]}"
        else:
            return "[Presentación vacía o sin contenido]"
            
    except Exception as e:
        print(f"❌ Error procesando PowerPoint: {e}")
        return None

# ----------------------------------------------------
# 8. PROCESADOR UNIVERSAL DE ARCHIVOS
# ----------------------------------------------------

def process_file(url: str, file_extension: str) -> str:
    """Procesa cualquier tipo de archivo soportado y retorna su contenido."""
    
    print(f"   📄 Tipo de archivo detectado: .{file_extension}")
    
    # Imágenes
    if file_extension in SUPPORTED_IMAGE_FORMATS:
        return process_image(url)
    
    # PDFs
    elif file_extension == 'pdf':
        return process_pdf(url)
    
    # Word
    elif file_extension in ['docx', 'doc']:
        return process_word(url)
    
    # Excel
    elif file_extension in ['xlsx', 'xls']:
        return process_excel(url)
    
    # PowerPoint
    elif file_extension in ['pptx', 'ppt']:
        return process_powerpoint(url)
    
    else:
        print(f"   ⚠️ Formato no soportado: .{file_extension}")
        return None

# ----------------------------------------------------
# 9. GENERACIÓN DE EMBEDDINGS
# ----------------------------------------------------

def create_and_upload_embedding(content: str, record_id: int):
    """Genera el embedding y actualiza el registro en Supabase."""
    
    try:
        # 1. Generar Embedding
        print(f"   [ID {record_id}] Generando embedding...")
        embedding_response = openai_client.embeddings.create(
            input=content,
            model="text-embedding-3-small"
        )
        embedding_vector = embedding_response.data[0].embedding

        # 2. Actualizar Supabase
        update_response = supabase.from_('mensajes_analisis').update({
            'embedding': embedding_vector,
            'procesado_ia': True
        }).eq('id', record_id).execute()

        if update_response.data:
            print(f"✔️ Actualizado ID {record_id} con embedding.")
            return True
        else:
            print(f"❌ Error al actualizar ID {record_id}.")
            return False
            
    except Exception as e:
        print(f"❌ Error procesando ID {record_id}: {e}")
        return False

# ----------------------------------------------------
# 10. LÓGICA PRINCIPAL
# ----------------------------------------------------

def main_processor():
    """Procesa mensajes pendientes de vectorización."""
    print("\n--- 🚀 Iniciando Proceso de Vectorización y Análisis de Documentos ---")

    try:
        # 1. Buscar registros sin vectorizar
        query_response = supabase.from_('mensajes_analisis').select("*").is_('embedding', 'null').order('fecha_hora', desc=False).limit(50).execute()
        
        pending_records = query_response.data if query_response.data else []

        if not pending_records:
            print("✅ No hay nuevos registros para procesar.")
            return

        print(f"🔎 Encontrados {len(pending_records)} registros pendientes.")

        # 2. Procesar cada registro
        processed_count = 0
        error_count = 0
        skipped_count = 0
        
        for record in pending_records:
            record_id = record.get('id')
            final_content = record.get('contenido_texto', '') or ""

            print(f"\n📝 Procesando ID {record_id}...")

            # A. Si tiene archivo adjunto (imagen, PDF, Office)
            if record.get('url_storage'):
                file_url = record['url_storage']
                file_extension = get_file_extension(file_url)
                
                if file_extension:
                    print(f"   [ID {record_id}] 📎 Procesando archivo adjunto...")
                    
                    # Procesar archivo según su tipo
                    file_content = process_file(file_url, file_extension)
                    
                    if file_content:
                        final_content = f"{final_content}\n\n{file_content}"
                        print(f"   [ID {record_id}] ✅ Archivo procesado exitosamente")
                    else:
                        print(f"   [ID {record_id}] ⚠️ No se pudo procesar el archivo")
                else:
                    print(f"   [ID {record_id}] ⚠️ Tipo de archivo no reconocido")

            # B. Vectorización del contenido final
            if final_content.strip():
                success = create_and_upload_embedding(final_content, record_id)
                if success:
                    processed_count += 1
                else:
                    error_count += 1
            else:
                print(f"   [ID {record_id}] ⚠️ Contenido vacío. Saltando.")
                skipped_count += 1

        print(f"\n📊 Resumen: {processed_count} procesados, {error_count} errores, {skipped_count} saltados")

    except Exception as e:
        print(f"❌ Error en main_processor: {e}")
        import traceback
        traceback.print_exc()

# ----------------------------------------------------
# 11. PUNTO DE ENTRADA
# ----------------------------------------------------

if __name__ == "__main__":
    print("="*70)
    print("🔧 Servicio de Vectorización y Análisis Multi-Formato iniciado")
    print(f"🌐 Conectado a Supabase: {SUPABASE_URL}")
    print(f"📁 Bucket: {BUCKET_NAME}")
    print(f"📄 Formatos soportados:")
    print(f"   • Imágenes: {', '.join(SUPPORTED_IMAGE_FORMATS)}")
    print(f"   • Documentos: {', '.join(SUPPORTED_DOCUMENT_FORMATS)}")
    print("="*70)
    
    # Bucle continuo para servicio 24/7
    cycle_count = 0
    while True:
        try:
            cycle_count += 1
            print(f"\n🔄 Ciclo #{cycle_count} - {time.strftime('%Y-%m-%d %H:%M:%S')}")
            main_processor()
            print(f"\n😴 Durmiendo 30 segundos antes de la siguiente búsqueda...")
            time.sleep(30)
        except KeyboardInterrupt:
            print("\n👋 Servicio detenido por el usuario")
            break
        except Exception as e:
            print(f"❌ Error crítico: {e}")
            import traceback
            traceback.print_exc()
            print("⏰ Esperando 60 segundos antes de reintentar...")
            time.sleep(60)
