import os
import time
import base64
import requests
from supabase import create_client, Client
from openai import OpenAI
from langchain_text_splitters import RecursiveCharacterTextSplitter

# ----------------------------------------------------
# 1. CONFIGURACIÓN E INICIALIZACIÓN (Usando Variables de Entorno de Railway)
# ----------------------------------------------------

# Inicializar clientes
# Estas variables se configuran directamente en Railway como "Secrets"
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

# Configuración de LangChain para el 'chunking'
text_splitter = RecursiveCharacterTextSplitter(
    chunk_size=1000,
    chunk_overlap=200,
    separators=["\n\n", "\n", ".", "!", "?", " ", ""],
)

# ----------------------------------------------------
# 2. FUNCIONES DE PROCESAMIENTO
# ----------------------------------------------------

def encode_image(image_url: str):
    """Descarga una imagen desde Supabase Storage y la codifica a Base64."""
    try:
        # Nota: La URL debe ser accesible (pública o usando la clave de servicio en la petición)
        response = requests.get(image_url, timeout=10)
        response.raise_for_status()  # Lanza un error para códigos de estado HTTP malos

        # Codificar binario a Base64
        return base64.b64encode(response.content).decode('utf-8')
    except Exception as e:
        print(f"Error al descargar o codificar imagen {image_url}: {e}")
        return None

def analyze_and_get_description(image_base64: str, file_type: str) -> str:
    """Usa GPT-4o para obtener una descripción textual de la imagen."""
    
    # Adaptar el prompt para el análisis de imágenes
    prompt = (
        "Actúa como un analista experto de inteligencia de negocios en sector minero. Describe concisamente la imagen. "
        "Identifica cualquier texto relevante, avance de proyecto (si aplica), o problema visible. "
        "El objetivo es convertir la imagen en contexto textual para un reporte ejecutivo. Máximo 50 palabras."
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
            max_tokens=100,
        )
        return response.choices[0].message.content.strip()

    except Exception as e:
        print(f"Error en la API de OpenAI para la imagen: {e}")
        return "ERROR: No se pudo generar descripción de la imagen."

def create_and_upload_embedding(content: str, record_id: int):
    """Genera el embedding y actualiza el registro en Supabase."""
    
    try:
        # 1. Generar Embedding
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
# 3. LÓGICA PRINCIPAL DEL PROCESO
# ----------------------------------------------------

def main_processor():
    """Procesa mensajes pendientes de vectorización."""
    print("--- 🚀 Iniciando Proceso de Vectorización y OCR ---")

    try:
        # 1. Buscar registros sin vectorizar
        query_response = supabase.from_('mensajes_analisis').select("*").is_('embedding', 'null').order('fecha_hora', desc=False).limit(50).execute()
        
        pending_records = query_response.data if query_response.data else []

        if not pending_records:
            print("✅ No hay nuevos registros para procesar.")
            return

        print(f"🔎 Encontrados {len(pending_records)} registros pendientes.")

        # 2. Procesar cada registro
        for record in pending_records:
            record_id = record.get('id')
            final_content = record.get('contenido_texto', '') or ""

            # A. Si es una imagen, hacer OCR Multimodal
            if record.get('es_imagen') and record.get('url_storage'):
                print(f"   [ID {record_id}] Procesando imagen...")
                
                file_type = "image/jpeg"  # Asumir JPEG o inferir del nombre/metadata
                base64_img = encode_image(record['url_storage'])
                
                if base64_img:
                    description = analyze_and_get_description(base64_img, file_type)
                    final_content = f"{final_content}\n[ANÁLISIS DE IMAGEN]: {description}"
                    print(f"   [ID {record_id}] Descripción: {description[:40]}...")

            # B. Procesamiento de Texto y Vectorización
            if final_content:
                create_and_upload_embedding(final_content, record_id)
            else:
                print(f"   [ID {record_id}] Contenido vacío. Saltando.")

    except Exception as e:
        print(f"❌ Error en main_processor: {e}")

# ----------------------------------------------------
# 4. PUNTO DE ENTRADA
# ----------------------------------------------------

if __name__ == "__main__":
    print("🔧 Servicio de Vectorización iniciado")
    print(f"🌐 Conectado a Supabase: {SUPABASE_URL}")
    
    # Bucle continuo para servicio 24/7
    while True:
        try:
            main_processor()
            print("😴 Durmiendo 30 segundos antes de la siguiente búsqueda...")
            time.sleep(30)
        except KeyboardInterrupt:
            print("\n👋 Servicio detenido por el usuario")
            break
        except Exception as e:
            print(f"❌ Error crítico: {e}")
            print("⏰ Esperando 60 segundos antes de reintentar...")
            time.sleep(60)
