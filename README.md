# analisis_caracteristicas_audio
Utilidad que recorre los archivos de audio de una carpeta y genera un excel con las caracteristicas de cada archivo analizado como durecion, frecuencia de muestreo, canales, bitrate y codec

⚖️ Contexto de uso

Esta herramienta ha sido desarrollada para su uso en informes periciales dentro del ámbito del análisis forense de audio.

Permite documentar de forma objetiva parámetros técnicos como:

Duración
Frecuencia de muestreo
Bitrate
Codec
Canales
Fecha de codificación (normalizada a UTC)
Nivel RMS (mono o por canal)

🚀 Funcionalidades
Análisis masivo de archivos de audio
Compatibilidad con múltiples formatos (mp3, wav, flac, m4a, etc.)
Extracción de metadatos mediante ffprobe
Cálculo de niveles RMS mediante ffmpeg
Exportación automática a Excel
Guardado progresivo del informe (robusto ante interrupciones)

🧰 Requisitos
🔹 Python
Python 3.8 o superior
🔹 Dependencias Python

Instalar desde requirements.txt:

pip install -r requirements.txt
🔹 FFmpeg (OBLIGATORIO)

Esta herramienta requiere que ffmpeg y ffprobe estén instalados en el sistema.

👉 Son necesarios para:

Lectura de metadatos (ffprobe)
Cálculo de niveles RMS (ffmpeg)
⚙️ Instalación de FFmpeg
🪟 Windows
Descargar desde:
https://ffmpeg.org/download.html
Descomprimir y añadir la carpeta bin al PATH del sistema
Verificar en consola:
ffmpeg -version
ffprobe -version
🐧 Linux (Ubuntu/Debian)
sudo apt update
sudo apt install ffmpeg
🍎 macOS (con Homebrew)
brew install ffmpeg

▶️ Uso

Ejecutar desde la carpeta donde están los audios:

python analisis_caracteristicas_audio.py "*"

🔍 Ejemplos
Analizar todos los audios del directorio actual
python analisis_caracteristicas_audio.py "*"
Analizar solo archivos que contengan un patrón
python analisis_caracteristicas_audio.py "*COMPLETO*"
Incluir subcarpetas
python analisis_caracteristicas_audio.py "*" -r
Especificar nombre del archivo de salida
python analisis_caracteristicas_audio.py "*" -o informe.xlsx

📊 Salida

Se genera un archivo Excel con:

Nombre del archivo
Fecha de codificación (UTC)
Duración
Frecuencia de muestreo
Canales
Bitrate
Codec
RMS (mono / canal izquierdo / canal derecho)

⚠️ Limitaciones
Depende de la disponibilidad y calidad de los metadatos del archivo
El cálculo RMS depende del comportamiento interno de ffmpeg
Puede variar ligeramente según el códec o la decodificación
No sustituye el análisis pericial experto

👨‍⚖️ Autor

DAVID SOTO ALVAREZ
Perito Informatico Forense
peritacionesinformaticas.es@gmail.com 
Responsable de Andalucia de la ANTPJI (Asociacion Nacional de Tasadores y Peritos Judiciales Informaticos)
Spain
