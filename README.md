# 📄 Imágenes a Word

Aplicación web para convertir múltiples imágenes a un documento Word (.docx) con un clic.

## ✨ Características

- Interfaz web moderna con drag & drop
- Soporta JPG, PNG, BMP, GIF, TIFF, WebP
- Las imágenes se ordenan alfabéticamente
- Cada imagen ocupa una página completa
- Márgenes optimizados para maximizar el espacio

---

## 🖥️ Instalación en Windows

### Requisitos
- Python 3.8 o superior ([Descargar Python](https://www.python.org/downloads/))
  - **IMPORTANTE:** Durante la instalación marca ✅ "Add Python to PATH"

### Pasos

1. **Descarga** o clona este repositorio

2. **Ejecuta el instalador** haciendo doble clic en:
   ```
   instalar.bat
   ```

3. **Inicia la aplicación** con:
   ```
   Imagenes a Word.bat
   ```

4. Se abrirá tu navegador en `http://localhost:5001`

---

## 🍎 Instalación en macOS

### Requisitos
- Python 3.8 o superior

### Pasos

1. **Clona** el repositorio:
   ```bash
   git clone https://github.com/MiguelSernaMon/Images-To-Word.git
   cd Images-To-Word
   ```

2. **Crea el entorno virtual e instala dependencias:**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

3. **Ejecuta la aplicación** con doble clic en:
   - `Imagenes a Word.command`
   - O `Imagenes a Word.app`

   > Primera vez: Si macOS lo bloquea, clic derecho → "Abrir"

---

## 🚀 Uso

1. Abre la aplicación (se abre el navegador automáticamente)
2. Arrastra tus imágenes o haz clic para seleccionarlas
3. Presiona **"Convertir a Word"**
4. El documento se descarga automáticamente

---

## 📁 Estructura del proyecto

```
├── app.py                    # Servidor Flask
├── images_to_word.py         # Script original (CLI)
├── templates/
│   └── index.html            # Interfaz web
├── requirements.txt          # Dependencias Python
├── instalar.bat              # Instalador Windows
├── Imagenes a Word.bat       # Ejecutable Windows
├── Imagenes a Word.command   # Ejecutable macOS
└── Imagenes a Word.app/      # App macOS
```

---

## 🛠️ Ejecución manual

Si prefieres ejecutar desde terminal:

```bash
# Activar entorno virtual
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# Ejecutar
python app.py
```

Luego abre `http://localhost:5001` en tu navegador.

---

## 📝 Licencia

MIT License - Usa libremente este proyecto.
