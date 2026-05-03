# Event Study Dashboard

Dashboard interactivo en Python/Streamlit para comparar el comportamiento histórico de activos financieros alrededor de fechas de eventos clave.

---

## Instalación rápida

```bash
pip install -r requirements.txt
streamlit run app.py
```

El dashboard se abre automáticamente en `http://localhost:8501`.

---

## Integración con Bloomberg

### Instalar blpapi
```bash
pip install --index-url=https://bcms.bloomberg.com/pip/simple/ blpapi
```
Bloomberg Terminal debe estar abierto en el mismo equipo. El dashboard mostrará ✅ cuando detecte blpapi.

Cuando se ejecuta localmente con Bloomberg disponible, los datos descargados se **guardan automáticamente** en `data/Data_historica.xlsx` para que sirvan como archivo default del repo. Súbelo a GitHub para que otros usuarios puedan usar las categorías sin Bloomberg.

---

## Formato CSV/Excel

| Date       | SPX Index | USGG10YR Index | USDMXN Curncy |
|------------|-----------|----------------|---------------|
| 2020-01-02 | 3257.85   | 1.88           | 18.87         |

- Primera columna: Fechas (YYYY-MM-DD u otro formato estándar)
- Columnas siguientes: una serie por columna, **nombre del ticker Bloomberg como encabezado** (ej. `SPX Index`, `GT10 Govt`)
- Formatos: `.csv`, `.xlsx`, `.xls`

Un CSV de muestra con datos sintéticos 2015-2024 está en `data/sample_data.csv`.

---

## Reporte PDF — modos de generación

El dashboard soporta **dos modos de PDF**, detectados automáticamente:

| Entorno | Modo | Requiere |
|---------|------|----------|
| 🖥️ Local con LaTeX | PDF *completo* (portada, índice, descripciones, gráficas) | `pdflatex` instalado |
| ☁️ Streamlit Cloud / sin LaTeX | PDF *simple* (una gráfica por página, sin texto) | Solo matplotlib |

Si `pdflatex` está en el `PATH`, la app usa el modo completo. Si no, cae al modo simple — funciona sin instalar nada extra.

### Instalar LaTeX para el PDF completo (opcional, solo local)

**macOS** (recomendado MacTeX, ~5 GB; o BasicTeX, ~100 MB):
```bash
# Opción ligera
brew install --cask basictex
sudo tlmgr update --self
sudo tlmgr install babel-spanish fancyhdr titlesec parskip

# Opción completa
brew install --cask mactex
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install texlive-latex-base texlive-latex-recommended \
                 texlive-latex-extra texlive-fonts-recommended \
                 texlive-lang-spanish
```

**Windows:** Descarga MiKTeX desde [miktex.org/download](https://miktex.org/download). Durante la instalación marca "Install missing packages on the fly = Yes". Reinicia la terminal después de instalar.

Verifica la instalación:
```bash
pdflatex --version
```

Si `pdflatex` aparece, la app generará el PDF completo automáticamente la próxima vez.

---

## Carpeta `data/`

Esta carpeta versiona archivos Excel/CSV que sirven como datos default. El usuario activa el checkbox **📦 Usar archivo default del repo** en el sidebar para cargarlos sin tener que subir nada.

El archivo default que se sobrescribe automáticamente al correr Bloomberg localmente está definido en `app.py`:
```python
DEFAULT_DATA_FILENAME = "Data_historica.xlsx"
```

Cambia esta constante si prefieres otro nombre de archivo.

---

## Estructura del proyecto

```
├── app.py                # Aplicación principal Streamlit
├── generate_report.py    # Generador de PDF (LaTeX o simple)
├── requirements.txt      # Dependencias Python
├── README.md
└── data/
    ├── Data_historica.xlsx       # Archivo default (auto-actualizado vía Bloomberg)
    ├── sample_data.csv           # Datos de ejemplo
    └── ...                       # Otros archivos históricos
```
