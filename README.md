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

---

## Formato CSV/Excel

| Date       | SPX Index | USGG10YR Index | USDMXN Curncy |
|------------|-----------|----------------|---------------|
| 2020-01-02 | 3257.85   | 1.88           | 18.87         |

- Primera columna: Fechas (YYYY-MM-DD u otro formato estándar)
- Columnas siguientes: una serie por columna, nombre del ticker como encabezado
- Formatos: `.csv`, `.xlsx`, `.xls`

Un CSV de muestra con datos sintéticos 2015-2024 está en `data/sample_data.csv`.

---

## Estructura del proyecto

```
├── app.py              # Aplicación principal
├── requirements.txt    # Dependencias
├── README.md
└── data/
    └── sample_data.csv # Datos de ejemplo
```
