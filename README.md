# Açaí Prime - Dashboard de Ventas y Comisiones

Este proyecto limpia, consolida y analiza las transacciones de ventas y productos de **Açaí Prime**, consolidando duplicados y generando un dashboard interactivo en HTML de alto impacto visual y funcional.

## Estructura del Proyecto

- `index.html`: El dashboard final interactivo y responsive (diseñado bajo la estética de [acaiprime.cl](https://acaiprime.cl/)).
- `generate_html_report.py`: Script en Python que limpia los datos de entrada, calcula las comisiones y genera el archivo `index.html`.
- `informe-ventas-2026-02-01_2026-05-25.xlsx`: Datos brutos de transacciones y formas de pago.
- `informe-productos-2026-02-01_2026-05-25.csv`: Base de datos de productos y sus categorías correspondientes.
- `.gitignore`: Archivos excluidos del control de versiones.

---

## Requisitos de Ejecución

Para volver a generar el reporte (en caso de que actualices los archivos `.csv` o `.xlsx`), necesitarás tener Python 3 e instalar las siguientes dependencias:

```bash
pip install pandas openpyxl
```

*(O puedes correrlo directamente usando el gestor `uv`):*
```bash
uv run --with pandas --with openpyxl generate_html_report.py
```

### Ejecutar la generación:
```bash
python generate_html_report.py
```
Esto procesará las planillas de datos y sobreescribirá el archivo `index.html` con las métricas actualizadas de inmediato.

---

## Funcionalidades del Dashboard (`index.html`)

1. **Dashboard y Ventas**: Métricas generales (Ventas brutas, comisiones totales estimadas, ingreso neto real, transacciones) y gráficos de ventas por día y categorías.
2. **Trazabilidad de Comisiones**: Detalle por forma de pago con las siguientes reglas aplicadas:
   - Débito: `1.55%`
   - Crédito: `2.50%`
   - Efectivo: `0.00%`
3. **Ventas por Producto Base**: Tabla interactiva con filtros dinámicos (búsqueda por texto, filtrado de categorías en tiempo real, mínimo de unidades vendidas y monto mínimo en CLP). Muestra un bloque con estadísticas del subconjunto seleccionado y permite ordenar de forma ascendente/descendente.
4. **Criterios de Consolidación**: Explicación detallada de cómo se limpiaron y agruparon variaciones del mismo producto (ej. aguas con/sin gas, jugos naturales, cafés con leche, variaciones de hamburguesas, etc.).
