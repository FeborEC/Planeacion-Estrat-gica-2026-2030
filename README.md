# Dashboard Plan Estratégico 2026 – 2030

Generador automático de dashboard HTML para el seguimiento del Plan Estratégico.

**Un solo script Python** lee el archivo Excel del modelo, procesa los datos mensualmente desde enero 2026 hasta diciembre 2030, calcula resultados, cumplimiento, semáforos, tendencias y alertas, y genera un **dashboard HTML interactivo, responsive y autocontenido** listo para publicar en GitHub Pages.

---

## Estructura del proyecto

```
proyecto/
├── "Plan Estrategico Dashboard.py"          ← ÚNICO script (con todo dentro)
├── chart_umd.js                  ← Chart.js (opcional, para offline)
├── Tablero_Indicadores_PE.xlsx
└── dashboard/
    └── index.html                ← salida (auto-generada)
```

Si no incluye `chart_umd.js`, el dashboard cargará Chart.js desde CDN (jsdelivr).
Si lo incluye, queda totalmente offline-capable.

## Requisitos

- Python 3.8+
- Una sola dependencia: `pip install openpyxl`

## Uso

```bash
python3 "Plan Estrategico Dashboard.py"
```

Opciones:

```bash
python3 "Plan Estrategico Dashboard.py" <ruta_excel> <ruta_salida>
```

Por defecto:
- Entrada: `Tablero_Indicadores_PE.xlsx`
- Salida: `dashboard/index.html`

## Flujo mensual

1. Abra el Excel y registre las variables del mes en las hojas `IND-XX` (celdas azules).
2. Guarde el Excel.
3. Ejecute `python3 "Plan Estrategico Dashboard.py"`.
4. Abra `dashboard/index.html` en el navegador.

## Estructura interna del script

El script `Plan Estrategico Dashboard.py` está dividido internamente en 8 secciones bien delimitadas por comentarios:

1. **Catálogo de indicadores** (datos del modelo: nombres, fórmulas, metas, tipos)
2. **Plantillas HTML / CSS / JS** (con la paleta KALLA)
3. **Extracción del Excel** (lectura via openpyxl)
4. **Cálculos** (resultados mensuales, cumplimiento, etc.)
5. **Validación**
6. **Construcción de payload JSON**
7. **Renderizado de HTML**
8. **Main**

Para personalizar, basta abrir el script y buscar la sección correspondiente.

## Características del dashboard

- **KPIs ejecutivos**: total / en meta / en riesgo / crítico / pendiente.
- **Filtros interactivos** por perspectiva y por año.
- **25 tarjetas de indicador** con mini-gráfico de evolución + meta de referencia.
- **Modal de detalle**: fórmula, descripción, 5 bloques de cumplimiento por año, gráfico de 60 meses, tabla mensual (sábana) completa, avisos de validación.
- **Responsive**: desktop, tablet, móvil.
- **Paleta KALLA**: verdes, teales y azules con accent por perspectiva.
- **Autocontenido**: un único HTML con Chart.js embebido.

## Publicación en GitHub Pages

```bash
git init
git add dashboard/index.html
git commit -m "Publicar dashboard"
git remote add origin https://github.com/<usuario>/<repo>.git
git push -u origin main
```

Settings → Pages → Source: Deploy from branch · Branch: `main` · Folder: `/dashboard` → URL pública: `https://<usuario>.github.io/<repo>/`.

## Cálculos implementados

| Métrica | Lógica |
|---|---|
| Resultado mensual | Depende del tipo: ratio (num/den), growth ((cur-prev)/prev), diff (a-b), avg (valor directo), count (valor directo). |
| % Cumpl. mensual | Ajustado por tendencia (Crecer/Reducir/Mantener). |
| % Cumpl. anual | Promedio de los % cumplimiento mensuales del año. |
| Semáforo | ≥95% 🟢 / 80–95% 🟡 / <80% 🔴 / sin datos ⚪. |
| Tendencia propia | Compara primer y último resultado de los últimos 12 meses con datos. |
| Avance global | Cumplimiento del último resultado disponible vs meta 2030. |
| Alertas | Texto descriptivo según el último año con datos. |

## Personalización

Todo está en el mismo archivo:

- Catálogo de indicadores: sección 1 (variable `IND`).
- Paleta de colores: sección 2 (variables `PALETTE`, `PERSP_COLOR`, CSS `:root`).
- Lógica de cálculo: sección 4 (funciones `compute_*`).
- HTML/CSS/JS: secciones 2 y 7.

## Licencia

Uso interno. Modificable libremente.
