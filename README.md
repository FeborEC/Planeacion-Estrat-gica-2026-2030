# Dashboard Plan Estratégico 2026 – 2030

Generador automático de dashboard HTML para el seguimiento del Plan Estratégico.

Dos scripts Python en orden de proceso: el primero toma los datos de las fuentes de información y llena el Excel; el segundo lee el Excel, calcula resultados, cumplimiento, semáforos, tendencias y alertas, y genera un **dashboard HTML interactivo y autocontenido** publicado en GitHub Pages.

---

## Estructura del proyecto (en orden de proceso)

```
Planeacion Estratégica 2026-2030/
├── 0_Actualizar_mes.bat            ← doble clic cada mes: corre 1, abre 2 para lo manual y corre 3
├── 1_Extraer_indicadores.py        ← fuentes de información → celdas azules del Excel
├── 2_Tablero_Indicadores_PE.xlsx   ← el modelo en Excel (aquí se digita lo manual)
├── 3_Generar_dashboard.py          ← Excel → docs/index.html y publicación en GitHub
├── Configuracion/
│   └── Parametros_Extraccion_PE.xlsx   ← rutas, cuentas contables, reglas, líneas de mayor rendimiento
├── Revision/
│   └── Auditoria_Extraccion_PE.xlsx    ← comparación digitado vs. fuentes y cálculo formulado (se regenera)
├── docs/
│   └── index.html                  ← dashboard publicado (GitHub Pages)
├── Soporte/                        ← chart_umd.js y material de apoyo
└── _respaldos/                     ← copias automáticas antes de escribir
```

Los Excel no se suben a GitHub (`.gitignore`); solo se publica `docs/index.html`.

## Requisitos

- Python 3.8+, `pip install openpyxl pandas pywin32` y Microsoft Excel (el paso 1 escribe a través de Excel para conservar fórmulas y comentarios)
- Acceso de lectura a la carpeta «Fuentes de información» del Modelo de Riesgos (ruta en `Configuracion/Parametros_Extraccion_PE.xlsx`)

## Flujo mensual

Doble clic en `0_Actualizar_mes.bat`, o paso a paso:

```bash
python 1_Extraer_indicadores.py          # llena las celdas vacías desde las fuentes (--comparar, --reemplazar, --hojas IND-13)
# digite en 2_Tablero_Indicadores_PE.xlsx lo que no sale de las fuentes (IPC, fallecidos, encuestas...), guarde y cierre
python 3_Generar_dashboard.py            # genera docs/index.html y lo publica (--no-publish para no publicar)
```

Revise `Revision/Auditoria_Extraccion_PE.xlsx` (hoja Comparacion): si al filtrar «Difiere» no sale nada, lo digitado cuadra con las fuentes.

## Estructura interna del script del dashboard

El script `3_Generar_dashboard.py` está dividido internamente en 8 secciones bien delimitadas por comentarios:

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
