#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Plan Estrategico Dashboard.py
=============================

Generador automático del dashboard HTML del Plan Estratégico 2026-2030.

Lee el archivo Excel del modelo (con las 25 hojas individuales de cada
indicador), procesa los datos mensualmente desde enero 2026 hasta diciembre
2030, calcula resultados, % de cumplimiento, semáforos, tendencias propias,
avance acumulado y alertas, y genera un dashboard HTML interactivo, moderno,
responsive y autocontenido, listo para publicar en GitHub Pages.

Este es un único script Python autocontenido. Solo necesita openpyxl:

    pip install openpyxl

Uso
---
    python3 "Plan Estrategico Dashboard.py"
    python3 "Plan Estrategico Dashboard.py" <ruta_excel>
    python3 "Plan Estrategico Dashboard.py" <ruta_excel> <ruta_salida_html>

Por defecto:
    Input:  Tablero_Indicadores_PE.xlsx
    Output: docs/index.html

Archivo auxiliar opcional en la misma carpeta:
    chart_umd.js   ← Chart.js (si existe se embebe; si no, se carga desde CDN)

Estructura interna del script
-----------------------------
    1.  CATÁLOGO DE INDICADORES        (datos del modelo)
    2.  PLANTILLAS HTML / CSS / JS     (paleta KALLA)
    3.  EXTRACCIÓN DEL EXCEL           (lectura openpyxl)
    4.  CÁLCULOS                       (resultados, cumplimiento, etc.)
    5.  VALIDACIÓN
    6.  CONSTRUCCIÓN DE PAYLOAD JSON
    7.  RENDERIZADO DE HTML
    8.  MAIN
"""

import sys
import json
import math
import re
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: falta openpyxl. Instale con:  pip install openpyxl")
    sys.exit(1)



# ============================================================================
# 1. CATÁLOGO DE INDICADORES
# ============================================================================

OBJ_GEN = {
    "FINANCIERA": "Mantener un margen financiero adecuado, buscando una distribución de los ingresos (50% gastos, 35% costos y 15% excedentes), mediante la eficiencia operativa, el control integral del riesgo y la optimización continua de la cartera, apoyado en políticas financieras prudentes y competitivas, herramientas de análisis y un equipo técnico fortalecido, con cumplimiento proyectado a diciembre de 2030.",
    "ASOCIADOS": "Aumentar la base social a 8.000 asociados al año 2030, mediante la implementación de estrategias de fidelización y campañas de atracción segmentadas, apoyado en una estructura comercial fortalecida, canales ágiles de comunicación y el uso de inteligencia de datos.",
    "PROCESOS INTERNOS": "Automatizar al menos el 70% de las tareas susceptibles de automatización asociadas a los procesos operativos y de servicio al asociado al finalizar el año 2030, apoyándose en tecnología, analítica, seguridad de la información y capacidades digitales organizacionales.",
    "APRENDIZAJE Y CRECIMIENTO": "Potenciar la formación financiera y solidaria de los asociados, así como el desarrollo profesional de los empleados, alcanzando al menos el 50% de participación de asociados y el 100% de empleados, mediante programas de formación continua, alianzas estratégicas y campañas de educación cooperativa, con recursos institucionales y seguimiento de indicadores de desarrollo y participación, a diciembre de 2030.",
}

PERSP_INFO = {
    "FINANCIERA":               {"short": "Financiera",   "icon": "💰", "accent": "1F6FB2"},
    "ASOCIADOS":                {"short": "Asociados",    "icon": "🤝", "accent": "00A14B"},
    "PROCESOS INTERNOS":        {"short": "Procesos",     "icon": "⚙️", "accent": "E08A00"},
    "APRENDIZAJE Y CRECIMIENTO":{"short": "Aprendizaje",  "icon": "🎓", "accent": "8E44AD"},
}

INI_ICON = {
    "Aumento de ingresos":"📈",
    "Reducción de mora":"🛡️",
    "Crecimiento y remuneración de depósitos":"🏦",
    "Rentabilidad y fortalecimiento patrimonial":"💎",
    "Crecimiento de la base social":"🌱",
    "Retención y profundización":"💗",
    "Operatividad digital y canales":"📱",
    "Automatización de procesos":"🤖",
    "Transformación digital del asociado":"💻",
    "Formación de asociados":"📚",
    "Formación y desarrollo del personal":"👥",
}

INI_DESC = {
    "Aumento de ingresos": "Aumentar los ingresos mediante la optimización de la cartera, la colocación de crédito en líneas de mayor rendimiento, la generación de comisiones por servicios no financieros y la rentabilidad de las inversiones.",
    "Reducción de mora": "Reducir y controlar el deterioro de cartera, fortaleciendo la gestión preventiva, la recuperación y la normalización de obligaciones.",
    "Crecimiento y remuneración de depósitos": "Fortalecer la captación y remuneración competitiva de depósitos como base de fondeo estable y rentable.",
    "Rentabilidad y fortalecimiento patrimonial": "Asegurar márgenes financieros sostenibles y consolidar el patrimonio mediante el cumplimiento de aportes y la generación de capital institucional.",
    "Crecimiento de la base social": "Impulsar el crecimiento de la base social mediante la vinculación de nuevos asociados, soportado en ofertas diferenciadas que respondan a las necesidades de los segmentos objetivo.",
    "Retención y profundización": "Reducir la desvinculación de asociados e incrementar el uso del portafolio de servicios mediante acciones de retención, seguimiento y profundización comercial.",
    "Operatividad digital y canales": "Optimizar los canales de atención y comunicación con los asociados, asegurando efectividad y satisfacción en cada interacción.",
    "Automatización de procesos": "Automatizar procesos operativos y de servicio al asociado para mejorar eficiencia, calidad y tiempos de respuesta.",
    "Transformación digital del asociado": "Promover la transformación digital de los asociados, incentivando el uso de los canales transaccionales digitales.",
    "Formación de asociados": "Fortalecer la formación financiera y solidaria de los asociados.",
    "Formación y desarrollo del personal": "Fortalecer el desarrollo profesional de los funcionarios y el clima organizacional.",
}

PERSP_ORDER = ["FINANCIERA","ASOCIADOS","PROCESOS INTERNOS","APRENDIZAJE Y CRECIMIENTO"]

# Indicadores
# tipo: 'ratio' | 'growth' | 'diff' | 'avg' | 'count'
# vars: lista de campos de entrada mensual con (codigo, etiqueta)
# metas: lista de 5 valores (2026..2030); None o str = "por definir" / referencia
# tipo_meta: 'pct' | 'bps' | 'ratio' | 'unidades' | 'texto'
IND = [
 # ============ FINANCIERA — Distribución del ingreso (IND-F) ============
 dict(p="FINANCIERA", ini="Distribución del ingreso",
   indicador="Gastos sobre Ingresos",
   desc="Mide la proporción de ingresos destinada a gastos operativos. Meta: ≤ 50%.",
   formula="Gastos operativos totales ÷ Ingresos totales",
   tipo="ratio", vars=[("num","Gastos operativos totales ($)"),("den","Ingresos totales ($)")],
   unidad="%", tipo_meta="pct", tend="Reducir", frec="Mensual",
   metas=[0.50,0.50,0.50,0.50,0.50], code="IND-F1"),
 dict(p="FINANCIERA", ini="Distribución del ingreso",
   indicador="Costos sobre Ingresos",
   desc="Mide la proporción de ingresos destinada a costos financieros. Meta: ≤ 35%.",
   formula="Costos financieros totales ÷ Ingresos totales",
   tipo="ratio", vars=[("num","Costos financieros totales ($)"),("den","Ingresos totales ($)")],
   unidad="%", tipo_meta="pct", tend="Reducir", frec="Mensual",
   metas=[0.35,0.35,0.35,0.35,0.35], code="IND-F2"),
 dict(p="FINANCIERA", ini="Distribución del ingreso",
   indicador="Excedente sobre Ingresos",
   desc="Mide la proporción de ingresos que se convierte en excedente. Meta: ≥ 15%.",
   formula="Excedente del período ÷ Ingresos totales",
   tipo="ratio", vars=[("num","Excedente del período ($)"),("den","Ingresos totales ($)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.15,0.15,0.15,0.15,0.15], code="IND-F3"),

 # ============ FINANCIERA — Aumento de ingresos ============
 dict(p="FINANCIERA", ini="Aumento de ingresos",
   indicador="Crecimiento de cartera en líneas de mayor rendimiento",
   desc="Mide la participación de las líneas de crédito con mayor tasa efectiva y contribución al margen financiero.",
   formula="Saldo de cartera de líneas de mayor rendimiento ÷ Saldo total de cartera",
   tipo="ratio", vars=[("num","Saldo cartera líneas mayor rendimiento ($)"),("den","Saldo total de cartera ($)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.08,0.10,0.12,0.14,0.16]),
 dict(p="FINANCIERA", ini="Aumento de ingresos",
   indicador="Crecimiento de ingresos operativos",
   desc="Refleja el impacto directo de la colocación sobre los ingresos financieros.",
   formula="(Ingresos mes actual − Ingresos mes anterior) ÷ Ingresos mes anterior",
   tipo="growth_acum", vars=[("acum","Ingresos Acumulados")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.10,0.10,0.10,0.10,0.10]),
 dict(p="FINANCIERA", ini="Aumento de ingresos",
   indicador="Crecimiento de ingresos por comisiones",
   desc="Mide el incremento de ingresos derivados del uso de servicios no financieros (convenios).",
   formula="(Ingresos comisiones mes actual − mes anterior) ÷ Ingresos mes anterior",
   tipo="growth_acum", vars=[("acum","Ingresos Acumulados")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.035,0.035,0.035,0.035,0.035]),
 dict(p="FINANCIERA", ini="Aumento de ingresos",
   indicador="Rentabilidad Fondo de Liquidez",
   desc="Mide la capacidad de generar rentabilidad por encima del mercado (IBR).",
   formula="Tasa promedio ponderada de rentabilidad del fondo de liquidez − IBR",
   tipo="diff", vars=[("tasa","Tasa promedio ponderada Fondo de Liquidez (%)"),("ref","IBR (%)")],
   unidad="Puntos básicos", tipo_meta="pts", tend="Crecer", frec="Mensual",
   metas=[-200,-200,-200,-200,-200]),

 # ============ FINANCIERA — Reducción de mora ============
 dict(p="FINANCIERA", ini="Reducción de mora",
   indicador="Índice de mora",
   desc="Mide el nivel de morosidad de la cartera por crédito, sin aplicación de arrastres o riesgo.",
   formula="Saldo de cartera ≥ B ÷ Saldo total de cartera",
   tipo="ratio", vars=[("num","Saldo de cartera en mora ≥ B ($)"),("den","Saldo total de cartera ($)")],
   unidad="%", tipo_meta="pct", tend="Reducir", frec="Mensual",
   metas=[0.05,0.05,0.05,0.05,0.05]),
 dict(p="FINANCIERA", ini="Reducción de mora",
   indicador="Tasa de recuperación",
   desc="Mide la proporción de cartera recuperada frente a la deteriorada.",
   formula="Saldo de cartera recuperada ÷ Saldo de cartera deteriorada",
   tipo="ratio", vars=[("num","Saldo cartera recuperada ($)"),("den","Saldo cartera deteriorada ($)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.30,0.30,0.30,0.30,0.30]),
 dict(p="FINANCIERA", ini="Reducción de mora",
   indicador="Tasa de normalización de cartera",
   desc="Mide la efectividad de acuerdos de pago y reestructuraciones.",
   formula="Saldo de normalizaciones al día ÷ Saldo total de normalizaciones",
   tipo="ratio", vars=[("num","Saldo normalizaciones al día ($)"),("den","Saldo total normalizaciones ($)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.80,0.80,0.80,0.80,0.80]),

 # ============ FINANCIERA — Crecimiento y remuneración de depósitos ============
 dict(p="FINANCIERA", ini="Crecimiento y remuneración de depósitos",
   indicador="Tasa de crecimiento de ahorros",
   desc="Mide el crecimiento de los depósitos vs el saldo base del año anterior.",
   formula="(Saldo mes − Saldo base año anterior) ÷ Saldo base año anterior",
   tipo="avg", vars=[("val","Saldo total de depósitos del mes ($)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[None,None,None,None,None]),
 dict(p="FINANCIERA", ini="Crecimiento y remuneración de depósitos",
   indicador="Rentabilidad real de los depósitos",
   desc="Asegura que la tasa de rendimiento de los depósitos supere la inflación.",
   formula="Tasa promedio ponderada ahorros − IPC",
   tipo="diff", vars=[("tasa","Tasa promedio ponderada ahorros (%)"),("ref","IPC VAR ANUAL (%)")],
   unidad="% vs IPC", tipo_meta="pct", tend="Mantener", frec="Mensual",
   metas=[0,0,0,0,0]),

 # ============ FINANCIERA — Rentabilidad y fortalecimiento patrimonial ============
 dict(p="FINANCIERA", ini="Rentabilidad y fortalecimiento patrimonial",
   indicador="Suficiencia del Margen",
   desc="Evalúa la capacidad del margen financiero bruto para cubrir los gastos operativos.",
   formula="Margen financiero bruto ÷ Gastos operativos",
   tipo="ratio", vars=[("num","Margen financiero bruto ($)"),("den","Gastos operativos ($)")],
   unidad="%", tipo_meta="pct", tend="Mantener", frec="Mensual",
   metas=[1.20,1.20,1.20,1.20,1.20]),
 dict(p="FINANCIERA", ini="Rentabilidad y fortalecimiento patrimonial",
   indicador="Cumplimiento de pago de aportes sociales",
   desc="Mide el grado de cumplimiento del recaudo mensual de aportes frente al esperado.",
   formula="Recaudo real de aportes ÷ Recaudo esperado de aportes",
   tipo="ratio", vars=[("num","Recaudo real de aportes ($)"),("den","Recaudo esperado de aportes ($)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.95,0.95,0.95,0.95,0.95]),
 dict(p="FINANCIERA", ini="Rentabilidad y fortalecimiento patrimonial",
   indicador="Capital institucional",
   desc="Mide la proporción del patrimonio permanente frente al total de activos.",
   formula="(Aportes amortizados + Reserva protección aportes + Excedente mes + Resultado ejercicios anteriores) ÷ Activo",
   tipo="multi_sum",
   vars=[("v1","Aportes amortizados ($)"),
         ("v2","Reserva protección aportes ($)"),
         ("v3","Excedente del mes ($)"),
         ("v4","Resultado de ejercicios anteriore"),
         ("den","Activo total ($)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.15,0.15,0.15,0.15,0.15]),
 dict(p="FINANCIERA", ini="Rentabilidad y fortalecimiento patrimonial",
   indicador="ROE",
   desc="Rentabilidad sobre el patrimonio. Meta dinámica = IPC corrido del año (modelo SES).",
   formula="Excedente del mes ÷ Patrimonio promedio (modelo SES)",
   tipo="ratio", vars=[("num","Excedente ($)"),("den","Patrimonio ($)")],
   unidad="%", tipo_meta="ipc", tend="Mantener", frec="Mensual",
   # Meta dinámica: se lee del IPC CORRIDO columna 12 del Excel (por período)
   metas=[None,None,None,None,None],
   code="IND-13"),

 # ============ ASOCIADOS ============
 dict(p="ASOCIADOS", ini="Crecimiento de la base social",
   indicador="Crecimiento de la base social",
   desc="Evalúa el incremento de asociados y el cumplimiento de la meta de base social.",
   # Fórmula Excel exacta (revisada):
   #   Resultado (col F) = E - base_dic_año_anterior        → acumulado desde enero
   #   Meta (col G) mes1 = meta_anual / 12                  → 481/12 = 40.08
   #   Meta (col G) mes_n = n × (meta_anual / 12)           → acumulada mes a mes
   #   % Cumpl = resultado_acumulado / meta_acumulada_mes_n
   #   2026: base=5365, meta_anual=481
   #   2027: base=5846, meta_anual=538 (+ajuste desvío dic2026)
   formula="Crecimiento real acumulado de asociados ÷ Crecimiento esperado acumulado",
   tipo="base_social",
   vars=[("val","Total de asociados al cierre del mes")],
   unidad="N° asociados", tipo_meta="abs", tend="Crecer", frec="Mensual",
   decimales=0,
   # Meta = crecimiento esperado anual (col D del Excel por año)
   metas=[481, 538, 539, 538, 539],
   # Base dic año anterior para calcular resultado acumulado
   base_social_bases={2026: 5365, 2027: 5846, 2028: 6384, 2029: 6923, 2030: 7461},
   code="IND-14"),
 dict(p="ASOCIADOS", ini="Retención y profundización",
   indicador="Tasa de retención de asociados anual",
   desc="Mide la capacidad de la cooperativa para mantener su base social activa.",
   # Fórmula Excel exacta (col L):
   #   Resultado = Base_que_se_mantiene / (Base_dic_año_anterior - Fallecidos)
   #   Col K (Tasa Retención simple) = Base_que_se_mantiene / Base_dic_año_anterior
   #   Resultado (col L) = Base_que_se_mantiene / (Base_dic_año_anterior - Fallecidos)
   #   % Cumpl = Resultado / Meta
   formula="(Base social que se mantiene) ÷ (Base social dic año anterior − Asociados fallecidos)",
   tipo="retencion",
   vars=[("base_ant",  "Base social (dic año anterior )"),
         ("base_act",  "Base social actual"),
         ("mantiene",  "Base social que se mantiene (estaban antes y ahora)"),
         ("nuevos",    "Asociados nuevos"),
         ("retiros",   "Retiros Totales"),
         ("fallecidos","Asociados Fallecidos")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   # frec=Mensual aunque la ficha diga "Anual": el Excel captura datos todos los meses
   metas=[0.98,0.98,0.98,0.98,0.98],
   decimales=0,   # variables de entrada sin decimales
   code="IND-15"),
 dict(p="ASOCIADOS", ini="Retención y profundización",
   indicador="Índice de satisfacción del asociado",
   desc="Evalúa la percepción del servicio para identificar oportunidades de mejora.",
   # Fórmula Excel exacta (col M):
   # = (NPS×0.4) + (CES/5×0.2) + ((FCR_WA+FCR_LL)/2×0.2) + (((SAT_OFI+SAT_ASESOR+SAT_CHAT+SAT_LL)/4)/5×0.2)
   # resultado_excel (data_only) es el valor correcto directamente
   formula="(NPS×40%) + (CES/5×20%) + ((FCR_WA+FCR_LL)/2×20%) + ((Sat_prom/5)×20%)",
   tipo="avg",
   vars=[("nps",     "NPS "),
         ("ces",     "CES"),
         ("fcr_wa",  "FCR WhatsApp"),
         ("fcr_ll",  "FCR Llamadas"),
         ("sat_ofi", "Satisfacción Presencial Visita Oficina"),
         ("sat_ase", "Satisfacción Presencial Respecto Asesor"),
         ("sat_chat","Satisfacción Multichat WhatsApp"),
         ("sat_ll",  "Satisfacción Reporte Llamadas")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   metas=[0.95,0.95,0.95,0.95,0.95],
   code="IND-16"),
 dict(p="ASOCIADOS", ini="Retención y profundización",
   indicador="Promedio de productos por asociado",
   desc="Mide el nivel de profundización comercial en la base social.",
   formula="Total de productos activos ÷ Total de asociados activos",
   # El resultado está pre-calculado en la columna Resultado del Excel (literal).
   # tipo='avg' prioriza resultado_excel; vars se muestran en composición del cálculo.
   tipo="avg",
   vars=[("num","Total productos activos"),("den","Total asociados activos")],
   unidad="Productos", tipo_meta="abs", tend="Crecer", frec="Mensual",
   metas=[3,3,3,3,3],
   hide_input_vars=True,
   code="IND-17"),

 # ============ PROCESOS ============
 dict(p="PROCESOS INTERNOS", ini="Automatización de procesos",
   indicador="Nivel de automatización de procesos",
   desc="Mide el porcentaje de procesos operativos y de servicio que han sido automatizados.",
   # PRIORIDAD: el valor de la columna Resultado (col G) del Excel manda siempre.
   #   col E = Tareas automatizadas (=Inventario!E110)   ← insumo informativo
   #   col F = Tareas identificadas  (=Inventario!F110)   ← insumo informativo
   #   col G = Resultado  → puede ser fórmula (E/F) o un valor que el usuario escriba
   #           manualmente; el script SIEMPRE respeta lo que haya en col G.
   #   % Cumpl (col I) es 'directo': el Resultado ES el % de cumplimiento.
   # tipo='avg' + hide_input_vars=True hace que el script lea resultado_excel literal
   # (mismo patrón que IND-17), de modo que se actualiza al cambiar el Excel.
   formula="Tareas automatizadas acumuladas ÷ Total tareas identificadas (Inventario de Automatización)",
   tipo="avg",
   vars=[("num","Tareas automatizadas (Inventario de automatización)"),
         ("den","Tareas  identificadas \n(Inventario de automatización)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Mensual",
   # frec=Mensual: todos los meses apuntan al mismo valor del inventario (acumulado anual)
   metas=[0.70,0.70,0.70,0.70,0.70],
   decimales=0,   # vars de entrada sin decimales
   hide_input_vars=True,
   code="IND-18"),
 dict(p="PROCESOS INTERNOS", ini="Transformación digital del asociado",
   indicador="Uso y adopción de canales transaccionales digitales",
   desc="Mide el porcentaje de transacciones realizadas por canales digitales.",
   # PRIORIDAD: columna Resultado (col I) del Excel = H/F (Transacc. digitales <70 ÷ Asociados <70).
   # tipo='avg' + hide_input_vars=True lee resultado_excel literal (mismo patrón que IND-18/20),
   # de modo que trae el dato real de cada mes y se actualiza al cambiar el Excel.
   formula="Transacciones digitales (<70 años) ÷ Total transacciones (<70 años)",
   tipo="avg", vars=[("num","Asociados con Transacciones digitales (<70 años)"),("den","Asociados (<70 años)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Trimestral",
   metas=[0.80,0.80,0.80,0.80,0.80],
   hide_input_vars=True,
   # Frecuencia oficial Trimestral, pero se capturan datos mensuales:
   # show_all_months=True habilita los 12 meses para mostrar todo dato disponible.
   show_all_months=True,
   code="IND-19"),
 dict(p="PROCESOS INTERNOS", ini="Transformación digital del asociado",
   indicador="Uso y adopción de App y CoopDigital",
   desc="Mide el uso de canales digitales (App) por asociados menores de 70 años.",
   # Fórmula Excel exacta (col I): Resultado = F / H
   #   col E = Asociados con Coop Digital (<70 años)   ← insumo informativo
   #   col F = Asociados con app (<70 años)            ← NUMERADOR
   #   col G = uso en el último mes                    ← insumo informativo
   #   col H = Total Asociados (<70 años)              ← DENOMINADOR
   #   El resultado está pre-calculado en col I (literal); tipo='avg' lo prioriza.
   formula="Asociados con App y CoopDigital (<70 años) ÷ Total de asociados (<70 años)",
   tipo="avg",
   vars=[("num","Asociados con app (<70 años)"),("den","Total Asociados (<70 años)")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Trimestral",
   metas=[0.80,0.80,0.80,0.80,0.80],
   hide_input_vars=True,
   # Frecuencia oficial Trimestral, pero se capturan datos mensuales:
   # show_all_months=True habilita los 12 meses para mostrar todo dato disponible.
   show_all_months=True,
   code="IND-20"),

 # ============ APRENDIZAJE ============
 dict(p="APRENDIZAJE Y CRECIMIENTO", ini="Formación de asociados",
   indicador="Formación de asociados por temáticas",
   desc="Mide el porcentaje de asociados capacitados en temas financieros y solidarios.",
   formula="Asociados capacitados ÷ Total de asociados",
   tipo="ratio", vars=[("num","Asociados capacitados (acumulado anual)"),("den","Total de asociados")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Trimestral",
   metas=[0.40,0.40,0.40,0.40,0.40],
   code="IND-21"),
 dict(p="APRENDIZAJE Y CRECIMIENTO", ini="Formación de asociados",
   indicador="Formación en economía solidaria (20 hrs mínimo)",
   desc="Mide el porcentaje de asociados que alcanzan las horas mínimas de formación solidaria.",
   formula="Asociados con horas mínimas de formación ÷ Total de asociados",
   tipo="ratio", vars=[("num","Asociados que cumplen 20h"),("den","Total de asociados")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Trimestral",
   metas=[0.40,0.40,0.40,0.40,0.40],
   code="IND-22"),
 dict(p="APRENDIZAJE Y CRECIMIENTO", ini="Formación y desarrollo del personal",
   indicador="Cumplimiento del plan de formación de personal",
   desc="Mide el grado de ejecución del plan de capacitación del personal.",
   formula="Actividades de formación ejecutadas ÷ Actividades planificadas",
   tipo="ratio", vars=[("num","Actividades de formación ejecutadas"),("den","Actividades planificadas")],
   unidad="%", tipo_meta="pct", tend="Crecer", frec="Semestral",
   metas=[1.0,1.0,1.0,1.0,1.0],
   code="IND-23"),
 dict(p="APRENDIZAJE Y CRECIMIENTO", ini="Formación y desarrollo del personal",
   indicador="Clima organizacional y compromiso del personal",
   desc="Mide el resultado de la encuesta de clima organizacional.",
   formula="Promedio de resultados de la encuesta de clima organizacional",
   tipo="avg", vars=[("val","Resultado encuesta clima organizacional (0–1)")],
   unidad="Puntos (escala)", tipo_meta="pct", tend="Mantener", frec="Anual",
   metas=[1.0,1.0,1.0,1.0,1.0],
   code="IND-24"),
]

# Sheet code for each indicator — respetar código personalizado si existe
_seq = 1
for d in IND:
    if 'code' not in d or not d['code']:
        d["code"] = f"IND-{_seq:02d}"
    if not d['code'].startswith('IND-F'):  # solo avanzar secuencia para indicadores estándar
        _seq += 1


# ============================================================================
# 2. PLANTILLAS HTML / CSS / JAVASCRIPT
# ============================================================================

import json


# ------------------------------------------------------------
# Paleta inspirada en la rueda KALLA (verdes, teales, azules)
# ------------------------------------------------------------
PALETTE = {
    # Marca
    'primary':       '#0876B2',   # azul KALLA
    'primary_dark':  '#075A8A',
    'primary_light': '#2989D0',
    'teal':          '#02A2C6',
    'teal_dark':     '#007C95',
    'green':         '#23A968',
    'green_light':   '#45BA76',
    # Semáforo
    'success':       '#23A968',
    'warning':       '#F0A030',
    'danger':        '#E04B4B',
    'neutral':       '#94A6B8',
    # Superficies
    'bg':            '#F4F8FB',
    'surface':       '#FFFFFF',
    'surface_alt':   '#EEF4F8',
    'border':        '#DCE6EE',
    # Texto
    'text_dark':     '#0E2A40',
    'text_med':      '#3F5A73',
    'text_lt':       '#7990A4',
}

# Color por perspectiva (dentro de la familia KALLA)
PERSP_COLOR = {
    'FINANCIERA':                '#0876B2',  # azul
    'ASOCIADOS':                 '#23A968',  # verde
    'PROCESOS INTERNOS':         '#02A2C6',  # teal
    'APRENDIZAJE Y CRECIMIENTO': '#1A4D8C',  # azul profundo
}


# ============================================================
# CSS
# ============================================================

CSS = r"""

:root {
  --primary:     #0876B2;
  --primary-dk:  #075A8A;
  --primary-lt:  #2989D0;
  --teal:        #02A2C6;
  --green:       #23A968;
  --green-lt:    #45BA76;
  --success:     #22C55E;
  --warning:     #F59E0B;
  --danger:      #EF4444;
  --neutral:     #94A3B8;
  --bg:          #F1F5F9;
  --surface:     #FFFFFF;
  --surface-2:   #F8FAFC;
  --surface-3:   #EEF4F8;
  --border:      #E2E8F0;
  --border-soft: #F1F5F9;
  --text-dark:   #0F172A;
  --text-med:    #475569;
  --text-lt:     #94A3B8;
  --shadow-xs:   0 1px 2px rgba(15,23,42,0.04);
  --shadow-sm:   0 1px 4px rgba(15,23,42,0.06), 0 1px 2px rgba(15,23,42,0.04);
  --shadow:      0 4px 14px rgba(15,23,42,0.07), 0 2px 4px rgba(15,23,42,0.04);
  --shadow-lg:   0 8px 28px rgba(15,23,42,0.10);
  --shadow-xl:   0 16px 48px rgba(15,23,42,0.18);
  --radius:      12px;
  --radius-sm:   8px;
  --radius-lg:   18px;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  background: var(--bg);
  color: var(--text-dark);
  font-size: 14px;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}

/* ─── Header ─────────────────────────────────────────────── */
/* ─── Header premium ─────────────────────────────────────── */
.header {
  background: linear-gradient(120deg, #0657A4 0%, #0B88C7 45%, #07A89E 80%, #2DC98E 100%);
  color: white;
  position: relative; z-index: 80;
  box-shadow: 0 4px 20px rgba(6,87,164,0.30);
  overflow: hidden;
}
/* Decoración circular de fondo */
.header::before {
  content: '';
  position: absolute; right: -80px; top: -80px;
  width: 340px; height: 340px; border-radius: 50%;
  background: rgba(255,255,255,0.06); pointer-events: none;
}
.header::after {
  content: '';
  position: absolute; right: 140px; bottom: -60px;
  width: 200px; height: 200px; border-radius: 50%;
  background: rgba(255,255,255,0.04); pointer-events: none;
}

.header-inner {
  max-width: 1400px; margin: 0 auto;
  padding: 20px 32px;
  display: flex; align-items: center; gap: 24px;
  flex-wrap: wrap; position: relative; z-index: 1;
}

/* Logo placeholder (texto estilizado cuando no hay imagen) */
.header-brand {
  display: flex; align-items: center; gap: 14px;
  flex-shrink: 0;
  border-right: 1px solid rgba(255,255,255,0.25);
  padding-right: 20px;
}
.header-logo-box {
  width: 52px; height: 52px; border-radius: 14px;
  background: rgba(255,255,255,0.15);
  border: 1.5px solid rgba(255,255,255,0.30);
  display: flex; align-items: center; justify-content: center;
  backdrop-filter: blur(8px);
}
.header-logo-box svg { width: 28px; height: 28px; color: white; }
.header-brand-text { display: flex; flex-direction: column; gap: 1px; }
.header-brand-name {
  font-size: 1.10rem; font-weight: 900; color: white;
  letter-spacing: -0.01em; line-height: 1;
}
.header-brand-sub {
  font-size: 0.66rem; color: rgba(255,255,255,0.80);
  font-weight: 500; letter-spacing: 0.04em;
  text-transform: uppercase;
}

/* Título y subtítulo */
.header-titles { flex: 1; min-width: 0; }
.header-titles h1 {
  font-size: 1.45rem; margin: 0; font-weight: 800;
  letter-spacing: -0.02em; color: white;
  display: flex; align-items: center; gap: 10px;
}
.brand-dot {
  display: inline-block; width: 10px; height: 10px; border-radius: 50%;
  background: #4ADE80;
  box-shadow: 0 0 0 3px rgba(74,222,128,0.30), 0 0 12px rgba(74,222,128,0.50);
  flex-shrink: 0;
}
.header-meta {
  display: flex; align-items: center; gap: 0;
  margin-top: 6px; flex-wrap: wrap;
}
.header-meta-item {
  display: inline-flex; align-items: center; gap: 5px;
  font-size: 0.80rem; color: rgba(255,255,255,0.85); font-weight: 500;
  padding: 0 12px; border-right: 1px solid rgba(255,255,255,0.25);
}
.header-meta-item:first-child { padding-left: 0; }
.header-meta-item:last-child  { border-right: none; }
.header-meta-item svg { width: 13px; height: 13px; opacity: 0.85; }

/* Botones header (contenedor) */
.header-btns { display: flex; gap: 10px; align-items: center; flex-shrink: 0; }

/* Botón Ruta Planeación Estratégica */
.btn-ruta-estrategica {
  display: inline-flex; align-items: center; gap: 10px;
  background: rgba(255,255,255,0.15); color: white;
  border: 1.5px solid rgba(255,255,255,0.40); border-radius: 12px;
  padding: 11px 20px; font-size: 0.88rem; font-weight: 700;
  cursor: pointer; transition: all .15s;
  text-decoration: none; flex-shrink: 0; white-space: nowrap;
}
.btn-ruta-estrategica:hover {
  background: rgba(255,255,255,0.26); border-color: rgba(255,255,255,0.65);
  color: white; text-decoration: none; transform: translateY(-2px);
}
.btn-ruta-estrategica .xl-icon {
  width: 32px; height: 32px; border-radius: 7px;
  background: rgba(255,255,255,0.20); display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.btn-ruta-estrategica .xl-icon svg { width: 18px; height: 18px; }

/* Botón Agregar Datos al Origen */
.btn-datos-origen {
  display: inline-flex; align-items: center; gap: 10px;
  background: rgba(255,255,255,0.15); color: white;
  border: 1.5px solid rgba(255,255,255,0.40); border-radius: 12px;
  padding: 11px 20px; font-size: 0.88rem; font-weight: 700;
  cursor: pointer; transition: all .15s;
  text-decoration: none; flex-shrink: 0; white-space: nowrap;
}
.btn-datos-origen:hover {
  background: rgba(255,255,255,0.26); border-color: rgba(255,255,255,0.65);
  color: white; text-decoration: none; transform: translateY(-2px);
}
.btn-datos-origen .xl-icon {
  width: 32px; height: 32px; border-radius: 7px;
  background: rgba(255,255,255,0.20); display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.btn-datos-origen .xl-icon svg { width: 18px; height: 18px; }

/* Botón Excel Consolidado */
.btn-excel-consolidado {
  display: inline-flex; align-items: center; gap: 10px;
  background: white; color: #166534;
  border: none; border-radius: 12px;
  padding: 11px 20px;
  font-size: 0.88rem; font-weight: 700;
  cursor: pointer; transition: all .15s;
  box-shadow: 0 4px 14px rgba(0,0,0,0.15);
  flex-shrink: 0; white-space: nowrap;
}
.btn-excel-consolidado:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 20px rgba(0,0,0,0.20);
  background: #F0FDF4;
}
.btn-excel-consolidado .xl-icon {
  width: 32px; height: 32px; border-radius: 7px;
  background: #16A34A; display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}
.btn-excel-consolidado .xl-icon svg { width: 18px; height: 18px; color: white; }

/* KPIs en el header */
.header-kpis {
  display: flex; gap: 10px; flex-wrap: wrap;
  max-width: 1400px; margin: 0 auto;
  padding: 0 32px 16px; position: relative; z-index: 1;
}
.hkpi {
  background: rgba(255,255,255,0.14);
  border: 1.5px solid rgba(255,255,255,0.22);
  border-radius: 14px;
  padding: 12px 18px;
  display: flex; align-items: center; gap: 14px;
  backdrop-filter: blur(8px);
  cursor: pointer; transition: all .15s;
  flex: 1; min-width: 100px;
}
.hkpi:hover { background: rgba(255,255,255,0.24); transform: translateY(-1px); }
.hkpi.clickable:hover { background: rgba(255,255,255,0.26); }
.hkpi-icon {
  width: 40px; height: 40px; border-radius: 11px;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}
.hkpi-icon svg { width: 20px; height: 20px; }
.hkpi-icon.total     { background: rgba(255,255,255,0.20); color: white; }
.hkpi-icon.cumple    { background: rgba(74,222,128,0.25);  color: #4ADE80; }
.hkpi-icon.riesgo    { background: rgba(253,230,138,0.25); color: #FDE68A; }
.hkpi-icon.critico   { background: rgba(252,165,165,0.25); color: #FCA5A5; }
.hkpi-icon.pendiente { background: rgba(255,255,255,0.12); color: rgba(255,255,255,0.60); }
.hkpi-right { display: flex; flex-direction: column; }
.hkpi-val { font-size: 1.55rem; font-weight: 900; line-height: 1; color: white; letter-spacing: -0.02em; }
.hkpi-lbl { font-size: 0.64rem; color: rgba(255,255,255,0.80); text-transform: uppercase; letter-spacing: 0.07em; margin-top: 3px; font-weight: 600; }
.hkpi-bar { height: 3px; border-radius: 999px; background: rgba(255,255,255,0.15); margin-top: 5px; overflow: hidden; width: 100%; }
.hkpi-bar-fill { height: 100%; border-radius: 999px; }
.hkpi-bar-fill.cumple    { background: #4ADE80; }
.hkpi-bar-fill.riesgo    { background: #FDE68A; }
.hkpi-bar-fill.critico   { background: #FCA5A5; }
.hkpi-bar-fill.pendiente { background: rgba(255,255,255,0.35); }
.hkpi-bar-fill.total     { background: white; }



/* ─── Container ──────────────────────────────────────────── */
.container { max-width: 1400px; margin: 0 auto; padding: 0 28px 32px; }

/* ─── Popover ────────────────────────────────────────────── */
.popover-overlay { position: fixed; inset: 0; z-index: 90; display: none; }
.popover-overlay.active { display: block; }
.popover {
  position: absolute; background: var(--surface);
  border: 1px solid var(--border); border-radius: var(--radius);
  box-shadow: var(--shadow-xl);
  min-width: 320px; max-width: 380px; max-height: 480px;
  overflow: hidden; display: flex; flex-direction: column;
  animation: popIn .12s ease-out;
}
@keyframes popIn { from { opacity:0; transform: translateY(-6px); } to { opacity:1; transform: translateY(0); } }
.popover-head {
  padding: 14px 18px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between; gap: 10px;
}
.popover-head .title { font-size: 0.95rem; font-weight: 700; color: var(--text-dark); }
.popover-head .count { font-size: 0.78rem; color: var(--text-lt); }
.popover-close {
  background: var(--surface-3); border: none; width: 26px; height: 26px;
  border-radius: 7px; cursor: pointer; color: var(--text-med);
  display: flex; align-items: center; justify-content: center; font-size: 1.1rem;
}
.popover-close:hover { background: #DCE6EE; }
.popover-list { overflow-y: auto; flex: 1; padding: 6px 0; }
.popover-item {
  display: flex; align-items: center; gap: 10px;
  padding: 9px 18px; cursor: pointer; transition: background .12s;
  border-left: 3px solid transparent;
}
.popover-item:hover { background: var(--surface-2); }
.popover-item .pop-code { font-size: 0.72rem; color: var(--text-lt); font-weight: 600; min-width: 50px; }
.popover-item .pop-name { font-size: 0.86rem; color: var(--text-dark); flex: 1; line-height: 1.35; }
.popover-item .pop-pct  { font-size: 0.82rem; font-weight: 700; }
.popover-item.cumple    { border-left-color: var(--success); } .popover-item.cumple .pop-pct { color: var(--success); }
.popover-item.riesgo    { border-left-color: var(--warning); } .popover-item.riesgo .pop-pct { color: var(--warning); }
.popover-item.critico   { border-left-color: var(--danger); }  .popover-item.critico .pop-pct { color: var(--danger); }
.popover-item.pendiente { border-left-color: var(--neutral); } .popover-item.pendiente .pop-pct { color: var(--neutral); }
.popover-empty { padding: 28px; text-align: center; color: var(--text-lt); font-size: 0.86rem; }

/* ─── Filters bar ────────────────────────────────────────── */
.filters-bar {
  background: var(--surface); border-top: 1px solid var(--border);
  padding: 12px 24px;
  display: flex; flex-wrap: wrap; gap: 12px; align-items: center;
  margin-bottom: 0; box-shadow: none;
}
.filter-group { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.filter-label {
  font-size: 0.72rem; font-weight: 700; color: var(--text-lt);
  text-transform: uppercase; letter-spacing: 0.08em; white-space: nowrap;
}

/* Tarjetas de perspectiva en la barra de filtros */
.persp-filter-row {
  display: flex; gap: 10px; flex-wrap: nowrap; align-items: stretch;
  flex: 1;
}
.persp-filter-card {
  display: flex; align-items: center; gap: 10px;
  background: var(--surface-2); border: 2px solid var(--border);
  border-radius: 14px; padding: 9px 14px;
  cursor: pointer; transition: all .15s; user-select: none;
  flex: 1; min-width: 0; max-width: none;
  position: relative; overflow: visible;
  box-sizing: border-box;
}
/* Barra inferior activa — pseudo-elemento con clip propio */
.persp-filter-card::after {
  content: ''; position: absolute; bottom: 0; left: 4px; right: 4px;
  height: 3px; opacity: 0; border-radius: 999px;
  transition: opacity .15s;
}
.persp-filter-card:hover { border-color: #C8DFF0; box-shadow: var(--shadow-sm); transform: translateY(-1px); }
.persp-filter-card.active { border-color: transparent; box-shadow: var(--shadow); }
.persp-filter-card.active::after { opacity: 1; }
.pfc-icon {
  width: 34px; height: 34px; border-radius: 9px; flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
  transition: transform .15s;
}
.persp-filter-card:hover .pfc-icon { transform: scale(1.06); }
.pfc-icon svg { width: 16px; height: 16px; }
.pfc-info {
  flex: 1; min-width: 0;
}
.pfc-name  {
  font-size: 0.76rem; font-weight: 700; color: var(--text-dark);
  line-height: 1.25; word-break: normal; hyphens: none;
  white-space: normal; overflow-wrap: break-word;
}
.pfc-count { font-size: 0.63rem; color: var(--text-lt); font-weight: 500; margin-top: 1px; white-space: nowrap; }
.pfc-pct   { font-size: 0.86rem; font-weight: 800; margin-top: 3px; }
.pfc-bar   { height: 3px; border-radius: 999px; background: #E2E8F0; margin-top: 4px; overflow: hidden; }
.pfc-bar-fill { height: 100%; border-radius: 999px; transition: width .5s; }

/* Divider entre grupos */
.filter-divider { width: 1px; background: var(--border); align-self: stretch; flex-shrink: 0; margin: 4px 0; }

/* Chips de año */
.chip {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 6px 14px; border-radius: 999px;
  background: var(--surface-3); color: var(--text-med);
  font-size: 0.82rem; font-weight: 500; cursor: pointer;
  border: 1px solid transparent; transition: all .15s; user-select: none;
}
.chip:hover { background: #E2EDF4; color: var(--text-dark); }
.chip.year.active { background: var(--text-dark); color: white; }
.chip.year.current::before {
  content: ""; width: 6px; height: 6px; border-radius: 50%;
  background: var(--green); display: inline-block;
}
.chip.year.current.active::before { background: #4ADE80; }
.chip svg { width: 15px; height: 15px; flex-shrink: 0; }

/* ─── Matriz card ────────────────────────────────────────── */
/* ─── Matriz layout ──────────────────────────────────────── */
.matriz-layout {
  display: flex; gap: 0; margin-bottom: 28px; align-items: flex-start;
}

/* Card de la tabla — full width */
.matriz-card {
  flex: 1; min-width: 0;
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius-lg); overflow: hidden;
  box-shadow: var(--shadow-sm);
}
.matriz-footer-bar {
  display: flex; align-items: center; justify-content: space-between;
  flex-wrap: wrap; gap: 16px;
  padding: 10px 18px; border-top: 1px solid var(--border);
  background: var(--surface-2);
}
.mleg { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
.mleg-item { display: inline-flex; align-items: center; gap: 6px; font-size: 0.76rem; color: var(--text-med); font-weight: 500; }
.mleg-dot2 { width: 10px; height: 10px; border-radius: 3px; flex-shrink: 0; }
.mleg-dot2.cumple    { background: #15803D; }
.mleg-dot2.riesgo    { background: #D97706; }
.mleg-dot2.critico   { background: #DC2626; }
.mleg-dot2.pendiente { background: #CBD5E1; }
.mtz-last-update { font-size: 0.74rem; color: var(--text-lt); display: flex; align-items: center; gap: 5px; }

/* Scroll */
.matriz-scroll {
  overflow-x: auto; -webkit-overflow-scrolling: touch;
  scroll-behavior: smooth;
}
.matriz-scroll::-webkit-scrollbar { height: 5px; }
.matriz-scroll::-webkit-scrollbar-track { background: var(--surface-2); }
.matriz-scroll::-webkit-scrollbar-thumb { background: #94A3B8; border-radius: 999px; }
.matriz-scroll::-webkit-scrollbar-thumb:hover { background: #64748B; }

/* ─── Matriz tabla ───────────────────────────────────────── */
.matriz-table {
  border-collapse: separate; border-spacing: 0;
  table-layout: fixed;
}
.matriz-table th, .matriz-table td {
  padding: 0; border-bottom: 1px solid var(--border-soft); vertical-align: middle;
  white-space: nowrap;
}

/* ── Encabezados matriz — paleta azul degradada ─────────── */
.matriz-table thead tr { height: 52px; }
.matriz-table thead th {
  padding: 0; font-size: 0.65rem; font-weight: 700;
  color: white; text-transform: uppercase; letter-spacing: 0.07em;
  background: #1976D2;
  border-bottom: none;
  text-align: center; position: sticky; top: 0; z-index: 5;
  transition: background .2s;
  height: 52px; max-height: 52px; overflow: hidden;
  vertical-align: middle;
}
/* Celda de mes: abbr + nombre */
.mtz-th-mes { padding: 0 !important; width: 72px; min-width: 72px; height: 52px; }
.mtz-th-mes-inner {
  display: flex; flex-direction: column; align-items: center;
  padding: 8px 4px 7px; height: 100%;
  border-bottom: 3px solid transparent;
  border-right: 1px solid rgba(255,255,255,0.10);
}
.mtz-th-mes-abbr { font-size: 0.70rem; font-weight: 800; letter-spacing: 0.06em; color: white; }
.mtz-th-mes-full { font-size: 0.57rem; color: rgba(255,255,255,0.72); margin-top: 3px; text-transform: capitalize; font-weight: 500; }
/* Mes actual — azul claro, texto blanco */
.mtz-th-mes.current { background: #42A5F5 !important; }
.mtz-th-mes.current .mtz-th-mes-inner { border-bottom-color: white; }
.mtz-th-mes.current .mtz-th-mes-abbr { color: white; font-size: 0.74rem; font-weight: 900; }
.mtz-th-mes.current .mtz-th-mes-full { color: rgba(255,255,255,0.92); }
/* Año header */
.mtz-th-year {
  background: #1565C0 !important;
  font-size: 0.70rem !important; font-weight: 900 !important;
  color: white !important; padding: 9px 10px !important;
  border-right: 2px solid rgba(255,255,255,0.20) !important;
  letter-spacing: 0.08em !important;
}
.mtz-th-year-current { background: #0D47A1 !important; }
/* Col Indicador */
.mtz-th-ind {
  text-align: center !important; width: 300px; min-width: 300px; max-width: 300px;
  padding: 6px 12px !important; position: sticky; left: 0; z-index: 10;
  background: #1565C0 !important;
  border-right: 3px solid rgba(255,255,255,0.25) !important;
  vertical-align: middle !important;
  overflow: hidden;
}
/* Col Cumpl */
.mtz-th-cumpl {
  width: 88px; min-width: 88px; max-width: 88px;
  background: #1565C0 !important;
  border-left: 3px solid rgba(255,255,255,0.20) !important;
  position: sticky; z-index: 10;
  padding: 0 !important;
}
/* Separador entre años */
.mtz-year-sep { border-left: 2px solid rgba(255,255,255,0.18) !important; }

/* Indicador cell */
.mtz-ind-cell {
  padding: 5px 12px 5px 16px !important; width: 300px; min-width: 300px; max-width: 300px;
  position: sticky; left: 0; z-index: 4;
  background: white; border-right: 3px solid var(--border) !important;
}
.mtz-ind-cell::after {
  content: ''; position: absolute; right: -1px; top: 0; bottom: 0;
  width: 8px; background: linear-gradient(90deg, rgba(0,0,0,0.04), transparent);
  pointer-events: none;
}
.mtz-ind-code { font-size: 0.64rem; font-weight: 800; letter-spacing: 0.06em; margin-bottom: 1px; }
.mtz-ind-name { font-size: 0.82rem; font-weight: 600; color: var(--text-dark); line-height: 1.3; white-space: normal; }
.mtz-ind-formula { font-size: 0.62rem; color: var(--text-lt); font-style: italic; margin-top: 3px; line-height: 1.3; white-space: normal; }

/* Columna META anual */
.mtz-th-meta {
  width: 62px; min-width: 62px; max-width: 62px;
  background: #0D47A1 !important;
  padding: 5px 3px !important;
  font-size: 0.58rem; font-weight: 800; letter-spacing: 0.05em;
  color: rgba(255,255,255,0.95); text-align: center;
  vertical-align: middle; overflow: hidden;
  border-left: 2px solid rgba(255,255,255,0.30) !important;
  border-right: 1px solid rgba(255,255,255,0.15) !important;
  line-height: 1.3;
}
.mtz-meta-cell {
  width: 62px; min-width: 62px; max-width: 62px;
  text-align: center; padding: 4px 3px !important;
  vertical-align: middle;
  border-left: 2px solid #BFDBFE !important;
  border-right: 1px solid var(--border) !important;
  background: #EFF6FF;
}
.mtz-meta-val {
  font-size: 0.68rem; font-weight: 700; color: #1D4ED8;
  white-space: nowrap;
}
/* Responsive móvil: ocultar meta */
@media (max-width: 480px) {
  .mtz-th-meta, .mtz-meta-cell { display: none; }
}

/* Columna tendencia */
.mtz-th-tend {
  width: 80px; min-width: 80px; max-width: 80px;
  background: #1565C0 !important;
  border-right: 2px solid rgba(255,255,255,0.20) !important;
  padding: 6px 4px !important;
  position: sticky; z-index: 10;
  font-size: 0.60rem; font-weight: 700; letter-spacing: 0.06em;
  color: white; text-align: center; vertical-align: middle;
  overflow: hidden;
}
.mtz-tend-cell {
  width: 80px; min-width: 80px; max-width: 80px;
  text-align: center; padding: 6px 4px !important;
  background: var(--surface-2);
  border-right: 2px solid var(--border) !important;
  vertical-align: middle;
}
/* mtz-tend-badge eliminado — ahora es texto plano */
/* Responsive móvil: ocultar tendencia */
@media (max-width: 480px) {
  .mtz-th-tend, .mtz-tend-cell { display: none; }
}

/* Semáforo mensual */
.mtz-month-cell { text-align: center; padding: 3px 3px !important; width: 72px; min-width: 72px; }
.mtz-year-sep { border-left: 2px solid var(--border) !important; }
.mtz-light {
  width: 62px; border-radius: 6px;
  display: inline-flex; flex-direction: column;
  align-items: center; justify-content: center;
  margin: 0 auto; cursor: pointer;
  transition: transform .12s, box-shadow .12s; position: relative;
  padding: 3px 2px 2px;
  gap: 1px;
}
.mtz-light:hover { transform: scale(1.08); box-shadow: var(--shadow); z-index: 2; }
.mtz-light svg { display: none; }
.mtz-val {
  font-size: 0.70rem; font-weight: 700; letter-spacing: -0.01em;
  white-space: nowrap; line-height: 1.2;
}
.mtz-estado {
  font-size: 0.52rem; font-weight: 600; letter-spacing: 0.01em;
  white-space: nowrap; opacity: 0.80; line-height: 1.1;
  overflow: hidden; max-width: 60px; text-overflow: ellipsis;
}
/* Pendiente: solo el guión, sin etiqueta */
.mtz-light.pendiente .mtz-estado { display: none; }
.mtz-light.cumple    { background: #B9F9CE; color: #15803D; }
.mtz-light.riesgo    { background: #FFD3A7; color: #D97706; }
.mtz-light.critico   { background: #FEE2E2; color: #DC2626; }
.mtz-light.pendiente { background: #F1F5F9; color: #94A3B8; }

/* Cumplimiento col */
.mtz-cumpl-cell {
  text-align: center; padding: 8px 8px !important;
  width: 88px; min-width: 88px; max-width: 88px;
  background: var(--surface-2);
  border-left: 3px solid var(--border) !important;
  position: sticky; z-index: 4;
}
.mtz-cumpl-pct { font-size: 0.86rem; font-weight: 800; }
.mtz-cumpl-pct.cumple    { color: #16A34A; }
.mtz-cumpl-pct.riesgo    { color: #D97706; }
.mtz-cumpl-pct.critico   { color: #DC2626; }
.mtz-cumpl-pct.pendiente { color: #CBD5E1; }
.mtz-mini-bar { width: 100%; height: 3px; border-radius: 999px; background: #E2E8F0; margin-top: 4px; overflow: hidden; }
.mtz-mini-bar-fill { height: 100%; border-radius: 999px; transition: width .6s; }

/* Rows */
.matriz-table tbody tr { transition: background .08s; }
.matriz-table tbody tr:hover > td { background: #F0F7FF !important; }
.mtz-first-row > td { border-top: 2px solid var(--border) !important; }

/* Responsive sidebar */



/* Tooltip */

/* ─── Gauges de cumplimiento ─────────────────────────────── */
.sticky-top {
  position: relative; z-index: 70;
  background: var(--bg);
}

/* ─── Barra compacta (aparece al bajar, reemplaza el encabezado fijo) ─── */
.cbar {
  position: fixed; top: 0; left: 0; right: 0; z-index: 85;
  background: linear-gradient(120deg, #0657A4 0%, #0B88C7 45%, #07A89E 80%, #2DC98E 100%);
  color: white; box-shadow: 0 4px 16px rgba(6,87,164,0.28);
  transform: translateY(-110%); transition: transform .2s ease;
}
.cbar.show { transform: none; }
.cbar-inner {
  max-width: 1400px; margin: 0 auto; padding: 8px 28px;
  display: flex; align-items: center; gap: 18px; min-height: 52px;
}
.cbar-title { font-size: 0.98rem; font-weight: 800; white-space: nowrap; display: flex; align-items: center; gap: 8px; cursor: pointer; }
.cbar-kpis { display: flex; gap: 6px; flex-wrap: nowrap; }
.cbar-kpi {
  display: inline-flex; align-items: center; gap: 5px;
  background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.25);
  border-radius: 999px; padding: 3px 10px; font-size: 0.74rem; font-weight: 700; white-space: nowrap;
}
.cbar-kpi.clickable { cursor: pointer; }
.cbar-kpi.clickable:hover { background: rgba(255,255,255,0.28); }
.cbar-dot { width: 8px; height: 8px; border-radius: 50%; }
.cbar-global {
  background: white; border-radius: 999px; padding: 3px 12px;
  font-size: 0.78rem; font-weight: 900; white-space: nowrap;
}
.cbar-global span { font-weight: 600; color: var(--text-lt); font-size: 0.68rem; margin-right: 4px; }
.cbar-persp {
  background: rgba(255,255,255,0.22); border-radius: 999px; padding: 3px 10px;
  font-size: 0.72rem; font-weight: 700; white-space: nowrap; cursor: pointer;
}
.cbar-years { display: flex; gap: 4px; margin-left: auto; flex-wrap: nowrap; }
.cbar-year {
  padding: 4px 10px; border-radius: 999px; font-size: 0.74rem; font-weight: 700;
  color: white; background: rgba(255,255,255,0.12); cursor: pointer; white-space: nowrap;
}
.cbar-year:hover { background: rgba(255,255,255,0.26); }
.cbar-year.active { background: white; color: #0657A4; }
@media (max-width: 1100px) { .cbar-kpi .cbar-lbl { display: none; } }
@media (max-width: 768px) { .cbar { display: none; } }
.gauges-section { margin-bottom: 0; }

/* ── Panel de cumplimiento: 1 fila horizontal compacta ── */
.gauges-strip {
  display: flex; align-items: stretch; gap: 0;
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius-lg); overflow: hidden;
  box-shadow: var(--shadow-sm);
}

/* Bloque global — izquierda */
.gs-global {
  display: flex; align-items: center; gap: 16px;
  padding: 14px 22px 14px 20px;
  background: linear-gradient(135deg,#F0FDF4,#F8FAFC);
  border-right: 1px solid var(--border);
  flex: 0 0 auto; min-width: 220px;
}
.gs-global-donut { flex-shrink: 0; }
.gs-global-info { display: flex; flex-direction: column; gap: 2px; }
.gs-global-label { font-size: 0.62rem; font-weight: 800; color: var(--text-lt); text-transform: uppercase; letter-spacing: 0.09em; }
.gs-global-pct   { font-size: 1.60rem; font-weight: 900; letter-spacing: -0.03em; line-height: 1; }
.gs-global-sub   { font-size: 0.70rem; color: var(--text-lt); margin-top: 1px; }

/* Separador de título */
.gs-sep {
  width: 1px; background: var(--border); flex-shrink: 0; align-self: stretch;
}

/* Cards de año — inline */
.gs-years {
  display: flex; flex: 1; align-items: stretch;
}
.gs-year {
  flex: 1; display: flex; flex-direction: column; align-items: center;
  justify-content: center; gap: 4px;
  padding: 10px 8px; cursor: pointer;
  border-right: 1px solid var(--border);
  transition: background .12s;
  position: relative;
}
.gs-year:last-child { border-right: none; }
.gs-year:hover { background: var(--surface-2); }
.gs-year.active { background: #EFF7FF; }
.gs-year.active::after {
  content: ''; position: absolute; bottom: 0; left: 20%; right: 20%;
  height: 3px; border-radius: 3px 3px 0 0;
}
.gs-year-label {
  font-size: 0.62rem; font-weight: 800; color: var(--text-lt);
  text-transform: uppercase; letter-spacing: 0.07em;
  display: flex; align-items: center; gap: 5px;
}
.gs-year-label .curr-dot {
  width: 5px; height: 5px; border-radius: 50%;
  background: #22C55E; display: inline-block;
}
.gs-year-donut { flex-shrink: 0; }
.gs-year-sub { font-size: 0.62rem; color: var(--text-lt); }

/* Filtros barra */
.filters-bar {
  background: var(--surface); border-top: 1px solid var(--border);
  padding: 9px 24px;
  display: flex; flex-wrap: wrap; gap: 14px; align-items: center;
  margin-bottom: 0; box-shadow: none;
  border-radius: 0; border-left: none; border-right: none;
}


.mtz-tooltip {
  position: fixed; z-index: 300;
  background: rgba(15,23,42,0.96); border-radius: var(--radius-sm);
  padding: 10px 14px; min-width: 180px; max-width: 250px;
  pointer-events: none; opacity: 0; transition: opacity .12s; color: white;
  font-size: 0.80rem; box-shadow: var(--shadow-lg);
}
.mtz-tooltip.visible { opacity: 1; }
.mtz-tooltip .t-month { font-size: 0.66rem; opacity: 0.65; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 3px; }
.mtz-tooltip .t-val   { font-size: 1.05rem; font-weight: 800; }
.mtz-tooltip .t-meta  { opacity: 0.70; margin-top: 2px; font-size: 0.76rem; }
.mtz-tooltip .t-cumpl { font-weight: 700; margin-top: 4px; font-size: 0.80rem; }
.mtz-tooltip .t-cumpl.cumple  { color: #4ADE80; }
.mtz-tooltip .t-cumpl.riesgo  { color: #FDE68A; }
.mtz-tooltip .t-cumpl.critico { color: #FCA5A5; }

/* ─── Globo de cálculo en composición ──────────────────── */
/* ─── Tooltip de cálculo paso a paso ─────────────────────── */
.calc-tooltip {
  position: fixed;
  z-index: 9999;
  background: #0F172A;
  color: white;
  border-radius: 14px;
  padding: 16px 18px;
  max-width: 340px; min-width: 260px;
  box-shadow: 0 16px 48px rgba(0,0,0,0.35), 0 4px 12px rgba(0,0,0,0.20);
  font-size: 0.80rem;
  pointer-events: auto;
  animation: ctIn .15s cubic-bezier(.34,1.4,.64,1);
}
@keyframes ctIn { from { opacity:0; transform:translateY(6px) scale(.96); } to { opacity:1; transform:translateY(0) scale(1); } }
/* Punta del tooltip */
.calc-tooltip::after {
  content: '';
  position: absolute; bottom: -7px; left: 28px;
  width: 14px; height: 14px;
  background: #0F172A;
  transform: rotate(45deg);
  border-radius: 2px;
}
.calc-tooltip.above::after {
  bottom: auto; top: -7px;
  transform: rotate(45deg);
}
.ct-month  { font-size: 0.65rem; color: rgba(255,255,255,0.55); font-weight: 700; text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 6px; }
.ct-formula{ font-family: "SF Mono",Menlo,Consolas,monospace; font-size: 0.74rem; color: rgba(255,255,255,0.70); background: rgba(255,255,255,0.08); border-radius: 6px; padding: 5px 9px; margin-bottom: 10px; line-height: 1.5; }
.ct-steps  { display: flex; flex-direction: column; gap: 5px; margin-bottom: 10px; }
.ct-step   { display: flex; align-items: flex-start; gap: 7px; font-size: 0.76rem; color: rgba(255,255,255,0.80); line-height: 1.45; }
.ct-step-n { width: 16px; height: 16px; border-radius: 50%; flex-shrink: 0; background: rgba(255,255,255,0.15); color: white; font-size: 0.58rem; font-weight: 800; display: flex; align-items: center; justify-content: center; margin-top: 1px; }
.ct-step strong { color: white; }
.ct-result {
  display: flex; align-items: center; justify-content: space-between;
  border-radius: 8px; padding: 8px 12px; margin-bottom: 8px;
}
.ct-result.cumple   { background: rgba(34,197,94,0.20); }
.ct-result.riesgo   { background: rgba(245,158,11,0.20); }
.ct-result.critico  { background: rgba(239,68,68,0.20); }
.ct-result.pendiente{ background: rgba(255,255,255,0.08); }
.ct-result-label { font-size: 0.74rem; font-weight: 600; }
.ct-result-label.cumple   { color: #4ADE80; }
.ct-result-label.riesgo   { color: #FDE68A; }
.ct-result-label.critico  { color: #FCA5A5; }
.ct-result-label.pendiente{ color: rgba(255,255,255,0.50); }
.ct-result-pct { font-size: 1.20rem; font-weight: 900; letter-spacing: -0.02em; }
.ct-result-pct.cumple   { color: #4ADE80; }
.ct-result-pct.riesgo   { color: #FDE68A; }
.ct-result-pct.critico  { color: #FCA5A5; }
.ct-interp { font-size: 0.70rem; color: rgba(255,255,255,0.55); line-height: 1.5; font-style: italic; }
.ct-close-hint { font-size: 0.64rem; color: rgba(255,255,255,0.35); margin-top: 8px; text-align: right; }
.cumpl-clickable {
  cursor: pointer; display: inline-block; transition: opacity .12s;
}
.cumpl-clickable:hover { opacity: 0.80; }
/* Detalle de cálculo inline dentro de la tabla */
.calc-inline-detail {
  padding: 14px 20px; font-size: 0.78rem;
  display: flex; flex-wrap: wrap; gap: 20px; align-items: flex-start;
}
.cid-formula {
  background: #EFF6FF; border-left: 3px solid var(--primary);
  border-radius: 7px; padding: 8px 12px;
  font-family: "SF Mono",Menlo,Consolas,monospace;
  font-size: 0.74rem; color: #1E40AF; white-space: nowrap;
}
.cid-steps { display: flex; flex-direction: column; gap: 4px; }
.cid-step  { display: flex; align-items: center; gap: 7px; color: var(--text-med); }
.cid-step-n{ width: 18px; height: 18px; border-radius: 50%; flex-shrink: 0;
  background: var(--primary); color: white;
  font-size: 0.58rem; font-weight: 800;
  display: flex; align-items: center; justify-content: center; }
.cid-step strong { color: var(--text-dark); }
.cid-result {
  display: inline-flex; align-items: center; gap: 10px;
  border-radius: 8px; padding: 8px 16px; font-weight: 700;
}
.cid-result.cumple   { background: #DCFCE7; color: #166534; }
.cid-result.riesgo   { background: #FEF3C7; color: #92400E; }
.cid-result.critico  { background: #FEE2E2; color: #991B1B; }
.cid-result.pendiente{ background: #F1F5F9; color: #64748B; }
.cid-result-pct { font-size: 1.15rem; font-weight: 900; }
.cid-interp { font-size: 0.72rem; color: var(--text-lt); font-style: italic; margin-top: 4px; }
/* Overlay invisible solo para cerrar */
.calc-tooltip-overlay {
  position: fixed; inset: 0; z-index: 9998; display: none;
}
.calc-tooltip-overlay.active { display: block; }

/* ─── Perspectivas (debajo de la matriz) ─────────────────── */
.section-title {
  font-size: 0.72rem; font-weight: 700; color: var(--text-lt);
  text-transform: uppercase; letter-spacing: 0.10em;
  margin: 0 0 14px;
  display: flex; align-items: center; gap: 8px;
}
.section-title::before {
  content: ""; width: 3px; height: 14px;
  background: linear-gradient(180deg, var(--primary) 0%, var(--teal) 100%);
  border-radius: 999px;
}
.perspective { margin-bottom: 36px; }
.persp-header {
  display: flex; align-items: center; gap: 16px;
  padding: 20px 26px; border-radius: var(--radius-lg);
  color: white; margin-bottom: 14px;
  box-shadow: var(--shadow); position: relative; overflow: hidden;
}
.persp-header::after {
  content: ""; position: absolute; right: -40px; top: -40px;
  width: 180px; height: 180px; border-radius: 50%;
  background: rgba(255,255,255,0.08); pointer-events: none;
}
.persp-header .persp-icon-wrap {
  width: 54px; height: 54px; border-radius: 14px;
  background: rgba(255,255,255,0.20);
  display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  border: 1px solid rgba(255,255,255,0.28);
}
.persp-header .persp-icon-wrap svg { width: 28px; height: 28px; color: white; }
.persp-header .persp-titles { flex: 1; }
.persp-header .persp-eyebrow { font-size: 0.72rem; font-weight: 600; opacity: 0.85; text-transform: uppercase; letter-spacing: 0.10em; }
.persp-header .persp-title   { font-weight: 700; font-size: 1.40rem; margin-top: 2px; letter-spacing: -0.01em; }
.persp-header .persp-stats   { display: flex; gap: 10px; align-items: center; font-size: 0.82rem; }
.persp-stat {
  background: rgba(255,255,255,0.18); padding: 5px 12px; border-radius: 999px;
  font-weight: 600; border: 1px solid rgba(255,255,255,0.20); white-space: nowrap;
}
.persp-stat strong { font-size: 1.02rem; font-weight: 700; }
.persp-objetivo {
  background: var(--surface); border: 1px solid var(--border);
  border-left: 4px solid currentColor;
  padding: 12px 18px; border-radius: var(--radius-sm);
  margin-bottom: 16px; font-size: 0.86rem; color: var(--text-med); line-height: 1.6;
}
.persp-objetivo strong { color: var(--text-dark); font-weight: 600; }

/* ─── Indicator cards ────────────────────────────────────── */
.ind-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); gap: 14px; }
.ind-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(340px, 1fr)); gap: 16px; }

/* ─── Indicator card ─────────────────────────────────────── */
.ind-card {
  background: var(--surface);
  border: 1.5px solid var(--border);
  border-radius: 18px;
  box-shadow: var(--shadow-xs);
  display: flex; flex-direction: column;
  cursor: pointer; position: relative; overflow: hidden;
  transition: transform .18s, box-shadow .18s, border-color .18s;
}
.ind-card:hover {
  transform: translateY(-3px);
  box-shadow: 0 8px 28px rgba(15,23,42,0.10);
  border-color: #C8DFF0;
}

/* Barra lateral de color de perspectiva */
.ind-card .accent-bar {
  position: absolute; left: 0; top: 0; bottom: 0; width: 5px; border-radius: 18px 0 0 18px;
}

/* Efecto iluminación al navegar desde la matriz */
@keyframes cardGlow {
  0%   { box-shadow: 0 0 0 0 rgba(255,220,50,0.7), 0 0 0 0 rgba(255,180,0,0.4); background: white; }
  20%  { box-shadow: 0 0 0 6px rgba(255,220,50,0.5), 0 0 30px rgba(255,200,0,0.25); background: #FFFEF0; }
  60%  { box-shadow: 0 0 0 10px rgba(255,220,50,0.15), 0 0 40px rgba(255,200,0,0.10); background: #FFFEF8; }
  100% { box-shadow: var(--shadow-xs); background: white; }
}
.ind-card-highlight {
  animation: cardGlow 2.2s cubic-bezier(.4,0,.2,1) forwards;
  border-color: #F59E0B !important;
  z-index: 2;
}

/* Head */
.ind-card-head {
  padding: 18px 20px 14px;
  display: flex; justify-content: space-between; align-items: flex-start; gap: 10px;
  padding-left: 22px;
}
.ind-code-badge {
  font-size: 0.68rem; font-weight: 800; letter-spacing: 0.10em;
  text-transform: uppercase; margin-bottom: 4px;
}
.ind-name {
  font-size: 1.05rem; font-weight: 800; color: var(--text-dark);
  line-height: 1.25; letter-spacing: -0.01em; margin: 0 0 5px;
}
.ind-ini {
  display: inline-flex; align-items: center; gap: 5px;
  font-size: 0.74rem; color: var(--text-med);
}
.status-pill {
  display: inline-flex; align-items: center; gap: 5px;
  padding: 5px 11px; border-radius: 999px;
  font-size: 0.70rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.05em; white-space: nowrap; flex-shrink: 0;
}
.status-pill.cumple    { background: #DCFCE7; color: #16A34A; border: 1px solid #BBF7D0; }
.status-pill.riesgo    { background: #FEF3C7; color: #D97706; border: 1px solid #FDE68A; }
.status-pill.critico   { background: #FEE2E2; color: #DC2626; border: 1px solid #FECACA; }
.status-pill.pendiente { background: var(--surface-3); color: var(--neutral); border: 1px solid var(--border); }

/* 3 métricas premium */
.ind-kpis {
  display: grid; grid-template-columns: 1fr 1fr 1fr;
  gap: 0; margin: 0 16px 14px; border-radius: 14px; overflow: hidden;
  border: 1px solid var(--border);
}
.ind-kpi {
  padding: 12px 10px; background: var(--surface-2);
  display: flex; flex-direction: column; align-items: center; text-align: center;
  position: relative;
}
.ind-kpi + .ind-kpi { border-left: 1px solid var(--border); }
.ind-kpi-icon {
  width: 30px; height: 30px; border-radius: 9px;
  display: flex; align-items: center; justify-content: center; margin-bottom: 6px;
}
.ind-kpi-icon svg { width: 15px; height: 15px; }
.ind-kpi-lbl { font-size: 0.60rem; font-weight: 700; color: var(--text-lt); text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 3px; line-height: 1.3; }
.ind-kpi-val { font-size: 1.05rem; font-weight: 800; color: var(--text-dark); letter-spacing: -0.01em; }
.ind-kpi-val.cumple    { color: #16A34A; }
.ind-kpi-val.riesgo    { color: #D97706; }
.ind-kpi-val.critico   { color: #DC2626; }
.ind-kpi-val.pendiente { color: var(--neutral); }
.ind-kpi-sub { font-size: 0.64rem; color: var(--text-lt); margin-top: 2px; }

/* Mini chart */
.chart-wrap { position: relative; height: 90px; margin: 0 16px 12px; pointer-events: none; }
.chart-wrap canvas { pointer-events: none; }

/* Metas por año (modo todos los años) */
.ind-metas-yr {
  margin: 0 16px 10px; display: flex; gap: 4px;
}
.ind-meta-yr-cell {
  flex: 1; background: var(--surface-2); border: 1px solid var(--border);
  border-radius: 8px; padding: 7px 6px; text-align: center;
}
.ind-meta-yr-cell.current { background: #EFF7FF; border-color: #BFDBFE; }
.ind-meta-yr-lbl { font-size: 0.58rem; font-weight: 700; color: var(--text-lt); text-transform: uppercase; letter-spacing: 0.05em; }
.ind-meta-yr-val { font-size: 0.78rem; font-weight: 800; color: var(--text-dark); margin-top: 2px; }
.ind-meta-yr-cumpl { font-size: 0.65rem; font-weight: 700; margin-top: 1px; }
.ind-meta-yr-cumpl.cumple    { color: #16A34A; }
.ind-meta-yr-cumpl.riesgo    { color: #D97706; }
.ind-meta-yr-cumpl.critico   { color: #DC2626; }
.ind-meta-yr-cumpl.pendiente { color: #CBD5E1; }

/* Footer con tendencia / avance / responsable */
.ind-foot {
  display: grid; grid-template-columns: 1fr 1fr 1fr;
  gap: 0; border-top: 1px solid var(--border);
}
.foot-kpi {
  padding: 10px 12px; display: flex; align-items: center; gap: 8px;
}
.foot-kpi + .foot-kpi { border-left: 1px solid var(--border); }
.foot-kpi-icon {
  width: 28px; height: 28px; border-radius: 8px; flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
}
.foot-kpi-icon svg { width: 13px; height: 13px; }
.foot-kpi-info .l { font-size: 0.60rem; color: var(--text-lt); font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; }
.foot-kpi-info .v { font-size: 0.80rem; font-weight: 700; color: var(--text-dark); margin-top: 1px; line-height: 1.2; }
.foot-kpi-info .s { font-size: 0.64rem; color: var(--text-lt); }

/* Alerta */
.ind-alert {
  margin: 0 14px 14px; padding: 9px 12px; border-radius: 10px;
  font-size: 0.78rem; font-weight: 600;
  display: flex; align-items: flex-start; gap: 8px; line-height: 1.4;
  pointer-events: none;
}
.ind-alert.ok       { background: #F0FDF4; color: #166534; border: 1px solid #BBF7D0; }
.ind-alert.warn     { background: #FFFBEB; color: #92400E; border: 1px solid #FDE68A; }
.ind-alert.critical { background: #FFF5F5; color: #991B1B; border: 1px solid #FECACA; }
.ind-alert.pending  { background: var(--surface-2); color: var(--text-med); border: 1px solid var(--border); }
.ind-alert .al-icon { font-size: 1rem; flex-shrink: 0; margin-top: 1px; }
.ind-alert .al-body .al-title { font-weight: 800; }
.ind-alert .al-body .al-sub   { font-size: 0.72rem; opacity: 0.8; margin-top: 2px; }

/* ─── Modal ──────────────────────────────────────────────── */
.modal-overlay {
  position: fixed; inset: 0; background: rgba(8,25,48,0.60);
  backdrop-filter: blur(4px);
  z-index: 100; display: none; align-items: flex-start; justify-content: center;
  padding: 32px 20px; overflow-y: auto; animation: fadeIn .18s;
}
.modal-overlay.active { display: flex; }
@keyframes fadeIn { from { opacity: 0; transform:translateY(8px); } to { opacity: 1; transform:translateY(0); } }
.modal {
  background: var(--surface); border-radius: 20px;
  max-width: 1060px; width: 100%;
  box-shadow: 0 24px 64px rgba(8,25,48,0.22), 0 4px 12px rgba(8,25,48,0.08);
  overflow: hidden;
}

/* Head con gradiente de la perspectiva */
.modal-head {
  padding: 24px 28px 20px;
  display: flex; gap: 18px; justify-content: space-between; align-items: flex-start;
  position: relative; overflow: hidden;
}
.modal-head-bg {
  position: absolute; inset: 0; opacity: 0.07; pointer-events: none;
}
.modal-head .m-code {
  font-size: 0.70rem; color: var(--text-lt); font-weight: 700;
  letter-spacing: 0.10em; text-transform: uppercase; margin-bottom: 4px;
}
.modal-head .m-name {
  font-size: 1.40rem; font-weight: 800; color: var(--text-dark);
  letter-spacing: -0.02em; line-height: 1.2; margin-bottom: 6px;
}
.modal-head .m-tags { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
.m-tag {
  display: inline-flex; align-items: center; gap: 5px;
  padding: 4px 10px; border-radius: 999px; font-size: 0.74rem; font-weight: 600;
  background: var(--surface-3); color: var(--text-med); border: 1px solid var(--border);
}
.m-tag.tend-reducir { background:#FEE2E2; color:#991B1B; border-color:#FECACA; }
.m-tag.tend-crecer  { background:#DCFCE7; color:#166534; border-color:#BBF7D0; }
.m-tag.tend-mantener{ background:#EFF6FF; color:#1D4ED8; border-color:#BFDBFE; }
.modal-close {
  background: var(--surface-2); border: 1px solid var(--border); width: 34px; height: 34px;
  border-radius: 10px; cursor: pointer; font-size: 1.2rem; color: var(--text-med);
  display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  transition: background .12s;
}
.modal-close:hover { background: #E2EDF4; color: var(--text-dark); }

.modal-body { padding: 0 28px 28px; }

/* Grid superior: info + cumplimiento por año */
.modal-top-grid {
  display: grid; grid-template-columns: 1fr 1.6fr; gap: 18px; margin-bottom: 20px;
  padding-top: 6px;
}


.m-info-card {
  background: var(--surface-2); border: 1px solid var(--border);
  border-radius: 14px; padding: 16px 18px; display: flex; flex-direction: column; gap: 14px;
}
.m-section-label {
  font-size: 0.66rem; font-weight: 800; color: var(--text-lt);
  text-transform: uppercase; letter-spacing: 0.10em;
  display: flex; align-items: center; gap: 6px; margin-bottom: 2px;
}
.m-section-label::before {
  content:''; width:3px; height:12px; border-radius:2px;
  background: linear-gradient(180deg, var(--primary), var(--teal));
}
.m-formula {
  font-family: "SF Mono", Menlo, Consolas, monospace;
  font-size: 0.82rem; color: var(--text-dark);
  background: white; border: 1px solid var(--border);
  border-left: 3px solid var(--primary);
  padding: 10px 13px; border-radius: 8px; line-height: 1.5;
}
.m-desc { font-size: 0.84rem; color: var(--text-med); line-height: 1.65; }
.m-attrs { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.m-attr { background: white; border: 1px solid var(--border); border-radius: 10px; padding: 10px 12px; }
.m-attr .lbl { font-size: 0.64rem; color: var(--text-lt); font-weight: 700; text-transform: uppercase; letter-spacing: 0.07em; }
.m-attr .val { font-size: 0.90rem; font-weight: 700; color: var(--text-dark); margin-top: 2px; }

/* Cumplimiento por año */
.m-year-panel {
  background: var(--surface-2); border: 1px solid var(--border);
  border-radius: 14px; padding: 16px 18px;
}
.m-year-grid { display: grid; grid-template-columns: repeat(5,1fr); gap: 8px; margin-top: 10px; }
.m-year-cell {
  background: white; border: 1px solid var(--border);
  border-radius: 12px; padding: 12px 8px; text-align: center;
  display: flex; flex-direction: column; align-items: center; gap: 4px;
  transition: transform .12s;
}
.m-year-cell:hover { transform: translateY(-2px); }
.m-year-cell.current {
  border: 2px solid var(--primary);
  box-shadow: 0 0 0 3px rgba(8,118,178,0.10);
  background: #F0F7FF;
}
.m-year-cell .y-label {
  font-size: 0.68rem; font-weight: 700; color: var(--text-lt);
  text-transform: uppercase; letter-spacing: 0.06em;
}
.m-year-cell.current .y-label { color: var(--primary); }
.m-year-cell .y-donut { margin: 4px 0; }
.m-year-cell .y-meta  { font-size: 0.68rem; font-weight: 600; color: #0876B2; opacity: 0.78; }
.m-year-cell .y-result{ font-size: 0.70rem; font-weight: 600; color: var(--text-med); }

/* Estadísticas del gráfico */
.m-chart-stats {
  display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 12px;
}
.m-stat-pill {
  background: var(--surface-2); border: 1px solid var(--border);
  border-radius: 10px; padding: 9px 14px;
  display: flex; flex-direction: column; gap: 2px; flex: 1; min-width: 100px;
}
.m-stat-pill .sp-label { font-size: 0.65rem; color: var(--text-lt); font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; }
.m-stat-pill .sp-val   { font-size: 1.02rem; font-weight: 800; color: var(--text-dark); }
.m-stat-pill .sp-sub   { font-size: 0.70rem; color: var(--text-lt); }
.m-stat-pill.positive .sp-val { color: #16A34A; }
.m-stat-pill.negative .sp-val { color: var(--danger); }
.m-stat-pill.neutral  .sp-val { color: var(--primary); }

.detail-chart-wrap { position: relative; height: 320px; }

/* Sábana mensual */
.tabla-wrap { overflow-x: auto; border-radius: 12px; border: 1px solid var(--border); }
.tabla { width: 100%; border-collapse: collapse; font-size: 0.80rem; }
.tabla th, .tabla td { padding: 8px 10px; border-bottom: 1px solid var(--border-soft); text-align: center; }
.tabla th { background: var(--surface-2); color: var(--text-lt); font-weight: 700;
  font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.05em; position: sticky; top: 0; }
.tabla td:first-child, .tabla th:first-child { text-align: left; font-weight: 700; padding-left: 14px; }
.tabla tbody tr:hover > td { background: #F0F7FF; }
.tabla tr.current-year > td { background: #EFF7FE; }
.tabla tr.current-year > td:first-child { font-weight: 800; color: var(--primary); }
.tabla .cell-has-data { font-weight: 600; }
.tabla .cell-cumple    { background: #F0FDF4; color: #16A34A; }
.tabla .cell-riesgo    { background: #FFFBEB; color: #D97706; }
.tabla .cell-critico   { background: #FFF5F5; color: #DC2626; }

.warnings-list { background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 10px;
  padding: 10px 14px; font-size: 0.80rem; color: #92400E; margin-bottom: 16px; }
.warnings-list ul { margin: 4px 0 0 16px; padding: 0; }

.m-chart-section { margin-bottom: 20px; }
.chart-win-btn:hover { border-color: var(--primary) !important; color: var(--primary) !important; }
#chart-slider { -webkit-appearance: none; appearance: none; height: 4px; border-radius: 999px;
  background: linear-gradient(to right, var(--primary) 0%, var(--primary) var(--val,0%), #E2E8F0 var(--val,0%)); outline: none; }
#chart-slider::-webkit-slider-thumb { -webkit-appearance: none; width: 16px; height: 16px;
  border-radius: 50%; background: var(--primary); border: 2.5px solid white;
  box-shadow: 0 1px 4px rgba(8,118,178,0.35); cursor: pointer; transition: transform .15s; }
#chart-slider::-webkit-slider-thumb:hover { transform: scale(1.2); }
#chart-prev:hover, #chart-next:hover { border-color: var(--primary) !important; color: var(--primary) !important; }
.footer { margin-top: 40px; padding: 22px 32px; text-align: center; font-size: 0.76rem; color: var(--text-lt); }

/* ─── Responsive ─────────────────────────────────────────── */
/* ── Responsive tablet ── */


/* ── Responsive móvil ── */

.hidden { display: none !important; }

/* ════════════════════════════════════════════════════════════
   RESPONSIVE — Mobile First
   ════════════════════════════════════════════════════════════ */

/* Tablet (≤768px) */
@media (max-width: 768px) {
  /* En tablet y móvil: sin sticky — evita pérdida de espacio visual */
  .header { position: relative !important; top: auto !important; }
  .sticky-top { position: relative !important; top: auto !important; box-shadow: none; }
  .header-inner { padding: 14px 16px; gap: 12px; }
  .header-brand { display: none; }
  .header-titles h1 { font-size: 1.18rem; }
  .header-meta { display: none; }
  .btn-ruta-estrategica .xl-label { display: none; }
  .btn-ruta-estrategica { padding: 9px 12px; }
  .btn-datos-origen .xl-label { display: none; }
  .btn-datos-origen { padding: 9px 12px; }
  .btn-excel-consolidado .xl-label { display: none; }
  .btn-excel-consolidado { padding: 9px 12px; }
  .header-kpis { padding: 0 16px 12px; gap: 7px; flex-wrap: wrap; }
  .hkpi { flex: 1; min-width: 82px; padding: 9px 10px; }
  .hkpi-val { font-size: 1.28rem; }
  .sticky-top > * { padding-left: 16px !important; padding-right: 16px !important; }
  .gauges-strip { overflow-x: auto; }
  .gs-global { min-width: 165px; padding: 10px 14px; }
  .filters-bar { padding: 8px 16px; flex-wrap: nowrap; overflow-x: auto; gap: 10px; }
  .filter-divider { display: none; }
  #year-chips { flex-wrap: nowrap; overflow-x: auto; }
  .persp-filter-card { min-width: 0; padding: 8px 10px; }
  .pfc-count { display: none; }
  .matriz-layout { margin-bottom: 20px; }
  .persp-header { padding: 14px 18px; }
  .persp-header .persp-stats { display: none; }
  .ind-grid { grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 12px; }
  .container { padding: 0 16px 24px; }
  .modal-top-grid { grid-template-columns: 1fr; gap: 14px; }
  .m-year-grid { grid-template-columns: repeat(5,1fr); gap: 6px; }
}

/* Móvil (≤480px) */
@media (max-width: 480px) {

  /* ── Header compacto ── */
  .header { position: relative !important; top: auto !important; }
  .header-inner { padding: 10px 12px; gap: 8px; }
  .header-brand { display: none; }
  .header-titles h1 { font-size: 1.02rem; }
  .header-meta { display: none; }
  .btn-ruta-estrategica { padding: 8px 10px; gap: 6px; }
  .btn-ruta-estrategica .xl-label { display: none; }
  .btn-ruta-estrategica .xl-icon { width: 28px; height: 28px; }
  .btn-ruta-estrategica .xl-icon svg { width: 16px; height: 16px; }
  .btn-datos-origen { padding: 8px 10px; gap: 6px; }
  .btn-datos-origen .xl-label { display: none; }
  .btn-datos-origen .xl-icon { width: 28px; height: 28px; }
  .btn-datos-origen .xl-icon svg { width: 16px; height: 16px; }
  .btn-excel-consolidado { padding: 8px 10px; gap: 6px; }
  .btn-excel-consolidado .xl-label { display: none; }
  .btn-excel-consolidado .xl-icon { width: 28px; height: 28px; }
  .btn-excel-consolidado .xl-icon svg { width: 16px; height: 16px; }

  /* KPIs: grid 3+2 compacto */
  .header-kpis {
    padding: 0 12px 10px;
    display: grid !important;
    grid-template-columns: repeat(3, 1fr);
    gap: 6px;
  }
  .header-kpis .hkpi:last-child {
    grid-column: span 1;
  }
  .hkpi {
    flex-direction: column; align-items: center;
    gap: 2px; padding: 8px 6px;
    border-radius: 10px; min-width: 0;
  }
  .hkpi-icon { width: 26px; height: 26px; border-radius: 7px; }
  .hkpi-icon svg { width: 13px; height: 13px; }
  .hkpi-right { align-items: center; width: 100%; }
  .hkpi-val { font-size: 1.15rem; line-height: 1; }
  .hkpi-lbl { font-size: 0.52rem; letter-spacing: 0.04em; text-align: center; opacity: 0.85; }
  .hkpi-bar { display: none; }

  /* ── Sticky ── */
  .sticky-top > * { padding-left: 12px !important; padding-right: 12px !important; }

  /* ── Gauges: columna vertical ── */
  .gauges-strip {
    flex-direction: column; border-radius: 12px;
    overflow: hidden;
  }
  .gs-sep { display: none; }
  .gs-global {
    flex-direction: row; align-items: center;
    gap: 12px; padding: 12px 14px;
    border-right: none;
    border-bottom: 1px solid var(--border);
    background: linear-gradient(135deg,#F0FDF4,#F8FAFC);
  }
  .gs-global-donut svg { width: 50px !important; height: 50px !important; }
  .gs-global-info { gap: 1px; }
  .gs-global-label { font-size: 0.60rem; }
  .gs-global-pct { font-size: 1.20rem; }
  .gs-global-sub { font-size: 0.66rem; }
  .gs-years {
    display: flex; overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    gap: 0;
  }
  .gs-year {
    flex: 0 0 20%; min-width: 70px;
    padding: 8px 5px;
    border-right: 1px solid var(--border);
  }
  .gs-year:last-child { border-right: none; }
  .gs-year-label { font-size: 0.58rem; gap: 3px; }
  .gs-year-donut svg { width: 44px !important; height: 44px !important; }
  .gs-year-sub { font-size: 0.58rem; }

  /* ── Filtros ── */
  .filters-bar {
    padding: 8px 12px;
    flex-direction: column;
    align-items: flex-start;
    gap: 8px;
  }
  .filter-group { width: 100%; }
  .filter-divider { display: none; }
  .persp-filter-row {
    display: grid !important;
    grid-template-columns: repeat(2, 1fr);
    gap: 7px; width: 100%;
    flex-wrap: unset;
  }
  .persp-filter-card {
    flex: none; width: 100%; max-width: none; min-width: 0;
    padding: 9px 10px;
  }
  .pfc-icon { width: 26px; height: 26px; border-radius: 7px; }
  .pfc-pct { font-size: 0.84rem; }
  .pfc-bar { height: 2px; }
  #year-chips { display: flex; flex-wrap: nowrap; overflow-x: auto; width: 100%; padding-bottom: 2px; }
  #year-chips .chip { flex-shrink: 0; padding: 5px 11px; font-size: 0.78rem; }

  /* ── Matriz: ocultar tabla, mostrar lista ── */
  .matriz-card .matriz-scroll { display: none; }
  .matriz-card .matriz-footer-bar {
    border-top: none; padding: 0;
    flex-direction: row; align-items: center; justify-content: flex-end;
  }
  .mleg { display: none; }
  .mtz-last-update { padding: 6px 12px; font-size: 0.68rem; }
  .mtz-mobile-list { display: block !important; }

  /* ── Perspectivas ── */
  .container { padding: 0 10px 24px; }
  .section-title { margin: 16px 0 10px; }
  .perspective { margin-bottom: 24px; }
  .persp-header { padding: 14px 16px; border-radius: 14px; gap: 10px; }
  .persp-header .persp-stats { display: none; }
  .persp-header .persp-title { font-size: 1.12rem; }
  .persp-objetivo { font-size: 0.79rem; padding: 10px 12px; border-radius: 10px; }
  .ind-grid { grid-template-columns: 1fr; gap: 10px; }

  /* ── Cards de indicador ── */
  .ind-card { border-radius: 14px; }
  .ind-card-head { padding: 14px 16px 10px; }
  .ind-name { font-size: 0.96rem; }
  .ind-kpis { margin: 0 12px 10px; }
  .ind-kpi { padding: 10px 8px; }
  .ind-kpi-lbl { font-size: 0.57rem; }
  .ind-kpi-val { font-size: 0.92rem; }
  .ind-kpi-icon { width: 28px; height: 28px; border-radius: 8px; }
  .chart-wrap { height: 75px; margin: 0 12px 10px; }
  .ind-foot { grid-template-columns: 1fr 1fr; }
  .foot-kpi:last-child { display: none; }
  .foot-kpi { padding: 8px 10px; }
  .foot-kpi-icon { width: 24px; height: 24px; }
  .foot-kpi-info .v { font-size: 0.76rem; }
  .foot-kpi-info .s { display: none; }
  .ind-alert { margin: 0 10px 12px; padding: 8px 10px; font-size: 0.74rem; }
  .al-sub { display: none; }

  /* ── Modal bottom-sheet ── */
  .modal-overlay { padding: 0; align-items: flex-end; }
  .modal {
    border-radius: 22px 22px 0 0;
    max-height: 92vh; overflow-y: auto;
    animation: slideUp .22s cubic-bezier(.34,1.1,.64,1);
  }
  @keyframes slideUp {
    from { transform: translateY(50px); opacity: 0; }
    to   { transform: translateY(0);    opacity: 1; }
  }
  /* Drag handle visual */
  .modal::before {
    content: '';
    display: block; width: 40px; height: 4px;
    background: #CBD5E1; border-radius: 2px;
    margin: 10px auto 4px;
  }
  .modal-head { padding: 10px 16px 10px; }
  .modal-head .m-name { font-size: 1.08rem; }
  .modal-body { padding: 0 14px 30px; }
  .modal-top-grid { grid-template-columns: 1fr; gap: 12px; }
  .m-year-grid { grid-template-columns: repeat(5,1fr); gap: 4px; }
  .m-year-cell { padding: 7px 3px; }
  .m-year-cell .y-label { font-size: 0.56rem; }
  .detail-chart-wrap { height: 200px; }
  .m-chart-stats { gap: 6px; }
  .m-stat-pill { padding: 7px 10px; min-width: 0; flex: 1; }
  .m-stat-pill .sp-label { font-size: 0.56rem; }
  .m-stat-pill .sp-val { font-size: 0.84rem; }
  .tabla { font-size: 0.70rem; }
  .tabla th, .tabla td { padding: 5px 5px; }
  .cb-formula-box { font-size: 0.68rem; }
  .calc-inline-detail { padding: 10px 10px; gap: 10px; flex-direction: column; }
  .cid-steps { gap: 4px; }
  .cid-step { font-size: 0.73rem; }

  /* ── Popover ── */
  .popover-overlay { pointer-events: auto; }
  .popover {
    position: fixed !important;
    left: 12px !important; right: 12px !important;
    top: auto !important; bottom: 0 !important;
    width: auto !important; max-width: none !important;
    border-radius: 20px 20px 0 0;
    max-height: 70vh;
  }
}

/* ── Landscape móvil (celular girado) ── */
@media (max-height: 500px) and (orientation: landscape) {
  .header { position: relative !important; top: auto !important; }
  .header-inner { padding: 8px 16px; gap: 10px; }
  .header-brand { display: none; }
  .header-meta { display: none; }
  .header-titles h1 { font-size: 1.0rem; }
  .header-titles .subtitle { display: none; }

  .header-btns { gap: 8px; }
  .btn-ruta-estrategica { padding: 7px 10px; gap: 6px; }
  .btn-ruta-estrategica .xl-label { display: none; }
  .btn-ruta-estrategica .xl-icon { width: 26px; height: 26px; }
  .btn-ruta-estrategica .xl-icon svg { width: 14px; height: 14px; }
  .btn-datos-origen { padding: 7px 10px; gap: 6px; }
  .btn-datos-origen .xl-label { display: none; }
  .btn-datos-origen .xl-icon { width: 26px; height: 26px; }
  .btn-datos-origen .xl-icon svg { width: 14px; height: 14px; }
  .btn-excel-consolidado { padding: 7px 10px; gap: 6px; }
  .btn-excel-consolidado .xl-label { display: none; }
  .btn-excel-consolidado .xl-icon { width: 26px; height: 26px; }
  .btn-excel-consolidado .xl-icon svg { width: 14px; height: 14px; }

  .header-kpis { padding: 0 16px 8px; gap: 6px; flex-wrap: nowrap; overflow-x: auto; }
  .hkpi { flex: 0 0 auto; min-width: 70px; padding: 7px 8px; }
  .hkpi-val { font-size: 1.1rem; }
  .hkpi-label { font-size: 0.58rem; }

  .sticky-top { position: relative !important; top: auto !important; box-shadow: none; }

  .ind-foot { grid-template-columns: 1fr 1fr 1fr; }
  .foot-kpi:last-child { display: flex; }
  .foot-kpi { padding: 7px 8px; gap: 6px; min-width: 0; }
  .foot-kpi-icon { width: 22px; height: 22px; flex-shrink: 0; }
  .foot-kpi-icon svg { width: 11px; height: 11px; }
  .foot-kpi-info { min-width: 0; overflow: hidden; }
  .foot-kpi-info .l { font-size: 0.54rem; }
  .foot-kpi-info .v {
    font-size: 0.70rem; white-space: nowrap;
    overflow: hidden; text-overflow: ellipsis; max-width: 100%;
  }
  .foot-kpi-info .s { display: none; }
  .ind-name {
    font-size: 0.88rem;
    display: -webkit-box; -webkit-line-clamp: 2;
    -webkit-box-orient: vertical; overflow: hidden;
  }
  .ind-alert { font-size: 0.72rem; padding: 7px 9px; margin: 0 10px 10px; }
  .al-sub { display: none; }

  .popover {
    position: fixed !important;
    left: 5% !important; right: 5% !important;
    top: 10px !important; bottom: 10px !important;
    width: 90% !important; max-width: none !important;
    border-radius: 16px; max-height: 95vh;
  }
}

/* ════════════════════════════════════════════════════════════
   TEMA NEUMÓRFICO CLARO — relieves suaves + ticks vibrantes
   ════════════════════════════════════════════════════════════ */
:root {
  --nm-bg: #E6ECF4; --nm-hi: #FFFFFF; --nm-lo: #C3CBD8; --nm-line: #D3DAE5;
  --nm-raised:    6px 6px 12px var(--nm-lo), -6px -6px 12px var(--nm-hi);
  --nm-raised-sm: 3px 3px 6px var(--nm-lo), -3px -3px 6px var(--nm-hi);
  --nm-inset:     inset 3px 3px 7px var(--nm-lo), inset -3px -3px 7px var(--nm-hi);
  --v-green: #12C97A; --v-amber: #FFAA1D; --v-red: #FF4D6D; --v-gray: #B8C2D3;
  --v-blue: #3A7BFD; --v-violet: #7B5CFF;
  --bg: var(--nm-bg);
}
body { background: var(--nm-bg); }

/* Ticks de estado */
.nm-tick {
  width: 22px; height: 22px; border-radius: 50%; flex-shrink: 0;
  display: inline-flex; align-items: center; justify-content: center; color: #fff;
}
.nm-tick svg { width: 13px; height: 13px; display: block; }
.nm-tick.cumple    { background: var(--v-green); }
.nm-tick.riesgo    { background: var(--v-amber); }
.nm-tick.critico   { background: var(--v-red); }
.nm-tick.pendiente { background: var(--v-gray); }

/* Anillos */
.nm-ring {
  width: var(--s); height: var(--s); border-radius: 50%; flex-shrink: 0;
  background: conic-gradient(var(--c) calc(var(--p) * 1%), #D5DCE7 0);
  display: flex; align-items: center; justify-content: center;
  box-shadow: var(--nm-raised-sm);
}
.nm-ring-in {
  width: 74%; height: 74%; border-radius: 50%; background: var(--nm-bg);
  box-shadow: var(--nm-inset);
  display: flex; flex-direction: column; align-items: center; justify-content: center; line-height: 1.05;
}
.nm-ring-pct { font-size: 1.35rem; font-weight: 900; letter-spacing: -0.02em; }
.nm-ring-pct.sm { font-size: 0.74rem; font-weight: 800; }
.nm-ring-lbl { font-size: 0.62rem; color: var(--text-med); font-weight: 600; }
.nm-counts { display: flex; gap: 6px; margin-top: 8px; }
.nm-cnt {
  min-width: 30px; height: 30px; border-radius: 50%; padding: 0 4px;
  display: inline-flex; align-items: center; justify-content: center;
  color: #fff; font-size: 0.78rem; font-weight: 800; box-shadow: var(--nm-raised-sm);
}
.nm-cnt.cumple { background: var(--v-green); } .nm-cnt.riesgo { background: var(--v-amber); }
.nm-cnt.critico { background: var(--v-red); } .nm-cnt.pendiente { background: var(--v-gray); }

/* Encabezado */
.header { background: var(--nm-bg); color: var(--text-dark); box-shadow: none; }
.header::before, .header::after { display: none; }
.header-brand { border-right-color: var(--nm-line); }
.header-logo-box { background: var(--nm-bg); border: none; box-shadow: var(--nm-raised-sm); }
.header-logo-box svg { color: var(--v-blue); }
.header-brand-name, .header-titles h1 { color: var(--text-dark); }
.header-brand-sub, .header-meta-item { color: var(--text-med); border-right-color: var(--nm-line); }
.btn-ruta-estrategica, .btn-datos-origen, .btn-excel-consolidado {
  background: var(--nm-bg); color: var(--text-dark); border: none; box-shadow: var(--nm-raised);
}
.btn-ruta-estrategica:hover, .btn-datos-origen:hover, .btn-excel-consolidado:hover {
  background: var(--nm-bg); color: var(--text-dark); box-shadow: var(--nm-raised-sm); transform: translateY(-1px);
}
.btn-ruta-estrategica .xl-icon { background: var(--v-blue); color: #fff; }
.btn-datos-origen .xl-icon { background: var(--v-violet); color: #fff; }
.btn-excel-consolidado .xl-icon { background: var(--v-green); }

/* KPIs del encabezado */
.hkpi { background: var(--nm-bg); border: none; box-shadow: var(--nm-raised); backdrop-filter: none; }
.hkpi:hover, .hkpi.clickable:hover { background: var(--nm-bg); box-shadow: var(--nm-raised-sm); }
.hkpi-icon { border-radius: 50%; color: #fff !important; }
.hkpi-icon.total { background: var(--v-blue); }
.hkpi-icon.cumple { background: var(--v-green); }
.hkpi-icon.riesgo { background: var(--v-amber); }
.hkpi-icon.critico { background: var(--v-red); }
.hkpi-icon.pendiente { background: var(--v-gray); }
.hkpi-val { color: var(--text-dark); }
.hkpi-lbl { color: var(--text-med); }
.hkpi-bar { background: #D5DCE7; box-shadow: inset 1px 1px 2px var(--nm-lo); }
.hkpi-bar-fill.total { background: var(--v-blue); }
.hkpi-bar-fill.cumple { background: var(--v-green); }
.hkpi-bar-fill.riesgo { background: var(--v-amber); }
.hkpi-bar-fill.critico { background: var(--v-red); }
.hkpi-bar-fill.pendiente { background: var(--v-gray); }

/* Panel de cumplimiento y filtros */
.sticky-top, .filters-bar { background: var(--nm-bg); border: none; }
.gauges-strip { background: var(--nm-bg); border: none; box-shadow: var(--nm-raised); }
.gs-global { background: transparent; border-right: 1px solid var(--nm-line); }
.gs-year { border-right-color: var(--nm-line); }
.gs-year:hover { background: rgba(255,255,255,0.35); }
.gs-year.active { background: var(--nm-bg) !important; box-shadow: var(--nm-inset); border-bottom: none !important; }
.persp-filter-card {
  background: var(--nm-bg); border: none; box-shadow: var(--nm-raised); padding: 10px 12px;
}
.persp-filter-card:hover { box-shadow: var(--nm-raised-sm); border: none; }
.persp-filter-card.active { box-shadow: var(--nm-inset); }
.persp-filter-card .nm-tick { width: 20px; height: 20px; }
.chip { background: var(--nm-bg); box-shadow: var(--nm-raised-sm); border: none; }
.chip:hover { background: var(--nm-bg); color: var(--text-dark); }
.chip.year.active { background: var(--nm-bg); color: var(--v-blue); box-shadow: var(--nm-inset); font-weight: 700; }

/* Matriz */
.matriz-card { background: var(--nm-bg); border: none; box-shadow: var(--nm-raised); }
.matriz-footer-bar { background: var(--nm-bg); border-top-color: var(--nm-line); }
.matriz-table th, .matriz-table td { border-bottom-color: var(--nm-line); }
.matriz-table thead th, .matriz-table thead th.mtz-th-ind, .matriz-table thead th.mtz-th-cumpl,
.matriz-table thead th.mtz-th-tend, .matriz-table thead th.mtz-th-meta, .matriz-table thead th.mtz-th-year {
  background: #DCE3EE !important; color: var(--text-med) !important;
  border-color: var(--nm-line) !important;
}
.mtz-th-mes-abbr { color: var(--text-dark); }
.mtz-th-mes-full { color: var(--text-lt); }
.matriz-table thead th.mtz-th-mes.current { background: var(--v-blue) !important; }
.mtz-th-mes.current .mtz-th-mes-abbr, .mtz-th-mes.current .mtz-th-mes-full { color: #fff; }
.mtz-ind-cell, .mtz-tend-cell, .mtz-cumpl-cell { background: var(--nm-bg); }
.mtz-ind-cell { border-right-color: var(--nm-line) !important; }
.mtz-tend-cell, .mtz-cumpl-cell { border-color: var(--nm-line) !important; }
.mtz-meta-cell { background: #DDE6F6; border-left-color: #C7D5EE !important; border-right-color: var(--nm-line) !important; }
.matriz-table tbody tr:hover > td { background: #EEF3FA !important; }

/* Celda mensual: pastilla en relieve con tick */
.mtz-light, .mtz-light.cumple, .mtz-light.riesgo, .mtz-light.critico, .mtz-light.pendiente {
  background: var(--nm-bg); flex-direction: row; justify-content: flex-start;
  width: 66px; height: 26px; padding: 0 6px 0 3px; gap: 4px;
  border-radius: 999px; box-shadow: var(--nm-raised-sm);
}
.mtz-light:hover { box-shadow: var(--nm-inset); transform: none; }
.mtz-light .nm-tick { width: 19px; height: 19px; }
.mtz-light .nm-tick svg { display: block; width: 11px; height: 11px; }
.mtz-light .mtz-val { font-size: 0.66rem; font-weight: 800; }
.mtz-light.cumple .mtz-val  { color: #0B9A5C; }
.mtz-light.riesgo .mtz-val  { color: #D98500; }
.mtz-light.critico .mtz-val { color: #E8324F; }
.mtz-light.pendiente { justify-content: center; box-shadow: var(--nm-inset); }
.mtz-light.pendiente .nm-tick { display: none; }
.mtz-light.pendiente .mtz-val { color: #A3AEC2; }

/* Botón y ventana de escalas */
.btn-escalas .xl-icon { background: var(--v-amber) !important; }
.header-titles { min-width: 360px; }
.header-titles h1 { white-space: nowrap; }
.header-btns { flex-wrap: wrap; justify-content: flex-end; }
.header-btns > * { padding: 8px 14px; font-size: 0.82rem; }
.header-btns .xl-icon { width: 26px; height: 26px; }
.header-btns .xl-icon svg { width: 15px; height: 15px; }
@media (max-width: 768px) { .header-titles { min-width: 0; } .header-titles h1 { white-space: normal; } }
.esc-link {
  border: none; background: var(--nm-bg); box-shadow: var(--nm-raised-sm); border-radius: 999px;
  padding: 4px 12px; font-size: 0.74rem; font-weight: 700; color: var(--v-blue); cursor: pointer;
}
.esc-link:hover { box-shadow: var(--nm-inset); }
#modal-content:has(.esc-groups) { max-width: 760px; background: var(--nm-bg); }
.esc-groups { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
.esc-card {
  display: flex; gap: 12px; align-items: flex-start; padding: 14px;
  background: var(--nm-bg); border-radius: 14px; box-shadow: var(--nm-raised);
}
.esc-card .nm-tick svg { width: 18px; height: 18px; }
.esc-title { font-size: 0.72rem; font-weight: 800; color: var(--text-med); text-transform: uppercase; letter-spacing: 0.06em; }
.esc-big { font-size: 1.35rem; font-weight: 900; line-height: 1.2; }
.esc-big.cumple { color: #0B9A5C; } .esc-big.riesgo { color: #D98500; }
.esc-big.critico { color: #E8324F; } .esc-big.pendiente { color: #94A3B8; }
.esc-desc { font-size: 0.78rem; color: var(--text-med); }
.esc-formulas { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-top: 8px; }
.esc-f {
  display: flex; flex-direction: column; gap: 4px; padding: 12px;
  border-radius: 12px; box-shadow: var(--nm-inset); font-size: 0.82rem;
}
.esc-f span { font-family: Consolas, monospace; font-size: 0.78rem; color: var(--v-violet); font-weight: 700; }
.esc-f small { color: var(--text-med); font-size: 0.72rem; }
.esc-table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 0.84rem; }
.esc-table th { text-align: left; font-size: 0.68rem; color: var(--text-med); text-transform: uppercase; letter-spacing: 0.06em; padding: 6px 8px; }
.esc-table td { padding: 6px 8px; border-top: 1px solid var(--nm-line); }
.esc-level { display: inline-block; padding: 3px 10px; border-radius: 8px; font-weight: 700; white-space: nowrap; }
.esc-ldesc { color: var(--text-med); font-size: 0.78rem; }
.esc-range { font-variant-numeric: tabular-nums; color: var(--text-med); }
.esc-pill {
  display: inline-flex; align-items: center; gap: 6px; padding: 2px 10px 2px 2px;
  border-radius: 999px; background: var(--nm-bg); box-shadow: var(--nm-raised-sm); font-weight: 700; font-size: 0.76rem;
}
.esc-pill .nm-tick { width: 20px; height: 20px; }
.esc-pill.cumple { color: #0B9A5C; } .esc-pill.riesgo { color: #D98500; }
.esc-pill.critico { color: #E8324F; }
.esc-note { margin-top: 16px; font-size: 0.76rem; color: var(--text-med); }
@media (max-width: 640px) { .esc-groups, .esc-formulas { grid-template-columns: 1fr; } }

/* Barra compacta en el mismo estilo */
.cbar { background: var(--nm-bg); color: var(--text-dark); box-shadow: 0 6px 14px var(--nm-lo); }
.cbar-kpi { background: var(--nm-bg); border: none; box-shadow: var(--nm-raised-sm); color: var(--text-dark); }
.cbar-kpi.clickable:hover { background: var(--nm-bg); box-shadow: var(--nm-inset); }
.cbar-global { background: var(--nm-bg); box-shadow: var(--nm-inset); }
.cbar-persp { background: var(--nm-bg); box-shadow: var(--nm-raised-sm); color: var(--v-blue); }
.cbar-year { background: var(--nm-bg); color: var(--text-med); box-shadow: var(--nm-raised-sm); }
.cbar-year:hover { background: var(--nm-bg); color: var(--text-dark); }
.cbar-year.active { background: var(--nm-bg); color: var(--v-blue); box-shadow: var(--nm-inset); }

"""


# ============================================================
# JAVASCRIPT
# ============================================================

JS = r"""


// ════════════════════════════════════════════════════════════
// Gauges de cumplimiento global y por año
// ════════════════════════════════════════════════════════════

// Escala de color verde→amarillo según porcentaje
// 100% = verde puro (#22C55E), 0% = amarillo/naranja (#F59E0B)
// Interpolación HSL: 142° (verde) → 45° (ámbar)
/* Colorimetría exacta del Excel — Parámetros semáforo */
const SEMAFORO_COLORS_DEFAULT = [
  { lim:1.00, bg:'#83F1B7', fg:'#006666', name:'Excelente'      },
  { lim:0.90, bg:'#B9F9CE', fg:'#15803D', name:'Sobresaliente'  },
  { lim:0.80, bg:'#B9F9CE', fg:'#15803D', name:'Satisfactorio'  },
  { lim:0.70, bg:'#FEF9C3', fg:'#A16207', name:'Aceptable'      },
  { lim:0.60, bg:'#FEF9C3', fg:'#A16207', name:'En seguimiento' },
  { lim:0.50, bg:'#FFD3A7', fg:'#D97706', name:'En riesgo'      },
  { lim:0.40, bg:'#FFD3A7', fg:'#EA580C', name:'En alerta'      },
  { lim:0.30, bg:'#FEE2E2', fg:'#DC2626', name:'Crítico'        },
  { lim:0.00, bg:'#FEC9C9', fg:'#B91C1C', name:'Muy crítico'    },
];

function getSemaforoColors() {
  return (typeof DATA !== 'undefined' && DATA.thresholds && DATA.thresholds.colors)
    ? DATA.thresholds.colors
    : SEMAFORO_COLORS_DEFAULT;
}

function getSemaforoEntry(pct) {
  if (pct === null || pct === undefined) {
    return (typeof DATA !== 'undefined' && DATA.thresholds && DATA.thresholds.pendiente)
      || { bg:'#F1F5F9', fg:'#64748B', name:'Pendiente' };
  }
  const colors = getSemaforoColors();
  if (pct < 0) return colors[colors.length - 1]; // negativo → Muy crítico
  for (const e of colors) { if (pct >= e.lim) return e; }
  return colors[colors.length - 1];
}

function pctToColor(pct) {
  if (pct === null || pct === undefined) return '#CBD5E1';
  return getSemaforoEntry(pct).fg;
}

function pctToBg(pct) { return getSemaforoEntry(pct).bg; }
function pctToName(pct) { return getSemaforoEntry(pct).name; }

// ── Tema neumórfico: ticks de estado y anillos ────────────
const NM_TICK = {
  cumple:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12l5 5 9-10"/></svg>',
  riesgo:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.2" stroke-linecap="round"><path d="M12 6v7"/><path d="M12 18h.01"/></svg>',
  critico:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.2" stroke-linecap="round"><path d="M7 7l10 10"/><path d="M17 7L7 17"/></svg>',
  pendiente: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round"><path d="M12 7v5l3 2"/></svg>',
};
const NM_PERSP_COLORS = ['#00C2A8', '#3A7BFD', '#FF8A00', '#FF3D8B'];

function nmStatus(pct) {
  if (pct === null || pct === undefined) return 'pendiente';
  return pct >= DATA.thresholds.cumple ? 'cumple' : pct >= DATA.thresholds.riesgo ? 'riesgo' : 'critico';
}

function nmRing(pct, color, size, inner) {
  const p = pct === null ? 0 : Math.max(0, Math.min(100, pct * 100));
  return `<div class="nm-ring" style="--p:${p.toFixed(1)};--c:${color};--s:${size}px">
    <div class="nm-ring-in">${inner}</div></div>`;
}

function buildGaugeDonut(pct, size, strokeW) {
  size   = size   || 110;
  strokeW = strokeW || 12;
  const vb = 100;
  const cx = 50, cy = 50;
  const r  = (vb / 2) - strokeW - 2;
  const sw = strokeW;
  const clamped = pct === null || pct === undefined ? 0 : Math.min(Math.max(pct, 0), 1);
  const circ = 2 * Math.PI * r;
  const dashOffset = circ * (1 - clamped);
  const color = getSemaforoEntry(pct).fg;
  const pctTxt = pct === null || pct === undefined ? '—' : Math.round(pct * 100) + '%';
  const fs = Math.round(r * 0.46);

  // Gradiente de arco: del color al color más claro
  const gradId = 'gd' + Math.random().toString(36).slice(2,7);

  return `<svg viewBox="0 0 ${vb} ${vb}" xmlns="http://www.w3.org/2000/svg"
       style="display:block;width:${size}px;height:${size}px;">
    <defs>
      <linearGradient id="${gradId}" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" stop-color="${color}" stop-opacity="0.7"/>
        <stop offset="100%" stop-color="${color}"/>
      </linearGradient>
    </defs>
    <circle cx="${cx}" cy="${cy}" r="${r}"
            fill="none" stroke="#E2E8F0" stroke-width="${sw}"/>
    ${clamped > 0 ? `<circle cx="${cx}" cy="${cy}" r="${r}"
            fill="none" stroke="url(#${gradId})" stroke-width="${sw}"
            stroke-linecap="round"
            stroke-dasharray="${circ}"
            stroke-dashoffset="${dashOffset}"
            transform="rotate(-90 ${cx} ${cy})"
            style="transition:stroke-dashoffset .8s cubic-bezier(.4,0,.2,1);"/>` : ''}
    <text x="${cx}" y="${cy + fs * 0.35}" text-anchor="middle"
          font-size="${fs}" font-weight="800" font-family="inherit"
          fill="${pct !== null ? color : '#94A3B8'}"
          letter-spacing="-0.03em">${pctTxt}</text>
  </svg>`;
}

function renderGauges() {
  const root = document.getElementById('gauges-panel');
  if (!root) return;
  const effYear = effectiveYear();
  const activeYr = effYear !== null ? effYear : CURRENT_YEAR;

  // Cumplimiento global del año activo
  let sumG = 0, nG = 0;
  getAllIndicators().forEach(({ind}) => {
    const yObj = ind.yearly.find(y => y.year === activeYr);
    if (yObj && yObj.cumplimiento !== null && yObj.cumplimiento !== undefined) {
      sumG += Math.min(yObj.cumplimiento, 1.5); nG++;
    }
  });
  const globalPct = nG > 0 ? sumG / nG : null;
  const globalColor = pctToColor(globalPct);
  const globalTxt = globalPct !== null ? Math.round(globalPct * 100) + '%' : '—';

  // Mini donut inline para años (más pequeño y compacto)
  function miniDonutInline(pct, sz, sw) {
    sz = sz || 46; sw = sw || 5;
    const r = sz/2 - sw - 1;
    const circ = 2 * Math.PI * r;
    const c = pct === null ? 0 : Math.min(Math.max(pct, 0), 1);
    const col = getSemaforoEntry(pct).fg;
    const txt = pct !== null ? Math.round(pct * 100) + '%' : '—';
    const fs = sz * 0.20;
    return `<svg viewBox="0 0 ${sz} ${sz}" width="${sz}" height="${sz}" style="display:block">
      <circle cx="${sz/2}" cy="${sz/2}" r="${r}" fill="none" stroke="#E2E8F0" stroke-width="${sw}"/>
      ${c > 0 ? `<circle cx="${sz/2}" cy="${sz/2}" r="${r}" fill="none" stroke="${col}"
        stroke-width="${sw}" stroke-linecap="round"
        stroke-dasharray="${circ}" stroke-dashoffset="${circ*(1-c)}"
        transform="rotate(-90 ${sz/2} ${sz/2})"/>` : ''}
      <text x="${sz/2}" y="${sz/2 + fs*0.38}" text-anchor="middle"
        font-size="${fs}" font-weight="800" fill="${pct!==null?col:'#94A3B8'}"
        font-family="inherit">${txt}</text>
    </svg>`;
  }

  // Años
  const yearsHTML = DATA.years.map(yr => {
    let sum = 0, n = 0;
    getAllIndicators().forEach(({ind}) => {
      const yObj = ind.yearly.find(y => y.year === yr);
      if (yObj && yObj.cumplimiento !== null && yObj.cumplimiento !== undefined) {
        sum += Math.min(yObj.cumplimiento, 1.5); n++;
      }
    });
    const pct = n > 0 ? sum / n : null;
    const isCurr = yr === CURRENT_YEAR;
    const isActive = activeYr === yr;
    const clr = getSemaforoEntry(pct).fg;
    return `<div class="gs-year ${isActive?'active':''}"
         style="${isActive?'background:'+hexA(clr,0.06)+';':''}${isActive?`border-bottom:3px solid ${clr};`:''}"
         data-yr="${yr}">
      <div class="gs-year-label" style="${isActive?'color:'+clr:''}">
        ${isCurr?'<span class="curr-dot"></span>':''}${yr}
      </div>
      <div class="gs-year-donut">${miniDonutInline(pct, 52, 5.5)}</div>
      <div class="gs-year-sub">${n} ind.</div>
    </div>`;
  }).join('');

  root.innerHTML = `
    <div class="gauges-strip">
      <div class="gs-global">
        <div class="gs-global-donut">
          ${nmRing(globalPct, '#7B5CFF', 92, `<span class="nm-ring-pct" style="color:#7B5CFF">${globalTxt}</span><span class="nm-ring-lbl">Global</span>`)}
        </div>
        <div class="gs-global-info">
          <div class="gs-global-label">Cumplimiento global</div>
          <div class="gs-global-sub">${nG} de ${DATA.summary.total} indicadores · ${activeYr}</div>
          <div class="nm-counts">
            <span class="nm-cnt cumple" title="En meta">${DATA.summary.cumple||0}</span>
            <span class="nm-cnt riesgo" title="En riesgo">${DATA.summary.riesgo||0}</span>
            <span class="nm-cnt critico" title="Crítico">${DATA.summary.critico||0}</span>
            <span class="nm-cnt pendiente" title="Sin datos">${DATA.summary.pendiente||0}</span>
          </div>
        </div>
      </div>
      <div class="gs-sep"></div>
      <div class="gs-years">${yearsHTML}</div>
    </div>`;

  // Click en año dentro de los gauges → cambiar filtro
  root.querySelectorAll('.gs-year').forEach(el => {
    el.addEventListener('click', () => {
      const yr = parseInt(el.dataset.yr);
      STATE.yearFilter = yr;
      renderFilters(); renderGauges(); renderMatriz(); renderPerspectives();
    });
  });

  renderCompactBar(globalTxt, globalColor, activeYr);
}

// ── Barra compacta: título + conteos + % global + años ────
function renderCompactBar(globalTxt, globalColor, activeYr) {
  const el = document.getElementById('cbar-inner');
  if (!el) return;
  const s = DATA.summary;
  const kpis = [
    { st:'cumple',    label:'En meta',   val:s.cumple||0,    dot:'#12C97A' },
    { st:'riesgo',    label:'En riesgo', val:s.riesgo||0,    dot:'#FFAA1D' },
    { st:'critico',   label:'Crítico',   val:s.critico||0,   dot:'#FF4D6D' },
    { st:'pendiente', label:'Sin datos', val:s.pendiente||0, dot:'#B8C2D3' },
  ];
  const persp = DATA.perspectives.find(p => p.key === STATE.perspFilter);
  const yearButtons = [{key:'all', label:'Todos'}].concat(DATA.years.map(y => ({key:y, label:String(y)})));
  el.innerHTML = `
    <div class="cbar-title" title="Volver arriba"><span class="brand-dot"></span> Plan Estratégico 2026 – 2030</div>
    <div class="cbar-kpis">
      <span class="cbar-kpi">${s.total||0} <span class="cbar-lbl">indicadores</span></span>
      ${kpis.map(k => `<span class="cbar-kpi ${k.val>0?'clickable':''}" data-status="${k.st}">
        <span class="cbar-dot" style="background:${k.dot}"></span>${k.val} <span class="cbar-lbl">${k.label}</span></span>`).join('')}
    </div>
    <div class="cbar-global" style="color:${globalColor}"><span>Global ${STATE.yearFilter==='all'?CURRENT_YEAR:activeYr}</span>${globalTxt}</div>
    ${persp ? `<div class="cbar-persp" title="Quitar filtro de perspectiva">${persp.name} ✕</div>` : ''}
    <div class="cbar-years">
      ${yearButtons.map(y => {
        const isActive = STATE.yearFilter==='all' ? y.key==='all'
          : STATE.yearFilter==='current' ? y.key===CURRENT_YEAR : y.key===STATE.yearFilter;
        return `<div class="cbar-year ${isActive?'active':''}" data-year="${y.key}">${y.label}</div>`;
      }).join('')}
    </div>`;
  el.querySelector('.cbar-title').addEventListener('click', () => window.scrollTo({top:0, behavior:'smooth'}));
  el.querySelectorAll('.cbar-kpi.clickable').forEach(k => k.addEventListener('click', e => {
    e.stopPropagation(); openPopoverForStatus(k.dataset.status, k);
  }));
  const pc = el.querySelector('.cbar-persp');
  if (pc) pc.addEventListener('click', () => {
    STATE.perspFilter = 'all';
    renderFilters(); renderGauges(); renderMatriz(); renderPerspectives();
  });
  el.querySelectorAll('.cbar-year').forEach(c => c.addEventListener('click', () => {
    const v = c.dataset.year;
    STATE.yearFilter = v==='all' ? 'all' : parseInt(v);
    renderFilters(); renderGauges(); renderMatriz(); renderPerspectives();
  }));
}

// Muestra la barra compacta al pasar los filtros y mantiene
// visible la fila de meses de la matriz justo debajo de ella.
function updateCompactState() {
  const bar = document.getElementById('cbar');
  const filters = document.querySelector('.filters-bar');
  if (!bar || !filters) return;
  const desktop = window.innerWidth >= 769;
  const show = desktop && filters.getBoundingClientRect().bottom < 0;
  bar.classList.toggle('show', show);

  const table = document.querySelector('.matriz-table');
  const thead = table && table.tHead;
  if (!thead) return;
  let off = 0;
  if (desktop) {
    const barH = show ? bar.offsetHeight : 0;
    const r = table.getBoundingClientRect();
    off = Math.max(0, Math.min(barH - r.top, r.height - thead.offsetHeight - 40));
  }
  thead.querySelectorAll('th').forEach(th => { th.style.transform = off ? `translateY(${off}px)` : ''; });
}

let _compactRaf = 0;
function onCompactScroll() {
  if (_compactRaf) return;
  _compactRaf = requestAnimationFrame(() => { _compactRaf = 0; updateCompactState(); });
}

// ════════════════════════════════════════════════════════════
// Estado global
// ════════════════════════════════════════════════════════════
const CURRENT_YEAR = new Date().getFullYear();
const STATE = { perspFilter: 'all', yearFilter: 'current', charts: {}, activeKpi: null };
const MONTH_ABBR = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const STATUS_LABELS = { cumple:'En meta', riesgo:'En riesgo', critico:'Crítico', pendiente:'Sin datos' };
function semaforoLabel(pct, status) {
  if (pct !== null && pct !== undefined && status !== 'pendiente') return getSemaforoEntry(pct).name;
  return status === 'pendiente' ? 'Sin datos' : (STATUS_LABELS[status] || 'Sin datos');
}
const PERSP_KEY_TO_SLUG = {
  'FINANCIERA':'financiera','ASOCIADOS':'asociados',
  'PROCESOS INTERNOS':'procesos','APRENDIZAJE Y CRECIMIENTO':'aprendizaje'
};

const ICONS = {
  total:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 14l4-4 4 4 5-5"/></svg>',
  cumple:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M9 12l2 2 4-4"/><circle cx="12" cy="12" r="10"/></svg>',
  riesgo:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4"/><path d="M12 17h.01"/><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/></svg>',
  critico:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M15 9l-6 6"/><path d="M9 9l6 6"/></svg>',
  pendiente:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>',
  FINANCIERA:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>',
  ASOCIADOS:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>',
  'PROCESOS INTERNOS':'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 01-2.83 2.83l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/></svg>',
  'APRENDIZAJE Y CRECIMIENTO':'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>',
};

// ── Formateo — estilo colombiano ─────────────────────────
// Formatear una variable de entrada según el label (no el tipo del indicador)
// Si el label contiene ($) → número en pesos colombianos (sin %)
function fmtVar(v, label, decimales) {
  if (v === null || v === undefined) return '—';
  if (typeof v !== 'number' || !isFinite(v)) return String(v);
  // Si decimales explícito: respetar eso (ej. IND-14 enteros)
  if (decimales !== undefined && decimales !== null) return numCO(v, decimales);
  const lbl = (label || '').toLowerCase();
  // Variables de porcentaje/tasa: convertir a %
  if (lbl.includes('%') || lbl.includes('tasa') || lbl.includes('porcentaje')) {
    return numCO(v * 100, 2) + '%';
  }
  // Variables en pesos: por label ($) o por magnitud >= 1,000,000
  const esPesos = lbl.includes('($)') || lbl.includes('pesos') || lbl.includes('cop')
                  || (Math.abs(v) >= 1000000 && !lbl.includes('%') && !lbl.includes('tasa'));
  if (esPesos) {
    // Siempre mostrar en millones (M) — nunca MM
    if (Math.abs(v) >= 1e6)  return numCO(v / 1e6,  2) + ' M';
    if (Math.abs(v) >= 1e3)  return numCO(v / 1e3,  2) + ' mil';
    return numCO(v, 2);
  }
  // Default: número con 2 decimales
  return numCO(v, 2);
}

// ── Formateo — estilo colombiano ─────────────────────────
// Convierte número a string con punto como separador de miles y coma decimal
function numCO(n, decimals) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  decimals = decimals !== undefined ? decimals : 2;
  // Redondear
  const factor = Math.pow(10, decimals);
  const rounded = Math.round(n * factor) / factor;
  // Separar parte entera y decimal
  const parts = rounded.toFixed(decimals).split('.');
  // Punto como separador de miles
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  // Coma como separador decimal (solo si hay decimales reales)
  if (decimals === 0) return parts[0];
  return parts[0] + ',' + parts[1];
}

// Millones: si unidad contiene "$" o "millones" y el valor >= 1.000.000
function shouldShowMillions(v, unidad) {
  const u = (unidad || '').toLowerCase();
  return (u.includes('$') || u.includes('mill') || u.includes('cop') || u.includes('pesos')) && Math.abs(v) >= 1000000;
}

// Formato principal para mostrar valores
function fmt(v, tipoMeta, unidad, decimales) {
  if (v === null || v === undefined || v === '') return '—';
  if (typeof v === 'string') return v;
  // decimales controla los decimales de las variables de entrada (fmtVar),
  // pero NO debe sobrepasar la lógica de tipo (pct/ipc necesitan *100 siempre)
  if (tipoMeta === 'pct' || tipoMeta === 'ipc') {
    // Para pct/ipc: siempre *100, luego aplicar decimales si se especifica (mín 2)
    const dec = (decimales !== undefined && decimales !== null) ? Math.max(decimales, 2) : 2;
    return numCO(v * 100, dec) + '%';
  }
  if (tipoMeta === 'bps' || tipoMeta === 'pts')
    return numCO(v, 0) + ' bps';
  if (tipoMeta === 'ratio') {
    const s = numCO(v, 2);
    return (unidad||'').toLowerCase().includes('x') ? s + 'x' : s;
  }
  if (tipoMeta === 'unidades') {
    if (shouldShowMillions(v, unidad)) return numCO(v / 1000000, 2) + ' M';
    return numCO(Math.round(v), 0);
  }
  if (tipoMeta === 'abs') {
    // Mostrar 2 decimales — si son ,00 fmtSmart los elimina
    return numCO(v, 2);
  }
  // texto u otro: intentar número con 2 decimales
  if (!isNaN(parseFloat(v))) {
    const n = parseFloat(v);
    if (shouldShowMillions(n, unidad)) return numCO(n / 1000000, 2) + ' M';
    return numCO(n, 2);
  }
  return String(v);
}

// Formato compacto para celdas de la matriz
function fmtMtz(v, tipoMeta, unidad, decimales) {
  if (v === null || v === undefined) return '—';
  if (typeof v === 'string') return v;
  if (tipoMeta === 'pct' || tipoMeta === 'ipc') {
    const dec = (decimales !== undefined && decimales !== null) ? Math.max(decimales, 2) : 2;
    return numCO(v * 100, dec).replace(/,00$/, '') + '%';
  }
  if (tipoMeta === 'bps' || tipoMeta === 'pts')
    return numCO(v, 0) + ' bps';
  if (tipoMeta === 'ratio') {
    const s = numCO(v, 2);
    return (unidad||'').toLowerCase().includes('x') ? s + 'x' : s;
  }
  if (tipoMeta === 'unidades') {
    if (shouldShowMillions(v, unidad)) return numCO(v / 1000000, 2) + 'M';
    return numCO(Math.round(v), 0);
  }
  if (tipoMeta === 'abs') {
    const s = numCO(v, 2);
    return s.replace(/,00$/, '');
  }
  return numCO(v, 2);
}

// Formato inteligente: igual que fmt() pero elimina decimales insignificantes (,00)
function fmtSmart(v, tipoMeta, unidad, decimales) {
  const s = fmt(v, tipoMeta, unidad, decimales);
  // Quitar ,00 antes de unidades (%, bps, M, etc.) o al final
  return s.replace(/,00(%|bps| M| mil|x|$)/, '$1')
          .replace(/,0(%|bps| M| mil|x|$)/, '$1');
}

// Formato porcentaje de cumplimiento
function fmtPct(v) {
  if (v === null || v === undefined) return '—';
  return numCO(v * 100, 0) + '%';
}

// ── Año efectivo ──────────────────────────────────────────
function effectiveYear() {
  if (STATE.yearFilter === 'all') return null;
  if (STATE.yearFilter === 'current') {
    return DATA.years.includes(CURRENT_YEAR) ? CURRENT_YEAR : DATA.years[0];
  }
  return STATE.yearFilter;
}

// ── Helpers de color/semáforo ─────────────────────────────
function monthCompliance(ind, yearIdx, monthIdx) {
  const flatIdx = yearIdx * 12 + monthIdx;
  const result = ind.monthly_results[flatIdx];
  if (result === null || result === undefined) return { status: 'pendiente', result: null, cumpl: null };
  // Usar meta_periodo (captura metas dinámicas: IPC mes, meta/11, etc.)
  const metaAnual = ind.metas[yearIdx];
  const metaPer   = (ind.meta_periodos && ind.meta_periodos[flatIdx] !== null &&
                     ind.meta_periodos[flatIdx] !== undefined)
                    ? ind.meta_periodos[flatIdx] : metaAnual;
  const meta = metaPer;
  if (meta === null || meta === undefined || typeof meta !== 'number') return { status: 'pendiente', result, cumpl: null };
  // Usar cumpl_formula_type del Excel si está disponible (más preciso que la tendencia)
  const cft  = ind.cumpl_formula_type || '';
  const tend = (ind.tend || '').toLowerCase().trim();
  let c;
  if (cft === 'directo') {
    c = result;  // El resultado ya ES el cumplimiento (ej. IND-18)
  } else if (cft === 'mantener' || (!cft && tend.startsWith('mantener'))) {
    c = meta !== 0 ? 1 - Math.abs(result - meta) / Math.max(Math.abs(meta), 1) : null;
  } else if (cft === 'reducir' || (!cft && (tend.startsWith('reducir') || tend.startsWith('disminuir')))) {
    c = meta !== 0 ? 1.0 - (result - meta) / meta : null;
  } else {
    c = meta !== 0 ? result / meta : null;
  }
  const status = c === null ? 'pendiente' : getSemaforoEntry(c).name !== 'Pendiente' ? (c>=DATA.thresholds.cumple?'cumple':c>=DATA.thresholds.riesgo?'riesgo':'critico') : 'critico';
  return { status, result, cumpl: c };
}

// ── KPIs en header ────────────────────────────────────────
function renderHeaderKPIs() {
  const s = DATA.summary;
  const el = document.getElementById('header-kpis');
  if (!el) return;

  const ICO_TOTAL = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 14l4-4 4 4 5-5"/></svg>';
  const ICO_CUMPLE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M9 12l2 2 4-4"/><circle cx="12" cy="12" r="10"/></svg>';
  const ICO_RIESGO = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4"/><path d="M12 17h.01"/><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/></svg>';
  const ICO_CRITICO = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M15 9l-6 6"/><path d="M9 9l6 6"/></svg>';
  const ICO_PEND = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="12" rx="10" ry="10"/><path d="M12 6v6l4 2"/></svg>';

  const total = s.total || 0;
  const items = [
    { key:'total',     label:'Indicadores', val: total,           st:'total',     ico: ICO_TOTAL,   clickable: false, barPct: 100 },
    { key:'cumple',    label:'En meta',     val: s.cumple||0,     st:'cumple',    ico: ICO_CUMPLE,  clickable: true,  barPct: total>0?(s.cumple||0)/total*100:0 },
    { key:'riesgo',    label:'En riesgo',   val: s.riesgo||0,     st:'riesgo',    ico: ICO_RIESGO,  clickable: true,  barPct: total>0?(s.riesgo||0)/total*100:0 },
    { key:'critico',   label:'Crítico',     val: s.critico||0,    st:'critico',   ico: ICO_CRITICO, clickable: true,  barPct: total>0?(s.critico||0)/total*100:0 },
    { key:'pendiente', label:'Sin datos',   val: s.pendiente||0,  st:'pendiente', ico: ICO_PEND,    clickable: true,  barPct: total>0?(s.pendiente||0)/total*100:0 },
  ];

  el.innerHTML = items.map(c => `
    <div class="hkpi ${c.clickable && c.val > 0 ? 'clickable' : ''}" data-status="${c.st}">
      <div class="hkpi-icon ${c.st}">${c.ico}</div>
      <div class="hkpi-right" style="flex:1">
        <div class="hkpi-val">${c.val}</div>
        <div class="hkpi-lbl">${c.label}</div>
        <div class="hkpi-bar"><div class="hkpi-bar-fill ${c.st}" style="width:${c.barPct.toFixed(0)}%"></div></div>
      </div>
    </div>
  `).join('');

  el.querySelectorAll('.hkpi.clickable').forEach(card => {
    card.addEventListener('click', e => { e.stopPropagation(); openPopoverForStatus(card.dataset.status, card); });
  });
}

// ── Popover ───────────────────────────────────────────────
function getAllIndicators() {
  const arr = [];
  DATA.perspectives.forEach(p => p.indicators.forEach(ind => arr.push({ind, persp: p})));
  return arr;
}
function openPopoverForStatus(status, anchorEl) {
  closePopover();
  const all = getAllIndicators().filter(x => x.ind.latest_status === status);
  const overlay = document.getElementById('popover-overlay');
  const pop = document.getElementById('popover');
  document.querySelectorAll('.hkpi.active').forEach(k => k.classList.remove('active'));
  anchorEl.classList.add('active');
  STATE.activeKpi = status;
  const items = all.length ? all.map(({ind, persp}) => {
    const cumpl = ind.latest_cumpl !== null && ind.latest_cumpl !== undefined ? fmtPct(ind.latest_cumpl) : '—';
    return `<div class="popover-item ${status}" data-code="${ind.code}">
      <span class="pop-code">${ind.code}</span>
      <span class="pop-name">${escapeHTML(ind.name)}<br><span style="font-size:0.72rem;color:var(--text-lt)">${persp.name}</span></span>
      <span class="pop-pct">${cumpl}</span>
    </div>`;
  }).join('') : '<div class="popover-empty">Sin indicadores en este estado.</div>';
  pop.innerHTML = `
    <div class="popover-head">
      <div><div class="title">${STATUS_LABELS[status]}</div><div class="count">${all.length} indicador${all.length===1?'':'es'}</div></div>
      <button class="popover-close">×</button>
    </div>
    <div class="popover-list">${items}</div>`;
  const rect = anchorEl.getBoundingClientRect();
  const popWidth = 360;
  let left = rect.left + window.scrollX;
  const maxLeft = window.innerWidth - popWidth - 16;
  if (left > maxLeft) left = maxLeft;
  if (left < 16) left = 16;
  pop.style.left = left + 'px';
  pop.style.top = (rect.bottom + window.scrollY + 8) + 'px';
  overlay.classList.add('active');
  pop.querySelector('.popover-close').addEventListener('click', closePopover);
  overlay.addEventListener('click', closePopover);
  pop.addEventListener('click', e => e.stopPropagation());
  pop.querySelectorAll('.popover-item').forEach(it => {
    it.addEventListener('click', () => { closePopover(); openModal(it.dataset.code); });
  });
}
function closePopover() {
  document.getElementById('popover-overlay').classList.remove('active');
  document.querySelectorAll('.hkpi.active').forEach(k => k.classList.remove('active'));
  STATE.activeKpi = null;
}

// ── Filtros ───────────────────────────────────────────────
function renderFilters() {
  const perspRow = document.getElementById('persp-filter-row');
  if (perspRow) {
    // Calcular cumplimiento por perspectiva
    const effYear = effectiveYear();
    const allYears = STATE.yearFilter === 'all';

    function perspCumpl(p) {
      let sum=0,n=0;
      p.indicators.forEach(ind => {
        const yrs = allYears ? DATA.years : [effYear || CURRENT_YEAR];
        yrs.forEach(yr => {
          const yObj = ind.yearly.find(y=>y.year===yr);
          if(yObj&&yObj.cumplimiento!==null){sum+=Math.min(yObj.cumplimiento,1.5);n++;}
        });
      });
      return n>0 ? sum/n : null;
    }

    let html = '';
    html += DATA.perspectives.map((p, pi) => {
      const cp = perspCumpl(p);
      const pct = cp!==null?Math.round(cp*100)+'%':'—';
      const vivid = NM_PERSP_COLORS[pi % NM_PERSP_COLORS.length];
      const st = nmStatus(cp);
      const isActive = STATE.perspFilter === p.key;
      return `<div class="persp-filter-card ${isActive?'active':''}" data-persp="${p.key}">
        ${nmRing(cp, vivid, 50, `<span class="nm-ring-pct sm" style="color:${vivid}">${pct}</span>`)}
        <div class="pfc-info">
          <div class="pfc-name" style="${isActive?'color:'+vivid:''}">${p.name}</div>
          <div class="pfc-count">${p.indicators.length} indicadores</div>
        </div>
        <span class="nm-tick ${st}">${NM_TICK[st]}</span>
      </div>`;
    }).join('');

    perspRow.innerHTML = html;
    perspRow.querySelectorAll('.persp-filter-card').forEach(card => {
      card.addEventListener('click', () => {
        STATE.perspFilter = card.dataset.persp;
        renderFilters(); renderGauges(); renderMatriz(); renderPerspectives();
      });
    });
  }

  const years = document.getElementById('year-chips');
  const yearButtons = [{key:'all', label:'Todos los años', isCurrent:false}].concat(
    DATA.years.map(y => ({key:y, label:y.toString(), isCurrent: y===CURRENT_YEAR}))
  );
  years.innerHTML = yearButtons.map(y => {
    const isActive = (STATE.yearFilter==='current' && y.isCurrent) || (STATE.yearFilter===y.key);
    return `<div class="chip year ${y.isCurrent?'current':''} ${isActive?'active':''}" data-year="${y.key}">
      ${y.label}${y.isCurrent?' (actual)':''}
    </div>`;
  }).join('');
  years.querySelectorAll('.chip').forEach(c => c.addEventListener('click', () => {
    const v = c.dataset.year;
    STATE.yearFilter = v==='all' ? 'all' : parseInt(v);
    renderFilters(); renderGauges(); renderMatriz(); renderPerspectives();
  }));
}

// ── MATRIZ mensual de semáforos ───────────────────────────
function renderMatriz() {
  const root = document.getElementById('matriz-panel');
  if (!root) return;
  const effYear = effectiveYear();
  const allYears = STATE.yearFilter === 'all';
  const singleYr = allYears ? null : (effYear !== null ? effYear : CURRENT_YEAR);

  // Años a mostrar en columnas
  const yearsToShow = allYears ? DATA.years : [singleYr];

  const visiblePersps = DATA.perspectives.filter(p =>
    STATE.perspFilter === 'all' || STATE.perspFilter === p.key
  );

  // ── Encabezado de la tabla ────────────────────────────
  // Fila 1: grupos de año (solo en modo todos los años)
  // Fila 2: meses
  let theadRow1 = '', theadRow2 = '';
  if (allYears) {
    const nowM = new Date().getMonth();
    theadRow1 = `<tr>
      <th class="mtz-th-ind" rowspan="2" style="vertical-align:middle;text-align:center;font-size:0.70rem;font-weight:800;letter-spacing:0.08em;">INDICADOR</th>
      <th class="mtz-th-tend" rowspan="2" style="vertical-align:middle;text-align:center;font-size:0.58rem;font-weight:700;letter-spacing:0.04em;line-height:1.3;">TEND.<br>ESPERADA</th>
      ${yearsToShow.map(yr => `
        <th class="mtz-th-meta" style="font-size:0.56rem;line-height:1.2;">META<br>${yr}</th>
        <th class="mtz-th-year ${yr===CURRENT_YEAR?'mtz-th-year-current':''}" colspan="12">
          ${yr}${yr===CURRENT_YEAR?' ★ actual':''}
        </th>`).join('')}
      <th class="mtz-th-cumpl" style="right:0;vertical-align:middle;font-size:0.62rem;font-weight:800;letter-spacing:0.06em;position:sticky;z-index:11;">
        CUMPL.
      </th>
    </tr>`;
    const MFULL_A = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    theadRow2 = `<tr>
      ${yearsToShow.map((yr,yi) => {
        const mesCells = MONTH_ABBR.map((m,mi) =>
          `<th class="mtz-th-mes ${(yr===CURRENT_YEAR&&mi===nowM)?'current':''}">
            <div class="mtz-th-mes-inner">
              <span class="mtz-th-mes-abbr">${m}</span>
              <span class="mtz-th-mes-full">${MFULL_A[mi]}</span>
            </div>
          </th>`
        ).join('');
        return `<td style="background:#0D47A1;border:none;width:62px;min-width:62px;"></td>${mesCells}`;
      }).join('')}
      <th class="mtz-th-cumpl" style="right:0;position:sticky;z-index:11;font-size:0;border-top:none;"></th>
    </tr>`;
  } else {
    const now = new Date();
    const curMonthIdx = now.getMonth(); // 0-based
    const MFULL_S = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    theadRow2 = `<tr>
      <th class="mtz-th-ind" style="vertical-align:middle;text-align:center;font-size:0.70rem;font-weight:800;letter-spacing:0.08em;">INDICADOR</th>
      <th class="mtz-th-tend" style="vertical-align:middle;text-align:center;font-size:0.58rem;font-weight:700;letter-spacing:0.04em;line-height:1.3;">TEND.<br>ESPERADA</th>
      <th class="mtz-th-meta">META<br>${singleYr}</th>
      ${MONTH_ABBR.map((m,mi)=>`<th class="mtz-th-mes ${mi===curMonthIdx&&singleYr===CURRENT_YEAR?'current':''}">
        <div class="mtz-th-mes-inner">
          <span class="mtz-th-mes-abbr">${m}</span>
          <span class="mtz-th-mes-full">${MFULL_S[mi]}</span>
        </div>
      </th>`).join('')}
      <th class="mtz-th-cumpl" style="right:0;vertical-align:middle;font-size:0.62rem;font-weight:800;letter-spacing:0.06em;line-height:1.3;">
        CUMPL.<br>${singleYr}
      </th>
    </tr>`;
  }

  // Calcular posición sticky de la columna Cumpl.
  // = width de col indicador (220) + 12*62 * nYears
  const cumplRight = 0;

  // ── Filas de la tabla ─────────────────────────────────
  let rows = '';
  visiblePersps.forEach(p => {
    p.indicators.forEach((ind, i) => {
      let monthCells = '';
      yearsToShow.forEach((yr, yi) => {
        const yIdx = yr - 2026;
        // ── Celda META del año (antes de los 12 meses) ──────────────
        const metaYr = ind.metas[yIdx];
        const metaYrFmt = (metaYr !== null && typeof metaYr === 'number')
          ? fmtSmart(metaYr, ind.tipo_meta, ind.unidad, ind.decimales) : '—';
        monthCells += `<td class="mtz-meta-cell">
          <div class="mtz-meta-val">${metaYrFmt}</div>
        </td>`;
        // ── 12 meses del año ─────────────────────────────────────────
        for (let m = 0; m < 12; m++) {
          const mc = monthCompliance(ind, yIdx, m);
          const resultFmt = mc.result!==null ? fmt(mc.result,ind.tipo_meta,ind.unidad,ind.decimales) : '—';
          const metaFmt = (ind.metas[yIdx]!==null&&typeof ind.metas[yIdx]==='number') ? fmtSmart(ind.metas[yIdx],ind.tipo_meta,ind.unidad,ind.decimales) : '—';
          const cumplFmt = mc.cumpl!==null ? (mc.cumpl*100).toFixed(0)+'%' : '—';
          const dispVal  = mc.result!==null ? fmtMtz(mc.result,ind.tipo_meta,ind.unidad,ind.decimales) : '–';
          // Etiqueta de estado usando colorimetría del semáforo
          // Colores exactos del semáforo (10 niveles)
          const semaEntry  = getSemaforoEntry(mc.cumpl);
          const estadoLbl  = mc.cumpl !== null ? semaEntry.name : '';
          monthCells += `<td class="mtz-month-cell">
            <div class="mtz-light ${mc.status}" data-estado="${estadoLbl}"
                 data-month="${MONTH_ABBR[m]} ${yr}" data-result="${escapeHTML(resultFmt)}"
                 data-meta="${escapeHTML(metaFmt)}" data-cumpl="${cumplFmt}"
                 data-status="${mc.status}" data-code="${ind.code}">
              <span class="nm-tick ${mc.status}">${NM_TICK[mc.status] || ''}</span>
              <span class="mtz-val">${dispVal}</span>
            </div></td>`;
        }
      });

      // Cumplimiento: promedio de los años mostrados
      let sumC=0,nC=0;
      yearsToShow.forEach(yr=>{
        const yObj=ind.yearly.find(y=>y.year===yr);
        if(yObj&&yObj.cumplimiento!==null){sumC+=Math.min(yObj.cumplimiento,1.5);nC++;}
      });
      const ac=nC>0?sumC/nC:null;
      const ast=ac===null?'pendiente':ac>=DATA.thresholds.cumple?'cumple':ac>=DATA.thresholds.riesgo?'riesgo':'critico';
        const astName=ac!==null?getSemaforoEntry(ac).name:'Sin datos';
      const apct=ac!==null?(ac*100).toFixed(0)+'%':'—';
      const abar=ac!==null?Math.min(100,Math.max(0,ac*100)):0;
      const aclr=ac!==null?getSemaforoEntry(ac).fg:'#CBD5E1';

      // Tendencia badge
      const tendLow = (ind.tend||'').toLowerCase();
      const tendIco   = tendLow.includes('crec') ? '↑' : tendLow.includes('reduc') ? '↓' : '↔';
      const tendLbl   = ind.tend || '—';
      const formulaShort = escapeHTML(ind.formula||'');

      rows += `<tr class="${i===0?'mtz-first-row':''}" data-code="${ind.code}">
        <td class="mtz-ind-cell" title="${formulaShort}">
          <div class="mtz-ind-code" style="color:${p.accent}">${ind.code}</div>
          <div class="mtz-ind-name">${escapeHTML(ind.name)}</div>
        </td>
        <td class="mtz-tend-cell">
          <span style="font-size:0.72rem;font-weight:600;color:var(--text-dark);">${tendIco} ${tendLbl}</span>
        </td>
        ${monthCells}
        <td class="mtz-cumpl-cell" style="right:0">
          <div class="mtz-cumpl-pct ${ast}">${apct}</div>
          <div class="mtz-mini-bar"><div class="mtz-mini-bar-fill" style="width:${abar}%;background:${aclr}"></div></div>
        </td>
      </tr>`;
    });
  });

  const now = new Date();
  const dateStr = now.toLocaleDateString('es-CO',{day:'2-digit',month:'2-digit',year:'numeric'})+' '+now.toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'});
  const totalCols = yearsToShow.length * 12;
  const tableW = 300 + 80 + yearsToShow.length * (62 + 12*72) + 88;  // ind+tend+meta_por_año+cumpl

  root.innerHTML = `
    <div class="matriz-layout">
      <div class="matriz-card">
        <div class="matriz-scroll">
          <table class="matriz-table" style="width:100%;min-width:${tableW}px;--mtz-hdr:${
            (() => {
              const p = DATA.perspectives.find(p=>p.key===STATE.perspFilter);
              return p ? p.accent : '#0657A4';
            })()
          };">
            <thead>
              ${theadRow1}
              ${theadRow2}
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
        <div class="matriz-footer-bar">
          <div class="mleg">
            <span class="mleg-item"><span class="mleg-dot2 cumple"></span> En meta (≥${Math.round(DATA.thresholds.cumple*100)}%)</span>
            <span class="mleg-item"><span class="mleg-dot2 riesgo"></span> En riesgo (${Math.round(DATA.thresholds.riesgo*100)}–${Math.round(DATA.thresholds.cumple*100)-1}%)</span>
            <span class="mleg-item"><span class="mleg-dot2 critico"></span> Crítico (&lt;${Math.round(DATA.thresholds.riesgo*100)}%)</span>
            <button class="esc-link" onclick="openEscalas()">Ver escalas</button>
            <span class="mleg-item"><span class="mleg-dot2 pendiente"></span> Sin datos</span>
          </div>
          <div class="mtz-last-update">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>
            ${dateStr}
          </div>
        </div>

        <!-- Vista de lista para móvil (visible solo en pantallas pequeñas) -->
        <div class="mtz-mobile-list" style="display:none;padding:8px 10px 4px;">
          ${(() => {
            const yr2 = allYears ? CURRENT_YEAR : (singleYr || CURRENT_YEAR);
            return visiblePersps.map(p => `
              <div style="margin-bottom:10px">
                <div style="font-size:0.62rem;font-weight:800;color:${p.accent};text-transform:uppercase;letter-spacing:0.08em;padding:4px 2px 5px;border-bottom:2px solid ${p.accent}20;margin-bottom:6px;">${p.name}</div>
                ${p.indicators.map(ind => {
                  const yObj2 = ind.yearly.find(y=>y.year===yr2)||ind.yearly[0];
                  const cp2 = yObj2?yObj2.cumplimiento:null;
                  const st2 = cp2===null?'pendiente':cp2>=DATA.thresholds.cumple?'cumple':cp2>=DATA.thresholds.riesgo?'riesgo':'critico';
                  const stLbl2 = cp2!==null ? getSemaforoEntry(cp2).name : 'Sin datos';
                  const stClr2 = {cumple:'#16A34A',riesgo:'#D97706',critico:'#DC2626',pendiente:'#94A3B8'}[st2];
                  const stBg2 = {cumple:'#DCFCE7',riesgo:'#FEF3C7',critico:'#FEE2E2',pendiente:'#F1F5F9'}[st2];
                  const lastD2 = findLastDataMonth(ind);
                  const pct2 = cp2!==null?Math.round(cp2*100)+'%':'—';
                  return `<div class="mtz-mob-row" data-code="${ind.code}"
                    style="background:white;border-radius:10px;padding:10px 12px;margin-bottom:6px;border:1px solid #E2E8F0;display:flex;align-items:center;gap:10px;cursor:pointer;box-shadow:0 1px 3px rgba(0,0,0,0.04);">
                    <div style="flex:1;min-width:0;">
                      <div style="font-size:0.68rem;font-weight:800;color:${p.accent};letter-spacing:0.05em;margin-bottom:1px;">${ind.code}</div>
                      <div style="font-size:0.83rem;font-weight:700;color:#0F172A;line-height:1.3;white-space:normal;">${escapeHTML(ind.name)}</div>
                      <div style="font-size:0.60rem;color:#64748B;font-style:italic;margin-top:2px;line-height:1.3;white-space:normal;">${escapeHTML(ind.formula||'')}</div>
                      ${lastD2?`<div style="font-size:0.62rem;color:#94A3B8;margin-top:2px;">Último: ${fmt(lastD2.val,ind.tipo_meta,ind.unidad)} · ${lastD2.label}</div>`:''}
                    </div>
                    <div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;flex-shrink:0;">
                      <span style="font-size:1.10rem;font-weight:900;color:${stClr2};line-height:1;">${pct2}</span>
                      <span style="background:${stBg2};color:${stClr2};font-size:0.60rem;font-weight:700;padding:2px 7px;border-radius:999px;">${stLbl2}</span>
                    </div>
                  </div>`;
                }).join('')}
              </div>
            `).join('');
          })()}
        </div>
      </div>
    </div>`;

  // ── Tooltip y clicks de celdas ────────────────────────
  const tip = document.getElementById('mtz-tooltip');
  if (!tip) return;
  root.querySelectorAll('.mtz-light').forEach(light => {
    light.addEventListener('mouseenter', () => {
      const d = light.dataset;
      tip.innerHTML = `<div class="t-month">${d.month}${d.estado ? ' · ' + d.estado : ''}</div><div class="t-val">${d.result}</div><div class="t-meta">Meta: ${d.meta}</div><div class="t-cumpl ${d.status}">Cumplimiento: ${d.cumpl}</div>`;
      tip.classList.add('visible');
    });
    light.addEventListener('mousemove', e => {
      const w=tip.offsetWidth,h=tip.offsetHeight;
      let x=e.clientX+14,y=e.clientY+14;
      if(x+w>window.innerWidth-12) x=e.clientX-w-14;
      if(y+h>window.innerHeight-12) y=e.clientY-h-14;
      tip.style.left=x+'px'; tip.style.top=y+'px';
    });
    light.addEventListener('mouseleave', () => tip.classList.remove('visible'));
  });
  // Clic en cualquier parte de la fila (celda mensual incluida) → abre el detalle
  root.querySelectorAll('tr[data-code]').forEach(row => {
    row.style.cursor = 'pointer';
    row.addEventListener('click', () => {
      tip.classList.remove('visible');
      openModal(row.dataset.code);
    });
  });
  root.querySelectorAll('.mtz-mob-row').forEach(row => {
    row.addEventListener('click', () => openModal(row.dataset.code));
  });
  updateCompactState();
}

// ── Perspectivas e indicadores ────────────────────────────
function renderPerspectives() {
  const root = document.getElementById('perspectives');
  Object.entries(STATE.charts).forEach(([k,ch]) => { if(k!=='detail'){try{ch.destroy();}catch(e){}} });
  STATE.charts = STATE.charts.detail ? {detail: STATE.charts.detail} : {};
  if (!root) return;  // sección de tarjetas retirada del tablero

  const visiblePersps = DATA.perspectives.filter(p =>
    STATE.perspFilter === 'all' || STATE.perspFilter === p.key
  );
  if (!visiblePersps.length) {
    root.innerHTML = '<div style="text-align:center;padding:40px;color:var(--text-lt)">Sin perspectivas que coincidan con el filtro.</div>';
    return;
  }
  const effYear = effectiveYear();
  root.innerHTML = visiblePersps.map(p => {
    const stats = computePerspStats(p);
    const cards = p.indicators.map(ind => renderCard(ind, p, effYear)).join('');
    return `<section class="perspective">
      <div class="persp-header" style="background:linear-gradient(135deg,${p.accent},${shade(p.accent,-18)})">
        <div class="persp-icon-wrap">${ICONS[p.key]||''}</div>
        <div class="persp-titles">
          <div class="persp-eyebrow">Perspectiva estratégica</div>
          <div class="persp-title">${p.name}</div>
        </div>
        <div class="persp-stats">
          <div class="persp-stat"><strong>${p.indicators.length}</strong> ind.</div>
          ${stats.cumple  ? `<div class="persp-stat">✅ ${stats.cumple}</div>` : ''}
          ${stats.riesgo  ? `<div class="persp-stat">⚠️ ${stats.riesgo}</div>` : ''}
          ${stats.critico ? `<div class="persp-stat">🔴 ${stats.critico}</div>` : ''}
        </div>
      </div>
      <div class="persp-objetivo" style="color:${p.accent}">
        <strong>Objetivo general:</strong> <span style="color:var(--text-med)">${escapeHTML(p.objetivo_general)}</span>
      </div>
      <div class="ind-grid">${cards}</div>
    </section>`;
  }).join('');
  // Registrar listener ANTES de buildMiniChart (que puede fallar si Chart.js no cargó)
  root.removeEventListener('click', root._cardClickHandler);
  root._cardClickHandler = (e) => {
    const card = e.target.closest('.ind-card');
    if (!card) return;
    if (e.target.closest('button') || e.target.closest('a')) return;
    openModal(card.dataset.code);
  };
  root.addEventListener('click', root._cardClickHandler);

  // buildMiniChart en try/catch para que un fallo no bloquee lo anterior
  visiblePersps.forEach(p => p.indicators.forEach(ind => {
    try { buildMiniChart(ind); } catch(e) { /* Chart.js no disponible */ }
  }));
}

function computePerspStats(persp) {
  const s = {cumple:0,riesgo:0,critico:0,pendiente:0};
  persp.indicators.forEach(ind => { if(s[ind.latest_status]!==undefined) s[ind.latest_status]++; });
  return s;
}

function findLastDataMonth(ind) {
  for (let i = ind.monthly_results.length - 1; i >= 0; i--) {
    const v = ind.monthly_results[i];
    if (v !== null && v !== undefined) {
      const yearIdx  = Math.floor(i / 12);
      const monthIdx = i % 12;
      return { year: 2026 + yearIdx, monthIdx, val: v, label: MONTH_ABBR[monthIdx] + ' ' + (2026 + yearIdx) };
    }
  }
  return null;
}

function renderCard(ind, persp, effYear) {
  const allYears = STATE.yearFilter === 'all';
  let dyObj = null;
  if (effYear !== null) { dyObj = ind.yearly.find(y => y.year === effYear); }
  else { for(let i=ind.yearly.length-1;i>=0;i--){ if(ind.yearly[i].cumplimiento!==null){dyObj=ind.yearly[i];break;} } }
  if (!dyObj) dyObj = ind.yearly[0];
  const status = dyObj.status;
  const meta   = fmtSmart(dyObj.meta_raw, ind.tipo_meta, ind.unidad, ind.decimales);
  const cumpl  = fmtPct(dyObj.cumplimiento);
  const yr     = dyObj.year;
  const acc    = ind.accumulated !== null ? fmtPct(ind.accumulated) : '—';
  const alert  = ind.alert || {level:'pending', text:'Sin datos cargados aún'};

  // Último dato real
  const lastData  = findLastDataMonth(ind);
  const lastVal   = lastData ? fmt(lastData.val, ind.tipo_meta, ind.unidad) : '—';
  const lastLabel = lastData ? lastData.label : 'Sin datos';

  // Tendencia
  const tendLower = (ind.tend||'').toLowerCase();
  const tendIcon  = tendLower.includes('reducir')||tendLower.includes('disminuir') ? '↓'
                  : tendLower.includes('crecer')||tendLower.includes('aumentar')   ? '↑' : '→';
  const tendSub   = tendLower.includes('reducir')||tendLower.includes('disminuir') ? 'Menor es mejor'
                  : tendLower.includes('crecer')||tendLower.includes('aumentar')   ? 'Mayor es mejor' : 'Estable';
  const tendColor = tendLower.includes('reducir')||tendLower.includes('disminuir') ? '#DC2626'
                  : tendLower.includes('crecer')||tendLower.includes('aumentar')   ? '#16A34A' : '#64748B';

  // Alerta
  const alertIcons = { ok:'✅', warn:'⚠️', critical:'🚨', pending:'ℹ️' };
  const alertTitles = {
    ok: 'En meta para ' + yr,
    warn: 'Atención en ' + yr + ': ' + cumpl + ' de la meta',
    critical: 'Crítico en ' + yr + ': ' + cumpl + ' de la meta',
    pending: 'Sin datos cargados aún'
  };
  // Construir explicación usando cumpl_formula_type (leído del Excel) — no la tendencia textual
  const cftExpl = (ind.cumpl_formula_type || '').toLowerCase();
  const tendExpl = (cftExpl === 'reducir' || (!cftExpl && (tendLower.includes('reducir')||tendLower.includes('disminuir'))))
    ? `Fórmula Excel: 1 − (Resultado − Meta) / Meta. La meta es ${fmtSmart(dyObj.meta_raw, ind.tipo_meta, ind.unidad, ind.decimales)}. Mientras el resultado sea menor a la meta, mayor el cumplimiento.`
    : (cftExpl === 'mantener' || (!cftExpl && tendLower.includes('mantener')))
    ? `Fórmula Excel: 1 − |Resultado − Meta| / MAX(|Meta|, 1). Meta: ${fmtSmart(dyObj.meta_raw, ind.tipo_meta, ind.unidad, ind.decimales)}. Entre más cercano a la meta, mayor el cumplimiento.`
    : cftExpl === 'directo'
    ? `Fórmula Excel: el resultado es directamente el % de cumplimiento. Meta referencial: ${fmtSmart(dyObj.meta_raw, ind.tipo_meta, ind.unidad, ind.decimales)}.`
    : `Fórmula Excel: Resultado ÷ Meta. Meta: ${fmtSmart(dyObj.meta_raw, ind.tipo_meta, ind.unidad, ind.decimales)}. Entre mayor el resultado respecto a la meta, mejor el cumplimiento.`;

  const alertSubs = {
    ok:       tendExpl,
    warn:     tendExpl,
    critical: tendExpl,
    pending:  `Fórmula de evaluación: ${ind.formula}. Ingresa los datos en el Excel y regenera el dashboard para ver el cumplimiento.`
  };

  // KPI icons
  const icoMeta   = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>`;
  const icoCal    = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`;
  const icoCheck  = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 12l2 2 4-4"/><circle cx="12" cy="12" r="10"/></svg>`;
  const icoTrend  = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/></svg>`;
  const icoGlobe  = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 014 10 15.3 15.3 0 01-4 10 15.3 15.3 0 01-4-10 15.3 15.3 0 014-10z"/></svg>`;
  const icoUser   = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>`;

  // Color status para último dato
  const lastStatus = lastData ? (() => {
    // buscar estado del mes del último dato
    const yi = lastData.year - 2026;
    const mc = monthCompliance(ind, yi, lastData.monthIdx);
    return mc.status;
  })() : 'pendiente';

  return `<div class="ind-card" data-code="${ind.code}">
    <div class="accent-bar" style="background:linear-gradient(180deg,${persp.accent},${persp.accent}99)"></div>

    <!-- HEAD -->
    <div class="ind-card-head">
      <div style="min-width:0;flex:1">
        <div class="ind-code-badge" style="color:${persp.accent}">${ind.code}</div>
        <h3 class="ind-name">${escapeHTML(ind.name)}</h3>
        <div class="ind-ini">${ind.ini_icon} <span style="color:var(--text-med)">${escapeHTML(ind.ini)}</span></div>
      </div>
      <span class="status-pill" style="${status!=='pendiente'&&ind.cumplimiento!==null?`background:${getSemaforoEntry(ind.cumplimiento).bg};color:${getSemaforoEntry(ind.cumplimiento).fg};border:1px solid ${getSemaforoEntry(ind.cumplimiento).fg}30;`:''}">${semaforoLabel(ind.cumplimiento,status)}</span>
    </div>

    <!-- 3 KPIs o tabla por año -->
    ${allYears ? `
    <div class="ind-kpis">
      <div class="ind-kpi">
        <div class="ind-kpi-icon" style="background:#EFF6FF;color:#2563EB">${icoMeta}</div>
        <div class="ind-kpi-lbl">Último dato</div>
        <div class="ind-kpi-val ${lastStatus}">${lastVal}</div>
        <div class="ind-kpi-sub">${lastLabel}</div>
      </div>
      <div class="ind-kpi">
        <div class="ind-kpi-icon" style="background:${hexA(tendColor,0.12)};color:${tendColor}">${icoCheck}</div>
        <div class="ind-kpi-lbl">Cumpl. ${yr}</div>
        <div class="ind-kpi-val ${status}">${cumpl}</div>
      </div>
      <div class="ind-kpi">
        <div class="ind-kpi-icon" style="background:#F0FDF4;color:#16A34A">${icoGlobe}</div>
        <div class="ind-kpi-lbl">Avance global</div>
        <div class="ind-kpi-val">${acc}</div>
      </div>
    </div>
    <div class="ind-metas-yr">
      ${ind.yearly.map(y => {
        const m = fmtSmart(y.meta_raw, ind.tipo_meta, ind.unidad, ind.decimales);
        const st = y.cumplimiento !== null ? (y.cumplimiento >= DATA.thresholds.cumple ? 'cumple' : y.cumplimiento >= DATA.thresholds.riesgo ? 'riesgo' : 'critico') : 'pendiente';
        const stNm = y.cumplimiento!==null ? getSemaforoEntry(y.cumplimiento).name : 'Sin datos';
        const cp = y.cumplimiento !== null ? Math.round(y.cumplimiento*100)+'%' : '—';
        return `<div class="ind-meta-yr-cell ${y.year===CURRENT_YEAR?'current':''}">
          <div class="ind-meta-yr-lbl">${y.year}</div>
          <div class="ind-meta-yr-val">${m}</div>
          <div class="ind-meta-yr-cumpl ${st}">${cp}</div>
        </div>`;
      }).join('')}
    </div>` : `
    <div class="ind-kpis">
      <div class="ind-kpi">
        <div class="ind-kpi-icon" style="background:#EFF6FF;color:#2563EB">${icoMeta}</div>
        <div class="ind-kpi-lbl">Meta ${yr}</div>
        <div class="ind-kpi-val">${meta}</div>
      </div>
      <div class="ind-kpi">
        <div class="ind-kpi-icon" style="background:${hexA(tendColor,0.12)};color:${tendColor}">${icoCal}</div>
        <div class="ind-kpi-lbl">Último dato<br>${lastLabel}</div>
        <div class="ind-kpi-val ${lastStatus}">${lastVal}</div>
      </div>
      <div class="ind-kpi">
        <div class="ind-kpi-icon" style="background:${status==='cumple'?'#DCFCE7':status==='riesgo'?'#FEF3C7':status==='critico'?'#FEE2E2':'#F1F5F9'};color:${status==='cumple'?'#16A34A':status==='riesgo'?'#D97706':status==='critico'?'#DC2626':'#94A3B8'}">${icoCheck}</div>
        <div class="ind-kpi-lbl">Cumplimiento ${yr}</div>
        <div class="ind-kpi-val ${status}">${cumpl}</div>
      </div>
    </div>`}

    <!-- Mini chart -->
    <div class="chart-wrap"><canvas id="chart-${ind.code}"></canvas></div>

    <!-- Footer -->
    <div class="ind-foot">
      <div class="foot-kpi">
        <div class="foot-kpi-icon" style="background:${hexA(tendColor,0.10)};color:${tendColor}">${icoTrend}</div>
        <div class="foot-kpi-info">
          <div class="l">Tendencia</div>
          <div class="v" style="color:${tendColor}">${tendIcon} ${ind.tend}</div>
          <div class="s">${tendSub}</div>
        </div>
      </div>
      <div class="foot-kpi">
        <div class="foot-kpi-icon" style="background:#EFF6FF;color:#2563EB">${icoGlobe}</div>
        <div class="foot-kpi-info">
          <div class="l">Avance global</div>
          <div class="v">${acc}</div>
          <div class="s">Respecto a la meta</div>
        </div>
      </div>
      <div class="foot-kpi">
        <div class="foot-kpi-icon" style="background:#F0FDF4;color:#16A34A">${icoUser}</div>
        <div class="foot-kpi-info">
          <div class="l">Responsable</div>
          <div class="v">${escapeHTML(ind.responsable||'Por definir')}</div>
        </div>
      </div>
    </div>

    <!-- Alerta -->
    <div class="ind-alert ${alert.level}">
      <span class="al-icon">${alertIcons[alert.level]||'ℹ️'}</span>
      <div class="al-body">
        <div class="al-title">${escapeHTML(alert.text || alertTitles[alert.level])}</div>
        <div class="al-sub">${alertSubs[alert.level]||''}</div>
      </div>
    </div>
  </div>`;
}


// ─── Función reutilizable de etiquetas inteligentes intercaladas ────────
function drawSmartLabels(chart, data, color, fmt_fn, tipo_meta, unidad, opts) {
  const meta0 = chart.getDatasetMeta(0);
  const {ctx: c} = chart;
  const area  = chart.chartArea;
  opts = opts || {};
  const FONT_SZ  = opts.fontSize  || 10;
  const LBL_H    = opts.lblH      || (FONT_SZ + 7);
  const LBL_PAD  = opts.lblPad    || 6;
  const MAX_LBL  = opts.maxLabels || 8;   // máximo etiquetas totales
  const MIN_PX_GAP = opts.minGap  || 44;  // px mínimo entre etiquetas en X

  const realIdxs = data.map((v,i)=>v!==null?i:-1).filter(i=>i>=0);
  if (!realIdxs.length) return;

  c.font = `700 ${FONT_SZ}px sans-serif`;

  // ── Paso 1: candidatos por prioridad ────────────────────────────
  const maxI  = realIdxs.reduce((a,b)=>data[a]>=data[b]?a:b);
  const minI  = realIdxs.reduce((a,b)=>data[a]<=data[b]?a:b);
  const first = realIdxs[0], last = realIdxs[realIdxs.length-1];

  // Inflexiones (picos/valles locales)
  const inflect = [];
  for (let k=1; k<realIdxs.length-1; k++) {
    const [pi,ci,ni] = [realIdxs[k-1], realIdxs[k], realIdxs[k+1]];
    const [pv,cv,nv] = [data[pi], data[ci], data[ni]];
    if (pv!==null && cv!==null && nv!==null &&
        ((cv>pv && cv>nv) || (cv<pv && cv<nv))) inflect.push(ci);
  }

  // Si hay pocos puntos, mostrarlos todos; si hay muchos, selección por separación X
  let selected;
  if (realIdxs.length <= MAX_LBL) {
    selected = [...realIdxs];
  } else {
    // Selección uniforme: tomar uno cada N puntos + siempre first/last/max/min
    const must = new Set([first, last, maxI, minI]);
    const step = Math.ceil(realIdxs.length / (MAX_LBL - must.size));
    for (let k=0; k<realIdxs.length; k+=step) must.add(realIdxs[k]);
    selected = [...must].sort((a,b)=>a-b);
    // Filtrar por separación mínima en X (greedy sobre los ya ordenados)
    const filtered = [selected[0]];
    for (let k=1; k<selected.length; k++) {
      const ptA = meta0.data[filtered[filtered.length-1]];
      const ptB = meta0.data[selected[k]];
      if (ptA && ptB && ptB.x - ptA.x >= MIN_PX_GAP) filtered.push(selected[k]);
      else if (selected[k] === last || selected[k] === maxI || selected[k] === minI) {
        // Siempre incluir max/min/last aunque estén cerca — reemplazar el vecino
        filtered[filtered.length-1] = selected[k];
      }
    }
    selected = filtered;
  }
  if (!selected.length) return;

  // ── Paso 2: construir objetos label con posición intercalada ───────
  // Intercalar arriba/abajo: par→arriba, impar→abajo
  const labels = selected.map((i, idx) => {
    const pt  = meta0.data[i];
    if (!pt) return null;
    const v   = data[i];
    const txt = fmt_fn(v, tipo_meta, unidad);
    const tw  = c.measureText(txt).width + LBL_PAD * 2;
    const above = (idx % 2 === 0);  // intercalado
    const cy_init = above
      ? pt.y - LBL_H - 6        // encima del punto
      : pt.y + 8;                // debajo del punto
    return { i, pt, txt, w: tw, h: LBL_H, cx: pt.x,
             cy: cy_init, above, anchor: pt.y };
  }).filter(Boolean);

  // ── Paso 3: anti-colisión vertical (solo dentro del mismo lado) ───
  for (let iter=0; iter<30; iter++) {
    let moved = false;
    for (let a=0; a<labels.length; a++) {
      for (let b=a+1; b<labels.length; b++) {
        const la=labels[a], lb=labels[b];
        const xOver = Math.abs(la.cx-lb.cx) < (la.w+lb.w)/2 + 3;
        const yOver = Math.abs(la.cy-lb.cy) < la.h + 3;
        if (xOver && yOver) {
          const push = (la.h + 3 - Math.abs(la.cy-lb.cy))/2 + 1;
          if (la.cy <= lb.cy) { la.cy -= push; lb.cy += push; }
          else                { la.cy += push; lb.cy -= push; }
          moved = true;
        }
      }
    }
    if (!moved) break;
  }

  // Clamp al área
  labels.forEach(lb => {
    const topLimit  = area.top + 2;
    const botLimit  = area.bottom - lb.h - 2;
    lb.cy = Math.max(topLimit, Math.min(botLimit, lb.cy));
  });

  // ── Paso 4: dibujar línea guía + caja + texto ─────────────────────
  labels.forEach(lb => {
    const {pt, txt, cx, cy, w, h, above} = lb;
    c.save();

    // Línea guía del label al punto
    const lineTop    = above ? cy + h + 1 : cy - 1;
    const lineBottom = above ? pt.y - 4   : pt.y + 4;
    if (Math.abs(lineTop - lineBottom) > 6) {
      c.beginPath();
      c.moveTo(cx, lineTop);
      c.lineTo(cx, lineBottom);
      c.strokeStyle = hexA(color, 0.28);
      c.lineWidth   = 1;
      c.setLineDash([2,3]);
      c.stroke();
      c.setLineDash([]);
    }

    // Punto de anclaje en la curva
    c.beginPath();
    c.arc(pt.x, pt.y, 3, 0, Math.PI*2);
    c.fillStyle = color;
    c.fill();
    c.strokeStyle = 'white';
    c.lineWidth = 1.5;
    c.stroke();

    // Caja con sombra sutil
    c.shadowColor = 'rgba(0,0,0,0.10)';
    c.shadowBlur  = 4;
    c.fillStyle   = 'white';
    c.beginPath();
    c.roundRect(cx - w/2, cy, w, h, 4);
    c.fill();
    c.shadowBlur  = 0;
    c.strokeStyle = hexA(color, 0.40);
    c.lineWidth   = 1;
    c.stroke();

    // Texto
    c.fillStyle    = color;
    c.font         = `700 ${FONT_SZ}px sans-serif`;
    c.textAlign    = 'center';
    c.textBaseline = 'middle';
    c.fillText(txt, cx, cy + h/2);

    c.restore();
  });
}

function buildMiniChart(ind) {
  const ctx = document.getElementById('chart-'+ind.code);
  if (!ctx) return;
  const persp = DATA.perspectives.find(p => p.indicators.includes(ind));
  const color = persp ? persp.accent : '#0876B2';
  const data  = ind.monthly_results.slice();
  const metaLine = [];
  for(let i=0;i<60;i++){const yi=Math.floor(i/12);const m=ind.metas[yi];metaLine.push(typeof m==='number'?m:null);}

  // Rango Y: incluir datos y meta. Si meta es >8x mayor, usar solo datos.
  const dataVals = data.filter(v=>v!==null);
  const metaVals = metaLine.filter(v=>v!==null);
  let yMin, yMax;
  if(dataVals.length || metaVals.length) {
    const dMn = dataVals.length ? Math.min(...dataVals) : null;
    const dMx = dataVals.length ? Math.max(...dataVals) : null;
    const mMn = metaVals.length ? Math.min(...metaVals) : null;
    const mMx = metaVals.length ? Math.max(...metaVals) : null;
    let useVals;
    if(dMx !== null && mMx !== null) {
      const dScale = Math.max(Math.abs(dMx), Math.abs(dMn||0));
      const mScale = Math.max(Math.abs(mMx), Math.abs(mMn||0));
      const ratio3 = dScale < 1e-6 ? 1 : mScale / dScale;
      useVals = (ratio3 > 8 && ind.tipo !== 'base_social') ? dataVals : [...dataVals, ...metaVals];
    } else {
      useVals = dataVals.length ? dataVals : metaVals;
    }
    yMin=Math.min(...useVals); yMax=Math.max(...useVals);
    const span=(yMax-yMin)||Math.abs(yMax)*0.5||0.5;
    yMin-=span*0.18; yMax+=span*0.22;
  }

  // Plugin separadores de año para el mini chart
  const miniYearSep = {
    id: 'miniYearSep',
    afterDraw(chart) {
      if (!STATE.yearFilter || STATE.yearFilter !== 'all') return;
      const {ctx: c, scales:{x,y}} = chart;
      if(!x||!y) return;
      [12,24,36,48].forEach(idx => {
        const xp = x.getPixelForValue(idx);
        c.save();
        c.strokeStyle = 'rgba(100,116,139,0.25)';
        c.lineWidth = 1; c.setLineDash([3,3]);
        c.beginPath(); c.moveTo(xp, y.top); c.lineTo(xp, y.bottom); c.stroke();
        c.setLineDash([]);
        c.fillStyle = '#94A3B8';
        c.font = '600 8px sans-serif'; c.textAlign = 'center';
        c.fillText((2026 + Math.floor(idx/12)).toString(), xp, y.top + 8);
        c.restore();
      });
      // Etiqueta 2026
      const x0 = x.getPixelForValue(0), x11 = x.getPixelForValue(11);
      const c2 = chart.ctx;
      c2.save();
      c2.fillStyle = '#94A3B8'; c2.font = '600 8px sans-serif'; c2.textAlign = 'center';
      c2.fillText('2026', (x0+x11)/2, y.top + 8);
      c2.restore();
    }
  };

  const miniSmartLabels = {
    id: 'miniSmartLabels',
    afterDatasetsDraw(chart) {
      drawSmartLabels(chart, chart.data.datasets[0].data, color, fmt, ind.tipo_meta, ind.unidad,
        { fontSize:9, lblH:15, maxLabels:6, minGap:36 });
    }
  };

  STATE.charts[ind.code] = new Chart(ctx, {
    type:'line',
    data:{
      labels: DATA.months_labels,
      datasets:[
        {label:'Resultado',data,borderColor:color,
         backgroundColor:(ctx2)=>{const g=ctx2.chart.ctx.createLinearGradient(0,0,0,ctx2.chart.height);g.addColorStop(0,hexA(color,0.22));g.addColorStop(1,hexA(color,0.00));return g;},
         tension:0.4,fill:true,pointRadius:3,pointHoverRadius:5,spanGaps:false,borderWidth:2.2},
        {label:'Meta',data:metaLine,borderColor:'#94A3B8',borderDash:[4,3],borderWidth:1.3,pointRadius:0,fill:false,spanGaps:true}
      ]},
    plugins:[miniYearSep, miniSmartLabels],
    options:{
      responsive:true,maintainAspectRatio:false,
      layout:{ padding:{ top: 18, bottom: 4, left: 2, right: 2 } },
      plugins:{legend:{display:false},tooltip:{
        backgroundColor:'rgba(15,23,42,0.90)',titleColor:'#94A3B8',bodyColor:'white',padding:10,cornerRadius:8,
        callbacks:{label:(c)=>c.raw===null?null:' '+c.dataset.label+': '+fmt(c.raw,ind.tipo_meta,ind.unidad)}}},
      scales:{
        x:{display:false},
        y:{display:false,min:yMin,max:yMax}
      }
    }
  });
}

// ── Modal ─────────────────────────────────────────────────
function openModal(code) {
  const ind = findIndicator(code);
  if (!ind) return;
  const persp = DATA.perspectives.find(p => p.indicators.includes(ind));
  const color = persp ? persp.accent : '#0876B2';
  const overlay = document.getElementById('modal-overlay');
  overlay.classList.add('active');

  // ── Tendencia tag ─────────────────────────────────────
  const tendLower = (ind.tend||'').toLowerCase();
  const tendClass = tendLower.includes('reducir')||tendLower.includes('disminuir') ? 'tend-reducir'
                  : tendLower.includes('crecer')||tendLower.includes('aumentar')    ? 'tend-crecer'
                  : 'tend-mantener';
  const tendIcon  = tendLower.includes('reducir')||tendLower.includes('disminuir') ? '↘'
                  : tendLower.includes('crecer')||tendLower.includes('aumentar')    ? '↗' : '→';

  // ── Mini donut SVG para año ───────────────────────────
  function miniDonut(pct, sz) {
    sz = sz || 44;
    const r = sz/2 - 5; const circ = 2*Math.PI*r;
    const c = pct===null ? 0 : Math.min(Math.max(pct,0),1);
    const dashOffset = circ*(1-c);
    const col = pctToColor(pct);
    const txt = pct===null ? '—' : Math.round(pct*100)+'%';
    const fs = sz*0.22;
    return `<svg viewBox="0 0 ${sz} ${sz}" width="${sz}" height="${sz}">
      <circle cx="${sz/2}" cy="${sz/2}" r="${r}" fill="none" stroke="#E2E8F0" stroke-width="4.5"/>
      ${c>0?`<circle cx="${sz/2}" cy="${sz/2}" r="${r}" fill="none" stroke="${col}" stroke-width="4.5"
        stroke-linecap="round" stroke-dasharray="${circ}" stroke-dashoffset="${dashOffset}"
        transform="rotate(-90 ${sz/2} ${sz/2})"/>`:''}
      <text x="${sz/2}" y="${sz/2+fs*0.38}" text-anchor="middle" font-size="${fs}" font-weight="800"
        fill="${pct!==null?col:'#94A3B8'}" font-family="inherit">${txt}</text>
    </svg>`;
  }

  // ── Cumplimiento por año ──────────────────────────────
  const yearCells = ind.yearly.map(y => {
    const isCurr = y.year === CURRENT_YEAR;
    return `<div class="m-year-cell ${isCurr?'current':''}">
      <div class="y-label">${y.year}${isCurr?' · actual':''}</div>
      <div class="y-donut">${miniDonut(y.cumplimiento, 52)}</div>
      <div class="y-meta">Meta: ${fmtSmart(y.meta_raw,ind.tipo_meta,ind.unidad,ind.decimales)}</div>
    </div>`;
  }).join('');

  // ── Estadísticas del período con datos ───────────────
  const validResults = ind.monthly_results.filter(r => r!==null);
  let maxV=null,minV=null,maxLbl='',minLbl='';
  let sumV=0, nV=0;
  ind.monthly_results.forEach((r,i) => {
    if (r===null) return;
    const yi=Math.floor(i/12), mi=i%12;
    const lbl = MONTH_ABBR[mi]+' '+(2026+yi);
    if (maxV===null||r>maxV){maxV=r;maxLbl=lbl;}
    if (minV===null||r<minV){minV=r;minLbl=lbl;}
    sumV+=r; nV++;
  });
  const avgV = nV>0 ? sumV/nV : null;
  const currYearObj = ind.yearly.find(y=>y.year===CURRENT_YEAR);
  const currMeta = currYearObj ? currYearObj.meta_raw : null;
  const vsMeta = (avgV!==null && typeof currMeta==='number') ? avgV - currMeta : null;

  const statsHTML = nV>0 ? `
    <div class="m-chart-stats">
      <div class="m-stat-pill"><div class="sp-label">↑ Máximo</div><div class="sp-val">${fmt(maxV,ind.tipo_meta,ind.unidad)}</div><div class="sp-sub">${maxLbl}</div></div>
      <div class="m-stat-pill"><div class="sp-label">↓ Mínimo</div><div class="sp-val">${fmt(minV,ind.tipo_meta,ind.unidad)}</div><div class="sp-sub">${minLbl}</div></div>
      <div class="m-stat-pill neutral"><div class="sp-label">⌀ Promedio ${CURRENT_YEAR}</div><div class="sp-val">${fmt(avgV,ind.tipo_meta,ind.unidad)}</div><div class="sp-sub">${nV} meses</div></div>
      ${vsMeta!==null?`<div class="m-stat-pill ${vsMeta>0?'negative':'positive'}"><div class="sp-label">vs Meta anual</div><div class="sp-val">${vsMeta>0?'+':''}${fmt(vsMeta,ind.tipo_meta,ind.unidad)}</div><div class="sp-sub">desviación</div></div>`:''}
    </div>` : '';

  // ── Sábana mensual con colores semáforo ──────────────
  let tablaRows = '';
  for(let yi=0;yi<5;yi++){
    const yr=2026+yi; const isCurr=yr===CURRENT_YEAR;
    const meta=ind.metas[yi];
    let row=`<tr class="${isCurr?'current-year':''}"><td>${yr}</td>`;
    for(let m=0;m<12;m++){
      const v=ind.monthly_results[yi*12+m];
      if(v===null){ row+=`<td><span style="color:var(--text-lt)">—</span></td>`; continue; }
      const mc = monthCompliance(ind, yi, m);
      row+=`<td class="cell-has-data cell-${mc.status}">${fmt(v,ind.tipo_meta,ind.unidad)}</td>`;
    }
    row+='</tr>'; tablaRows+=row;
  }

  const warningsHTML = ind.warnings&&ind.warnings.length
    ? `<div class="warnings-list"><strong>⚠ Avisos:</strong><ul>${ind.warnings.map(w=>`<li>${escapeHTML(w)}</li>`).join('')}</ul></div>` : '';

  document.getElementById('modal-content').innerHTML = `
    <div class="modal-head" style="border-bottom:3px solid ${color}20">
      <div class="modal-head-bg" style="background:linear-gradient(135deg,${color},transparent)"></div>
      <div style="flex:1;min-width:0;position:relative">
        <div class="m-code">${ind.code} · ${persp.name}</div>
        <div class="m-name">${escapeHTML(ind.name)}</div>
        <div class="m-tags">
          <span class="m-tag ${tendClass}">${tendIcon} Tendencia: ${ind.tend}</span>
          <span class="m-tag">${ind.ini_icon} ${escapeHTML(ind.ini)}</span>
          <span class="m-tag">📅 ${ind.frec}</span>
          <span class="m-tag">👤 ${escapeHTML(ind.responsable||'—')}</span>
        </div>
      </div>
      <div style="display:flex;gap:8px;align-items:center;flex-shrink:0">
        <button id="modal-export-btn" onclick="exportIndicadorXlsx()" style="display:inline-flex;align-items:center;gap:6px;background:var(--primary);border:none;color:white;border-radius:9px;padding:7px 14px;font-size:0.78rem;font-weight:700;cursor:pointer;transition:background .15s;" onmouseover="this.style.background='#075A8A'" onmouseout="this.style.background='var(--primary)'">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          Excel
        </button>
        <button class="modal-close" id="modal-close">×</button>
      </div>
    </div>

    <div class="modal-body">
      ${warningsHTML}
      <div class="modal-top-grid">
        <div class="m-info-card">
          <div>
            <div class="m-section-label">Fórmula</div>
            <div class="m-formula">${escapeHTML(ind.formula)}</div>
          </div>
          <div>
            <div class="m-section-label">Descripción</div>
            <div class="m-desc">${escapeHTML(ind.desc)}</div>
          </div>
          <div class="m-attrs">
            <div class="m-attr"><div class="lbl">Unidad</div><div class="val">${escapeHTML(ind.unidad)}</div></div>
            <div class="m-attr"><div class="lbl">Frecuencia</div><div class="val">${ind.frec}</div></div>
          </div>
        </div>
        <div class="m-year-panel">
          <div class="m-section-label">Cumplimiento por año</div>
          <div class="m-year-grid">${yearCells}</div>
        </div>
      </div>

      <div class="m-chart-section">
        <div class="m-section-label" style="margin-bottom:10px">Resultados mensuales 2026–2030</div>
        ${statsHTML}
        <!-- Controles de período -->
        <div id="chart-controls" style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:8px;">
          <div style="font-size:0.76rem;color:var(--text-med);font-weight:500;" id="chart-period-label">Ene 2026 – Dic 2026</div>
          <div style="display:flex;gap:4px;">
            ${['6M','12M','18M','24M'].map(w=>`<button class="chart-win-btn" data-win="${parseInt(w)}" style="padding:3px 10px;font-size:0.72rem;font-weight:600;border-radius:6px;border:1.5px solid #E2E8F0;background:white;color:var(--text-med);cursor:pointer;transition:all .15s;">${w}</button>`).join('')}
            <button class="chart-win-btn" data-win="60" style="padding:3px 10px;font-size:0.72rem;font-weight:600;border-radius:6px;border:1.5px solid #E2E8F0;background:white;color:var(--text-med);cursor:pointer;transition:all .15s;">Todo</button>
          </div>
        </div>
        <div class="detail-chart-wrap" style="height:320px;"><canvas id="detail-chart"></canvas></div>
        <!-- Slider de navegación -->
        <div id="chart-slider-wrap" style="margin-top:12px;padding:0 4px;">
          <div style="display:flex;align-items:center;gap:10px;">
            <button id="chart-prev" style="width:28px;height:28px;border-radius:50%;border:1.5px solid #E2E8F0;background:white;cursor:pointer;font-size:14px;display:flex;align-items:center;justify-content:center;color:var(--text-med);transition:all .15s;">‹</button>
            <div style="flex:1;position:relative;">
              <input type="range" id="chart-slider" min="0" max="48" value="0" step="1"
                style="width:100%;accent-color:var(--primary);cursor:pointer;height:4px;">
            </div>
            <button id="chart-next" style="width:28px;height:28px;border-radius:50%;border:1.5px solid #E2E8F0;background:white;cursor:pointer;font-size:14px;display:flex;align-items:center;justify-content:center;color:var(--text-med);transition:all .15s;">›</button>
          </div>
          <div style="text-align:center;margin-top:4px;font-size:0.70rem;color:var(--text-lt);" id="chart-slider-info"></div>
        </div>
      </div>

      <div>
        <div class="m-section-label" style="margin-bottom:10px">Sábana mensual</div>
        <div class="tabla-wrap">
          <table class="tabla">
            <thead><tr><th>Año</th>${MONTH_ABBR.map(m=>`<th>${m}</th>`).join('')}</tr></thead>
            <tbody>${tablaRows}</tbody>
          </table>
        </div>
      </div>

      <div style="margin-top:20px">
        <div class="m-section-label" style="margin-bottom:10px">Composición del cálculo</div>
        <div id="modal-composicion" style="font-size:0.82rem;color:var(--text-med)">
          <div style="text-align:center;padding:20px;color:var(--text-lt)">Cargando...</div>
        </div>
      </div>
    </div>`;

  MODAL_IND = ind;
  document.getElementById('modal-close').addEventListener('click', closeModal);
  overlay.addEventListener('click', e => { if(e.target===overlay) closeModal(); });

  // Sección composición del cálculo (renderizada después del chart)
  // Chart y composición en bloques independientes para evitar que un fallo en uno bloquee el otro
  setTimeout(() => {
    try { buildDetailChart(ind, color); } catch(e) { console.warn('Chart.js no disponible:', e.message); }
  }, 80);

  setTimeout(() => {
    const compEl = document.getElementById('modal-composicion');
    if (!compEl) return;
    try {
      compEl.innerHTML = buildComposicion(ind);
      requestAnimationFrame(() => attachComposicionBubbles(ind));
    } catch(e) {
      compEl.innerHTML = '<div style="padding:16px;color:var(--text-lt)">No se pudo cargar la composición.</div>';
      console.warn('buildComposicion error:', e);
    }
  }, 100);
}

function buildDetailChart(ind, color) {
  const ctx = document.getElementById('detail-chart');
  if (!ctx) return;
  if (STATE.charts.detail) try { STATE.charts.detail.destroy(); } catch(e) {}

  const allData   = ind.monthly_results.map(r => r);    // 60 valores
  const allMeta   = [];
  const allMetaPer = ind.meta_periodos || [];
  for(let i=0;i<60;i++){
    const yi=Math.floor(i/12);
    // Usar meta_periodo si existe, sino meta anual
    const mp = (allMetaPer[i]!==null && allMetaPer[i]!==undefined) ? allMetaPer[i] : ind.metas[yi];
    allMeta.push(typeof mp==='number' ? mp : null);
  }

  // ── Estado del control ──────────────────────────────
  let winSize  = 12;   // meses visibles
  let winStart = 0;    // índice de inicio

  // Ajustar inicio al último dato disponible (ventana centrada al final)
  const lastDataIdx = allData.reduceRight((acc, v, i) => acc===-1 && v!==null ? i : acc, -1);
  if(lastDataIdx >= 0) winStart = Math.max(0, Math.min(lastDataIdx - winSize + 1, 60 - winSize));

  function computeYRange(start, size) {
    // Siempre incluir AMBOS (datos y meta) en el rango para consistencia visual.
    // Excepción: si meta es >8x mayor que datos (IND-14: meta=5846 vs datos=13),
    //            usar solo datos y dibujar meta como referencia fuera del área principal.
    const seg  = allData.slice(start, start+size).filter(v=>v!==null);
    const segM = allMeta.slice(start, start+size).filter(v=>v!==null);
    if(!seg.length && !segM.length) return [undefined, undefined];

    const dataMn = seg.length  ? Math.min(...seg)  : null;
    const dataMx = seg.length  ? Math.max(...seg)  : null;
    const metaMn = segM.length ? Math.min(...segM) : null;
    const metaMx = segM.length ? Math.max(...segM) : null;

    let mn, mx;
    if(dataMx !== null && metaMx !== null) {
      const dataSpan  = Math.max(Math.abs(dataMx), Math.abs(dataMn||0));
      const metaSpan  = Math.max(Math.abs(metaMx),  Math.abs(metaMn||0));
      // base_social: meta acumulada crece continuamente → siempre eje independiente
      const ratio = dataSpan < 1e-6 ? 1 : metaSpan / dataSpan;
      if(ratio > 8 && ind.tipo !== 'base_social') {
        // Meta enormemente mayor (no base_social): solo datos en eje
        mn = dataMn; mx = dataMx;
      } else {
        // Incluir ambos en un solo eje — muestra cumplimiento claramente
        mn = Math.min(dataMn, metaMn);
        mx = Math.max(dataMx, metaMx);
      }
    } else if(dataMx !== null) {
      mn = dataMn; mx = dataMx;
    } else {
      mn = metaMn; mx = metaMx;
    }
    // Margen superior/inferior para que los puntos no queden pegados al borde
    const span = (mx - mn) || Math.abs(mx) * 0.5 || 1;
    return [mn - span*0.18, mx + span*0.22];
  }

  function updatePeriodLabel(start, size) {
    const end = Math.min(start+size-1, 59);
    const lbl = document.getElementById('chart-period-label');
    if(lbl) lbl.textContent = DATA.months_labels[start] + ' – ' + DATA.months_labels[end];
    const info = document.getElementById('chart-slider-info');
    const filled = allData.slice(start, start+size).filter(v=>v!==null).length;
    if(info) info.textContent = filled + ' / ' + size + ' meses con datos';
  }

  function applyWindow(start, size) {
    start = Math.max(0, Math.min(start, 60-size));
    winStart = start; winSize = size;
    const chart = STATE.charts.detail;
    if(!chart) return;
    const labels  = DATA.months_labels.slice(start, start+size);
    const dataSlice = allData.slice(start, start+size);
    const metaSlice = allMeta.slice(start, start+size);
    chart.data.labels = labels;
    chart.data.datasets[0].data = dataSlice;
    chart.data.datasets[1].data = metaSlice;
    const [mn,mx] = computeYRange(start, size);
    chart.options.scales.y.min = mn;
    chart.options.scales.y.max = mx;


    // Actualizar meta label en leyenda con valor del período
    // Para base_social la meta es acumulada (cambia cada mes), no mostrar valor fijo
    if(ind.tipo !== 'base_social') {
      const metaVal = metaSlice.find(v=>v!==null);
      chart.data.datasets[1].label = 'Meta anual'+(metaVal!==null&&metaVal!==undefined ? ' ('+fmtSmart(metaVal,ind.tipo_meta,ind.unidad,ind.decimales)+')' : '');
    }
    chart.update('active');
    // Slider
    const sl = document.getElementById('chart-slider');
    if(sl){ sl.max = 60-size; sl.value = start; }
    updatePeriodLabel(start, size);
    // Botones activos
    document.querySelectorAll('.chart-win-btn').forEach(b=>{
      const w = parseInt(b.dataset.win);
      b.style.background = (w===size||(w===60&&size===60)) ? 'var(--primary)' : 'white';
      b.style.color = (w===size||(w===60&&size===60)) ? 'white' : 'var(--text-med)';
      b.style.borderColor = (w===size||(w===60&&size===60)) ? 'var(--primary)' : '#E2E8F0';
    });
  }

  // Plugin: etiquetas en TODOS los puntos con datos
  const dataLabelsPlugin = {
    id: 'smartLabels',
    afterDatasetsDraw(chart) {
      const ds = chart.data.datasets[0].data;
      drawSmartLabels(chart, ds, color, (v,tm,u)=>fmt(v,tm,u,ind.decimales), ind.tipo_meta, ind.unidad,
        { fontSize:10, lblH:17, maxLabels: winSize <= 12 ? 99 : 8, minGap: winSize <= 12 ? 0 : 44 });
    }
  };

  const [yMin0, yMax0] = computeYRange(winStart, winSize);

  STATE.charts.detail = new Chart(ctx, {
    type: 'line',
    data: {
      labels: DATA.months_labels.slice(winStart, winStart+winSize),
      datasets: [
        {
          label: 'Resultado mensual',
          data: allData.slice(winStart, winStart+winSize),
          borderColor: color,
          backgroundColor: (ctx2) => {
            const g = ctx2.chart.ctx.createLinearGradient(0, 0, 0, ctx2.chart.height);
            g.addColorStop(0, hexA(color, 0.28));
            g.addColorStop(0.55, hexA(color, 0.10));
            g.addColorStop(1, hexA(color, 0.00));
            return g;
          },
          tension: 0.35, fill: true,
          pointRadius: (ctx2) => allData.slice(winStart,winStart+winSize)[ctx2.dataIndex]!==null ? 5 : 0,
          pointHoverRadius: 7,
          pointBackgroundColor: 'white',
          pointBorderColor: color,
          pointBorderWidth: 2.5,
          spanGaps: false,
          borderWidth: 2.5,
          yAxisID: 'y',
          order: 1,
        },
        {
          label: ind.tipo==='base_social' ? 'Meta acumulada' : 'Meta anual',
          data: allMeta.slice(winStart, winStart+winSize),
          borderColor: '#94A3B8',
          backgroundColor: 'transparent',
          borderDash: [6, 4],
          borderWidth: 1.6,
          pointRadius: (ctx2) => {
            // Para base_social mostrar puntos pequeños solo donde hay dato
            if(ind.tipo!=='base_social') return 0;
            return allMeta.slice(winStart,winStart+winSize)[ctx2.dataIndex]!==null ? 3 : 0;
          },
          pointBorderColor: '#94A3B8',
          pointBackgroundColor: 'white',
          fill: false,
          spanGaps: true,
          yAxisID: 'y',
          order: 2,
        }
      ]
    },
    plugins: [dataLabelsPlugin],
    options: {
      responsive: true, maintainAspectRatio: false,
      animation: { duration: 320, easing: 'easeInOutQuart' },
      layout: { padding: { top: 26, right: 18, bottom: 4, left: 6 } },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: 'top', align: 'end',
          labels: {
            font: { size: 11, weight: '600' }, padding: 16,
            usePointStyle: true, pointStyleWidth: 18,
          }
        },
        tooltip: {
          backgroundColor: 'rgba(15,23,42,0.92)',
          titleColor: '#94A3B8', bodyColor: 'white',
          padding: 12, cornerRadius: 10,
          callbacks: {
            label: (c) => {
              if(c.raw===null) return null;
              return ' '+c.dataset.label+': '+fmt(c.raw, ind.tipo_meta, ind.unidad, ind.decimales);
            }
          }
        }
      },
      scales: {
        x: {
          ticks: {
            autoSkip: false,
            maxRotation: 45, minRotation: 0,
            font: { size: 10 }, color: '#64748B',
            callback: (val, i) => DATA.months_labels.slice(winStart, winStart+winSize)[i] || ''
          },
          grid: { color: '#F8FAFC', drawBorder: false }
        },
        y: {
          min: yMin0, max: yMax0,
          ticks: {
            callback: (v) => fmt(v, ind.tipo_meta, ind.unidad, ind.decimales),
            font: { size: 10 }, color: '#94A3B8', maxTicksLimit: 6
          },
          grid: { color: '#F1F5F9', drawBorder: false }
        },

      }
    }
  });

  // ── Conectar controles ─────────────────────────────
  updatePeriodLabel(winStart, winSize);

  // Botones 6M / 12M / 18M / 24M / Todo
  document.querySelectorAll('.chart-win-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const w = parseInt(btn.dataset.win);
      const newSize = Math.min(w, 60);
      // Mantener el extremo derecho si es posible
      const newStart = Math.max(0, Math.min(winStart + winSize - newSize, 60 - newSize));
      applyWindow(newStart, newSize);
    });
  });

  // Slider
  const slider = document.getElementById('chart-slider');
  if(slider) {
    slider.max = 60 - winSize;
    slider.value = winStart;
    slider.addEventListener('input', () => applyWindow(parseInt(slider.value), winSize));
  }

  // Flechas prev/next (desplazan 1 mes)
  const btnPrev = document.getElementById('chart-prev');
  const btnNext = document.getElementById('chart-next');
  if(btnPrev) btnPrev.addEventListener('click', () => applyWindow(winStart-1, winSize));
  if(btnNext) btnNext.addEventListener('click', () => applyWindow(winStart+1, winSize));

  // Activar botón 12M por defecto
  applyWindow(winStart, winSize);
}

// Scroll a la tarjeta del indicador en la sección de perspectivas y hace highlight
function scrollToIndicator(code, monthLabel) {
  // Asegurar que la perspectiva del indicador sea visible
  const ind = findIndicator(code);
  if (!ind) return;
  const persp = DATA.perspectives.find(p => p.indicators.includes(ind));
  if (!persp) return;

  // Si hay filtro de perspectiva activo que no incluye este indicador, resetear
  if (STATE.perspFilter !== 'all' && STATE.perspFilter !== persp.key) {
    STATE.perspFilter = 'all';
    renderFilters(); renderGauges(); renderMatriz(); renderPerspectives();
  }

  // Esperar render y luego hacer scroll
  setTimeout(() => {
    const card = document.querySelector(`.ind-card[data-code="${code}"]`);
    if (!card) return;

    // Calcular offset de la barra compacta (solo desktop)
    const barEl = document.getElementById('cbar');
    const isDesktop = window.innerWidth >= 769;
    const stickyH = isDesktop && barEl ? barEl.offsetHeight : 0;

    const cardTop = card.getBoundingClientRect().top + window.scrollY - stickyH - 16;
    window.scrollTo({ top: cardTop, behavior: 'smooth' });

    // Highlight: iluminación amarilla tipo "spotlight"
    card.classList.remove('ind-card-highlight');
    void card.offsetWidth; // reflow
    card.classList.add('ind-card-highlight');
    // Atenuar cards vecinas temporalmente
    document.querySelectorAll('.ind-card').forEach(c => {
      if(c !== card) { c.style.transition = 'opacity .3s'; c.style.opacity = '0.45'; }
    });
    setTimeout(() => {
      card.classList.remove('ind-card-highlight');
      document.querySelectorAll('.ind-card').forEach(c => { c.style.opacity = ''; });
    }, 2200);

    // Si se pasó un mes específico, resaltar esa columna en el mini-chart
    if (monthLabel) {
      const monthIdx = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
        .findIndex(m => monthLabel.startsWith(m));
      if (monthIdx >= 0) {
        const yr = parseInt(monthLabel.split(' ')[1]) || CURRENT_YEAR;
        const flatIdx = (yr - 2026) * 12 + monthIdx;
        highlightChartPoint(code, flatIdx);
      }
    }
  }, STATE.perspFilter !== 'all' ? 120 : 20);
}

// Resalta un punto en el sparkline de la card
function highlightChartPoint(code, flatIdx) {
  const chart = STATE.charts[code];
  if (!chart) return;
  const meta = {
    datasetIndex: 0,
    index: flatIdx,
  };
  chart.tooltip.setActiveElements([meta], {x:0, y:0});
  chart.setDatasetVisibility(0, true);
  chart.update('none');
  setTimeout(() => { try { chart.tooltip.setActiveElements([], {}); chart.update('none'); } catch(e){} }, 2500);
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('active');
  if(STATE.charts.detail){try{STATE.charts.detail.destroy();}catch(e){}delete STATE.charts.detail;}
}

// ── Ventana de escalas de cumplimiento ────────────────────
function openEscalas() {
  const tCumple = DATA.thresholds.cumple, tRiesgo = DATA.thresholds.riesgo;
  const pct = v => Math.round(v * 100) + '%';
  const groups = [
    { st:'cumple',    title:'En meta',   range:`≥ ${pct(tCumple)}`, desc:'Se alcanzó o superó la meta.' },
    { st:'riesgo',    title:'En riesgo', range:`${pct(tRiesgo)} – ${Math.round(tCumple*100)-1}%`, desc:'Por debajo de la meta; requiere seguimiento.' },
    { st:'critico',   title:'Crítico',   range:`< ${pct(tRiesgo)}`, desc:'Muy lejos de la meta; requiere acción.' },
    { st:'pendiente', title:'Sin datos', range:'—', desc:'Aún no hay resultado registrado.' },
  ];

  // Niveles del semáforo leídos de la hoja Parámetros (mayor → menor límite)
  const levels = getSemaforoColors();
  const pend = DATA.thresholds.pendiente;
  const levelRow = (lv, range, st) => `<tr>
      <td><span class="esc-level" style="background:${lv.bg};color:${lv.fg}">${lv.emoji ? lv.emoji + ' ' : ''}${escapeHTML(lv.name)}</span></td>
      <td class="esc-range">${range}</td>
      <td class="esc-ldesc">${escapeHTML(lv.desc || '')}</td>
      <td><span class="esc-pill ${st}"><span class="nm-tick ${st}">${NM_TICK[st]}</span>${groups.find(g=>g.st===st).title}</span></td>
    </tr>`;
  const levelRows = levels.map((lv, i) => {
    const upper = i === 0 ? null : levels[i - 1].lim;
    const range = upper === null ? `≥ ${pct(lv.lim)}`
      : lv.lim < 0 ? `< ${pct(upper)}` : `${pct(lv.lim)} – ${Math.round(upper*100)-1}%`;
    return levelRow(lv, range, nmStatus(lv.lim < 0 ? 0 : lv.lim));
  }).join('') + (pend ? levelRow(pend, '—', 'pendiente') : '');

  const content = document.getElementById('modal-content');
  content.innerHTML = `
    <div class="modal-head">
      <div style="flex:1;min-width:0">
        <div class="m-code">Guía del tablero</div>
        <div class="m-name">Escalas de cumplimiento de la meta</div>
        <div style="font-size:0.84rem;color:var(--text-med)">El color de cada mes y de cada indicador depende del % de cumplimiento frente a su meta.</div>
      </div>
      <button class="modal-close" id="escalas-close" aria-label="Cerrar">×</button>
    </div>
    <div class="modal-body">
      <div class="esc-groups">
        ${groups.map(g => `<div class="esc-card">
          <span class="nm-tick ${g.st}" style="width:34px;height:34px">${NM_TICK[g.st]}</span>
          <div><div class="esc-title">${g.title}</div><div class="esc-big ${g.st}">${g.range}</div><div class="esc-desc">${g.desc}</div></div>
        </div>`).join('')}
      </div>

      <div class="m-section-label" style="margin-top:22px">Cómo se calcula el cumplimiento</div>
      <div class="esc-formulas">
        <div class="esc-f"><b>↑ Crecer</b><span>Resultado ÷ Meta</span><small>Meta 15%, resultado 12% → 80% · En riesgo</small></div>
        <div class="esc-f"><b>↓ Reducir</b><span>1 − (Resultado − Meta) ÷ Meta</span><small>Meta 5%, resultado 6% → 80% · En riesgo</small></div>
        <div class="esc-f"><b>↔ Mantener</b><span>1 − |Resultado − Meta|</span><small>Se penaliza alejarse de la meta hacia arriba o hacia abajo</small></div>
      </div>

      ${levelRows ? `<div class="m-section-label" style="margin-top:22px">Niveles del semáforo</div>
      <table class="esc-table">
        <thead><tr><th>Nivel</th><th>Cumplimiento</th><th>Descripción</th><th>Estado en el tablero</th></tr></thead>
        <tbody>${levelRows}</tbody>
      </table>` : ''}
      <div class="esc-note">Los cortes se leen de la hoja <b>Parámetros</b> del Excel. El % anual es el promedio de los meses del año (tope 150%).</div>
    </div>`;

  const overlay = document.getElementById('modal-overlay');
  overlay.classList.add('active');
  overlay.onclick = e => { if (e.target === overlay) closeModal(); };
  document.getElementById('escalas-close').addEventListener('click', closeModal);
}

// ── Helpers ───────────────────────────────────────────────
function findIndicator(code) {
  for(const p of DATA.perspectives) for(const ind of p.indicators) if(ind.code===code) return ind;
  return null;
}
function escapeHTML(s) {
  if(s===null||s===undefined) return '';
  return String(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);
}
function hexA(hex, alpha) {
  const h=hex.replace('#','');
  const r=parseInt(h.substring(0,2),16),g=parseInt(h.substring(2,4),16),b=parseInt(h.substring(4,6),16);
  return `rgba(${r},${g},${b},${alpha})`;
}
function shade(hex, percent) {
  const h=hex.replace('#','');
  let r=parseInt(h.substring(0,2),16),g=parseInt(h.substring(2,4),16),b=parseInt(h.substring(4,6),16);
  const f=percent/100;
  r=Math.max(0,Math.min(255,Math.round(r+255*f)));
  g=Math.max(0,Math.min(255,Math.round(g+255*f)));
  b=Math.max(0,Math.min(255,Math.round(b+255*f)));
  return '#'+[r,g,b].map(v=>v.toString(16).padStart(2,'0')).join('');
}

// ── Init ──────────────────────────────────────────────────

// ════════════════════════════════════════════════════════════
// Composición del cálculo + Exportación Excel
// ════════════════════════════════════════════════════════════

// Indicador actualmente abierto en el modal
let MODAL_IND = null;

// ── Filas futuras (sin datos) para la tabla de composición ──────────────
function buildFutureRows(ind, existingRows, hideVars, showMetaPer) {
  const MONTH_NAMES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  const existingKeys = new Set(existingRows.map(r => `${r.year}-${r.month}`));
  let rows = '';
  let rowIdx = existingRows.length;
  for (let yr = 2026; yr <= 2030; yr++) {
    for (let mo = 1; mo <= 12; mo++) {
      if (existingKeys.has(`${yr}-${mo}`)) continue;
      const yi = yr - 2026;
      const mi = mo - 1;
      const flatIdx = yi * 12 + mi;
      const metaAnual = ind.metas ? ind.metas[yi] : null;
      const metaPer = (ind.meta_periodos && ind.meta_periodos[flatIdx] !== null &&
                       ind.meta_periodos[flatIdx] !== undefined)
                      ? ind.meta_periodos[flatIdx] : metaAnual;
      const isEven = rowIdx % 2 === 0;
      const varCells = hideVars ? '' : (ind.vars||[]).map(() =>
        `<td style="text-align:center;color:var(--text-lt)">—</td>`).join('');
      const metaPerCell = showMetaPer
        ? `<td style="text-align:center;background:#EFF6FF;color:#1D4ED8;font-weight:600">
            ${typeof metaPer==='number' ? fmtSmart(metaPer,ind.tipo_meta,ind.unidad,ind.decimales) : '—'}
           </td>` : '';
      rows += `<tr style="${isEven?'':'background:var(--surface-2)'}opacity:0.7">
        <td style="font-weight:600;color:var(--text-med)">${yr}</td>
        <td style="color:var(--text-med)">${MONTH_NAMES[mi]}</td>
        ${varCells}
        <td style="text-align:center;background:var(--surface-3);color:var(--text-lt)">—</td>
        ${metaPerCell}
        <td style="text-align:center;background:var(--surface-3);color:var(--text-med)">
          ${typeof metaAnual==='number' ? fmtSmart(metaAnual,ind.tipo_meta,ind.unidad,ind.decimales) : '—'}
        </td>
        <td style="text-align:center;background:var(--surface-3);color:var(--text-lt)">—</td>
      </tr>`;
      rowIdx++;
    }
  }
  return rows;
}

function buildExpandBtn(ind, existingRows) {
  const existingKeys = new Set(existingRows.map(r => `${r.year}-${r.month}`));
  let futureCount = 0;
  for (let yr = 2026; yr <= 2030; yr++)
    for (let mo = 1; mo <= 12; mo++)
      if (!existingKeys.has(`${yr}-${mo}`)) futureCount++;
  if (futureCount === 0) return '';
  return `
    <div style="margin-top:10px;text-align:center">
      <button id="btn-expand-future"
        onclick="(function(){
          const tb = document.getElementById('comp-body-future');
          const btn = document.getElementById('btn-expand-future');
          const open = tb.style.display !== 'none';
          tb.style.display = open ? 'none' : '';
          btn.innerHTML = open
            ? '<svg width=\\'13\\' height=\\'13\\' viewBox=\\'0 0 24 24\\' fill=\\'none\\' stroke=\\'currentColor\\' stroke-width=\\'2.2\\' stroke-linecap=\\'round\\' stroke-linejoin=\\'round\\'><polyline points=\\'6 9 12 15 18 9\\'/></svg> Ver meses pendientes hasta 2030 (${futureCount})'
            : '<svg width=\\'13\\' height=\\'13\\' viewBox=\\'0 0 24 24\\' fill=\\'none\\' stroke=\\'currentColor\\' stroke-width=\\'2.2\\' stroke-linecap=\\'round\\' stroke-linejoin=\\'round\\'><polyline points=\\'18 15 12 9 6 15\\'/></svg> Ocultar meses pendientes';
        })()"
        style="display:inline-flex;align-items:center;gap:6px;
               background:var(--surface-2);border:1px solid var(--border);
               color:var(--text-med);border-radius:8px;padding:7px 16px;
               font-size:0.74rem;font-weight:600;cursor:pointer;transition:background .15s;">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        Ver meses pendientes hasta 2030 (${futureCount})
      </button>
    </div>`;
}

function buildComposicion(ind) {
  if (!ind.vars || !ind.vars.length) return '';
  const MONTH_NAMES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

  const hideInputVars = ind.hide_input_vars === true;

  // Filas con datos (monthly_inputs)
  const rows = (ind.monthly_inputs || []).filter(rec => {
    if (hideInputVars) {
      // Para indicadores sin vars visibles, mostrar filas con resultado calculado
      const flatIdx = (rec.year - 2026) * 12 + (rec.month - 1);
      return ind.monthly_results[flatIdx] !== null && ind.monthly_results[flatIdx] !== undefined;
    }
    return ind.vars.some(v => rec[v.code] !== null && rec[v.code] !== undefined);
  });

  if (!rows.length) {
    return `<div style="padding:16px;background:var(--surface-2);border-radius:10px;font-size:0.84rem;color:var(--text-lt);text-align:center">
      Sin variables de entrada registradas aún.
    </div>`;
  }

  // Tabla de variables de entrada
  // Mostrar columna Meta período si ALGÚN valor difiere de la meta anual
  const showMetaPer = ind.meta_periodos && ind.meta_periodos.some((mp, i) => {
    if (mp === null || mp === undefined) return false;
    const yi = Math.floor(i / 12);
    const ma = ind.metas[yi];
    return ma !== null && ma !== undefined && Math.abs(mp - ma) > 1e-9;
  });

  // Adaptar labels de variables: agregar "(M $)" si son en pesos
  const varLabels = ind.vars.map(v => {
    if (!ind.vars_en_pesos) return escapeHTML(v.label);
    // Reemplazar ($) por ($ M) o añadir unidad si no tiene
    const lbl = v.label.replace(/\(\$\)/g, '($ M)').replace(/\(M \$\)/g, '($ M)');
    return lbl.includes('($') ? escapeHTML(lbl) : escapeHTML(v.label + (v.label.includes('$') ? '' : ' ($ M)'));
  });

  const hideVars = ind.hide_input_vars === true;
  const thead = `<tr>
    <th style="text-align:left">Año</th>
    <th style="text-align:left">Mes</th>
    ${hideVars ? '' : varLabels.map(lbl=>`<th>${lbl}</th>`).join('')}
    <th style="background:var(--surface-3)">Resultado</th>
    ${showMetaPer ? `<th style="background:#EFF6FF;color:#1D4ED8;font-weight:700">Meta período</th>` : ''}
    <th style="background:var(--surface-3)">Meta</th>
    <th style="background:var(--surface-3)">Cumpl.</th>
  </tr>`;

  const tbody = rows.map((rec, i) => {
    const yi = rec.year - 2026;
    const mi = rec.month - 1;
    const flatIdx  = yi * 12 + mi;
    const result   = ind.monthly_results[flatIdx];
    const metaAnual = ind.metas[yi];
    // Meta del período: si existe y es distinta, usarla para el cumplimiento
    const metaPer  = (ind.meta_periodos && ind.meta_periodos[flatIdx] !== null &&
                      ind.meta_periodos[flatIdx] !== undefined)
                     ? ind.meta_periodos[flatIdx] : metaAnual;
    const meta = metaPer;  // usar meta_periodo para el cálculo
    const cft2  = ind.cumpl_formula_type || '';
    const tend  = (ind.tend||'').toLowerCase();
    let cumpl = null;
    if (result !== null && typeof meta === 'number') {
      if (cft2 === 'directo') {
        cumpl = result;  // El resultado ya ES el cumplimiento
      } else if (cft2 === 'mantener' || (!cft2 && tend.includes('mantener'))) {
        cumpl = meta!==0 ? 1 - Math.abs(result-meta)/Math.max(Math.abs(meta),1) : null;
      } else if (cft2 === 'reducir' || (!cft2 && (tend.includes('reducir')||tend.includes('disminuir')))) {
        cumpl = meta!==0 ? 1.0-(result-meta)/meta : null;
      } else if (!cft2 && meta < 0) {
        cumpl = meta!==0 ? 1.0-(result-meta)/meta : null;
      } else {
        cumpl = meta!==0 ? result/meta : null;
      }
    }
    const st = cumpl===null?'pendiente':cumpl>=(DATA.thresholds?.cumple||0.95)?'cumple':cumpl>=(DATA.thresholds?.riesgo||0.80)?'riesgo':'critico';
    const _stE    = getSemaforoEntry(cumpl);
    const stColor = cumpl!==null ? _stE.fg : '#94A3B8';
    const stBg    = cumpl!==null ? _stE.bg : '#F1F5F9';
    // stColor ya declarado arriba
    const isEven = i%2===0;
    return `<tr style="${isEven?'':'background:var(--surface-2)'}">
      <td style="font-weight:600">${rec.year}</td>
      <td>${MONTH_NAMES[mi]}</td>
      ${hideVars ? '' : ind.vars.map(v=>{
        const val = rec[v.code];
        return `<td style="text-align:center">${val!==null&&val!==undefined?fmtVar(val,v.label,ind.decimales):'—'}</td>`;
      }).join('')}
      <td style="text-align:center;background:var(--surface-3);font-weight:700">
        ${result!==null?fmt(result,ind.tipo_meta,ind.unidad,ind.decimales):'—'}
      </td>
      ${showMetaPer ? `<td style="text-align:center;background:#EFF6FF;color:#1D4ED8;font-weight:600">
        ${typeof metaPer==='number'?fmtSmart(metaPer,ind.tipo_meta,ind.unidad,ind.decimales):'—'}
      </td>` : ''}
      <td style="text-align:center;background:var(--surface-3)">
        ${typeof metaAnual==='number'?fmtSmart(metaAnual,ind.tipo_meta,ind.unidad,ind.decimales):'—'}
      </td>
      <td style="text-align:center;background:var(--surface-3)">
        <span class="cumpl-clickable" style="font-weight:800;color:${stColor};cursor:pointer;border-bottom:2px dashed ${stColor};padding-bottom:1px;"
              data-row-id="cr-${i}-${mi}"
              data-result="${result}"
              data-meta="${meta}"
              data-cumpl="${cumpl}"
              data-month="${MONTH_NAMES[mi]} ${rec.year}">
          ${cumpl!==null?numCO(cumpl*100,0)+'%':'—'}
        </span>
      </td>
    </tr>
    <tr id="cr-${i}-${mi}-detail" style="display:none">
      <td colspan="${2 + (hideVars?0:ind.vars.length) + 4 + (showMetaPer?1:0)}"
          style="padding:0;background:#F8FAFC;border-bottom:1px solid var(--border);">
        <div class="calc-inline-detail" id="cr-${i}-${mi}-box"></div>
      </td>
    </tr>
  `;
  }).join('');

  // Descripción de variables (ocultar si hide_input_vars)
  const varDesc = hideInputVars ? '' : `<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px">
    ${ind.vars.map((v,i)=>{
      const colors=['#EFF6FF:#2563EB','#F0FDF4:#16A34A','#FFF7ED:#D97706','#FDF4FF:#7C3AED','#FFF5F5:#DC2626'];
      const [bg,clr]=colors[i%colors.length].split(':');
      return `<div style="background:${bg};border:1px solid ${clr}30;border-radius:8px;padding:6px 12px;font-size:0.76rem">
        <span style="font-weight:800;color:${clr}">${v.code.toUpperCase()}</span>
        <span style="color:var(--text-med);margin-left:6px">${escapeHTML(v.label)}</span>
      </div>`;
    }).join('')}
  </div>`;

  const html = `
    ${varDesc}
    <div style="overflow-x:auto;border-radius:10px;border:1px solid var(--border)">
      <table class="tabla" style="font-size:0.76rem">
        <thead>${thead}</thead>
        <tbody id="comp-body-data">${tbody}</tbody>
        <tbody id="comp-body-future" style="display:none">${buildFutureRows(ind, rows, hideVars, showMetaPer)}</tbody>
      </table>
    </div>
    ${buildExpandBtn(ind, rows)}
    <div style="margin-top:8px;font-size:0.70rem;color:var(--text-lt);display:flex;align-items:center;gap:5px">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/></svg>
      Pulsa sobre cualquier valor de cumplimiento para ver el cálculo detallado paso a paso.
    </div>`;

  // Necesitamos retornar el html y luego attachar los eventos
  // Lo hacemos con un container y un timeout
  return html;
}

// Attachar eventos de burbuja de cálculo en la sección de composición
function attachComposicionBubbles(ind) {
  const container = document.getElementById('modal-composicion');
  if (!container) return;

  container.querySelectorAll('.cumpl-clickable').forEach(span => {
    span.addEventListener('click', e => {
      e.stopPropagation();
      const rowId  = span.dataset.rowId;
      const result = parseFloat(span.dataset.result);
      const meta   = parseFloat(span.dataset.meta);
      const cumpl  = span.dataset.cumpl === 'null' || !span.dataset.cumpl ? null : parseFloat(span.dataset.cumpl);
      const month  = span.dataset.month;
      const detailRow = document.getElementById(rowId + '-detail');
      const detailBox = document.getElementById(rowId + '-box');

      // Toggle: si ya está abierto, cerrar
      if (detailRow && detailRow.style.display !== 'none') {
        detailRow.style.display = 'none';
        span.style.background = '';
        return;
      }

      // Cerrar cualquier otro abierto
      container.querySelectorAll('tr[id$="-detail"]').forEach(r => r.style.display = 'none');
      container.querySelectorAll('.cumpl-clickable').forEach(s => s.style.background = '');

      if (!detailRow || !detailBox) return;

      // Construir contenido
      const { formulaLine, steps, interp, pct, st } = buildCalcExplanation(ind, result, meta, cumpl, month);
      const stLbl = cumpl!==null ? getSemaforoEntry(cumpl).name : 'Sin datos';

      detailBox.innerHTML = `
        <div class="cid-formula">${escapeHTML(formulaLine)}</div>
        <div class="cid-steps">
          ${steps.map((s,i) => `<div class="cid-step"><div class="cid-step-n">${i+1}</div><div>${s}</div></div>`).join('')}
        </div>
        <div class="cid-result ${st}">
          <span>${stLbl}</span>
          <span class="cid-result-pct">${pct}</span>
        </div>
        <div style="flex-basis:100%"><div class="cid-interp">${escapeHTML(interp)}</div></div>
      `;

      detailRow.style.display = '';
      // Highlight del span activo
      const stBg = cumpl!==null ? getSemaforoEntry(parseFloat(span.dataset.cumpl||0)).bg : '#F1F5F9';
      span.style.background = stBg;
      span.style.borderRadius = '4px';
      span.style.padding = '2px 5px';

      // Scroll suave a la fila de detalle
      detailRow.scrollIntoView({behavior:'smooth', block:'nearest'});
    });
  });
}

// ── Globo explicativo de cálculo ─────────────────────────
// Usa cumpl_formula_type (leído directamente del Excel) para mostrar
// la fórmula exacta que aplica el Excel, no la tendencia textual.
function buildCalcExplanation(ind, result, meta, cumpl, monthLabel) {
  // Prioridad: cumpl_formula_type (leído del Excel) > tendencia textual
  const cft  = (ind.cumpl_formula_type || '').toLowerCase();
  const tend = (ind.tend || '').toLowerCase();
  const isReducir  = cft === 'reducir'  || (!cft && (tend.includes('reducir')  || tend.includes('disminuir')));
  const isMantener = cft === 'mantener' || (!cft && tend.includes('mantener'));
  const isDirecto  = cft === 'directo';
  // crecer: todo lo que no sea reducir/mantener/directo

  const r   = fmtSmart(result, ind.tipo_meta, ind.unidad, ind.decimales);
  const m   = fmtSmart(meta,   ind.tipo_meta, ind.unidad, ind.decimales);
  const pct = cumpl !== null ? numCO(cumpl * 100, 2) + '%' : '—';
  const st  = cumpl === null ? 'pendiente'
            : getSemaforoEntry(cumpl).name !== 'Pendiente'
              ? (cumpl >= DATA.thresholds.cumple ? 'cumple' : cumpl >= DATA.thresholds.riesgo ? 'riesgo' : 'critico')
              : 'critico';
  const stLabel = cumpl !== null ? getSemaforoEntry(cumpl).name : 'Sin datos';

  let formulaLine, steps, interp;

  if (isReducir) {
    // Fórmula Excel: 1 − (Resultado − Meta) / Meta
    const dev    = result - meta;
    const devPct = meta !== 0 ? dev / meta : 0;
    formulaLine  = `Cumpl = 1 − (Resultado − Meta) / Meta`;
    steps = [
      `Desviación = ${r} − ${m} = <strong>${fmt(dev, ind.tipo_meta, ind.unidad)}</strong>`,
      `Ratio = ${fmt(dev, ind.tipo_meta, ind.unidad)} ÷ ${m} = <strong>${numCO(devPct * 100, 2)}%</strong>`,
      `Cumplimiento = 100% − ${numCO(devPct * 100, 2)}% = <strong>${pct}</strong>`,
    ];
    interp = result > meta
      ? `El resultado superó la meta en ${numCO(Math.abs(devPct)*100,2)}%, lo que reduce el cumplimiento a ${pct}.`
      : `El resultado está por debajo de la meta. Cumplimiento: ${pct}.`;

  } else if (isMantener) {
    // Fórmula Excel: 1 − |Resultado − Meta| / MAX(|Meta|, 1)
    const absDev = Math.abs(result - meta);
    const denom  = Math.max(Math.abs(meta), 1);
    const ratio  = absDev / denom;
    formulaLine  = `Cumpl = 1 − |Resultado − Meta| / MAX(|Meta|, 1)`;
    steps = [
      `Desviación absoluta = |${r} − ${m}| = <strong>${fmt(absDev, ind.tipo_meta, ind.unidad)}</strong>`,
      `Denominador = MAX(|${m}|, 1) = <strong>${numCO(denom * 100, 2)}%</strong>`,
      `Ratio = ${fmt(absDev, ind.tipo_meta, ind.unidad)} ÷ ${numCO(denom * 100, 2)}% = <strong>${numCO(ratio * 100, 2)}%</strong>`,
      `Cumplimiento = 100% − ${numCO(ratio * 100, 2)}% = <strong>${pct}</strong>`,
    ];
    interp = `Desviación de ${numCO(ratio*100,2)}% respecto a la meta. Cumplimiento: ${pct}.`;

  } else if (isDirecto) {
    // Fórmula Excel: el resultado ya es el cumplimiento
    formulaLine = `Cumpl = Resultado (el resultado es el % de cumplimiento)`;
    steps = [
      `Resultado obtenido: <strong>${r}</strong>`,
      `Cumplimiento = ${r} = <strong>${pct}</strong>`,
    ];
    interp = `El resultado representa directamente el nivel de cumplimiento: ${pct}.`;

  } else {
    // Fórmula Excel: G/H = Resultado ÷ Meta (crecer)
    const ratio = meta !== 0 ? result / meta : 0;
    formulaLine = `Cumpl = Resultado ÷ Meta`;
    steps = [
      `Resultado obtenido: <strong>${r}</strong>`,
      `Meta del período: <strong>${m}</strong>`,
      `Cumplimiento = ${r} ÷ ${m} = <strong>${pct}</strong>`,
    ];
    interp = result >= meta
      ? `El resultado supera o iguala la meta. Cumplimiento: ${pct}.`
      : `El resultado está por debajo de la meta en ${numCO((1 - ratio)*100,2)}%. Cumplimiento: ${pct}.`;
  }

  return { formulaLine, steps, interp, pct, st, stLabel };
}

function showCalcTooltip(ind, result, meta, cumpl, monthLabel, anchorEl) {
  // Cerrar cualquier tooltip abierto
  closeCalcTooltip();

  if (result === null || result === undefined) return;

  const { formulaLine, steps, interp, pct, st } = buildCalcExplanation(ind, result, meta, cumpl, monthLabel);
  const stLabels = { cumple:'✓ En meta', riesgo:'⚠ En riesgo', critico:'✗ Crítico', pendiente:'Sin datos' }; // legacy

  const tip = document.createElement('div');
  tip.className = 'calc-tooltip';
  tip.id = 'calc-tooltip-el';
  tip.innerHTML = `
    <div class="ct-month">${escapeHTML(monthLabel)}</div>
    <div class="ct-formula">${escapeHTML(formulaLine)}</div>
    <div class="ct-steps">
      ${steps.map((s,i) => `<div class="ct-step"><div class="ct-step-n">${i+1}</div><div>${s}</div></div>`).join('')}
    </div>
    <div class="ct-result ${st}">
      <span class="ct-result-label ${st}">${stLabels[st]||''}</span>
      <span class="ct-result-pct ${st}">${pct}</span>
    </div>
    <div class="ct-interp">${escapeHTML(interp)}</div>
    <div class="ct-close-hint">Clic fuera para cerrar</div>
  `;

  document.body.appendChild(tip);

  // Posicionar encima o debajo del elemento
  const rect  = anchorEl.getBoundingClientRect();
  const tw    = 320;
  const th    = tip.offsetHeight || 220;
  let   left  = rect.left + rect.width/2 - tw/2;
  const spaceBelow = window.innerHeight - rect.bottom - 12;
  const spaceAbove = rect.top - 12;
  let   top, above = false;

  if (spaceBelow >= th || spaceBelow >= spaceAbove) {
    top = rect.bottom + window.scrollY + 10;
  } else {
    top = rect.top + window.scrollY - th - 10;
    above = true;
    tip.classList.add('above');
  }

  left = Math.max(12, Math.min(left, window.innerWidth - tw - 12));
  tip.style.left  = left + 'px';
  tip.style.top   = top + 'px';
  tip.style.width = tw + 'px';

  // Overlay invisible para cerrar
  const ov = document.getElementById('calc-tooltip-overlay');
  ov.classList.add('active');
  ov.onclick = () => closeCalcTooltip();
}

function closeCalcTooltip() {
  const existing = document.getElementById('calc-tooltip-el');
  if (existing) existing.remove();
  const ov = document.getElementById('calc-tooltip-overlay');
  if (ov) { ov.classList.remove('active'); ov.onclick = null; }
}

// Alias para compatibilidad
function showCalcBubble(ind, result, meta, cumpl, monthLabel, anchorEl) {
  showCalcTooltip(ind, result, meta, cumpl, monthLabel, anchorEl);
}

// ── Exportar indicador individual ─────────────────────────
// ════════════════════════════════════════════════════════════
// Utilidades de exportación Excel con estilos (xlsx-js-style)
// ════════════════════════════════════════════════════════════

// Paleta de perspectivas para Excel
const PERSP_PALETTE = {
  'FINANCIERA':              { header:'0D5FA6', sub:'1A7CC4', band1:'EBF4FF', band2:'D6EAFA', accent:'0D5FA6', text:'0D3B6B' },
  'ASOCIADOS':               { header:'16704A', sub:'1E9860', band1:'EDFAF3', band2:'D0F5E4', accent:'16704A', text:'0D4530' },
  'PROCESOS INTERNOS':       { header:'0A7A86', sub:'0DA6B5', band1:'E8FAFB', band2:'C8F2F6', accent:'0A7A86', text:'064F57' },
  'APRENDIZAJE Y CRECIMIENTO':{ header:'5B21B6', sub:'7C3AED', band1:'F5F0FF', band2:'EDE5FF', accent:'5B21B6', text:'3B1078' },
};
function getPalette(perspKey) {
  return PERSP_PALETTE[perspKey] || { header:'334155', sub:'475569', band1:'F8FAFC', band2:'F1F5F9', accent:'334155', text:'0F172A' };
}
function hexXL(hex) { return (hex||'#CCCCCC').replace(/^#/,'').toUpperCase().padEnd(6,'0'); }

// ── Constructores de celda (seguros, sin valores nulos) ──────────────────
// SheetJS requiere que t:'n' siempre tenga v numérico; t:'s' siempre v string.
// Nunca pasar v:null con t:'n' — corrompe el XML.
function cs(v, s) {
  // Celda de texto — siempre string
  const val = (v === null || v === undefined) ? '' : String(v);
  return { v: val, t: 's', s: s || {} };
}
function cn(v, s, z) {
  // Celda numérica — si no hay número válido, devolver celda de texto '—'
  if (typeof v !== 'number' || !isFinite(v)) return { v: '—', t: 's', s: s || {} };
  return { v, t: 'n', z: z || '#,##0.00', s: s || {} };
}
function cpct(v, s) {
  // Celda porcentaje (v es 0-1) — si no es número válido, texto '—'
  if (typeof v !== 'number' || !isFinite(v)) return { v: '—', t: 's', s: s || {} };
  return { v, t: 'n', z: '0.00%', s: s || {} };
}
function cf(formula, fallback, s, z) {
  // Celda con fórmula Excel real (auditable) — fallback para cuando no aplica
  if (!formula) return cn(fallback, s, z);
  const cell = { f: formula, t: 'n', z: z || '0.00%', s: s || {} };
  if (typeof fallback === 'number' && isFinite(fallback)) cell.v = fallback;
  return cell;
}

// ── Estilos base ──────────────────────────────────────────
function sTitle(bgHex, fgHex) {
  return { font:{ bold:true, sz:14, color:{rgb: hexXL(fgHex||'#FFFFFF')}, name:'Calibri' },
           fill:{ fgColor:{ rgb: hexXL(bgHex) } },
           alignment:{ horizontal:'left', vertical:'center', indent:1 },
           border:{} };
}
function sSecTitle(bgHex, fgHex) {
  return { font:{ bold:true, sz:11, color:{rgb: hexXL(fgHex||'#FFFFFF')}, name:'Calibri' },
           fill:{ fgColor:{ rgb: hexXL(bgHex) } },
           alignment:{ horizontal:'left', vertical:'center', indent:1 } };
}
function sHdr(bgHex, fgHex, center) {
  return { font:{ bold:true, sz:10, color:{rgb: hexXL(fgHex||'#FFFFFF')}, name:'Calibri' },
           fill:{ fgColor:{ rgb: hexXL(bgHex) } },
           alignment:{ horizontal: center?'center':'left', vertical:'center', wrapText:true },
           border:{ bottom:{ style:'medium', color:{rgb:'CCCCCC'} },
                    top:   { style:'thin',   color:{rgb:'CCCCCC'} },
                    left:  { style:'thin',   color:{rgb:'CCCCCC'} },
                    right: { style:'thin',   color:{rgb:'CCCCCC'} } } };
}
function sLbl(bgHex) {
  return { font:{ bold:true, sz:9, color:{rgb:'475569'}, name:'Calibri' },
           fill:{ fgColor:{ rgb: hexXL(bgHex||'#F8FAFC') } },
           alignment:{ horizontal:'left', vertical:'center', indent:1 },
           border:{ bottom:{ style:'thin', color:{rgb:'E2E8F0'} },
                    right: { style:'thin', color:{rgb:'E2E8F0'} } } };
}
function sVal(bgHex, bold, center) {
  return { font:{ bold:!!bold, sz:10, color:{rgb:'0F172A'}, name:'Calibri' },
           fill:{ fgColor:{ rgb: hexXL(bgHex||'#FFFFFF') } },
           alignment:{ horizontal: center?'center':'left', vertical:'center', wrapText:true, indent: center?0:1 },
           border:{ bottom:{ style:'thin', color:{rgb:'E2E8F0'} },
                    right: { style:'thin', color:{rgb:'E2E8F0'} } } };
}
function sNum(bgHex, bold) {
  return { font:{ bold:!!bold, sz:10, color:{rgb:'0F172A'}, name:'Calibri' },
           fill:{ fgColor:{ rgb: hexXL(bgHex||'#FFFFFF') } },
           alignment:{ horizontal:'center', vertical:'center' },
           border:{ bottom:{ style:'thin', color:{rgb:'E2E8F0'} },
                    right: { style:'thin', color:{rgb:'E2E8F0'} } } };
}
function sStat(status) {
  const m = { cumple:{bg:'D1FAE5',fg:'065F46'}, riesgo:{bg:'FEF3C7',fg:'92400E'},
               critico:{bg:'FEE2E2',fg:'991B1B'}, pendiente:{bg:'F1F5F9',fg:'64748B'},
               'Cumple':{bg:'D1FAE5',fg:'065F46'}, 'En riesgo':{bg:'FEF3C7',fg:'92400E'},
               'Crítico':{bg:'FEE2E2',fg:'991B1B'}, 'Sin datos':{bg:'F1F5F9',fg:'64748B'} };
  const c = m[status] || m['Sin datos'];
  return { font:{ bold:true, sz:10, color:{rgb:c.fg}, name:'Calibri' },
           fill:{ fgColor:{ rgb:c.bg } },
           alignment:{ horizontal:'center', vertical:'center' },
           border:{ bottom:{ style:'thin', color:{rgb:'E2E8F0'} },
                    right: { style:'thin', color:{rgb:'E2E8F0'} } } };
}
function sPct(cumpl) {
  const bg = cumpl===null?'F1F5F9':cumpl>=0.95?'D1FAE5':cumpl>=0.80?'FEF3C7':'FEE2E2';
  const fg = cumpl===null?'64748B':cumpl>=0.95?'065F46':cumpl>=0.80?'92400E':'991B1B';
  return { font:{ bold:true, sz:10, color:{rgb:fg}, name:'Calibri' },
           fill:{ fgColor:{ rgb:bg } },
           alignment:{ horizontal:'center', vertical:'center' },
           border:{ bottom:{ style:'thin', color:{rgb:'E2E8F0'} },
                    right: { style:'thin', color:{rgb:'E2E8F0'} } } };
}
function sBand(yi, bgLight, bgDark) {
  const bg = yi%2===0 ? bgLight : bgDark;
  return { font:{ sz:10, color:{rgb:'0F172A'}, name:'Calibri' },
           fill:{ fgColor:{ rgb:bg } },
           alignment:{ horizontal:'center', vertical:'center' },
           border:{ bottom:{ style:'thin', color:{rgb:'D1D5DB'} },
                    right: { style:'thin', color:{rgb:'D1D5DB'} } } };
}
function sEmpty() {
  return { font:{sz:10,name:'Calibri'}, fill:{fgColor:{rgb:'FFFFFF'}} };
}

// ── Escribir celda en worksheet ───────────────────────────
function wc(ws, r, c, cell) {
  if (!cell) return;
  const addr = XLSX.utils.encode_cell({r, c});
  // Sanitizar: nunca v:null/undefined con tipo numérico
  if ((cell.t === 'n') && (cell.v === null || cell.v === undefined || !isFinite(cell.v))) {
    cell = { v: '—', t: 's', s: cell.s || {} };
  }
  // Sanitizar: texto siempre string
  if (cell.t === 's' && typeof cell.v !== 'string') {
    cell = { ...cell, v: cell.v === null || cell.v === undefined ? '' : String(cell.v) };
  }
  ws[addr] = cell;
}
function merge(ws, r1, c1, r2, c2) {
  ws['!merges'] = ws['!merges'] || [];
  ws['!merges'].push({s:{r:r1,c:c1}, e:{r:r2,c:c2}});
}
function setRef(ws, rows, cols) {
  if (rows > 0 && cols > 0) {
    ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:rows-1,c:cols-1}});
  }
}

// ── Calcular cumplimiento mensual ─────────────────────────
function calcCumpl(result, meta, tend) {
  if (result===null||result===undefined||typeof meta!=='number') return null;
  const t = (tend||'').toLowerCase();
  if (t.includes('reducir')||t.includes('disminuir')) return meta!==0 ? 1.0-(result-meta)/meta : null;
  if (t.includes('mantener')) return meta!==0 ? 1-Math.abs(result-meta)/Math.abs(meta) : null;
  return meta!==0 ? result/meta : null;
}

// ── Hoja mensual de un indicador ──────────────────────────
function buildIndSheet(ind, pal) {
  // Replica la estructura exacta de las hojas IND-XX del Excel origen:
  // Fila 1:  Título indicador
  // Fila 2:  Perspectiva · Objetivo específico
  // Fila 4:  Fórmula
  // Fila 5:  Descripción
  // Fila 6:  Unidad
  // Fila 7:  Tendencia objetivo
  // Fila 8:  Frecuencia
  // Fila 9:  Responsable
  // Fila 11: 🎯 METAS ANUALES
  // Fila 12: Encabezados metas (Año | Meta | ...)
  // Filas 13-17: Una fila por año (2026-2030)
  // Fila 19: 📝 CAPTURA MENSUAL
  // Fila 20: Instrucción
  // Fila 21: Encabezados tabla mensual
  // Filas 22+: Datos (60 filas, una por mes)

  const MNAMES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const nVars  = ind.hide_input_vars ? 0 : ind.vars.length;
  // Cols: # | Año | Mes | [vars...] | Resultado | Meta | % Cumpl. | Estado
  const NCOLS  = 4 + nVars + 4;
  const ws     = { '!merges':[], '!rows':[], '!cols':[] };
  const persp  = DATA.perspectives.find(p => p.indicators.includes(ind));

  // ── Fila estilos helpers locales ─────────────────────────
  const sBand = (c) => c%2===0 ? pal.band1 : pal.band2;
  const R0 = 0; // base-0 row counter

  let r = R0;

  // ── R0: Título ──────────────────────────────────────────
  const titleText = `${ind.code}  ·  ${ind.name.toUpperCase()}`;
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs(c===0?titleText:'', sTitle(pal.header)));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:30}; r++;

  // ── R1: Perspectiva · Objetivo ─────────────────────────
  const perspLine = `Perspectiva: ${persp?.name||''}   ·   Objetivo específico: ${ind.ini_icon||'🎯'} ${ind.ini}`;
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs(c===0?perspLine:'', sSecTitle(pal.sub)));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:20}; r++;

  // ── R2: vacía ──────────────────────────────────────────
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs('', sEmpty()));
  ws['!rows'][r] = {hpt:6}; r++;

  // ── R3-R8: Atributos (replicando filas 5-10 del Excel) ──
  const attrs = [
    ['Fórmula',           ind.formula],
    ['Descripción',       ind.desc],
    ['Unidad',            ind.unidad],
    ['Tendencia objetivo',ind.tend],
    ['Frecuencia',        ind.frec],
    ['Responsable',       ind.responsable||'Por definir'],
  ];
  attrs.forEach(([lbl, val]) => {
    wc(ws, r, 0, cs('', sVal(pal.band1)));  // col A vacía
    wc(ws, r, 1, cs(lbl, {...sLbl(pal.band2), alignment:{horizontal:'right',vertical:'center'}}));
    merge(ws, r, 1, r, 2);
    wc(ws, r, 3, cs(val, {...sVal(pal.band1), font:{sz:10,color:{rgb:'0F172A'},name:'Calibri',
      italic: lbl==='Fórmula'}}));
    merge(ws, r, 3, r, NCOLS-1);
    for(let c=NCOLS;c<NCOLS+1;c++) wc(ws, r, c, cs('', sVal(pal.band1)));
    ws['!rows'][r] = {hpt:18}; r++;
  });

  // ── R9: vacía ──────────────────────────────────────────
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs('', sEmpty()));
  ws['!rows'][r] = {hpt:6}; r++;

  // ── R10: Título metas ───────────────────────────────────
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs(c===0?'🎯  METAS ANUALES  (editable)':'', sSecTitle(pal.sub)));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:22}; r++;

  // ── R11: Encabezados metas ─────────────────────────────
  const metaHdrs = ['', 'Año', 'Meta', 'Crec. esperado / Notas', '', '', ''];
  ['','Año','Meta','Notas','','',''].forEach((h,c)=>{
    if(c < NCOLS) wc(ws, r, c, cs(h, sHdr(pal.header,'#FFFFFF', c>0)));
  });
  // extra cols
  for(let c=4;c<NCOLS;c++) wc(ws, r, c, cs('', sHdr(pal.header,'#FFFFFF',true)));
  ws['!rows'][r] = {hpt:20}; r++;
  const metaHdrRow = r; // para referencias de fórmula (fila Excel base-1 de los datos de metas)

  // ── R12-R16: Una fila por año ───────────────────────────
  // En el Excel origen, las metas están en filas 14-18 (base-1)
  // Aquí empiezan en fila Excel (r+1) base-1
  const metaDataStartXL = r + 1;  // fila Excel base-1 del primer año (2026)
  DATA.years.forEach((yr, yi) => {
    const bg  = sBand(yi);
    const m   = ind.metas[yi];
    const isCurr = yr === new Date().getFullYear();
    const yrSt = {...sVal(bg,true,true), font:{bold:true,sz:11,
      color:{rgb: isCurr?hexXL(pal.accent):'0F172A'},name:'Calibri'}};
    wc(ws, r, 0, cs('', sVal(bg)));
    wc(ws, r, 1, cs(yr.toString(), yrSt));
    wc(ws, r, 2, m!==null ? cn(m, {...sNum(bg,true), font:{bold:true,sz:11,color:{rgb:hexXL(pal.accent)},name:'Calibri'}}) : cs('—', sNum(bg)));
    // Notas / cumplimiento año
    const y = ind.yearly[yi];
    const cp = y?.cumplimiento;
    wc(ws, r, 3, cp!=null ? cpct(cp, sPct(cp)) : cs(y?.n_months>0?'—':'Pendiente', sPct(null)));
    const lbl = y ? semaforoLabel(y.cumplimiento, y.status) : 'Pendiente';
    wc(ws, r, 4, cs(lbl, sStat(lbl)));
    for(let c=5;c<NCOLS;c++) wc(ws, r, c, cs('', sVal(bg)));
    ws['!rows'][r] = {hpt:20}; r++;
  });

  // ── R17: vacía ──────────────────────────────────────────
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs('', sEmpty()));
  ws['!rows'][r] = {hpt:6}; r++;

  // ── R18: Título captura ─────────────────────────────────
  for(let c=0;c<NCOLS;c++) wc(ws, r, c,
    cs(c===0?'📝  CAPTURA MENSUAL DE VARIABLES  (Ene-2026 → Dic-2030)':'', sSecTitle(pal.sub)));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:22}; r++;

  // ── R19: Instrucción ───────────────────────────────────
  const instrText = 'Ingrese los valores en las celdas azules. El resultado, % cumplimiento y semáforo se calculan automáticamente.';
  for(let c=0;c<NCOLS;c++) wc(ws, r, c,
    cs(c===0?instrText:'', {...sVal(pal.band1), font:{sz:9,italic:true,color:{rgb:'64748B'},name:'Calibri'}}));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:28}; r++;

  // ── R20: Encabezados tabla mensual ─────────────────────
  // Replicando Excel: # | Año | Mes | [vars] | Resultado | Meta | % Cumpl. | Estado
  const hdrs = ['#','Año','Mes',
    ...(ind.hide_input_vars ? [] : ind.vars.map(v=>v.label)),
    'Resultado','Meta','% Cumpl.','Estado'];
  hdrs.forEach((h,c) => wc(ws, r, c, cs(h, sHdr(c<3?pal.header:c<3+nVars?pal.sub:pal.header, '#FFFFFF', c>0))));
  ws['!rows'][r] = {hpt:22};
  const hdrRowXL = r + 1;  // fila Excel base-1 del encabezado
  r++;

  // ── R21+: 60 filas de datos con FÓRMULAS EXCEL REALES ──
  // Columnas (base-0): 0=# 1=Año 2=Mes [3..3+nVars-1]=vars 3+nVars=Res 3+nVars+1=Meta 3+nVars+2=Cumpl 3+nVars+3=Estado
  const cRes   = XLSX.utils.encode_col(3 + nVars);      // columna Resultado (letra)
  const cMeta  = XLSX.utils.encode_col(3 + nVars + 1);  // columna Meta
  const cCumpl = XLSX.utils.encode_col(3 + nVars + 2);  // columna % Cumpl.
  const cEst   = XLSX.utils.encode_col(3 + nVars + 3);  // columna Estado
  const cft    = ind.cumpl_formula_type || '';
  const tLwr   = (ind.tend||'').toLowerCase();

  // Fila Excel base-1 donde empieza la tabla (primera fila de datos)
  const dataStartXL = r + 1;

  DATA.months_labels.forEach((_, i) => {
    const rec    = (ind.monthly_inputs||[])[i] || {};
    const yi     = Math.floor(i/12), mi = i%12;
    const bg     = sBand(yi);
    const result = ind.monthly_results[i];
    const metaAn = ind.metas[yi];
    const metaPer= (ind.meta_periodos && ind.meta_periodos[i]!=null) ? ind.meta_periodos[i] : metaAn;
    const cumpl  = (result!=null && metaPer!=null) ? calcCumpl(result, metaPer, ind.tend) : null;
    const xlRow  = r + 1;  // fila Excel base-1 de esta fila de datos

    // Columnas de variables (letras Excel, base-0 → col 3,4,...)
    const varCols = ind.vars.map((_,vi) => XLSX.utils.encode_col(3 + vi));

    // ── # | Año | Mes ──────────────────────────────────────
    wc(ws, r, 0, cn(i+1, sNum(bg)));
    wc(ws, r, 1, cs((2026+yi).toString(), {...sVal(bg,false,true), font:{bold:mi===0,sz:10,
      color:{rgb:mi===0?hexXL(pal.accent):'475569'},name:'Calibri'}}));
    wc(ws, r, 2, cs(MNAMES[mi], sVal(bg)));

    // ── Variables de entrada ──────────────────────────────
    if(!ind.hide_input_vars) {
      ind.vars.forEach((v,vi) => {
        const val = rec[v.code];
        if(val==null) {
          wc(ws, r, 3+vi, cs('', {...sNum(bg), font:{sz:10,color:{rgb:'94A3B8'},name:'Calibri'}}));
        } else {
          const lbl = (v.label||'').toLowerCase();
          const isDollar = lbl.includes('($)') || lbl.includes('pesos') || lbl.includes('cop');
          const isPct    = lbl.includes('%') || lbl.includes('tasa') || lbl.includes('porcentaje');
          // Marcar celdas con datos como editables (fondo azul claro como en Excel origen)
          const inputSt = {...sNum(bg), fill:{fgColor:{rgb:'DBEAFE'}},
            font:{sz:10,color:{rgb:'1D4ED8'},name:'Calibri',bold:false}};
          wc(ws, r, 3+vi, cn(val, inputSt, isDollar?'#,##0.00':isPct?'0.00%':'#,##0.00'));
        }
      });
    }

    // ── Resultado: fórmula según tipo ──────────────────────
    // Replicando col G del Excel origen
    let resFormula = null;
    if(!ind.hide_input_vars && nVars >= 2 && ind.tipo === 'ratio') {
      const numC = XLSX.utils.encode_col(3);
      const denC = XLSX.utils.encode_col(3 + nVars - 1);
      resFormula = `IFERROR(IF(${denC}${xlRow}=0,"",${numC}${xlRow}/${denC}${xlRow}),"")`;
    } else if(!ind.hide_input_vars && nVars >= 1 && ind.tipo === 'ratio') {
      const numC = XLSX.utils.encode_col(3);
      resFormula = `IFERROR(IF(${numC}${xlRow}="","",${numC}${xlRow}),"")`;
    }
    // Para base_social: resultado = E - base_dic_año_anterior
    if(ind.tipo === 'base_social' && !ind.hide_input_vars && nVars>=1) {
      const valC = XLSX.utils.encode_col(3);
      // base dic año anterior está en metaDataStartXL + yi - 1 (fila año anterior en tabla metas)
      // Simplificamos: usar valor directo (auditable via valor)
      resFormula = result!=null ? null : null; // usar valor directo
    }

    const resSt = {...sNum(bg), font:{bold:true,sz:10,color:{rgb:hexXL(pal.accent)},name:'Calibri'}};
    if(result==null) {
      wc(ws, r, 3+nVars, cs('', {...sNum(bg), font:{sz:10,color:{rgb:'94A3B8'},name:'Calibri'}}));
    } else if(resFormula) {
      wc(ws, r, 3+nVars, cf(resFormula, result, resSt, '#,##0.00######'));
    } else {
      wc(ws, r, 3+nVars, cn(result, resSt, '#,##0.00######'));
    }

    // ── Meta: fórmula VLOOKUP-style referenciando tabla metas ──
    // Replicando col H del Excel origen: =IFERROR(VLOOKUP(C_row,$B$metaStart:$C$metaEnd,2,FALSE),"")
    // Aquí usamos valor directo con referencia a la celda de meta del año
    const metaMetaRow = metaDataStartXL + yi;  // fila Excel base-1 de la meta del año en la tabla metas
    const metaFormula = `IFERROR(INDEX(C${metaDataStartXL}:C${metaDataStartXL+4},MATCH(B${xlRow},B${metaDataStartXL}:B${metaDataStartXL+4},0)),"")`;
    const metaSt = {...sNum(bg), font:{sz:10,color:{rgb:'1D4ED8'},name:'Calibri'}};
    if(metaAn==null) {
      wc(ws, r, 3+nVars+1, cs('', {...sNum(bg)}));
    } else {
      // Para meta periodo acumulada usar valor; para meta fija usar fórmula referenciando tabla
      const metaDisplayVal = metaPer != null ? metaPer : metaAn;
      wc(ws, r, 3+nVars+1, cf(metaFormula, metaDisplayVal, metaSt, '#,##0.######'));
    }

    // ── % Cumpl.: fórmula Excel real ──────────────────────
    // Replicando col I del Excel origen: =IFERROR(IF(OR(G="",NOT(ISNUMBER(H))),"",G/H),"")
    const resRef  = `${cRes}${xlRow}`;
    const metaRef = `${cMeta}${xlRow}`;
    let cumplFormula;
    if(cft==='directo') {
      cumplFormula = `IFERROR(IF(OR(${resRef}="",NOT(ISNUMBER(${metaRef}))),"",${resRef}),"")`;
    } else if(cft==='reducir' || (!cft && (tLwr.includes('reducir')||tLwr.includes('disminuir')))) {
      cumplFormula = `IFERROR(IF(OR(${resRef}="",NOT(ISNUMBER(${metaRef}))),"",1-(${resRef}-${metaRef})/${metaRef}),"")`;
    } else if(cft==='mantener' || (!cft && tLwr.includes('mantener'))) {
      cumplFormula = `IFERROR(IF(OR(${resRef}="",NOT(ISNUMBER(${metaRef}))),"",1-ABS(${resRef}-${metaRef})/MAX(ABS(${metaRef}),1)),"")`;
    } else {
      // Crecer — igual que Excel: =IFERROR(IF(OR(G="",NOT(ISNUMBER(H))),"",G/H),"")
      cumplFormula = `IFERROR(IF(OR(${resRef}="",NOT(ISNUMBER(${metaRef}))),"",${resRef}/${metaRef}),"")`;
    }
    const cumplCell = cf(cumplFormula, cumpl, sPct(cumpl), '0.00%');
    wc(ws, r, 3+nVars+2, cumplCell);

    // ── Estado: fórmula condicional referenciando % Cumpl. ──
    // Replicando col J del Excel origen: =IFERROR(IF(I="","Pendiente",VLOOKUP(I,Params,2,TRUE)),"Pendiente")
    // Simplificamos: semáforo basado en umbrales del dashboard
    const cumplRef = `${cCumpl}${xlRow}`;
    const cumple = DATA.thresholds?.cumple || 0.95;
    const riesgo = DATA.thresholds?.riesgo || 0.80;
    const stFormula = `IFERROR(IF(${cumplRef}="","Pendiente",IF(${cumplRef}>=${cumple},"Cumple",IF(${cumplRef}>=${riesgo},"En riesgo","Crítico"))),"Pendiente")`;
    const stLabel = cumpl===null?'Pendiente':cumpl>=cumple?'Cumple':cumpl>=riesgo?'En riesgo':'Crítico';
    wc(ws, r, 3+nVars+3, cf(stFormula, stLabel, sStat(stLabel), '@'));

    ws['!rows'][r] = {hpt:17};
    r++;
  });

  setRef(ws, r, NCOLS);

  // ── Anchos de columna (replicando Excel origen) ──────────
  ws['!cols'] = [
    {wch:5},   // #
    {wch:7},   // Año
    {wch:13},  // Mes
    ...(ind.hide_input_vars ? [] : ind.vars.map(()=>({wch:22}))),
    {wch:14},  // Resultado
    {wch:14},  // Meta
    {wch:11},  // % Cumpl.
    {wch:13},  // Estado
  ];

  return ws;
}


// ── Hoja ficha resumen de un indicador ───────────────────
function buildFichaSheet(ind, pal) {
  const ws = { '!merges':[], '!rows':[], '!cols':[] };
  const persp = DATA.perspectives.find(p=>p.indicators.includes(ind));
  const NCOLS = 6;
  let r = 0;

  // Título
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs(c===0 ? `FICHA  ·  ${ind.code}  ·  ${ind.name}` : '', sTitle(pal.header)));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:34};
  r++;

  // Subtítulo perspectiva
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs(c===0 ? (persp?.name||'') : '', sSecTitle(pal.sub)));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:22};
  r++;

  // Fila vacía
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs('', sEmpty()));
  ws['!rows'][r] = {hpt:8};
  r++;

  // Atributos
  const attrs = [
    ['Fórmula de cálculo', ind.formula],
    ['Descripción', ind.desc],
    ['Unidad de medida', ind.unidad],
    ['Tendencia objetivo', ind.tend],
    ['Frecuencia de medición', ind.frec],
    ['Área responsable', ind.responsable||'—'],
  ];
  attrs.forEach(([lbl, val])=>{
    wc(ws, r, 0, cs(lbl, sLbl(pal.band2)));
    wc(ws, r, 1, cs(val, {...sVal(pal.band1), font:{sz:10,color:{rgb:'0F172A'},name:'Calibri'}}));
    merge(ws, r, 1, r, NCOLS-1);
    ws['!rows'][r] = {hpt:18};
    r++;
  });

  // Fila vacía
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs('', sEmpty()));
  ws['!rows'][r] = {hpt:8};
  r++;

  // Cumplimiento por año
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs(c===0?'CUMPLIMIENTO POR AÑO':'', sSecTitle(pal.sub)));
  merge(ws, r, 0, r, NCOLS-1);
  ws['!rows'][r] = {hpt:22};
  r++;

  // Encabezados
  ['Año','Meta','Resultado promedio','Cumplimiento','Estado','Meses con datos'].forEach((h,c)=>{
    wc(ws, r, c, cs(h, sHdr(pal.header,'#FFFFFF',c>0)));
  });
  ws['!rows'][r] = {hpt:22};
  r++;

  // Filas por año
  ind.yearly.forEach((y,yi)=>{
    const bg  = yi%2===0 ? pal.band1 : pal.band2;
    const cp  = y.cumplimiento;
    const lbl = semaforoLabel(y.cumplimiento, y.status);
    const isCurr = y.year === new Date().getFullYear();
    const yrSt = {...sVal(bg, true, true), font:{bold:true, sz:11, color:{rgb: isCurr?hexXL(pal.accent):'0F172A'}, name:'Calibri'}};
    wc(ws, r, 0, cs(y.year.toString(), yrSt));
    wc(ws, r, 1, y.meta_raw!=null ? cn(y.meta_raw, sNum(bg)) : cs('—', sNum(bg)));
    wc(ws, r, 2, y.avg_result!=null ? cn(y.avg_result, sNum(bg)) : cs('—', sNum(bg)));
    wc(ws, r, 3, cp!=null ? cpct(cp, sPct(cp)) : cs('—', sPct(null)));
    wc(ws, r, 4, cs(lbl, sStat(lbl)));
    wc(ws, r, 5, cn(y.n_months, sNum(bg)));
    ws['!rows'][r] = {hpt:20};
    r++;
  });

  // ── Nota de fórmula de cumplimiento ──────────────────
  for(let c=0;c<NCOLS;c++) wc(ws, r, c, cs('', sEmpty()));
  ws['!rows'][r] = {hpt:8}; r++;

  const tend_lwr2 = (ind.tend||'').toLowerCase();
  const cft2 = ind.cumpl_formula_type || '';
  const formulaCumpl = cft2==='reducir'||tend_lwr2.includes('reducir')
    ? '=1 − (Resultado − Meta) / Meta'
    : cft2==='mantener'||tend_lwr2.includes('mantener')
    ? '=1 − |Resultado − Meta| / MAX(|Meta|; 1)'
    : cft2==='directo' ? '= Resultado (directo como %)'
    : '= Resultado ÷ Meta';

  wc(ws, r, 0, cs('Fórmula % Cumpl.', sLbl(pal.band2)));
  wc(ws, r, 1, cs(formulaCumpl, {...sVal(pal.band1), font:{italic:true,sz:10,color:{rgb:'1D4ED8'},name:'Calibri',bold:true}}));
  merge(ws, r, 1, r, NCOLS-1);
  ws['!rows'][r] = {hpt:20}; r++;

  wc(ws, r, 0, cs('Hoja con fórmulas', sLbl(pal.band2)));
  wc(ws, r, 1, cs('Ver hoja "Datos Mensuales" — columna % Cumpl. contiene fórmulas auditables referenciando Resultado y Meta de cada fila.', {...sVal(pal.band1), font:{sz:9,color:{rgb:'64748B'},name:'Calibri'}}));
  merge(ws, r, 1, r, NCOLS-1);
  ws['!rows'][r] = {hpt:28}; r++;

  setRef(ws, r, NCOLS);
  ws['!cols'] = [{wch:24},{wch:16},{wch:20},{wch:14},{wch:14},{wch:16}];
  return ws;
}

// ── Exportar indicador individual ────────────────────────
function exportIndicadorXlsx() {
  const ind = MODAL_IND;
  if (!ind) return;

  const fname = `${ind.code} ${ind.name.substring(0,30).replace(/[/\\?*[\]]/g,'_')}.xlsx`;

  // Prioridad: usar la hoja copiada fielmente del Excel origen (con fórmulas y estilos originales)
  if (ind.xlsx_b64 && ind.xlsx_b64.length > 0) {
    try {
      // Decodificar base64 y descargar directamente
      const binary = atob(ind.xlsx_b64);
      const bytes  = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url; a.download = fname; a.click();
      URL.revokeObjectURL(url);
      return;
    } catch(e) {
      console.warn('Error descargando xlsx_b64, usando fallback SheetJS:', e);
    }
  }

  // Fallback: construir con SheetJS si no hay b64 embebido
  if (typeof XLSX === 'undefined') {
    alert('xlsx-js-style no disponible. Verifica tu conexión a internet.'); return;
  }
  const persp = DATA.perspectives.find(p=>p.indicators.includes(ind));
  const pal   = getPalette(persp?.key||'');
  const wb    = XLSX.utils.book_new();
  wb.Props    = { Title: ind.code + ' — ' + ind.name, Author: 'FEBOR Plan Estratégico 2026-2030', Company: 'FEBOR Entidad Cooperativa' };
  XLSX.utils.book_append_sheet(wb, buildFichaSheet(ind, pal), 'Ficha');
  XLSX.utils.book_append_sheet(wb, buildIndSheet(ind, pal), 'Datos Mensuales');
  XLSX.writeFile(wb, fname);
}

// ── Exportar consolidado ──────────────────────────────────
function exportConsolidado() {
  // Intentar descargar el Excel fuente original (100% auditable)
  if (typeof __XLSX_SOURCE_B64__ !== 'undefined' && __XLSX_SOURCE_B64__ && __XLSX_SOURCE_B64__.length > 100) {
    try {
      const byteStr = atob(__XLSX_SOURCE_B64__);
      const arr = new Uint8Array(byteStr.length);
      for (let i = 0; i < byteStr.length; i++) arr[i] = byteStr.charCodeAt(i);
      const blob = new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      const today = new Date().toISOString().slice(0,10);
      a.href     = url;
      a.download = `Tablero_Indicadores_PE_${today}.xlsx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 500);
      return;
    } catch(e) {
      console.warn('Error descargando Excel fuente:', e);
    }
  }
  // Fallback: generar Excel con SheetJS si no hay fuente embebida
  if (typeof XLSX === 'undefined') {
    alert('xlsx-js-style no disponible. Verifica tu conexión a internet.'); return;
  }
  const wb = XLSX.utils.book_new();
  wb.Props = { Title:'Tablero Plan Estratégico FEBOR 2026-2030', Author:'FEBOR', Company:'FEBOR Entidad Cooperativa' };

  // ── Hoja Resumen General ────────────────────────────────
  const wsRes   = { '!merges':[], '!rows':[], '!cols':[] };
  const HDR_BG  = '0657A4';
  const HDR_BG2 = '334155';
  const RCOLS   = 7 + DATA.years.length*3 + 1;
  let r = 0;

  // Título principal
  for(let c=0;c<RCOLS;c++) wc(wsRes, r, c, cs(c===0?'PLAN ESTRATÉGICO FEBOR 2026 – 2030  ·  RESUMEN GENERAL':'', sTitle(HDR_BG)));
  merge(wsRes, r, 0, r, RCOLS-1);
  wsRes['!rows'][r] = {hpt:36};
  r++;

  // Sub-encabezados de año
  const yHdrs = ['','Indicador','Perspectiva','Tendencia','Unidad','Frec.','Responsable'];
  yHdrs.forEach((h,c)=> wc(wsRes, r, c, cs(h, sHdr(HDR_BG,'#FFFFFF',c>0))));
  DATA.years.forEach((yr,yi)=>{
    const yrBg = yi%2===0 ? '1E4A8A' : '163A70';
    wc(wsRes, r, 7+yi*3, cs(yr.toString()+(yr===new Date().getFullYear()?' ★':''), sHdr(yrBg,'#FFFFFF',true)));
    wc(wsRes, r, 7+yi*3+1, cs('', sHdr(yrBg,'#FFFFFF',true)));
    wc(wsRes, r, 7+yi*3+2, cs('', sHdr(yrBg,'#FFFFFF',true)));
    merge(wsRes, r, 7+yi*3, r, 7+yi*3+2);
  });
  wc(wsRes, r, 7+DATA.years.length*3, cs('Avance Global', sHdr('0E9F8F','#FFFFFF',true)));
  wsRes['!rows'][r] = {hpt:20};
  r++;

  // Sub-cols por año
  for(let c=0;c<7;c++) wc(wsRes, r, c, cs('', sHdr(HDR_BG,'#FFFFFF',true)));
  DATA.years.forEach((yr,yi)=>{
    const yrBg = yi%2===0 ? '1E4A8A' : '163A70';
    ['Meta','Cumplim.','Estado'].forEach((h,hi)=>{
      wc(wsRes, r, 7+yi*3+hi, cs(h, sHdr(yrBg,'#FFFFFF',true)));
    });
  });
  wc(wsRes, r, 7+DATA.years.length*3, cs('', sHdr('0E9F8F','#FFFFFF',true)));
  wsRes['!rows'][r] = {hpt:18};
  r++;

  // Filas de indicadores
  let lastPerspKey = null;
  getAllIndicators().forEach(({ind, persp})=>{
    const pal    = getPalette(persp.key);
    const isNew  = persp.key !== lastPerspKey;
    if (isNew) lastPerspKey = persp.key;
    const rowBg  = pal.band1;

    wc(wsRes, r, 0, cs(ind.code, {...sNum(isNew?pal.band2:rowBg, true), font:{bold:true,sz:10,color:{rgb:pal.accent},name:'Calibri'}}));
    wc(wsRes, r, 1, cs(ind.name, sVal(rowBg)));
    wc(wsRes, r, 2, cs(persp.name, {...sVal(isNew?pal.band2:rowBg, true), font:{bold:isNew,sz:10,color:{rgb:pal.accent},name:'Calibri'}}));
    wc(wsRes, r, 3, cs(ind.tend, sVal(rowBg)));
    wc(wsRes, r, 4, cs(ind.unidad, sVal(rowBg, false, true)));
    wc(wsRes, r, 5, cs(ind.frec, sVal(rowBg, false, true)));
    wc(wsRes, r, 6, cs(ind.responsable||'—', sVal(rowBg)));
    ind.yearly.forEach((y,yi)=>{
      const cp = y.cumplimiento;
      const lbl= semaforoLabel(y.cumplimiento, y.status);
      wc(wsRes, r, 7+yi*3,   y.meta_raw!=null ? cn(y.meta_raw, sNum(rowBg)) : cs('—', sNum(rowBg)));
      wc(wsRes, r, 7+yi*3+1, cp!=null ? cpct(cp, sPct(cp)) : cs('—', sPct(null)));
      wc(wsRes, r, 7+yi*3+2, cs(lbl, sStat(lbl)));
    });
    const ac = ind.accumulated;
    wc(wsRes, r, 7+DATA.years.length*3, ac!=null ? cpct(ac, sPct(ac)) : cs('—', sPct(null)));
    wsRes['!rows'][r] = {hpt:18};
    r++;
  });

  setRef(wsRes, r, RCOLS);
  wsRes['!cols'] = [
    {wch:9},{wch:34},{wch:18},{wch:12},{wch:10},{wch:10},{wch:18},
    ...DATA.years.flatMap(()=>[{wch:13},{wch:12},{wch:12}]),
    {wch:14}
  ];
  XLSX.utils.book_append_sheet(wb, wsRes, 'Resumen General');

  // ── Una hoja por indicador ────────────────────────────
  getAllIndicators().forEach(({ind, persp})=>{
    const pal = getPalette(persp.key);
    XLSX.utils.book_append_sheet(wb, buildIndSheet(ind, pal), ind.code);
  });

  const today = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `FEBOR Tablero PE 2026-2030 ${today}.xlsx`);
}


document.addEventListener('DOMContentLoaded', () => {
  renderHeaderKPIs();
  renderGauges();
  renderFilters();
  renderMatriz();
  renderPerspectives();
  updateCompactState();
  window.addEventListener('scroll', onCompactScroll, { passive: true });
  window.addEventListener('resize', onCompactScroll);
  document.addEventListener('keydown', e => {
    if(e.key==='Escape'){ if(STATE.activeKpi) closePopover(); else closeModal(); }
  });
});

"""


# ============================================================
# HTML TEMPLATE
# ============================================================

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0">
  <title>Plan Estratégico 2026–2030 · Dashboard</title>
  <script>__CHARTJS__</script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
  <style>
__CSS__
  </style>
</head>
<body>
  <header class="header">
    <!-- Fila superior: marca + título + botón -->
    <div class="header-inner">
      <!-- Brand FEBOR -->
      <div class="header-brand">
        <div class="header-logo-box">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/>
            <polyline points="9 22 9 12 15 12 15 22"/>
          </svg>
        </div>
        <div class="header-brand-text">
          <div class="header-brand-name">Febor</div>
          <div class="header-brand-sub">Entidad Cooperativa</div>
        </div>
      </div>

      <!-- Título -->
      <div class="header-titles">
        <h1><span class="brand-dot"></span> Plan Estratégico 2026 – 2030</h1>
        <div class="header-meta">
          <span class="header-meta-item">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
            Seguimiento mensual
          </span>
          <span class="header-meta-item">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 14l4-4 4 4 5-5"/></svg>
            25 indicadores
          </span>
          <span class="header-meta-item">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/></svg>
            4 perspectivas estratégicas
          </span>
        </div>
      </div>

      <!-- Botones header -->
      <div class="header-btns">
        <button class="btn-datos-origen btn-escalas" onclick="openEscalas()">
          <div class="xl-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20a8 8 0 1 1 8-8"/><path d="M12 12l4-4"/><path d="M4 12h1"/><path d="M12 4v1"/><path d="M19 12h1"/></svg>
          </div>
          <span class="xl-label">Escalas de la meta</span>
        </button>
        <a href="https://feborec.github.io/Ruta-Plan-Estrat-gico-2026-2030/" target="_blank" class="btn-ruta-estrategica">
          <div class="xl-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12h18"/><path d="M13 5l7 7-7 7"/></svg>
          </div>
          <span class="xl-label">Ir a la Ruta de Planeación Estratégica</span>
        </a>
        <a href="https://feborentidad.sharepoint.com/:f:/s/Riesgos/IgCyJeu9SCtOQK7ZPM6BcfJlASf2_0UdUDIXTZgCNjUiAH4?e=Og8tzB" target="_blank" class="btn-datos-origen">
          <div class="xl-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="M5 12h14"/></svg>
          </div>
          <span class="xl-label">Agregar datos al origen</span>
        </a>
        <button class="btn-excel-consolidado" onclick="exportConsolidado()">
          <div class="xl-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
          </div>
          <span class="xl-label">Excel Consolidado</span>
        </button>
      </div>
    </div>

    <!-- Fila inferior: KPIs -->
    <div class="header-kpis" id="header-kpis"></div>
  </header>

  <!-- Barra compacta: aparece al bajar -->
  <div class="cbar" id="cbar"><div class="cbar-inner" id="cbar-inner"></div></div>

  <div class="container">

    <div class="sticky-top">
      <div id="gauges-panel" class="gauges-section" style="padding:12px 24px 0;"></div>
      <div class="filters-bar">
        <div id="persp-filter-row" class="persp-filter-row"></div>
        <div class="filter-divider"></div>
        <div class="filter-group">
          <span class="filter-label">Año:</span>
          <div id="year-chips" style="display:flex;gap:6px;flex-wrap:wrap"></div>
        </div>
      </div>
    </div>

    <div style="padding: 16px 24px 0; max-width:1400px; margin: 0 auto;" id="main-content">
    <div id="matriz-panel"></div>
    </div><!-- /main-content -->

    <div class="footer" style="max-width:1400px;margin:0 auto;padding:22px 32px;">
      Generado automáticamente desde Tablero_Indicadores_PE.xlsx<br>
      Para actualizar: <code>python "Plan Estrategico Dashboard.py"</code>
    </div>
    <!-- Excel fuente embebido -->
    <script>const __XLSX_SOURCE_B64__ = '__XLSX_B64__';</script>
  </div>

  <div class="mtz-tooltip" id="mtz-tooltip"></div>
  <div class="calc-tooltip-overlay" id="calc-tooltip-overlay"></div>

  <div class="modal-overlay" id="modal-overlay">
    <div class="modal" id="modal-content" onclick="event.stopPropagation()"></div>
  </div>

  <div class="popover-overlay" id="popover-overlay">
    <div class="popover" id="popover"></div>
  </div>

  <script>
    const DATA = __DATA_JSON__;
__JS__
  </script>
</body>
</html>

"""


# ============================================================
# RENDER
# ============================================================

def _embed_xlsx(xlsx_path):
    """Lee el Excel fuente y lo retorna como base64 para embeber en el HTML."""
    import base64, os
    try:
        with open(str(xlsx_path), 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode('utf-8')
    except Exception:
        return ''


def _copy_sheet_to_wb(ws_src, wb_out, title, hidden=False):
    """Copia una hoja completa (fórmulas, estilos, merges) a un workbook destino."""
    from copy import copy as _copy
    ws_dst = wb_out.create_sheet(title)

    for col_letter, col_dim in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col_letter].width  = col_dim.width
        ws_dst.column_dimensions[col_letter].hidden = col_dim.hidden

    for row_num, row_dim in ws_src.row_dimensions.items():
        ws_dst.row_dimensions[row_num].height = row_dim.height

    for row in ws_src.iter_rows():
        for cell in row:
            dst = ws_dst.cell(row=cell.row, column=cell.column)
            dst.value = cell.value
            if cell.has_style:
                dst.font          = _copy(cell.font)
                dst.fill          = _copy(cell.fill)
                dst.border        = _copy(cell.border)
                dst.alignment     = _copy(cell.alignment)
                dst.number_format = cell.number_format
                dst.protection    = _copy(cell.protection)

    for merge_range in ws_src.merged_cells.ranges:
        ws_dst.merge_cells(str(merge_range))

    if hidden:
        ws_dst.sheet_state = 'hidden'

    return ws_dst


def _build_ind_xlsx_b64(wb_src, code):
    """Copia la hoja del indicador + hoja Parámetros (oculta) del Excel origen.
    Preserva 100%: fórmulas, formatos, estilos, merges, anchos, altos.
    La hoja Parámetros se incluye oculta para resolver vínculos internos
    (semáforo, responsables, etc.) sin referencias a archivos externos."""
    import base64, io
    from openpyxl import Workbook

    if code not in wb_src.sheetnames:
        return ''
    try:
        wb_out = Workbook()
        wb_out.remove(wb_out.active)

        # 1. Hoja principal del indicador (visible)
        _copy_sheet_to_wb(wb_src[code], wb_out, code, hidden=False)

        # 2. Hoja Parámetros (oculta) — resuelve VLOOKUP de semáforo y responsables
        if 'Parámetros' in wb_src.sheetnames:
            _copy_sheet_to_wb(wb_src['Parámetros'], wb_out, 'Parámetros', hidden=True)

        # 3. Para IND-18: incluir hoja Inventario_Automatización (visible)
        #    Es la fuente de datos del indicador y debe estar presente para auditoría
        if code == 'IND-18' and 'Inventario_Automatización' in wb_src.sheetnames:
            _copy_sheet_to_wb(wb_src['Inventario_Automatización'], wb_out, 'Inventario_Automatización', hidden=False)

        buf = io.BytesIO()
        wb_out.save(buf)
        return base64.b64encode(buf.getvalue()).decode('utf-8')
    except Exception as e:
        print(f"  ⚠️  No se pudo generar xlsx para {code}: {e}")
        return ''


def _load_chartjs():
    """Carga Chart.js desde el archivo local si existe; si no, deja un fallback
    que carga desde CDN (útil cuando el dashboard se publica en GitHub Pages)."""
    import os
    here = os.path.dirname(os.path.abspath(__file__))
    chart_path = os.path.join(here, 'chart_umd.js')
    if os.path.exists(chart_path):
        with open(chart_path, 'r', encoding='utf-8') as f:
            return f.read()
    # Fallback: inyectar un loader que carga Chart.js desde CDN
    return (
        "/* Chart.js no embebido — cargando desde CDN */\n"
        "(function(){var s=document.createElement('script');"
        "s.src='https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js';"
        "s.onload=function(){console.log('Chart.js CDN cargado.');};"
        "document.head.appendChild(s);})();"
    )


def render_dashboard_html(payload, xlsx_path=None):
    """Renderiza el dashboard como HTML autocontenido (incluye Chart.js inline)."""
    json_str = json.dumps(payload, ensure_ascii=False, indent=2)
    xlsx_b64 = _embed_xlsx(xlsx_path) if xlsx_path else ''
    html = (HTML_TEMPLATE
            .replace('__CHARTJS__', _load_chartjs())
            .replace('__CSS__', CSS)
            .replace('__JS__', JS)
            .replace('__DATA_JSON__', json_str)
            .replace('__XLSX_B64__', xlsx_b64))
    return html


# ============================================================================
# 3-8. EXTRACCIÓN, CÁLCULO, VALIDACIÓN, PAYLOAD, RENDER Y MAIN
# ============================================================================


# openpyxl es la única dependencia externa
# Imports locales
# (data_model y dashboard_templates están incluidos arriba en este mismo archivo)


# ============================================================
# CONFIGURACIÓN
# ============================================================

YEARS = [2026, 2027, 2028, 2029, 2030]
MONTHS = []
for y in YEARS:
    for m in range(1, 13):
        MONTHS.append((y, m))

MES_NAMES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MES_ABBR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
            "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def freq_months(frec):
    """Devuelve el conjunto de meses (1-12) en que se mide el indicador."""
    f = (frec or "").lower()
    if "mensual" in f:
        return set(range(1, 13))
    if "trimestral" in f:
        return {3, 6, 9, 12}
    if "semestral" in f:
        return {6, 12}
    if "anual" in f:
        return {12}
    return set(range(1, 13))


# ============================================================
# EXTRACCIÓN DE DATOS DEL EXCEL
# ============================================================

def _find_data_first_row(ws):
    """Localiza la primera fila de la tabla mensual (donde col B = 1)."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == 1:
            return r
    return None


def _read_metas(ws, data_first):
    """Lee las 5 metas anuales desde la tabla de metas (sobre data_first)."""
    metas = []
    for r in range(1, data_first):
        if ws.cell(r, 2).value == 2026:
            for i in range(5):
                metas.append(ws.cell(r + i, 3).value)
            break
    return metas


def _read_responsables(wb):
    """Extrae el catálogo de responsables de la hoja Parámetros."""
    responsables = {}
    if 'Parámetros' not in wb.sheetnames:
        return responsables
    pws = wb['Parámetros']
    for r in range(1, pws.max_row + 1):
        if pws.cell(r, 2).value == 'Código':
            for rr in range(r + 1, r + 1 + len(IND) + 5):
                code = pws.cell(rr, 2).value
                resp = pws.cell(rr, 6).value
                if code and isinstance(code, str) and code.startswith('IND-'):
                    responsables[code] = resp
            break
    return responsables


def extract_indicator(wb, d, wb_dt=None):
    """Extrae datos crudos de la hoja de un indicador.
    wb     = workbook data_only=False (fórmulas)
    wb_dt  = workbook data_only=True  (valores calculados, opcional)
    """
    sheet_name = d['code']
    if sheet_name not in wb.sheetnames:
        print(f"  ⚠️  Hoja {sheet_name} no encontrada — se omite")
        return None
    ws = wb[sheet_name]

    data_first = _find_data_first_row(ws)
    if data_first is None:
        print(f"  ⚠️  No se localizó la tabla mensual en {sheet_name}")
        return None

    metas_raw = _read_metas(ws, data_first)
    if len(metas_raw) != 5:
        metas_raw = list(d['metas'])  # fallback al catálogo

    # Para base_social (IND-14): leer bases dinámicamente del Excel
    # La base de cada año = meta total del año anterior (col C, fila del año anterior)
    if d.get('tipo') == 'base_social':
        bases_dinamicas = {}
        for r in range(1, data_first):
            yr_cell = ws.cell(r, 2).value
            if yr_cell == 2025:  # Fila base (cierre dic año anterior al primero)
                for offset, year in enumerate(range(2026, 2031)):
                    base_val = ws.cell(r + offset, 3).value  # Col C = meta total
                    if isinstance(base_val, (int, float)):
                        bases_dinamicas[year] = base_val
                break
        if bases_dinamicas:
            d = dict(d)  # copia para no mutar el catálogo global
            d['base_social_bases'] = bases_dinamicas

    hdr_row_idx = data_first - 1

    # ── Mapear columnas de entrada por nombre de encabezado ──────────────
    # Construir mapa hdr_texto → columna
    hdr_map = {}
    for c in range(2, 20):
        v = str(ws.cell(hdr_row_idx, c).value or '').strip()
        if v:
            hdr_map[v.lower()] = c

    # Intentar mapear cada var del catálogo al encabezado más parecido
    col_map = {}
    for var_code, var_label in d['vars']:
        label_l = var_label.strip().lower()
        best_col = None
        META_SKIP = {'resultado','meta año','% cumpl.','estado','#','año','mes','meta'}
        # 1. Coincidencia exacta
        if label_l in hdr_map and label_l not in META_SKIP:
            best_col = hdr_map[label_l]
        else:
            # 2. Primeros 15 chars (tolera sufijos como "90 días", "VAR ANUAL")
            for hdr_txt, col in hdr_map.items():
                if hdr_txt.lower() in META_SKIP:
                    continue
                if label_l[:15] in hdr_txt or hdr_txt[:15] in label_l:
                    best_col = col
                    break
            # 3. Palabras clave si sigue sin match
            if best_col is None:
                # Keywords: incluir palabras >= 3 chars que sean siglas/términos únicos
                kws = [w.strip('()%') for w in label_l.split()
                       if len(w.strip('()%')) >= 3
                       and w.strip('()%') not in ('tasa','promedio','ponderada','total',
                                                   'saldo','valor','resultado','margen',
                                                   'indice','índice','cartera','ahorros',
                                                   'deposito','depósito','del','de','la',
                                                   'los','las','por','que')]
                kws = [k for k in kws if k]
                for hdr_txt, col in hdr_map.items():
                    if hdr_txt.lower() in META_SKIP:
                        continue
                    if any(k in hdr_txt for k in kws):
                        best_col = col
                        break
        col_map[var_code] = best_col

    # Fallback secuencial si alguna var no se mapeó
    # Encontrar primera col de entrada (saltando #, Año, Mes, Resultado, Meta, Cumpl, Estado)
    SKIP = {'#','año','mes','resultado','meta año','% cumpl.','estado',
            'resultado encuesta','resultado de encuesta',
            'meta 2026','meta 2027','meta 2028','meta 2029','meta 2030',
            'cumplimiento','cumpl.','estado del indicador'}
    first_input_col = 5
    for c in range(5, 14):
        h = str(ws.cell(hdr_row_idx, c).value or '').strip().lower()
        if h and h not in SKIP:
            first_input_col = c
            break

    none_vars = [vc for vc, _ in d['vars'] if col_map.get(vc) is None]
    for j, vc in enumerate(none_vars):
        col_map[vc] = first_input_col + j

    # Detectar columna Resultado precalculado (Excel fórmula → data_only=True)
    resultado_col = hdr_map.get('resultado')

    # Detectar columna Meta y su divisor mensual (ej. =10%/11 → divisor=11)
    meta_per_col = None
    meta_divisor = 1
    for c_m in range(2, 18):
        h_m = str(ws.cell(hdr_row_idx, c_m).value or '').strip().lower()
        if h_m in ('meta', 'meta año', 'meta '):
            meta_per_col = c_m
            _fm = str(ws.cell(data_first, c_m).value or '')
            if '/11' in _fm: meta_divisor = 11
            elif '/12' in _fm: meta_divisor = 12
            break
    # Para indicadores con tipo_meta="ipc" (ej. IND-13 ROE):
    # La meta dinámica es el IPC CORRIDO del año (columna "IPC CORRIDO")
    if d.get('tipo_meta') == 'ipc' and meta_per_col is None:
        for c_m in range(2, 20):
            h_m = str(ws.cell(hdr_row_idx, c_m).value or '').strip().lower()
            if 'ipc corrido' in h_m or 'ipc_corrido' in h_m:
                meta_per_col = c_m
                break

    # Leer tendencia y fórmula directamente del Excel (fila 8 y 5)
    tend_excel = None
    formula_excel = None
    try:
        for r_ficha in range(5, 12):
            lbl = str(ws.cell(r_ficha, 2).value or '').lower()
            val_d = ws.cell(r_ficha, 4).value
            val_c = ws.cell(r_ficha, 3).value
            val = val_d or val_c
            if 'tendencia' in lbl and val:
                tend_excel = str(val).strip()
            if ('fórmula' in lbl or 'formula' in lbl) and val:
                formula_excel = str(val).strip()
    except Exception:
        pass

    # Hoja data_only=True para leer valores calculados (refs externas, fórmulas)
    ws_vals = wb_dt[sheet_name] if (wb_dt and sheet_name in wb_dt.sheetnames) else ws

    inputs = []
    for i, (y, m) in enumerate(MONTHS):
        row = data_first + i
        rec = {'year': y, 'month': m}
        for var_code, _ in d['vars']:
            c = col_map.get(var_code, first_input_col)
            # Intentar primero data_only=True (tiene valores de fórmulas y refs externas)
            v = ws_vals.cell(row, c).value
            if not isinstance(v, (int, float)):
                # Fallback: data_only=False (valores literales)
                v = ws.cell(row, c).value
            rec[var_code] = v if isinstance(v, (int, float)) else None
        # Resultado precalculado: leer del workbook data_only=True si está disponible
        if resultado_col:
            ws_dt = wb_dt[sheet_name] if (wb_dt and sheet_name in wb_dt.sheetnames) else ws
            rv = ws_dt.cell(row, resultado_col).value
            rec['resultado_excel'] = rv if isinstance(rv, (int, float)) else None
        # Meta del período: meta_anual / divisor (o meta_anual si divisor=1)
        yr_i = i // 12
        m_anual = d['metas'][yr_i] if yr_i < len(d['metas']) else None
        if meta_per_col:
            # Leer meta del período directamente del Excel con valores calculados
            _ws_vals = wb_dt[sheet_name] if (wb_dt and sheet_name in wb_dt.sheetnames) else ws
            mv_per = _ws_vals.cell(row, meta_per_col).value
            if isinstance(mv_per, (int, float)):
                rec['meta_periodo'] = mv_per
            elif m_anual is not None and meta_divisor > 1:
                rec['meta_periodo'] = m_anual / meta_divisor
            else:
                rec['meta_periodo'] = m_anual
        else:
            rec['meta_periodo'] = m_anual / meta_divisor if (m_anual is not None and meta_divisor > 1) else m_anual
        inputs.append(rec)

    # Extraer meta_periodos de los inputs
    meta_periodos_raw = [rec.get('meta_periodo') for rec in inputs]

    # ── Detectar tipo de fórmula % Cumpl. directamente del Excel ─────────
    # Lee la fórmula de la celda % Cumpl. fila 23 para determinar si es:
    # 'crecer'  → G/H o F/G (ratio simple)
    # 'reducir' → 1-(G-H)/H
    # 'mantener'→ 1-ABS(G-H)/MAX(ABS(H),1)
    cumpl_formula_type = None
    try:
        hdr_lower = {c: str(ws.cell(hdr_row_idx, c).value or '').lower().strip()
                     for c in range(2, 16)}
        # Buscar columna % Cumpl. — exacta, no confundir con vars que contengan 'cumpl'
        cumpl_c = next((c for c, h in hdr_lower.items()
                        if h in ('% cumpl.', '% cumpl', 'cumpl.', '% cumplimiento',
                                 '% de cumplimiento', 'cumplimiento')), None)
        if cumpl_c:
            # Leer fórmula del workbook data_only=False
            raw_f = str(ws.cell(data_first, cumpl_c).value or '').upper().replace(' ', '')
            if '1-ABS(' in raw_f:
                cumpl_formula_type = 'mantener'
            elif '1-(' in raw_f and 'ABS' not in raw_f:
                cumpl_formula_type = 'reducir'
            elif '"N/A"' in raw_f or raw_f.endswith(',"")') and '/' not in raw_f and '1-' not in raw_f:
                # Fórmula retorna el resultado directamente como cumplimiento (IND-18)
                cumpl_formula_type = 'directo'
            elif raw_f and '/' in raw_f and '1-' not in raw_f and 'ABS' not in raw_f:
                cumpl_formula_type = 'crecer'
    except Exception:
        pass

    # Detectar si las variables de entrada están en pesos (para mostrar en millones)
    # Incluye vars con ($) y también vars cuyo primer valor real sea >= 1,000,000
    vars_en_pesos = any('($)' in v[1] or 'pesos' in v[1].lower() for v in d['vars'])
    if not vars_en_pesos:
        # Verificar por magnitud: si alguna var tiene valores >= 1M
        for rec in inputs[:6]:  # revisar primeros 6 meses
            for vc, _ in d['vars']:
                val = rec.get(vc)
                if isinstance(val, (int, float)) and abs(val) >= 1_000_000:
                    vars_en_pesos = True
                    break
            if vars_en_pesos:
                break

    return {
        'code':         d['code'],
        'name':         d['indicador'],
        'formula':      formula_excel or d['formula'],
        'tend':         tend_excel or d['tend'],
        'cumpl_formula_type': cumpl_formula_type,
        'persp':        d['p'],
        'ini':          d['ini'],
        'desc':         d['desc'],
        'unidad':       d['unidad'],
        'tipo':         d['tipo'],
        'tipo_meta':    d['tipo_meta'],
        'frec':         d['frec'],
        'metas_raw':    metas_raw,
        'inputs':       inputs,
        'vars':         d['vars'],
        'allowed_months': sorted(range(1, 13)) if d.get('show_all_months') else sorted(freq_months(d['frec'])),
        'meta_periodos_raw': meta_periodos_raw,
        'vars_en_pesos': vars_en_pesos,
        'hide_input_vars': d.get('hide_input_vars', False),
        'decimales':       d.get('decimales', None),
    }


# ============================================================
# CÁLCULOS
# ============================================================

def compute_monthly_results(ind):
    """Calcula el resultado de cada uno de los 60 meses según el tipo de fórmula."""
    tipo = ind['tipo']
    inputs = ind['inputs']
    allowed = set(ind['allowed_months'])
    results = []

    for i, rec in enumerate(inputs):
        if rec['month'] not in allowed:
            results.append(None)
            continue

        var_keys = [k for k, _ in ind['vars']]
        result = None

        if tipo == 'ratio':
            # Último var = denominador; los anteriores se suman como numerador
            den = rec[var_keys[-1]]
            num_vals = [rec[k] for k in var_keys[:-1]]
            if den is not None and den != 0 and all(v is not None for v in num_vals):
                result = sum(num_vals) / den

        elif tipo == 'growth':
            cur = rec[var_keys[0]]
            if i == 0 or cur is None:
                result = None
            else:
                prev = inputs[i - 1][var_keys[0]]
                if prev is not None and prev != 0:
                    result = (cur - prev) / prev

        elif tipo == 'growth_acum':
            # Lógica exacta del Excel:
            # E[i] = acum[i] - acum[i-1]  (ingreso del mes)
            # F[i] = (E[i] - E[i-1]) / E[i-1]  (crecimiento del ingreso mensual)
            # resultado = F[i]
            acum = rec.get('acum')
            if acum is None:
                result = None
            else:
                mes_en_año = i % 12
                prev_acum = inputs[i-1].get('acum') if i > 0 else None
                # Ingreso del mes actual = diferencia de acumulados
                e_cur = acum if mes_en_año == 0 else (
                    (acum - prev_acum) if prev_acum is not None else None
                )
                # Ingreso del mes anterior
                if i >= 2:
                    acum_2 = inputs[i-2].get('acum')
                    mes_ant = (i-1) % 12
                    e_prev = inputs[i-1].get('acum') if mes_ant == 0 else (
                        (prev_acum - acum_2) if (prev_acum is not None and acum_2 is not None) else None
                    )
                elif i == 1 and prev_acum is not None:
                    e_prev = prev_acum  # Enero del año
                else:
                    e_prev = None
                # Crecimiento = (e_cur - e_prev) / e_prev
                if e_cur is not None and e_prev is not None and e_prev != 0:
                    result = (e_cur - e_prev) / e_prev
                else:
                    result = None

        elif tipo == 'diff':
            # Prioridad 1: resultado precalculado del Excel (más fiable — evita mapeo erróneo de vars)
            if rec.get('resultado_excel') is not None:
                result = rec['resultado_excel']
            else:
                # Prioridad 2: calcular a - b desde las variables de entrada
                a = rec.get(var_keys[0])
                b = rec.get(var_keys[1]) if len(var_keys) > 1 else None
                if a is not None and b is not None and a != b:
                    diff = a - b
                    result = diff * 10000 if ind['tipo_meta'] in ('bps','pts') else diff

        elif tipo in ('avg', 'count'):
            v_avg = rec.get(var_keys[0])
            res_xl = rec.get('resultado_excel')
            # Cuando hide_input_vars=True (IND-17, IND-18, IND-20): la columna Resultado
            # MANDA siempre. Si Resultado está vacío (sin fórmula/dato) → None (sin dato),
            # nunca se cae al valor de las vars de entrada. Así una celda vacía no muestra 0.
            if ind.get('hide_input_vars'):
                result = res_xl if isinstance(res_xl, (int, float)) else None
            # Prioridad 1: resultado_excel cuando el primer var es grande (IND-17)
            # o cuando hay múltiples vars con resultado pre-calculado (IND-16)
            elif res_xl is not None and (
                    len(var_keys) > 1 or
                    (v_avg is not None and isinstance(v_avg,(int,float)) and abs(v_avg) > 1_000)):
                # Si resultado_excel==0 pero TODAS las vars de entrada son None → sin datos
                if res_xl == 0 and len(var_keys) > 1:
                    all_none = all(rec.get(k) is None for k in var_keys)
                    result = None if all_none else res_xl
                else:
                    result = res_xl
            else:
                result = v_avg

        elif tipo == 'multi_sum':
            # Suma de todas las vars excepto la última (denominador)
            den = rec.get(var_keys[-1])
            num_vals = [rec.get(k) for k in var_keys[:-1]]
            # Permitir None en numeradores (ignorar, no bloquear)
            non_null_nums = [v for v in num_vals if v is not None]
            if den is not None and den != 0 and non_null_nums:
                result = sum(non_null_nums) / den

        elif tipo == 'retencion':
            # Fórmula Excel IND-15 (col L):
            #   Resultado = mantiene / (base_ant - fallecidos)
            # Prioridad: resultado_excel si disponible (data_only=True)
            if rec.get('resultado_excel') is not None:
                result = rec['resultado_excel']
            else:
                mantiene   = rec.get('mantiene')
                base_ant   = rec.get('base_ant')
                fallecidos = rec.get('fallecidos') or 0
                denom = (base_ant or 0) - fallecidos
                if mantiene is not None and base_ant is not None and denom != 0:
                    result = mantiene / denom

        elif tipo == 'base_social':
            # Fórmula Excel IND-14 (col F = E - $C$base_año):
            #   Resultado = total_asociados_mes - base_dic_año_anterior  → ACUMULADO
            #   TODOS los meses del año usan la misma base fija (no diferencia mes a mes)
            #   Prioridad: resultado_excel si está disponible
            total = rec.get('val')
            if total is not None:
                if rec.get('resultado_excel') is not None:
                    result = rec['resultado_excel']
                else:
                    year = rec.get('year', 2026)
                    bases = ind.get('base_social_bases', {})
                    base = bases.get(year, 5365)
                    result = total - base

        results.append(result)

    return results


def get_meta_numeric(ind, year):
    """Devuelve la meta numérica del año, o None si es texto."""
    m = ind['metas_raw'][YEARS.index(year)]
    return m if isinstance(m, (int, float)) else None


def compute_compliance(result, meta, tend, tipo=None):
    """Replica exactamente la fórmula % Cumpl. de cada hoja IND del Excel.

    Según la tendencia:
      Reducir  → 1 − (Resultado − Meta) / Meta
                 Ej: IND-F1/F2, IND-05 (índice de mora)
      Mantener → 1 − |Resultado − Meta| / MAX(|Meta|, 1)
                 Ej: IND-09/10/11/12/23
      Crecer   → Resultado / Meta
                 Ej: IND-F3, IND-01/04/06/07...
                 Nota: si Meta es negativo (ej. IND-04: meta=−200 bps)
                 y Resultado también es negativo, la división da positivo.
                 −5 bps / −200 bps = 2.5% ← INCORRECTO según Excel
                 El Excel usa 1−(G−H)/H para IND-04 implícitamente
                 porque la meta negativa con tendencia Crecer significa
                 "quiero acercarme a cero desde negativo"
    """
    if result is None or meta is None:
        return None
    # meta o resultado que no son número (p. ej. «#VALUE!» cuando falta el IPC del mes, o «>= IPC»): cumplimiento pendiente, como en el Excel
    if not isinstance(meta, (int, float)) or isinstance(meta, bool) or not isinstance(result, (int, float)) or isinstance(result, bool):
        return None

    # Si se pasa cumpl_formula_type como tipo, usarlo directamente
    if tipo == 'directo':
        return result  # El resultado YA es el cumplimiento (IND-18)
    if tipo == 'crecer':
        return result / meta if meta != 0 else None
    if tipo == 'reducir':
        return (1.0 - (result - meta) / meta) if meta != 0 else None
    if tipo == 'mantener':
        return 1.0 - abs(result - meta) / max(abs(meta), 1.0)

    # Caso especial: meta negativa con tendencia Crecer
    if tend == 'Crecer' and meta < 0:
        if meta == 0:
            return None
        return 1.0 - (result - meta) / meta

    if tend == 'Reducir':
        if meta == 0:
            return None
        return 1.0 - (result - meta) / meta

    if tend == 'Mantener':
        # Excel: 1 - ABS(Resultado-Meta) / MAX(ABS(Meta), 1)
        # Si meta < 1 el denominador es 1; si meta >= 1 el denominador es meta
        return 1.0 - abs(result - meta) / max(abs(meta), 1.0)

    # Crecer con meta positiva → Resultado / Meta
    if meta == 0:
        return None
    return result / meta


def status_from_compliance(c):
    """Mapea cumplimiento a estado usando umbrales parametrizables del Excel."""
    if c is None:
        return 'pendiente'
    if c >= THRESHOLD_CUMPLE:
        return 'cumple'
    if c >= THRESHOLD_RIESGO:
        return 'riesgo'
    return 'critico'


def compute_yearly_metrics(ind, monthly_results):
    """Para cada año (2026-2030): meta, resultado promedio, % cumplimiento, estado."""
    out = []
    meta_periodos = ind.get('meta_periodos_raw') or []
    tipo = ind.get('tipo', '')

    for yi, year in enumerate(YEARS):
        seg   = monthly_results[yi * 12: (yi + 1) * 12]
        valid = [r for r in seg if r is not None]
        meta  = get_meta_numeric(ind, year)

        if tipo == 'base_social' and valid:
            # Para base_social la meta de la tarjeta = meta acumulada del último mes con datos
            # y el cumplimiento = resultado_último / meta_acumulada_último
            last_idx_in_seg = max(i for i, r in enumerate(seg) if r is not None)
            flat_idx = yi * 12 + last_idx_in_seg
            meta_acum_ultimo = (meta_periodos[flat_idx]
                                if flat_idx < len(meta_periodos) and meta_periodos[flat_idx] is not None
                                else meta)
            last_result = seg[last_idx_in_seg]
            cumpl = compute_compliance(last_result, meta_acum_ultimo,
                                       ind['tend'], tipo=ind.get('cumpl_formula_type'))
            out.append({
                'year': year,
                'meta': meta_acum_ultimo,
                'meta_raw': meta_acum_ultimo,   # lo que muestra la tarjeta
                'avg_result': last_result,
                'cumplimiento': cumpl,
                'status': status_from_compliance(cumpl),
                'n_months': len(valid),
            })
            continue

        avg_result = sum(valid) / len(valid) if valid else None

        # Promedio de cumplimientos mensuales del año
        cmpls = []
        for r in valid:
            c = compute_compliance(r, meta, ind['tend'], tipo=ind.get('cumpl_formula_type'))
            if c is not None:
                cmpls.append(c)
        cumpl = sum(cmpls) / len(cmpls) if cmpls else None

        out.append({
            'year': year,
            'meta': meta,
            'meta_raw': ind['metas_raw'][yi],
            'avg_result': avg_result,
            'cumplimiento': cumpl,
            'status': status_from_compliance(cumpl),
            'n_months': len(valid),
        })
    return out


def compute_trend(monthly_results, ind):
    """Tendencia propia del indicador: 'up' / 'down' / 'flat' / None.

    Se calcula comparando el primer y último resultado disponibles en
    los últimos 12 meses con datos.
    """
    valid = [r for r in monthly_results if r is not None]
    if len(valid) < 2:
        return None
    window = valid[-12:] if len(valid) >= 12 else valid
    first, last = window[0], window[-1]
    if abs(last - first) / max(abs(first), 1e-9) < 0.02:
        return 'flat'
    return 'up' if last > first else 'down'


def compute_accumulated(monthly_results, ind):
    """Avance acumulado: cumplimiento del último resultado disponible vs meta 2030."""
    meta_2030 = get_meta_numeric(ind, 2030)
    if meta_2030 is None:
        return None
    latest = next((r for r in reversed(monthly_results) if r is not None), None)
    if latest is None:
        return None
    return compute_compliance(latest, meta_2030, ind['tend'])


def compute_alert(yearly_metrics):
    """Genera la alerta de desviación basada en el último año con datos."""
    for ym in reversed(yearly_metrics):
        if ym['cumplimiento'] is not None:
            c = ym['cumplimiento']
            y = ym['year']
            if c >= 0.95:
                return {'level': 'ok', 'text': f'En meta en {y} ({c*100:.0f}%)'}
            if c >= 0.80:
                return {'level': 'warn', 'text': f'Atención en {y}: {c*100:.0f}% de la meta'}
            return {'level': 'critical', 'text': f'Crítico en {y}: {c*100:.0f}% — revisar'}
    return {'level': 'pending', 'text': 'Sin datos cargados aún'}


# ============================================================
# VALIDACIÓN
# ============================================================

def validate_indicator(ind):
    """Valida los datos de un indicador y devuelve una lista de avisos."""
    warnings = []
    # ¿Las metas tienen el tipo esperado?
    for i, m in enumerate(ind['metas_raw']):
        if m is None or (isinstance(m, str) and 'definir' in m.lower()):
            warnings.append(f"Meta {YEARS[i]} sin definir")
    # ¿Hay meses con datos parciales (algunas variables sí, otras no)?
    var_keys = [k for k, _ in ind['vars']]
    for rec in ind['inputs']:
        if rec['month'] not in set(ind['allowed_months']):
            continue
        present = [rec[k] for k in var_keys]
        n_filled = sum(1 for v in present if v is not None)
        if 0 < n_filled < len(var_keys):
            warnings.append(f"Datos parciales en {MES_ABBR[rec['month']-1]} {rec['year']}")
    return warnings


# ============================================================
# PROCESAMIENTO COMPLETO
# ============================================================

# Valores de respaldo si la tabla del semáforo no se puede leer del Excel
SEMAFORO_LEVELS_DEFAULT = [
    {'lim':1.00,'bg':'#83F1B7','fg':'#006666','name':'Excelente','emoji':'🏆','desc':'Supera ampliamente el objetivo'},
    {'lim':0.90,'bg':'#BBF9D0','fg':'#15803D','name':'Sobresaliente','emoji':'🟢','desc':'Muy por encima del objetivo'},
    {'lim':0.80,'bg':'#BBF9D0','fg':'#15803D','name':'Satisfactorio','emoji':'🔵','desc':'Dentro del objetivo'},
    {'lim':0.70,'bg':'#FEF9C3','fg':'#A16207','name':'Aceptable','emoji':'🟦','desc':'Cerca del objetivo'},
    {'lim':0.60,'bg':'#FEF9C3','fg':'#A16207','name':'En seguimiento','emoji':'🟡','desc':'Por debajo — requiere acción'},
    {'lim':0.50,'bg':'#FFD3A7','fg':'#EA580C','name':'En riesgo','emoji':'🟠','desc':'Brecha significativa'},
    {'lim':0.40,'bg':'#FFD3A7','fg':'#EA580C','name':'En alerta','emoji':'🔶','desc':'Plan de mejora inmediato'},
    {'lim':0.30,'bg':'#FEE2E2','fg':'#DC2626','name':'Crítico','emoji':'🔴','desc':'Incumplimiento grave'},
    {'lim':-100000,'bg':'#FEC9C9','fg':'#B91C1C','name':'Muy crítico','emoji':'⛔','desc':'Incumplimiento severo'},
]
SEMAFORO_PENDIENTE_DEFAULT = {'bg':'#F1F5F9','fg':'#64748B','name':'Pendiente','emoji':'⚪','desc':'Sin datos registrados'}
SEMAFORO_LEVELS = list(SEMAFORO_LEVELS_DEFAULT)
SEMAFORO_PENDIENTE = dict(SEMAFORO_PENDIENTE_DEFAULT)


def _cell_rgb(color):
    """Color ARGB de openpyxl → '#RRGGBB' (None si es de tema o no está definido)."""
    try:
        rgb = color.rgb if color is not None else None
    except Exception:
        return None
    if isinstance(rgb, str) and len(rgb) in (6, 8):
        return '#' + rgb[-6:].upper()
    return None


def _clean_level_desc(text):
    """Quita el rango entre paréntesis del final, p. ej. '(30–39%)': el tablero lo calcula."""
    if not isinstance(text, str):
        return ''
    return re.sub(r'\s*\([^)]*%\)\s*$', '', text).strip()


def load_semaforo_levels(wb):
    """Lee la tabla del semáforo de Parámetros (debajo del título 'SEMÁFORO'):
    col B = límite inferior, C = nivel (su relleno y color de letra = colores
    del nivel), D = emoji, E = descripción. La fila 'Pendiente' define el estado
    sin datos. Retorna (niveles de mayor a menor límite, nivel pendiente)."""
    wp = wb['Parámetros']
    start = None
    for r in range(1, wp.max_row + 1):
        v = wp.cell(r, 2).value
        if isinstance(v, str) and 'SEMÁFORO' in v.upper():
            start = r
            break
    if start is None:
        return None, None
    defaults = {lv['name']: lv for lv in SEMAFORO_LEVELS_DEFAULT}
    levels, pendiente = [], None
    for r in range(start + 1, min(start + 30, wp.max_row + 1)):
        name = wp.cell(r, 3).value
        if not isinstance(name, str) or not name.strip():
            continue
        name = name.strip()
        c = wp.cell(r, 3)
        fill = _cell_rgb(c.fill.fgColor) if c.fill is not None and c.fill.fill_type else None
        font = _cell_rgb(c.font.color) if c.font is not None else None
        dflt = defaults.get(name, SEMAFORO_PENDIENTE_DEFAULT if name == 'Pendiente' else {})
        entry = {
            'name':  name,
            'emoji': wp.cell(r, 4).value or dflt.get('emoji', ''),
            'desc':  _clean_level_desc(wp.cell(r, 5).value) or dflt.get('desc', ''),
            'bg':    fill or dflt.get('bg', '#F1F5F9'),
            'fg':    font or dflt.get('fg', '#475569'),
        }
        lim = wp.cell(r, 2).value
        if isinstance(lim, (int, float)) and not isinstance(lim, bool):
            entry['lim'] = float(lim)
            levels.append(entry)
        elif name == 'Pendiente':
            pendiente = entry
            break
    if len(levels) < 3:
        return None, None
    levels.sort(key=lambda lv: lv['lim'], reverse=True)
    return levels, pendiente or dict(SEMAFORO_PENDIENTE_DEFAULT)


def load_semaforo_thresholds(wb):
    """Carga los niveles del semáforo desde Parámetros y deriva los cortes:
    verde (en meta) desde el nivel 'Excelente' y ámbar desde 'En riesgo'.
    Retorna (dict_deciles, threshold_en_meta, threshold_en_riesgo)."""
    global SEMAFORO_LEVELS, SEMAFORO_PENDIENTE
    try:
        levels, pendiente = load_semaforo_levels(wb)
    except Exception:
        levels, pendiente = None, None
    if levels:
        SEMAFORO_LEVELS, SEMAFORO_PENDIENTE = levels, pendiente
    else:
        print("   ⚠️  Semáforo: no se pudo leer la tabla de Parámetros (valores por defecto)")
        SEMAFORO_LEVELS = list(SEMAFORO_LEVELS_DEFAULT)
        SEMAFORO_PENDIENTE = dict(SEMAFORO_PENDIENTE_DEFAULT)
    deciles = {lv['name']: lv['lim'] for lv in SEMAFORO_LEVELS}
    # Verde solo cuando se alcanza o supera la meta (nivel 'Excelente')
    t_s = deciles.get('Excelente', SEMAFORO_LEVELS[0]['lim'])
    t_r = deciles.get('En riesgo', 0.50)
    return deciles, t_s, t_r


THRESHOLD_CUMPLE = 1.00
THRESHOLD_RIESGO = 0.50
THRESHOLD_DECILES = {
    'Excelente':1.0,'Sobresaliente':0.9,'Satisfactorio':0.8,
    'Aceptable':0.7,'En seguimiento':0.6,'En riesgo':0.5,
    'En alerta':0.4,'Crítico':0.3,'Muy crítico':0.1,'Inaceptable':0.0
}


def process_workbook(xlsx_path):
    """Procesa el workbook completo y devuelve la lista de indicadores enriquecidos."""
    wb    = load_workbook(xlsx_path, data_only=False)   # fórmulas
    wb_dt = load_workbook(xlsx_path, data_only=True)    # valores calculados
    global THRESHOLD_CUMPLE, THRESHOLD_RIESGO, THRESHOLD_DECILES
    THRESHOLD_DECILES, THRESHOLD_CUMPLE, THRESHOLD_RIESGO = load_semaforo_thresholds(wb)
    print(f"   ✓ Semáforo: {len(THRESHOLD_DECILES)} deciles · En meta ≥ {THRESHOLD_CUMPLE*100:.0f}%  ·  En riesgo ≥ {THRESHOLD_RIESGO*100:.0f}%")
    responsables = _read_responsables(wb)

    indicators = []
    for d in IND:
        ind = extract_indicator(wb, d, wb_dt)
        if ind is None:
            continue
        ind['responsable'] = responsables.get(ind['code']) or 'Por definir'
        ind['monthly_results'] = compute_monthly_results(ind)
        ind['meta_periodos']    = ind.get('meta_periodos_raw', [None]*60)
        ind['yearly'] = compute_yearly_metrics(ind, ind['monthly_results'])
        ind['trend'] = compute_trend(ind['monthly_results'], ind)
        ind['accumulated'] = compute_accumulated(ind['monthly_results'], ind)
        ind['alert'] = compute_alert(ind['yearly'])
        ind['warnings'] = validate_indicator(ind)
        indicators.append(ind)
    return indicators


# ============================================================
# CONSTRUCCIÓN DE PAYLOAD PARA EL DASHBOARD
# ============================================================

def _to_jsonable(v):
    if v is None:
        return None
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        return v
    if isinstance(v, (int, str, bool)):
        return v
    return str(v)


def build_payload(indicators, wb_src=None):
    """Construye el JSON consolidado que consume el dashboard."""
    perspectives = {}
    for p in PERSP_ORDER:
        perspectives[p] = {
            'key': p,
            'name': PERSP_INFO[p]['short'],
            'icon': PERSP_INFO[p]['icon'],
            'accent': '#' + PERSP_INFO[p]['accent'],
            'objetivo_general': OBJ_GEN[p],
            'indicators': [],
        }

    summary = {'total': 0, 'cumple': 0, 'riesgo': 0, 'critico': 0, 'pendiente': 0}

    for ind in indicators:
        # Estado consolidado del indicador (último año con datos)
        latest_status = 'pendiente'
        latest_cumpl = None
        latest_year = None
        for ym in reversed(ind['yearly']):
            if ym['cumplimiento'] is not None:
                latest_status = ym['status']
                latest_cumpl = ym['cumplimiento']
                latest_year = ym['year']
                break

        payload_ind = {
            'code': ind['code'],
            'name': ind['name'],
            'ini': ind['ini'],
            'ini_icon': INI_ICON.get(ind['ini'], '🎯'),
            'ini_desc': INI_DESC.get(ind['ini'], ''),
            'formula': ind['formula'],
            'desc': ind['desc'],
            'unidad': ind['unidad'],
            'tipo': ind['tipo'],
            'tipo_meta': ind['tipo_meta'],
            'tend': ind['tend'],
            'cumpl_formula_type': ind.get('cumpl_formula_type'),
            'frec': ind['frec'],
            'responsable': ind['responsable'],
            'vars': [{'code': v[0], 'label': v[1]} for v in ind['vars']],
            'vars_en_pesos': ind.get('vars_en_pesos', False),
            'hide_input_vars': ind.get('hide_input_vars', False),
            'decimales': ind.get('decimales', None),
            'metas': [_to_jsonable(m) for m in ind['metas_raw']],
            'monthly_results': [_to_jsonable(r) for r in ind['monthly_results']],
            'meta_periodos':   [_to_jsonable(m) for m in ind.get('meta_periodos', [None]*60)],
            'yearly': [{
                'year': y['year'],
                'meta': _to_jsonable(y['meta']),
                'meta_raw': _to_jsonable(y['meta_raw']),
                'avg_result': _to_jsonable(y['avg_result']),
                'cumplimiento': _to_jsonable(y['cumplimiento']),
                'status': y['status'],
                'n_months': y['n_months'],
            } for y in ind['yearly']],
            'trend': ind['trend'],
            'accumulated': _to_jsonable(ind['accumulated']),
            'alert': ind['alert'],
            'warnings': ind['warnings'],
            'allowed_months': ind['allowed_months'],
            'latest_status': latest_status,
            'latest_cumpl': _to_jsonable(latest_cumpl),
            'latest_year': latest_year,
            'xlsx_b64': _build_ind_xlsx_b64(wb_src, ind['code']) if wb_src else '',
            'monthly_inputs': [
                {k: (_to_jsonable(v) if k not in ('year','month') else v)
                 for k, v in rec.items()}
                for rec in ind['inputs']
            ],
        }
        perspectives[ind['persp']]['indicators'].append(payload_ind)
        summary['total'] += 1
        summary[latest_status] = summary.get(latest_status, 0) + 1

    return {
        'months_labels': [f"{MES_ABBR[m-1]} {y}" for (y, m) in MONTHS],
        'years': YEARS,
        'perspectives': list(perspectives.values()),
        'summary': summary,
        'thresholds': {
            'cumple': THRESHOLD_CUMPLE,
            'riesgo': THRESHOLD_RIESGO,
            'deciles': THRESHOLD_DECILES,
            # Niveles leídos de Parámetros (mayor → menor límite) y estado sin datos
            'colors': SEMAFORO_LEVELS,
            'pendiente': SEMAFORO_PENDIENTE,
        },
    }


# ============================================================
# AUTOPUBLICACIÓN A GITHUB
# ============================================================

def _github_api(method, path, token, data=None, timeout=20):
    """Llamada genérica a la GitHub REST API. Devuelve (status, json_dict)."""
    import urllib.request
    import urllib.error
    url = f'https://api.github.com{path}'
    body = None
    if data is not None:
        body = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=body,
        method=method,
        headers={
            'Accept': 'application/vnd.github+json',
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json',
            'X-GitHub-Api-Version': '2022-11-28',
            'User-Agent': 'Plan-Estrategico-Dashboard',
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read().decode('utf-8') or '{}'
            return r.status, json.loads(raw)
    except urllib.error.HTTPError as e:
        try:
            err_body = e.read().decode('utf-8')
            err_json = json.loads(err_body) if err_body else {}
        except Exception:
            err_json = {'message': str(e)}
        return e.code, err_json


def setup_github(token, repo_name, git_root, out_path):
    """Configuración inicial completa: crea el repo en GitHub, configura remote,
    hace push inicial y activa GitHub Pages. Todo vía API + git CLI."""
    import subprocess

    print(f"\n🔧 Configurando GitHub...")

    # 1. Validar token y obtener usuario
    status, user_info = _github_api('GET', '/user', token)
    if status != 200:
        print(f"   ❌ Token inválido o sin permisos ({status}).")
        print(f"      Verifique que el token tiene el scope 'repo'.")
        print(f"      Genere uno nuevo en: https://github.com/settings/tokens/new")
        return None
    username = user_info['login']
    print(f"   ✓ Autenticado como: {username}")

    # 2. Verificar si el repo existe; crear si no
    status, _ = _github_api('GET', f'/repos/{username}/{repo_name}', token)
    if status == 404:
        print(f"   ⏳ Creando repo {username}/{repo_name}...")
        status, resp = _github_api('POST', '/user/repos', token, data={
            'name': repo_name,
            'private': False,
            'auto_init': False,
            'description': 'Dashboard del Plan Estratégico 2026-2030',
            'has_issues': False,
            'has_wiki': False,
        })
        if status not in (200, 201):
            print(f"   ❌ No se pudo crear el repo: {resp.get('message', 'error')}")
            for err in resp.get('errors', []):
                print(f"      - {err.get('message', err)}")
            return None
        print(f"   ✓ Repo creado: https://github.com/{username}/{repo_name}")
    elif status == 200:
        print(f"   ✓ Repo existente: {username}/{repo_name}")
    else:
        print(f"   ❌ Error consultando el repo ({status})")
        return None

    # 3. Configurar remote 'origin' con token embebido (queda sólo en .git/config local)
    remote_url = f'https://x-access-token:{token}@github.com/{username}/{repo_name}.git'
    safe_url = f'https://github.com/{username}/{repo_name}.git'
    r = subprocess.run(['git', 'remote'], capture_output=True, text=True, cwd=str(git_root))
    remotes = r.stdout.split()
    if 'origin' in remotes:
        subprocess.run(['git', 'remote', 'set-url', 'origin', remote_url], cwd=str(git_root))
        print(f"   ✓ Remote 'origin' actualizado")
    else:
        subprocess.run(['git', 'remote', 'add', 'origin', remote_url], cwd=str(git_root))
        print(f"   ✓ Remote 'origin' agregado: {safe_url}")

    # 4. Asegurar que la rama sea 'main'
    subprocess.run(['git', 'branch', '-M', 'main'], capture_output=True, cwd=str(git_root))

    # 5. Si no hay commits aún, hacer commit inicial con el dashboard
    r = subprocess.run(['git', 'log', '-1', '--oneline'], capture_output=True, text=True, cwd=str(git_root))
    if r.returncode != 0:
        rel = str(out_path.resolve().relative_to(git_root))
        subprocess.run(['git', 'add', rel], cwd=str(git_root))
        subprocess.run(['git', 'commit', '-m', 'Inicial: dashboard del Plan Estratégico'],
                       capture_output=True, cwd=str(git_root))
        print(f"   ✓ Commit inicial creado")

    # 6. Push inicial
    print(f"   ⏳ Push inicial a GitHub...")
    r = subprocess.run(['git', 'push', '-u', 'origin', 'main'],
                       capture_output=True, text=True, cwd=str(git_root))
    if r.returncode != 0:
        err = r.stderr.strip() or r.stdout.strip()
        # Ocultar el token si aparece en el error
        err = err.replace(token, '***')
        print(f"   ❌ Push falló:")
        for line in err.split('\n')[:6]:
            print(f"      {line}")
        return None
    print(f"   ✓ Push exitoso")

    # 7. Activar GitHub Pages (branch main, folder /docs)
    print(f"   ⏳ Activando GitHub Pages...")
    status, resp = _github_api('POST', f'/repos/{username}/{repo_name}/pages', token, data={
        'source': {'branch': 'main', 'path': '/docs'},
    })
    if status in (200, 201):
        print(f"   ✓ GitHub Pages activado (branch main, folder /docs)")
    elif status == 409:
        print(f"   ℹ️  GitHub Pages ya estaba activado")
    else:
        print(f"   ⚠️  No se pudo activar Pages automáticamente ({status}: {resp.get('message','')})")
        print(f"      Actívelo manualmente en:")
        print(f"      https://github.com/{username}/{repo_name}/settings/pages")

    pages_url = f'https://{username}.github.io/{repo_name}/'
    print(f"\n🎉 ¡Configuración completa!")
    print(f"   📦 Repo:      https://github.com/{username}/{repo_name}")
    print(f"   🌐 Dashboard: {pages_url}")
    print(f"   ⏱️  La URL estará disponible en 1-2 minutos.")
    print(f"\n   Desde ahora, cada `python \"Plan Estrategico Dashboard.py\"`")
    print(f"   actualizará el dashboard publicado automáticamente.")
    return pages_url


def _find_git_root(start_path):
    """Busca la raíz del repo Git subiendo desde start_path."""
    p = Path(start_path).resolve()
    for _ in range(8):
        if (p / '.git').exists():
            return p
        if p.parent == p:
            break
        p = p.parent
    return None


def _try_gh_setup(git_root, out_path):
    """Intenta crear el repo + push + Pages usando GitHub CLI (gh).

    Retorna True si tuvo éxito. False si gh no está instalado o no está
    autenticado (el caller decidirá qué hacer).
    """
    import subprocess

    # 1. ¿gh está instalado?
    try:
        r = subprocess.run(['gh', '--version'], capture_output=True, text=True, timeout=10)
        if r.returncode != 0:
            raise FileNotFoundError
    except (FileNotFoundError, subprocess.TimeoutExpired):
        print(f"\n💡 GitHub CLI no está instalado en este equipo.")
        print(f"   Para publicación 100% automática (recomendado):")
        print(f"      1. Instale gh:  https://cli.github.com  (instalador para Windows)")
        print(f"      2. En la terminal:  gh auth login")
        print(f"      3. Ejecute de nuevo: python \"Plan Estrategico Dashboard.py\"")
        print(f"\n   Alternativa con Personal Access Token:")
        print(f"      python \"Plan Estrategico Dashboard.py\" --setup-github")
        return False

    # 2. ¿gh está autenticado?
    r = subprocess.run(['gh', 'auth', 'status'], capture_output=True, text=True)
    if r.returncode != 0:
        print(f"\n💡 GitHub CLI instalado pero no autenticado.")
        print(f"   Ejecute una vez:    gh auth login")
        print(f"   Después:             python \"Plan Estrategico Dashboard.py\"")
        return False

    # 3. Obtener el nombre de usuario
    r = subprocess.run(['gh', 'api', 'user', '-q', '.login'],
                       capture_output=True, text=True)
    username = r.stdout.strip()
    if not username:
        print(f"   ⚠️  No se pudo obtener el usuario de GitHub")
        return False

    # 4. Nombre del repo (= nombre de carpeta sanitizado)
    repo_name = git_root.name.lower()
    repo_name = ''.join(c if c.isalnum() or c in '-_' else '-' for c in repo_name)
    repo_name = repo_name.strip('-') or 'plan-estrategico'

    print(f"\n🔧 Configuración automática vía GitHub CLI")
    print(f"   Usuario: {username}")
    print(f"   Repo:    {repo_name}")

    # 5. Asegurar branch 'main'
    subprocess.run(['git', 'branch', '-M', 'main'],
                   capture_output=True, cwd=str(git_root))

    # 6. Asegurar al menos un commit (para que haya algo que pushear)
    r = subprocess.run(['git', 'log', '-1', '--oneline'],
                       capture_output=True, cwd=str(git_root))
    if r.returncode != 0:
        subprocess.run(['git', 'add', '.'], cwd=str(git_root))
        subprocess.run(['git', 'commit', '-m', 'Inicial: dashboard del Plan Estratégico'],
                       capture_output=True, cwd=str(git_root))
        print(f"   ✓ Commit inicial creado")

    # 7. Crear el repo Y hacer push (en un solo paso)
    print(f"   ⏳ Creando repo y subiendo archivos...")
    r = subprocess.run(
        ['gh', 'repo', 'create', repo_name,
         '--public', '--source=.', '--remote=origin', '--push'],
        capture_output=True, text=True, cwd=str(git_root),
    )
    if r.returncode != 0:
        err = (r.stderr or r.stdout).strip()
        if 'already exists' in err.lower() or 'name already exists' in err.lower():
            # El repo ya existe — solo conectar el remote y pushear
            print(f"   ℹ️  El repo {repo_name} ya existe. Conectando...")
            remote_url = f'https://github.com/{username}/{repo_name}.git'
            # Si origin ya existe, set-url; si no, add
            r_rem = subprocess.run(['git', 'remote'], capture_output=True, text=True, cwd=str(git_root))
            if 'origin' in r_rem.stdout.split():
                subprocess.run(['git', 'remote', 'set-url', 'origin', remote_url],
                               capture_output=True, cwd=str(git_root))
            else:
                subprocess.run(['git', 'remote', 'add', 'origin', remote_url],
                               capture_output=True, cwd=str(git_root))
            # Push
            r_push = subprocess.run(['git', 'push', '-u', 'origin', 'main'],
                                    capture_output=True, text=True, cwd=str(git_root))
            if r_push.returncode != 0:
                print(f"   ❌ Push falló: {(r_push.stderr or r_push.stdout).strip()}")
                return False
        else:
            print(f"   ❌ gh repo create falló:")
            for line in err.split('\n')[:6]:
                print(f"      {line}")
            return False
    print(f"   ✓ Repo creado y archivos subidos")

    # 8. Activar GitHub Pages (branch main, folder /docs)
    print(f"   ⏳ Activando GitHub Pages...")
    r = subprocess.run(
        ['gh', 'api', '--method', 'POST',
         f'/repos/{username}/{repo_name}/pages',
         '-f', 'source[branch]=main',
         '-f', 'source[path]=/docs'],
        capture_output=True, text=True,
    )
    if r.returncode == 0:
        print(f"   ✓ GitHub Pages activado (branch main, folder /docs)")
    else:
        err = (r.stderr or r.stdout).strip()
        if '409' in err or 'already' in err.lower():
            print(f"   ℹ️  GitHub Pages ya estaba activado")
        else:
            print(f"   ⚠️  No se pudo activar Pages automáticamente.")
            print(f"      Actívelo manualmente en:")
            print(f"      https://github.com/{username}/{repo_name}/settings/pages")

    pages_url = f"https://{username}.github.io/{repo_name}/"
    print(f"\n🎉 ¡Todo listo!")
    print(f"   📦 Repo:      https://github.com/{username}/{repo_name}")
    print(f"   🌐 Dashboard: {pages_url}")
    print(f"   ⏱️  La URL pública estará disponible en 1-2 minutos.")
    print(f"\n   Próximas veces ejecute solo:")
    print(f"      python \"Plan Estrategico Dashboard.py\"")
    return True


def _ensure_gitignore(git_root):
    """Crea o actualiza .gitignore para proteger archivos locales."""
    gi_path = git_root / '.gitignore'
    entries_needed = [
        '*.xlsx', '*.xlsm', '*.xls',
        '~$*',
        '__pycache__/', '*.pyc', '.env', 'venv/',
    ]
    existing = gi_path.read_text(encoding='utf-8') if gi_path.exists() else ''
    additions = [e for e in entries_needed if e not in existing]
    if additions:
        with open(gi_path, 'a', encoding='utf-8') as f:
            f.write('\n# Excel y temporales — gestionado por Plan_Estrategico_Dashboard.py\n')
            for e in additions:
                f.write(e + '\n')


def publish_to_github(out_path):
    """Publica SOLO el dashboard (docs/index.html) a GitHub.

    Flujo simple y seguro:
      1. git add docs/index.html   ← solo el dashboard, nunca el Excel
      2. git commit                ← si hubo cambios
      3. git push                  ← publicar
    No hace stash, no toca otros archivos, no interfiere con el Excel.
    """
    import subprocess

    git_root = _find_git_root(out_path.parent)
    if git_root is None:
        try:
            subprocess.run(['git', 'init'], capture_output=True, cwd=str(Path.cwd()))
            subprocess.run(['git', 'branch', '-M', 'main'],
                           capture_output=True, cwd=str(Path.cwd()))
            git_root = Path.cwd()
            print(f"\n   ✓ Repositorio Git inicializado en {git_root.name}/")
        except Exception as e:
            print(f"\n   ⚠️  No se pudo inicializar Git: {e}")
            return False

    # Proteger siempre el Excel y temporales
    _ensure_gitignore(git_root)

    try:
        subprocess.run(['git', '--version'], capture_output=True, check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("   ⚠️  Git no está instalado. Saltando publicación.")
        return False

    # Ruta relativa del dashboard dentro del repo
    try:
        rel_path = str(out_path.resolve().relative_to(git_root.resolve()))
    except ValueError:
        rel_path = str(out_path)

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
    print(f"\n📤 Publicando en GitHub...")

    # ── PASO 1: git add SOLO el dashboard ────────────────────────────
    r = subprocess.run(['git', 'add', rel_path],
                       capture_output=True, text=True, cwd=str(git_root))
    if r.returncode != 0:
        print(f"   ❌ git add falló: {r.stderr.strip() or r.stdout.strip()}")
        return False

    # ── PASO 2: commit si hay cambios ────────────────────────────────
    r = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=str(git_root))
    has_changes = (r.returncode != 0)

    if has_changes:
        msg = f"Dashboard actualizado — {timestamp}"
        r = subprocess.run(['git', 'commit', '-m', msg],
                           capture_output=True, text=True, cwd=str(git_root))
        if r.returncode != 0:
            print(f"   ❌ git commit falló: {r.stderr.strip() or r.stdout.strip()}")
            return False
        print(f"   ✓ Commit: {msg}")
    else:
        # Verificar si hay commits sin pushear aunque no haya nuevos cambios
        r = subprocess.run(['git', 'status', '-sb'], capture_output=True,
                           text=True, cwd=str(git_root))
        if 'ahead' not in r.stdout:
            print("   ℹ️  El dashboard no cambió. Nada que publicar.")
            return True

    # ── Verificar remote ─────────────────────────────────────────────
    r = subprocess.run(['git', 'remote'], capture_output=True, text=True, cwd=str(git_root))
    if not r.stdout.strip():
        return _try_gh_setup(git_root, out_path)

    # ── PASO 3: push directo — sin stash, sin pull, sin tocar archivos ─
    r = subprocess.run(['git', 'push'], capture_output=True, text=True, cwd=str(git_root))
    if r.returncode != 0:
        err = r.stderr.strip() or r.stdout.strip()
        # Si falla por divergencia, intentar con --force-with-lease (safe force)
        if 'fetch first' in err or 'rejected' in err or 'non-fast-forward' in err:
            print("   ⚠️  Repositorio divergente. Intentando push seguro...")
            r2 = subprocess.run(['git', 'push', '--force-with-lease'],
                                capture_output=True, text=True, cwd=str(git_root))
            if r2.returncode == 0:
                print("   ✓ Dashboard publicado a GitHub")
                print("   🌐 Disponible en GitHub Pages en unos segundos.")
                return True
        print(f"   ❌ git push falló:")
        for line in err.split('\n')[:6]:
            print(f"      {line}")
        print("   ℹ️  Dashboard generado en docs/index.html  ✓")
        print("   ℹ️  Para publicar manualmente:")
        print("      git add docs/index.html && git commit -m 'update' && git push")
        return False

    print("   ✓ Dashboard publicado a GitHub")
    print("   🌐 Disponible en GitHub Pages en unos segundos.")
    return True


def _interactive_setup_github(git_root, out_path):
    """Pide token y nombre de repo interactivamente y llama a setup_github."""
    import getpass
    import os

    print("\n╔══════════════════════════════════════════════════════════════╗")
    print("║   CONFIGURACIÓN INICIAL DE GITHUB (sólo una vez)             ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print()
    print("Necesita un Personal Access Token (PAT) de GitHub.")
    print()
    print("Si no tiene uno, créelo aquí (toma 30 segundos):")
    print("  → https://github.com/settings/tokens/new")
    print()
    print("  Configure así:")
    print("    Note:        Plan Estratégico Dashboard")
    print("    Expiration:  No expiration (o 90 días)")
    print("    Scopes:      marque la casilla 'repo' (todo el grupo)")
    print()
    print("  Clic 'Generate token' y copie el token (empieza con 'ghp_').")
    print()

    # Token: variable de entorno > prompt
    token = os.environ.get('GITHUB_TOKEN', '').strip()
    if token:
        print(f"✓ Usando token de la variable de entorno GITHUB_TOKEN")
    else:
        try:
            token = getpass.getpass("Pegue su Personal Access Token (no se mostrará): ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nCancelado.")
            return None
    if not token:
        print("❌ No se ingresó token. Cancelado.")
        return None

    # Nombre del repo: default = nombre de la carpeta sanitizado
    default_repo = git_root.name.lower()
    default_repo = ''.join(c if c.isalnum() or c in '-_' else '-' for c in default_repo)
    default_repo = default_repo.strip('-') or 'plan-estrategico'
    try:
        repo_name = input(f"Nombre del repo (Enter para usar '{default_repo}'): ").strip()
    except (KeyboardInterrupt, EOFError):
        print("\nCancelado.")
        return None
    if not repo_name:
        repo_name = default_repo

    return setup_github(token, repo_name, git_root, out_path)


# ============================================================
# MAIN
# ============================================================

def _audit_vs_excel(indicators, xlsx_path):
    """Auditoría automática: compara Python vs celdas calculadas del Excel.
    Se ejecuta cada vez que se genera el dashboard.
    """
    try:
        from openpyxl import load_workbook as _lw
        wb_t = _lw(str(xlsx_path), data_only=True)
    except Exception as e:
        print(f"   ⚠️  Auditoría no ejecutada: {e}")
        return

    errores = []; total_ok = 0; total_rows = 0

    for ind in indicators:
        sn = ind['code']
        if sn not in wb_t.sheetnames:
            continue
        ws = wb_t[sn]
        hdrs = {c: str(ws.cell(22, c).value or '').strip().lower() for c in range(2, 16)}
        res_col   = next((c for c, h in hdrs.items() if h == 'resultado'), None)
        cumpl_col = next((c for c, h in hdrs.items()
                          if h in ('% cumpl.', '% cumpl', 'cumpl.',
                                   '% cumplimiento', '% de cumplimiento')), None)
        if not res_col:
            continue

        for r in range(23, 83):
            if not ws.cell(r, 4).value:
                break
            res_xl = ws.cell(r, res_col).value
            if not isinstance(res_xl, (int, float)):
                continue

            m_idx    = r - 23
            cumpl_xl = ws.cell(r, cumpl_col).value if cumpl_col else None
            res_py   = ind['monthly_results'][m_idx]                        if m_idx < len(ind['monthly_results']) else None
            yr_i     = m_idx // 12
            meta     = ind['metas_raw'][yr_i] if yr_i < len(ind['metas_raw']) else None
            mp       = ind.get('meta_periodos', [])[m_idx]                        if m_idx < len(ind.get('meta_periodos', [])) else None
            meta_use = mp if mp is not None else meta
            cumpl_py = compute_compliance(
                res_py, meta_use, ind['tend'],
                tipo=ind.get('cumpl_formula_type')
            ) if res_py is not None and meta_use is not None else None

            total_rows += 1
            d_res   = abs(res_py - res_xl)                       if res_py is not None and res_xl is not None else None
            d_cumpl = abs(cumpl_py - cumpl_xl) * 100                       if cumpl_py is not None and cumpl_xl is not None else None

            ok = (d_res is None or d_res < 1e-4) and (d_cumpl is None or d_cumpl < 0.05)
            if ok:
                total_ok += 1
            else:
                errores.append((sn, str(ws.cell(r, 4).value),
                                 d_res, d_cumpl, res_py, res_xl, cumpl_py, cumpl_xl))

    if not errores:
        print(f"   ✅ Auditoría Excel: {total_ok}/{total_rows} filas — "
              f"modelo 100% consistente con Excel")
    else:
        print(f"   ❌ Auditoría Excel: {total_ok}/{total_rows} correctas — "
              f"{len(errores)} diferencia(s):")
        print(f"   {'IND':<8} {'Mes':<12} {'Δ Resultado':>14} {'Δ Cumpl (pp)':>14}")
        print(f"   {'─'*52}")
        for sn, mes, dr, dc, rp, rx, cp, cx in errores[:20]:
            dr_s = f"{dr:.6f}" if dr is not None else f"Py=— XL={rx:.4f}" if rx else "—"
            dc_s = f"{dc:.4f}" if dc is not None else f"Py=— XL={cx:.4f}" if cx else "—"
            print(f"   {sn:<8} {mes:<12} {dr_s:>14} {dc_s:>14} pp")
        if len(errores) > 20:
            print(f"   ... y {len(errores)-20} diferencia(s) más")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = [a for a in sys.argv[1:] if a.startswith('--')]
    no_publish = '--no-publish' in flags
    do_setup = '--setup-github' in flags

    xlsx_path = Path(args[0]) if args else Path('Tablero_Indicadores_PE.xlsx')
    out_path = Path(args[1]) if len(args) > 1 else Path('docs/index.html')

    if not xlsx_path.exists():
        print(f"❌ ERROR: archivo no encontrado: {xlsx_path}")
        sys.exit(1)

    print(f"📖 Leyendo  {xlsx_path}")
    indicators = process_workbook(xlsx_path)
    print(f"   ✓ {len(indicators)} indicadores procesados")

    total_warnings = sum(len(i['warnings']) for i in indicators)
    if total_warnings:
        print(f"   ⚠️  {total_warnings} avisos de validación (visibles en el dashboard)")

    # Auditoría automática: comparar Python vs Excel
    _audit_vs_excel(indicators, xlsx_path)

    # Abrir wb_src para copiar hojas con fidelidad total
    try:
        wb_src_copy = load_workbook(str(xlsx_path)) if xlsx_path else None
    except Exception:
        wb_src_copy = None
    payload = build_payload(indicators, wb_src=wb_src_copy)
    print(f"   ✓ Resumen: {payload['summary']}")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    html = render_dashboard_html(payload, xlsx_path=xlsx_path)
    out_path.write_text(html, encoding='utf-8')

    size_kb = out_path.stat().st_size / 1024
    print(f"✅ Dashboard generado: {out_path}  ({size_kb:.1f} KB)")

    # Setup inicial de GitHub si se solicita
    if do_setup:
        import subprocess
        git_root = _find_git_root(out_path.parent)
        if git_root is None:
            # Inicializar git en la carpeta actual
            print(f"\n🔧 Inicializando repositorio Git en {Path.cwd()}...")
            subprocess.run(['git', 'init'], capture_output=True, cwd=str(Path.cwd()))
            subprocess.run(['git', 'branch', '-M', 'main'], capture_output=True, cwd=str(Path.cwd()))
            git_root = Path.cwd()
            # Commit inicial con todo
            subprocess.run(['git', 'add', '.'], cwd=str(git_root))
            subprocess.run(['git', 'commit', '-m', 'Inicial'], capture_output=True, cwd=str(git_root))
        _interactive_setup_github(git_root, out_path)
        return

    # Auto-publicar a GitHub (a menos que se use --no-publish)
    if not no_publish:
        publish_to_github(out_path)
    else:
        print(f"\n   (Auto-publicación desactivada con --no-publish)")



if __name__ == '__main__':
    main()
