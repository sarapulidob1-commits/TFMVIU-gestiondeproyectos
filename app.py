import streamlit as st
import pandas as pd
import os
import re
from datetime import date, timedelta, datetime

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
ARCHIVO_EXCEL = "Base_Datos_TFM.xlsx"

COLS_PROYECTOS = [
    "ID_Proyecto", "Nombre_Proyecto", "Descripcion", "Cliente_Area", "Tecnologia",
    "Lider_Proyecto", "Fecha_Inicio", "Fecha_Fin_Planificada", "Fecha_Fin_Real",
    "Presupuesto_Planificado", "Presupuesto_Ejecutado", "Beneficio_Esperado",
    "Avance_Pct", "Estado_Proyecto", "Prioridad"
]

DATOS_FICTICIOS_PROYECTOS = [
    {
        "ID_Proyecto": "PRJ-001", "Nombre_Proyecto": "Implementación Power BI",
        "Cliente_Area": "Finanzas", "Lider_Proyecto": "Ana Gómez",
        "Fecha_Inicio": "2026-03-02", "Fecha_Fin_Planificada": "2026-09-30",
        "Presupuesto_Planificado": 25000, "Presupuesto_Ejecutado": 14200,
        "Avance_Pct": 55, "Estado_Proyecto": "En curso", "Prioridad": "Alta"
    },
    {
        "ID_Proyecto": "PRJ-002", "Nombre_Proyecto": "Automatización de Facturas",
        "Cliente_Area": "Operaciones", "Lider_Proyecto": "Carlos Ruiz",
        "Fecha_Inicio": "2026-08-03", "Fecha_Fin_Planificada": "2026-11-16",
        "Presupuesto_Planificado": 18000, "Presupuesto_Ejecutado": 0,
        "Avance_Pct": 0, "Estado_Proyecto": "No iniciado", "Prioridad": "Media"
    },
    {
        "ID_Proyecto": "PRJ-003", "Nombre_Proyecto": "Migración CRM a la Nube",
        "Cliente_Area": "Comercial", "Lider_Proyecto": "Luisa Mendoza",
        "Fecha_Inicio": "2026-01-12", "Fecha_Fin_Planificada": "2026-07-15",
        "Presupuesto_Planificado": 42000, "Presupuesto_Ejecutado": 46800,
        "Avance_Pct": 70, "Estado_Proyecto": "En riesgo", "Prioridad": "Alta"
    },
    {
        "ID_Proyecto": "PRJ-004", "Nombre_Proyecto": "Portal de Autoservicio RRHH",
        "Cliente_Area": "Talento Humano", "Lider_Proyecto": "Jorge Castaño",
        "Fecha_Inicio": "2026-04-06", "Fecha_Fin_Planificada": "2026-10-30",
        "Presupuesto_Planificado": 15500, "Presupuesto_Ejecutado": 6100,
        "Avance_Pct": 40, "Estado_Proyecto": "En curso", "Prioridad": "Media"
    },
    {
        "ID_Proyecto": "PRJ-005", "Nombre_Proyecto": "Capacitación en Analítica de Datos",
        "Cliente_Area": "Toda la empresa", "Lider_Proyecto": "Ana Gómez",
        "Fecha_Inicio": "2026-02-02", "Fecha_Fin_Planificada": "2026-05-29",
        "Presupuesto_Planificado": 8000, "Presupuesto_Ejecutado": 7650,
        "Avance_Pct": 100, "Estado_Proyecto": "Finalizado", "Prioridad": "Baja"
    },
    {
        "ID_Proyecto": "PRJ-006", "Nombre_Proyecto": "Rediseño de Marca Corporativa",
        "Cliente_Area": "Marketing", "Lider_Proyecto": "Camila Torres",
        "Fecha_Inicio": "2025-09-01", "Fecha_Fin_Planificada": "2026-01-30",
        "Presupuesto_Planificado": 12000, "Presupuesto_Ejecutado": 12500,
        "Avance_Pct": 100, "Estado_Proyecto": "Finalizado", "Prioridad": "Media"
    },
    {
        "ID_Proyecto": "PRJ-007", "Nombre_Proyecto": "Implementación ERP Logístico",
        "Cliente_Area": "Logística", "Lider_Proyecto": "Daniel Rojas",
        "Fecha_Inicio": "2026-05-10", "Fecha_Fin_Planificada": "2027-01-20",
        "Presupuesto_Planificado": 60000, "Presupuesto_Ejecutado": 22000,
        "Avance_Pct": 30, "Estado_Proyecto": "En curso", "Prioridad": "Alta"
    },
    {
        "ID_Proyecto": "PRJ-008", "Nombre_Proyecto": "Optimización de Infraestructura TI",
        "Cliente_Area": "TI", "Lider_Proyecto": "Paula Sánchez",
        "Fecha_Inicio": "2026-06-01", "Fecha_Fin_Planificada": "2026-08-15",
        "Presupuesto_Planificado": 20000, "Presupuesto_Ejecutado": 21500,
        "Avance_Pct": 85, "Estado_Proyecto": "En riesgo", "Prioridad": "Alta"
    },
    {
        "ID_Proyecto": "PRJ-009", "Nombre_Proyecto": "Campaña Digital Q4",
        "Cliente_Area": "Marketing", "Lider_Proyecto": "Camila Torres",
        "Fecha_Inicio": "2026-09-01", "Fecha_Fin_Planificada": "2026-12-15",
        "Presupuesto_Planificado": 9500, "Presupuesto_Ejecutado": 0,
        "Avance_Pct": 0, "Estado_Proyecto": "No iniciado", "Prioridad": "Baja"
    },
    {
        "ID_Proyecto": "PRJ-010", "Nombre_Proyecto": "Auditoría de Procesos Internos",
        "Cliente_Area": "Operaciones", "Lider_Proyecto": "Carlos Ruiz",
        "Fecha_Inicio": "2025-11-03", "Fecha_Fin_Planificada": "2026-02-28",
        "Presupuesto_Planificado": 7000, "Presupuesto_Ejecutado": 6800,
        "Avance_Pct": 100, "Estado_Proyecto": "Finalizado", "Prioridad": "Media"
    },
    {
        "ID_Proyecto": "PRJ-011", "Nombre_Proyecto": "Plataforma de Autoservicio Comercial",
        "Cliente_Area": "Comercial", "Lider_Proyecto": "Luisa Mendoza",
        "Fecha_Inicio": "2026-07-01", "Fecha_Fin_Planificada": "2027-02-28",
        "Presupuesto_Planificado": 35000, "Presupuesto_Ejecutado": 9000,
        "Avance_Pct": 15, "Estado_Proyecto": "En curso", "Prioridad": "Alta"
    },
    {
        "ID_Proyecto": "PRJ-012", "Nombre_Proyecto": "Programa de Bienestar Laboral",
        "Cliente_Area": "Talento Humano", "Lider_Proyecto": "Jorge Castaño",
        "Fecha_Inicio": "2026-03-15", "Fecha_Fin_Planificada": "2026-06-30",
        "Presupuesto_Planificado": 6000, "Presupuesto_Ejecutado": 5900,
        "Avance_Pct": 100, "Estado_Proyecto": "Finalizado", "Prioridad": "Baja"
    },
]

COLS_COLABORADORES = [
    "ID_Colaborador", "Nombre", "Rol", "Area", "Email",
    "Fecha_Ingreso", "Capacidad_Horas_Semana", "Costo_Hora", "Proyectos_Asignados", "Estado"
]

DATOS_FICTICIOS_COLABORADORES = [
    {
        "ID_Colaborador": "COL-001", "Nombre": "Ana Gómez", "Rol": "Líder de Proyecto",
        "Area": "Finanzas", "Email": "ana.gomez@empresa.com", "Fecha_Ingreso": "2025-01-15",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-001,PRJ-005", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-002", "Nombre": "Carlos Ruiz", "Rol": "Analista de Datos",
        "Area": "Operaciones", "Email": "carlos.ruiz@empresa.com", "Fecha_Ingreso": "2025-03-01",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-002,PRJ-010", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-003", "Nombre": "Luisa Mendoza", "Rol": "Consultora BI",
        "Area": "Comercial", "Email": "luisa.mendoza@empresa.com", "Fecha_Ingreso": "2024-11-10",
        "Capacidad_Horas_Semana": 35, "Proyectos_Asignados": "PRJ-003,PRJ-011", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-004", "Nombre": "Jorge Castaño", "Rol": "Desarrollador BI",
        "Area": "Talento Humano", "Email": "jorge.castano@empresa.com", "Fecha_Ingreso": "2025-02-20",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-001,PRJ-003,PRJ-004,PRJ-012", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-005", "Nombre": "Marta Ortiz", "Rol": "Analista Junior",
        "Area": "Finanzas", "Email": "marta.ortiz@empresa.com", "Fecha_Ingreso": "2025-05-05",
        "Capacidad_Horas_Semana": 20, "Proyectos_Asignados": "", "Estado": "Inactivo",
    },
    {
        "ID_Colaborador": "COL-006", "Nombre": "Camila Torres", "Rol": "Líder de Proyecto",
        "Area": "Marketing", "Email": "camila.torres@empresa.com", "Fecha_Ingreso": "2024-06-15",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-006,PRJ-009", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-007", "Nombre": "Daniel Rojas", "Rol": "Líder de Proyecto",
        "Area": "Logística", "Email": "daniel.rojas@empresa.com", "Fecha_Ingreso": "2023-09-01",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-007", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-008", "Nombre": "Paula Sánchez", "Rol": "Arquitecta de Infraestructura",
        "Area": "TI", "Email": "paula.sanchez@empresa.com", "Fecha_Ingreso": "2024-02-10",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-008", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-009", "Nombre": "Andrés Vargas", "Rol": "Desarrollador BI",
        "Area": "TI", "Email": "andres.vargas@empresa.com", "Fecha_Ingreso": "2025-04-01",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-007,PRJ-008", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-010", "Nombre": "Valentina Rey", "Rol": "Analista Financiera",
        "Area": "Finanzas", "Email": "valentina.rey@empresa.com", "Fecha_Ingreso": "2024-08-20",
        "Capacidad_Horas_Semana": 30, "Proyectos_Asignados": "PRJ-001", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-011", "Nombre": "Felipe Muñoz", "Rol": "Consultor de Procesos",
        "Area": "Operaciones", "Email": "felipe.munoz@empresa.com", "Fecha_Ingreso": "2023-11-05",
        "Capacidad_Horas_Semana": 40, "Proyectos_Asignados": "PRJ-002,PRJ-010,PRJ-011", "Estado": "Activo",
    },
    {
        "ID_Colaborador": "COL-012", "Nombre": "Isabella Duarte", "Rol": "Practicante de Datos",
        "Area": "Comercial", "Email": "isabella.duarte@empresa.com", "Fecha_Ingreso": "2026-01-10",
        "Capacidad_Horas_Semana": 15, "Proyectos_Asignados": "", "Estado": "Activo",
    },
]

COLS_SALUD = [
    "ID_Registro", "Semana", "ID_Colaborador", "Nombre_Colaborador",
    "ID_Proyecto", "Nombre_Proyecto", "Horas_Asignadas", "Horas_Reales",
    "Carga_Percibida", "Capacidad_Horas_Semana", "Comentario"
]

# Registros de muestra: (semana, colaborador, proyecto, asignadas, reales, percibida, comentario)
# 4 semanas de historia con casos saludables, en precaución, críticos y subutilizados.
_REGISTROS_SALUD_BASE = [
    # Ana Gómez (40 h) — saludable estable
    ("2026-06-15", "COL-001", "PRJ-001", 32, 30, 3, ""),
    ("2026-06-22", "COL-001", "PRJ-001", 32, 33, 3, ""),
    ("2026-06-29", "COL-001", "PRJ-001", 32, 34, 3, "Cierre de sprint"),
    ("2026-07-06", "COL-001", "PRJ-001", 32, 32, 3, ""),
    # Carlos Ruiz (40 h) — subutilizado (proyecto aún no inicia)
    ("2026-06-15", "COL-002", "PRJ-002", 20, 18, 2, "Preparación del proyecto"),
    ("2026-06-22", "COL-002", "PRJ-002", 20, 21, 2, ""),
    ("2026-06-29", "COL-002", "PRJ-002", 20, 22, 2, ""),
    ("2026-07-06", "COL-002", "PRJ-002", 20, 19, 2, ""),
    # Luisa Mendoza (35 h) — escalando de saludable a crítico
    ("2026-06-15", "COL-003", "PRJ-003", 30, 31, 4, ""),
    ("2026-06-22", "COL-003", "PRJ-003", 30, 33, 4, "Retrasos por dependencias"),
    ("2026-06-29", "COL-003", "PRJ-003", 32, 34, 5, "Semana crítica de migración"),
    ("2026-07-06", "COL-003", "PRJ-003", 25, 26, 5, ""),
    ("2026-07-06", "COL-003", "PRJ-011", 10, 9, 4, "Arranque del nuevo proyecto"),
    # Jorge Castaño (40 h) — crítico sostenido, 4 proyectos
    ("2026-06-15", "COL-004", "PRJ-001", 12, 13, 5, ""),
    ("2026-06-15", "COL-004", "PRJ-003", 10, 11, 5, ""),
    ("2026-06-15", "COL-004", "PRJ-004", 14, 13, 4, ""),
    ("2026-06-15", "COL-004", "PRJ-012", 6, 6, 3, ""),
    ("2026-06-22", "COL-004", "PRJ-001", 12, 14, 5, ""),
    ("2026-06-22", "COL-004", "PRJ-003", 10, 12, 5, "Apoyo extra por riesgo"),
    ("2026-06-22", "COL-004", "PRJ-004", 14, 14, 5, ""),
    ("2026-06-22", "COL-004", "PRJ-012", 6, 5, 3, ""),
    ("2026-06-29", "COL-004", "PRJ-001", 12, 13, 5, ""),
    ("2026-06-29", "COL-004", "PRJ-003", 10, 11, 5, ""),
    ("2026-06-29", "COL-004", "PRJ-004", 14, 15, 5, ""),
    ("2026-06-29", "COL-004", "PRJ-012", 6, 4, 3, "Cierre del programa"),
    ("2026-07-06", "COL-004", "PRJ-001", 14, 15, 5, ""),
    ("2026-07-06", "COL-004", "PRJ-003", 12, 13, 5, ""),
    ("2026-07-06", "COL-004", "PRJ-004", 14, 16, 5, ""),
    # Camila Torres (40 h) — subutilizada (campaña en planeación)
    ("2026-06-15", "COL-006", "PRJ-009", 15, 12, 2, "Planeación de campaña"),
    ("2026-06-22", "COL-006", "PRJ-009", 15, 14, 2, ""),
    ("2026-06-29", "COL-006", "PRJ-009", 15, 16, 2, ""),
    ("2026-07-06", "COL-006", "PRJ-009", 15, 15, 2, ""),
    # Daniel Rojas (40 h) — saludable, un pico de precaución
    ("2026-06-15", "COL-007", "PRJ-007", 34, 33, 3, ""),
    ("2026-06-22", "COL-007", "PRJ-007", 34, 35, 3, ""),
    ("2026-06-29", "COL-007", "PRJ-007", 34, 36, 4, "Semana de integración"),
    ("2026-07-06", "COL-007", "PRJ-007", 34, 34, 3, ""),
    # Paula Sánchez (40 h) — escalando hacia crítico (proyecto en riesgo)
    ("2026-06-15", "COL-008", "PRJ-008", 36, 36, 4, ""),
    ("2026-06-22", "COL-008", "PRJ-008", 36, 38, 4, ""),
    ("2026-06-29", "COL-008", "PRJ-008", 36, 39, 5, "Incidentes en producción"),
    ("2026-07-06", "COL-008", "PRJ-008", 38, 41, 5, "Sobrecarga por estabilización"),
    # Andrés Vargas (40 h) — bordeando precaución/crítico
    ("2026-06-15", "COL-009", "PRJ-007", 20, 19, 3, ""),
    ("2026-06-15", "COL-009", "PRJ-008", 18, 18, 4, ""),
    ("2026-06-22", "COL-009", "PRJ-007", 20, 21, 4, ""),
    ("2026-06-22", "COL-009", "PRJ-008", 18, 17, 4, ""),
    ("2026-06-29", "COL-009", "PRJ-007", 20, 20, 4, ""),
    ("2026-06-29", "COL-009", "PRJ-008", 18, 18, 4, ""),
    ("2026-07-06", "COL-009", "PRJ-007", 20, 19, 4, ""),
    ("2026-07-06", "COL-009", "PRJ-008", 18, 17, 4, ""),
    # Valentina Rey (30 h) — saludable
    ("2026-06-15", "COL-010", "PRJ-001", 22, 20, 2, ""),
    ("2026-06-22", "COL-010", "PRJ-001", 22, 22, 3, ""),
    ("2026-06-29", "COL-010", "PRJ-001", 22, 23, 3, ""),
    ("2026-07-06", "COL-010", "PRJ-001", 22, 21, 2, ""),
    # Felipe Muñoz (40 h) — en precaución/crítico con 2 frentes
    ("2026-06-15", "COL-011", "PRJ-002", 16, 17, 4, ""),
    ("2026-06-15", "COL-011", "PRJ-011", 20, 20, 4, "Preparación de kickoff"),
    ("2026-06-22", "COL-011", "PRJ-002", 16, 16, 4, ""),
    ("2026-06-22", "COL-011", "PRJ-011", 20, 22, 4, ""),
    ("2026-06-29", "COL-011", "PRJ-002", 16, 18, 5, ""),
    ("2026-06-29", "COL-011", "PRJ-011", 22, 21, 4, ""),
    ("2026-07-06", "COL-011", "PRJ-002", 14, 15, 4, ""),
    ("2026-07-06", "COL-011", "PRJ-011", 24, 22, 5, "Inicio oficial del proyecto"),
]

_MAPA_NOMBRE_COL_SEED = {c["ID_Colaborador"]: c["Nombre"] for c in DATOS_FICTICIOS_COLABORADORES}
_MAPA_CAPACIDAD_SEED = {c["ID_Colaborador"]: c["Capacidad_Horas_Semana"] for c in DATOS_FICTICIOS_COLABORADORES}
_MAPA_NOMBRE_PROY_SEED = {p["ID_Proyecto"]: p["Nombre_Proyecto"] for p in DATOS_FICTICIOS_PROYECTOS}

DATOS_FICTICIOS_SALUD = [
    {
        "ID_Registro": f"SAL-{i:03d}",
        "Semana": semana,
        "ID_Colaborador": id_col,
        "Nombre_Colaborador": _MAPA_NOMBRE_COL_SEED[id_col],
        "ID_Proyecto": id_proy,
        "Nombre_Proyecto": _MAPA_NOMBRE_PROY_SEED[id_proy],
        "Horas_Asignadas": asignadas,
        "Horas_Reales": reales,
        "Carga_Percibida": percibida,
        "Capacidad_Horas_Semana": _MAPA_CAPACIDAD_SEED[id_col],
        "Comentario": comentario,
    }
    for i, (semana, id_col, id_proy, asignadas, reales, percibida, comentario)
    in enumerate(_REGISTROS_SALUD_BASE, start=1)
]

# --- Enriquecimiento de los datos de muestra de Proyectos y Colaboradores ---
_DESCRIPCIONES_PROY = {
    "PRJ-001": "Tableros financieros en Power BI para la dirección.",
    "PRJ-002": "Automatizar el registro y aprobación de facturas de proveedores.",
    "PRJ-003": "Migrar el CRM local a la nube y capacitar al equipo comercial.",
    "PRJ-004": "Portal para que los empleados gestionen sus trámites de RRHH.",
    "PRJ-005": "Programa de formación en analítica para toda la empresa.",
    "PRJ-006": "Actualización de la identidad visual y aplicaciones de marca.",
    "PRJ-007": "Implementación del ERP para la operación logística nacional.",
    "PRJ-008": "Renovación de servidores y migración a infraestructura híbrida.",
    "PRJ-009": "Campaña digital de fin de año para captación de clientes.",
    "PRJ-010": "Auditoría interna de procesos operativos y plan de mejoras.",
    "PRJ-011": "Plataforma e-commerce B2B de autoservicio para clientes.",
    "PRJ-012": "Programa de bienestar y clima laboral para el equipo.",
}
_FIN_REAL_PROY = {
    "PRJ-005": "2026-06-05",   # terminó tarde
    "PRJ-006": "2026-01-28",   # terminó antes
    "PRJ-010": "2026-03-10",   # terminó tarde
    "PRJ-012": "2026-06-30",   # terminó a tiempo
}
_TECNOLOGIA_PROY = {
    "PRJ-001": "Analítica de datos", "PRJ-002": "Automatización",
    "PRJ-003": "Nube / CRM", "PRJ-004": "Desarrollo web",
    "PRJ-005": "Formación", "PRJ-006": "Diseño / Marca",
    "PRJ-007": "ERP", "PRJ-008": "Infraestructura TI",
    "PRJ-009": "Marketing digital", "PRJ-010": "Procesos",
    "PRJ-011": "E-commerce", "PRJ-012": "Talento humano",
}
# Beneficio esperado ($): permite calcular el ROI del portafolio en Power BI
# (ROI = (Beneficio_Esperado - Presupuesto) / Presupuesto).
_BENEFICIO_PROY = {
    "PRJ-001": 60000, "PRJ-002": 30000, "PRJ-003": 80000, "PRJ-004": 25000,
    "PRJ-005": 15000, "PRJ-006": 20000, "PRJ-007": 120000, "PRJ-008": 35000,
    "PRJ-009": 22000, "PRJ-010": 12000, "PRJ-011": 70000, "PRJ-012": 10000,
}
for _p in DATOS_FICTICIOS_PROYECTOS:
    _p["Descripcion"] = _DESCRIPCIONES_PROY.get(_p["ID_Proyecto"], "")
    _p["Fecha_Fin_Real"] = _FIN_REAL_PROY.get(_p["ID_Proyecto"], "")
    _p["Tecnologia"] = _TECNOLOGIA_PROY.get(_p["ID_Proyecto"], "")
    _p["Beneficio_Esperado"] = _BENEFICIO_PROY.get(_p["ID_Proyecto"], 0)

_COSTO_HORA = {
    "COL-001": 45, "COL-002": 35, "COL-003": 40, "COL-004": 38,
    "COL-005": 18, "COL-006": 42, "COL-007": 45, "COL-008": 50,
    "COL-009": 32, "COL-010": 28, "COL-011": 36, "COL-012": 12,
}
for _c in DATOS_FICTICIOS_COLABORADORES:
    _c["Costo_Hora"] = _COSTO_HORA.get(_c["ID_Colaborador"], 30)

# --- Módulo 4: Presupuesto detallado ---
COLS_PRESUPUESTO = [
    "ID_Gasto", "ID_Proyecto", "Nombre_Proyecto", "Fecha",
    "Categoria", "Concepto", "Tipo", "Monto"
]

CATEGORIAS_GASTO = [
    "Personal", "Software / Licencias", "Hardware", "Consultoría",
    "Capacitación", "Viáticos", "Otros"
]

# (proyecto, fecha, categoría, concepto, tipo, monto)
# Modelo de partidas: las filas "Planificado" son las partidas presupuestales por
# rubro de cada proyecto; las filas "Real" son los gastos ejecutados contra esas
# partidas. Las sumas por proyecto cuadran con Presupuesto_Planificado y
# Presupuesto_Ejecutado de la hoja Proyectos, e incluyen rubros excedidos,
# al límite y sin ejecutar para demostrar el control plan vs. real.
_GASTOS_BASE = [
    # PRJ-001 · Implementación Power BI — en curso (plan 25,000 / real 14,200)
    ("PRJ-001", "2026-03-02", "Personal", "Horas del equipo BI", "Planificado", 9000),
    ("PRJ-001", "2026-03-02", "Personal", "Horas de apoyo del área financiera", "Planificado", 3000),
    ("PRJ-001", "2026-03-02", "Software / Licencias", "Licencias Power BI Pro", "Planificado", 5000),
    ("PRJ-001", "2026-03-02", "Software / Licencias", "Conectores y componentes premium", "Planificado", 1000),
    ("PRJ-001", "2026-03-02", "Capacitación", "Formación de usuarios finales", "Planificado", 4000),
    ("PRJ-001", "2026-03-02", "Consultoría", "Acompañamiento experto en modelado", "Planificado", 2500),
    ("PRJ-001", "2026-03-02", "Viáticos", "Visitas a las áreas usuarias", "Planificado", 500),
    ("PRJ-001", "2026-03-15", "Software / Licencias", "Licencias Power BI Pro (compra anual)", "Real", 5000),
    ("PRJ-001", "2026-03-20", "Software / Licencias", "Conector premium de datos", "Real", 500),
    ("PRJ-001", "2026-04-12", "Software / Licencias", "Componente de visualización", "Real", 300),
    ("PRJ-001", "2026-03-15", "Personal", "Horas del equipo — 1.ª quincena de marzo", "Real", 1000),
    ("PRJ-001", "2026-03-31", "Personal", "Horas del equipo — 2.ª quincena de marzo", "Real", 1050),
    ("PRJ-001", "2026-04-15", "Personal", "Horas del equipo — 1.ª quincena de abril", "Real", 1000),
    ("PRJ-001", "2026-04-30", "Personal", "Horas del equipo — 2.ª quincena de abril", "Real", 1100),
    ("PRJ-001", "2026-05-15", "Personal", "Horas del equipo — 1.ª quincena de mayo", "Real", 1050),
    ("PRJ-001", "2026-05-31", "Personal", "Horas del equipo — 2.ª quincena de mayo", "Real", 1050),
    ("PRJ-001", "2026-06-15", "Personal", "Horas del equipo — 1.ª quincena de junio", "Real", 1100),
    ("PRJ-001", "2026-06-30", "Personal", "Horas del equipo — 2.ª quincena de junio", "Real", 1050),
    # PRJ-002 · Automatización de Facturas — no iniciado (solo partidas)
    ("PRJ-002", "2026-08-03", "Software / Licencias", "Plataforma de automatización de facturas", "Planificado", 8500),
    ("PRJ-002", "2026-08-03", "Software / Licencias", "Licencias de lectura automática (OCR)", "Planificado", 1500),
    ("PRJ-002", "2026-08-03", "Personal", "Horas del equipo de operaciones", "Planificado", 6000),
    ("PRJ-002", "2026-08-03", "Capacitación", "Entrenamiento en la herramienta", "Planificado", 1000),
    ("PRJ-002", "2026-08-03", "Capacitación", "Material de formación", "Planificado", 500),
    ("PRJ-002", "2026-08-03", "Otros", "Gestión del cambio con proveedores", "Planificado", 500),
    # PRJ-003 · Migración CRM — en riesgo (plan 42,000 / real 46,800, varios rubros excedidos)
    ("PRJ-003", "2026-01-12", "Consultoría", "Consultoría de migración — fases 1 y 2", "Planificado", 12000),
    ("PRJ-003", "2026-01-12", "Consultoría", "Consultoría de migración — fase 3", "Planificado", 8000),
    ("PRJ-003", "2026-01-12", "Software / Licencias", "Suscripción CRM en la nube", "Planificado", 13000),
    ("PRJ-003", "2026-01-12", "Software / Licencias", "Módulos adicionales del CRM", "Planificado", 2000),
    ("PRJ-003", "2026-01-12", "Capacitación", "Talleres del equipo comercial", "Planificado", 3000),
    ("PRJ-003", "2026-01-12", "Capacitación", "Material y plataforma de formación", "Planificado", 2000),
    ("PRJ-003", "2026-01-12", "Viáticos", "Visitas a sedes regionales", "Planificado", 1500),
    ("PRJ-003", "2026-01-12", "Otros", "Imprevistos de la migración", "Planificado", 500),
    ("PRJ-003", "2026-01-31", "Consultoría", "Consultoría — arranque y diagnóstico", "Real", 2800),
    ("PRJ-003", "2026-02-15", "Consultoría", "Consultoría — diseño de la migración", "Real", 3000),
    ("PRJ-003", "2026-02-28", "Consultoría", "Consultoría — preparación de datos", "Real", 3200),
    ("PRJ-003", "2026-03-31", "Consultoría", "Consultoría — migración piloto", "Real", 3100),
    ("PRJ-003", "2026-04-15", "Consultoría", "Consultoría — correcciones del piloto", "Real", 3000),
    ("PRJ-003", "2026-04-30", "Consultoría", "Consultoría — migración masiva", "Real", 3200),
    ("PRJ-003", "2026-05-31", "Consultoría", "Consultoría — validación de datos", "Real", 3100),
    ("PRJ-003", "2026-06-20", "Consultoría", "Consultoría — soporte a pruebas", "Real", 3100),
    ("PRJ-003", "2026-04-01", "Software / Licencias", "Suscripción CRM (anualidad)", "Real", 13500),
    ("PRJ-003", "2026-05-10", "Software / Licencias", "Módulos adicionales del CRM", "Real", 1500),
    ("PRJ-003", "2026-06-05", "Software / Licencias", "Almacenamiento extra en la nube", "Real", 1200),
    ("PRJ-003", "2026-06-25", "Software / Licencias", "Soporte premium del proveedor", "Real", 900),
    ("PRJ-003", "2026-05-08", "Capacitación", "Taller regional — zona norte", "Real", 1000),
    ("PRJ-003", "2026-05-20", "Capacitación", "Taller regional — zona centro", "Real", 1100),
    ("PRJ-003", "2026-06-01", "Capacitación", "Taller regional — zona sur", "Real", 1050),
    ("PRJ-003", "2026-06-10", "Capacitación", "Talleres de adopción del CRM", "Real", 1050),
    ("PRJ-003", "2026-06-24", "Capacitación", "Material y plataforma de formación", "Real", 1000),
    # PRJ-004 · Portal de Autoservicio RRHH — en curso (plan 15,500 / real 6,100)
    ("PRJ-004", "2026-04-06", "Personal", "Horas de desarrollo del portal", "Planificado", 8000),
    ("PRJ-004", "2026-04-06", "Personal", "Horas de diseño de experiencia", "Planificado", 2000),
    ("PRJ-004", "2026-04-06", "Software / Licencias", "Hosting del portal", "Planificado", 1500),
    ("PRJ-004", "2026-04-06", "Software / Licencias", "Componentes y librerías", "Planificado", 2000),
    ("PRJ-004", "2026-04-06", "Otros", "Pruebas con usuarios", "Planificado", 1500),
    ("PRJ-004", "2026-04-06", "Capacitación", "Guías y formación a empleados", "Planificado", 500),
    ("PRJ-004", "2026-04-15", "Personal", "Horas de diseño — 1.ª quincena de abril", "Real", 700),
    ("PRJ-004", "2026-04-30", "Personal", "Horas de diseño — 2.ª quincena de abril", "Real", 750),
    ("PRJ-004", "2026-05-15", "Personal", "Horas de desarrollo — 1.ª quincena de mayo", "Real", 800),
    ("PRJ-004", "2026-05-31", "Personal", "Horas de desarrollo — 2.ª quincena de mayo", "Real", 800),
    ("PRJ-004", "2026-06-15", "Personal", "Horas de desarrollo — 1.ª quincena de junio", "Real", 750),
    ("PRJ-004", "2026-06-30", "Personal", "Horas de desarrollo — 2.ª quincena de junio", "Real", 800),
    ("PRJ-004", "2026-05-05", "Software / Licencias", "Hosting del portal (semestre)", "Real", 750),
    ("PRJ-004", "2026-06-15", "Software / Licencias", "Componentes del portal", "Real", 750),
    # PRJ-005 · Capacitación en Analítica — finalizado (plan 8,000 / real 7,650, al límite)
    ("PRJ-005", "2026-02-02", "Capacitación", "Instructores del programa", "Planificado", 5500),
    ("PRJ-005", "2026-02-02", "Capacitación", "Material de formación", "Planificado", 1500),
    ("PRJ-005", "2026-02-02", "Capacitación", "Plataforma de aprendizaje", "Planificado", 1000),
    ("PRJ-005", "2026-02-15", "Capacitación", "Instructores — arranque del programa", "Real", 900),
    ("PRJ-005", "2026-02-28", "Capacitación", "Material del primer ciclo", "Real", 1000),
    ("PRJ-005", "2026-03-15", "Capacitación", "Instructores — marzo", "Real", 1100),
    ("PRJ-005", "2026-03-31", "Capacitación", "Plataforma de aprendizaje (trimestre)", "Real", 1000),
    ("PRJ-005", "2026-04-15", "Capacitación", "Instructores — abril", "Real", 1000),
    ("PRJ-005", "2026-04-30", "Capacitación", "Material del segundo ciclo", "Real", 950),
    ("PRJ-005", "2026-05-15", "Capacitación", "Instructores — mayo", "Real", 850),
    ("PRJ-005", "2026-05-29", "Capacitación", "Instructores — cierre del programa", "Real", 850),
    # PRJ-006 · Rediseño de Marca — finalizado (plan 12,000 / real 12,500, consultoría excedida)
    ("PRJ-006", "2025-09-01", "Consultoría", "Agencia de diseño de marca", "Planificado", 7000),
    ("PRJ-006", "2025-09-01", "Consultoría", "Fotografía y banco de imágenes", "Planificado", 2000),
    ("PRJ-006", "2025-09-01", "Otros", "Impresión de material corporativo", "Planificado", 1800),
    ("PRJ-006", "2025-09-01", "Otros", "Señalética de sedes", "Planificado", 1200),
    ("PRJ-006", "2025-09-20", "Consultoría", "Agencia — anticipo", "Real", 2000),
    ("PRJ-006", "2025-10-30", "Consultoría", "Agencia — conceptos de marca", "Real", 2000),
    ("PRJ-006", "2025-11-30", "Consultoría", "Agencia — propuesta final", "Real", 2000),
    ("PRJ-006", "2025-12-15", "Consultoría", "Fotografía corporativa", "Real", 1900),
    ("PRJ-006", "2026-01-15", "Consultoría", "Agencia — manual de marca", "Real", 1900),
    ("PRJ-006", "2026-01-10", "Otros", "Impresión de papelería", "Real", 900),
    ("PRJ-006", "2026-01-20", "Otros", "Impresión de material comercial", "Real", 900),
    ("PRJ-006", "2026-01-28", "Otros", "Señalética sede principal", "Real", 900),
    # PRJ-007 · ERP Logístico — en curso, el más grande (plan 60,000 / real 22,000)
    ("PRJ-007", "2026-05-10", "Software / Licencias", "Licenciamiento ERP", "Planificado", 24000),
    ("PRJ-007", "2026-05-10", "Software / Licencias", "Módulo de gestión de bodegas", "Planificado", 4000),
    ("PRJ-007", "2026-05-10", "Software / Licencias", "Módulo de transporte", "Planificado", 2000),
    ("PRJ-007", "2026-05-10", "Consultoría", "Implementador ERP", "Planificado", 18000),
    ("PRJ-007", "2026-05-10", "Consultoría", "Gestión del cambio", "Planificado", 4000),
    ("PRJ-007", "2026-05-10", "Capacitación", "Formación de usuarios clave", "Planificado", 3000),
    ("PRJ-007", "2026-05-10", "Capacitación", "Material de formación", "Planificado", 1000),
    ("PRJ-007", "2026-05-10", "Viáticos", "Acompañamiento en bodegas", "Planificado", 2500),
    ("PRJ-007", "2026-05-10", "Hardware", "Terminales de radiofrecuencia", "Planificado", 1000),
    ("PRJ-007", "2026-05-10", "Otros", "Imprevistos de implementación", "Planificado", 500),
    ("PRJ-007", "2026-05-20", "Software / Licencias", "Anticipo de licenciamiento", "Real", 12000),
    ("PRJ-007", "2026-06-10", "Software / Licencias", "Módulo de bodegas — anticipo", "Real", 2000),
    ("PRJ-007", "2026-06-28", "Software / Licencias", "Módulo de transporte — anticipo", "Real", 1000),
    ("PRJ-007", "2026-05-31", "Consultoría", "Implementador — arranque", "Real", 1500),
    ("PRJ-007", "2026-06-15", "Consultoría", "Implementador — análisis de procesos", "Real", 2000),
    ("PRJ-007", "2026-06-25", "Consultoría", "Implementador — diseño de la solución", "Real", 2000),
    ("PRJ-007", "2026-07-05", "Consultoría", "Gestión del cambio — talleres iniciales", "Real", 1500),
    # PRJ-008 · Infraestructura TI — en riesgo (plan 20,000 / real 21,500, rubros excedidos)
    ("PRJ-008", "2026-06-01", "Hardware", "Servidores", "Planificado", 10000),
    ("PRJ-008", "2026-06-01", "Hardware", "Almacenamiento", "Planificado", 3000),
    ("PRJ-008", "2026-06-01", "Hardware", "Equipos de red", "Planificado", 1000),
    ("PRJ-008", "2026-06-01", "Personal", "Horas del equipo TI", "Planificado", 6000),
    ("PRJ-008", "2026-06-15", "Hardware", "Compra de servidores", "Real", 11400),
    ("PRJ-008", "2026-06-18", "Hardware", "Almacenamiento adicional", "Real", 2000),
    ("PRJ-008", "2026-06-22", "Hardware", "Discos y repuestos", "Real", 600),
    ("PRJ-008", "2026-06-25", "Hardware", "Tarjetas y equipos de red", "Real", 1200),
    ("PRJ-008", "2026-06-15", "Personal", "Horas TI — 1.ª quincena de junio", "Real", 2000),
    ("PRJ-008", "2026-06-30", "Personal", "Horas TI — 2.ª quincena de junio", "Real", 2100),
    ("PRJ-008", "2026-07-08", "Personal", "Horas TI — estabilización", "Real", 2200),
    # PRJ-009 · Campaña Digital Q4 — no iniciado (solo partidas)
    ("PRJ-009", "2026-09-01", "Otros", "Pauta digital", "Planificado", 5500),
    ("PRJ-009", "2026-09-01", "Otros", "Colaboraciones con influencers", "Planificado", 1500),
    ("PRJ-009", "2026-09-01", "Consultoría", "Agencia de medios", "Planificado", 1500),
    ("PRJ-009", "2026-09-01", "Consultoría", "Producción de piezas publicitarias", "Planificado", 500),
    ("PRJ-009", "2026-09-01", "Software / Licencias", "Herramienta de analítica de campaña", "Planificado", 500),
    # PRJ-010 · Auditoría de Procesos — finalizado (plan 7,000 / real 6,800)
    ("PRJ-010", "2025-11-03", "Personal", "Horas del equipo auditor", "Planificado", 5000),
    ("PRJ-010", "2025-11-03", "Viáticos", "Visitas a plantas y sedes", "Planificado", 2000),
    ("PRJ-010", "2025-11-15", "Personal", "Horas auditoría — 1.ª quincena de noviembre", "Real", 800),
    ("PRJ-010", "2025-11-30", "Personal", "Horas auditoría — 2.ª quincena de noviembre", "Real", 850),
    ("PRJ-010", "2025-12-15", "Personal", "Horas auditoría — 1.ª quincena de diciembre", "Real", 900),
    ("PRJ-010", "2025-12-31", "Personal", "Horas auditoría — 2.ª quincena de diciembre", "Real", 850),
    ("PRJ-010", "2026-01-31", "Personal", "Horas auditoría — enero", "Real", 850),
    ("PRJ-010", "2026-02-20", "Personal", "Horas del informe final", "Real", 850),
    ("PRJ-010", "2025-11-20", "Viáticos", "Visitas a sedes — noviembre", "Real", 400),
    ("PRJ-010", "2025-12-10", "Viáticos", "Visitas a plantas — diciembre", "Real", 450),
    ("PRJ-010", "2026-01-15", "Viáticos", "Visitas a sedes — enero", "Real", 400),
    ("PRJ-010", "2026-02-10", "Viáticos", "Visitas de cierre — febrero", "Real", 450),
    # PRJ-011 · Plataforma Comercial — recién iniciado (plan 35,000 / real 9,000)
    ("PRJ-011", "2026-07-01", "Personal", "Horas de desarrollo", "Planificado", 12000),
    ("PRJ-011", "2026-07-01", "Personal", "Horas de análisis comercial", "Planificado", 3000),
    ("PRJ-011", "2026-07-01", "Software / Licencias", "Plataforma e-commerce", "Planificado", 10000),
    ("PRJ-011", "2026-07-01", "Software / Licencias", "Pasarela de pagos", "Planificado", 2000),
    ("PRJ-011", "2026-07-01", "Consultoría", "Especialista en pagos en línea", "Planificado", 6000),
    ("PRJ-011", "2026-07-01", "Capacitación", "Formación del equipo comercial", "Planificado", 1500),
    ("PRJ-011", "2026-07-01", "Otros", "Imprevistos", "Planificado", 500),
    ("PRJ-011", "2026-07-05", "Software / Licencias", "Anticipo plataforma e-commerce", "Real", 2500),
    ("PRJ-011", "2026-07-08", "Software / Licencias", "Configuración inicial de la plataforma", "Real", 1500),
    ("PRJ-011", "2026-07-03", "Personal", "Levantamiento de requisitos", "Real", 1500),
    ("PRJ-011", "2026-07-07", "Personal", "Diseño funcional", "Real", 1700),
    ("PRJ-011", "2026-07-09", "Personal", "Horas de desarrollo — arranque", "Real", 1800),
    # PRJ-012 · Bienestar Laboral — finalizado (plan 6,000 / real 5,900, al límite)
    ("PRJ-012", "2026-03-15", "Otros", "Actividades de bienestar", "Planificado", 3000),
    ("PRJ-012", "2026-03-15", "Otros", "Feria de bienestar", "Planificado", 1000),
    ("PRJ-012", "2026-03-15", "Capacitación", "Talleres de manejo del estrés", "Planificado", 2000),
    ("PRJ-012", "2026-04-10", "Otros", "Jornada deportiva", "Real", 900),
    ("PRJ-012", "2026-04-25", "Otros", "Actividades de integración", "Real", 1000),
    ("PRJ-012", "2026-05-15", "Otros", "Jornadas de bienestar", "Real", 1000),
    ("PRJ-012", "2026-06-20", "Otros", "Feria de bienestar", "Real", 1000),
    ("PRJ-012", "2026-04-20", "Capacitación", "Taller de manejo del estrés I", "Real", 1000),
    ("PRJ-012", "2026-06-10", "Capacitación", "Taller de manejo del estrés II", "Real", 1000),
]

DATOS_FICTICIOS_PRESUPUESTO = [
    {
        "ID_Gasto": f"GTO-{i:03d}", "ID_Proyecto": p,
        "Nombre_Proyecto": _MAPA_NOMBRE_PROY_SEED[p],
        "Fecha": f, "Categoria": cat, "Concepto": con, "Tipo": t, "Monto": m,
    }
    for i, (p, f, cat, con, t, m) in enumerate(_GASTOS_BASE, start=1)
]

# --- Módulo 5: Tareas e hitos ---
COLS_TAREAS = [
    "ID_Tarea", "ID_Proyecto", "Nombre_Proyecto", "Tipo", "Nombre_Tarea",
    "ID_Responsable", "Nombre_Responsable", "Fecha_Planificada", "Fecha_Real", "Estado"
]

ESTADOS_TAREA = ["Pendiente", "En curso", "Completada", "Retrasada"]

# (proyecto, tipo, nombre, responsable, fecha_plan, fecha_real, estado)
# Cronograma de muestra: cubre los 12 proyectos con hitos y tareas en todos los
# estados posibles — completadas a tiempo y con retraso, vencidas sin completar,
# por vencer en la semana y planificadas a futuro — para demostrar el control
# de cronograma (la referencia temporal de la demo es julio de 2026).
_TAREAS_BASE = [
    # PRJ-001
    ("PRJ-001", "Hito", "Kickoff del proyecto", "COL-001", "2026-03-02", "2026-03-02", "Completada"),
    ("PRJ-001", "Tarea", "Levantamiento de indicadores financieros", "COL-010", "2026-03-20", "2026-03-20", "Completada"),
    ("PRJ-001", "Tarea", "Conexión a las fuentes contables", "COL-004", "2026-04-05", "2026-04-08", "Completada"),
    ("PRJ-001", "Tarea", "Modelo de datos financiero", "COL-010", "2026-04-15", "2026-04-20", "Completada"),
    ("PRJ-001", "Tarea", "Validación de cifras con Finanzas", "COL-010", "2026-05-05", "2026-05-05", "Completada"),
    ("PRJ-001", "Tarea", "Diseño de tableros para la dirección", "COL-004", "2026-05-20", "2026-05-18", "Completada"),
    ("PRJ-001", "Tarea", "Tablero de ingresos v1", "COL-004", "2026-05-28", "2026-06-02", "Completada"),
    ("PRJ-001", "Tarea", "Pruebas de usuario del tablero v1", "COL-010", "2026-06-08", "2026-06-12", "Completada"),
    ("PRJ-001", "Hito", "Entrega del tablero v1", "COL-001", "2026-06-15", "2026-06-22", "Completada"),
    ("PRJ-001", "Tarea", "Capacitación a usuarios finales", "COL-010", "2026-07-14", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Tablero de costos v2", "COL-004", "2026-08-10", "", "En curso"),
    ("PRJ-001", "Hito", "Cierre del proyecto", "COL-001", "2026-09-30", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Comité mensual del proyecto — 06/03", "COL-001", "2026-03-06", "2026-03-06", "Completada"),
    ("PRJ-001", "Tarea", "Comité mensual del proyecto — 03/04", "COL-001", "2026-04-03", "2026-04-03", "Completada"),
    ("PRJ-001", "Tarea", "Comité mensual del proyecto — 08/05", "COL-001", "2026-05-08", "2026-05-08", "Completada"),
    ("PRJ-001", "Tarea", "Comité mensual del proyecto — 05/06", "COL-001", "2026-06-05", "2026-06-05", "Completada"),
    ("PRJ-001", "Tarea", "Comité mensual del proyecto — 03/07", "COL-001", "2026-07-03", "2026-07-03", "Completada"),
    ("PRJ-001", "Tarea", "Comité mensual del proyecto — 07/08", "COL-001", "2026-08-07", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Comité mensual del proyecto — 04/09", "COL-001", "2026-09-04", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Revisión quincenal de avance — 17/07", "COL-004", "2026-07-17", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Revisión quincenal de avance — 31/07", "COL-004", "2026-07-31", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Revisión quincenal de avance — 14/08", "COL-004", "2026-08-14", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Revisión quincenal de avance — 28/08", "COL-004", "2026-08-28", "", "Pendiente"),
    ("PRJ-001", "Tarea", "Revisión quincenal de avance — 11/09", "COL-004", "2026-09-11", "", "Pendiente"),
    # PRJ-002
    ("PRJ-002", "Hito", "Kickoff del proyecto", "COL-002", "2026-08-03", "", "Pendiente"),
    ("PRJ-002", "Tarea", "Levantamiento del proceso de facturas", "COL-011", "2026-08-21", "", "Pendiente"),
    ("PRJ-002", "Tarea", "Selección de la plataforma", "COL-002", "2026-09-04", "", "Pendiente"),
    ("PRJ-002", "Tarea", "Diseño del flujo de aprobación", "COL-011", "2026-09-18", "", "Pendiente"),
    ("PRJ-002", "Tarea", "Configuración de la lectura automática (OCR)", "COL-002", "2026-10-02", "", "Pendiente"),
    ("PRJ-002", "Tarea", "Integración con el sistema contable", "COL-011", "2026-10-16", "", "Pendiente"),
    ("PRJ-002", "Tarea", "Pruebas piloto con 5 proveedores", "COL-011", "2026-10-30", "", "Pendiente"),
    ("PRJ-002", "Hito", "Piloto aprobado", "COL-002", "2026-11-06", "", "Pendiente"),
    ("PRJ-002", "Tarea", "Capacitación a cuentas por pagar", "COL-002", "2026-11-10", "", "Pendiente"),
    ("PRJ-002", "Hito", "Salida en vivo", "COL-002", "2026-11-16", "", "Pendiente"),
    # PRJ-003
    ("PRJ-003", "Hito", "Kickoff del proyecto", "COL-003", "2026-01-12", "2026-01-12", "Completada"),
    ("PRJ-003", "Tarea", "Diagnóstico del CRM actual", "COL-003", "2026-01-30", "2026-02-04", "Completada"),
    ("PRJ-003", "Tarea", "Diseño de la arquitectura en la nube", "COL-004", "2026-02-20", "2026-02-27", "Completada"),
    ("PRJ-003", "Tarea", "Depuración de datos de clientes", "COL-004", "2026-03-13", "2026-03-25", "Completada"),
    ("PRJ-003", "Hito", "Migración piloto", "COL-003", "2026-03-31", "2026-04-10", "Completada"),
    ("PRJ-003", "Tarea", "Corrección de errores del piloto", "COL-004", "2026-04-24", "2026-05-06", "Completada"),
    ("PRJ-003", "Hito", "Migración de datos", "COL-003", "2026-04-30", "2026-05-15", "Completada"),
    ("PRJ-003", "Tarea", "Configuración de integraciones", "COL-004", "2026-05-15", "2026-05-29", "Completada"),
    ("PRJ-003", "Tarea", "Validación de datos migrados", "COL-003", "2026-06-05", "2026-06-17", "Completada"),
    ("PRJ-003", "Tarea", "Pruebas integrales", "COL-004", "2026-06-20", "", "Retrasada"),
    ("PRJ-003", "Tarea", "Capacitación del equipo comercial", "COL-003", "2026-07-03", "", "En curso"),
    ("PRJ-003", "Hito", "Go-live del CRM", "COL-003", "2026-07-10", "", "En curso"),
    ("PRJ-003", "Hito", "Cierre y estabilización", "COL-003", "2026-07-15", "", "Pendiente"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 23/01", "COL-003", "2026-01-23", "2026-01-23", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 06/02", "COL-003", "2026-02-06", "2026-02-07", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 20/02", "COL-003", "2026-02-20", "2026-02-20", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 06/03", "COL-003", "2026-03-06", "2026-03-08", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 20/03", "COL-003", "2026-03-20", "2026-03-20", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 03/04", "COL-003", "2026-04-03", "2026-04-04", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 17/04", "COL-003", "2026-04-17", "2026-04-17", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 01/05", "COL-003", "2026-05-01", "2026-05-03", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 15/05", "COL-003", "2026-05-15", "2026-05-15", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 29/05", "COL-003", "2026-05-29", "2026-05-30", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 12/06", "COL-003", "2026-06-12", "2026-06-12", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 26/06", "COL-003", "2026-06-26", "2026-06-28", "Completada"),
    ("PRJ-003", "Tarea", "Comité quincenal de seguimiento — 10/07", "COL-003", "2026-07-10", "", "Pendiente"),
    # PRJ-004
    ("PRJ-004", "Hito", "Kickoff del proyecto", "COL-004", "2026-04-06", "2026-04-06", "Completada"),
    ("PRJ-004", "Tarea", "Levantamiento de trámites de RRHH", "COL-004", "2026-04-24", "2026-04-28", "Completada"),
    ("PRJ-004", "Tarea", "Arquitectura y seguridad del portal", "COL-004", "2026-05-08", "2026-05-08", "Completada"),
    ("PRJ-004", "Tarea", "Diseño del portal", "COL-004", "2026-05-15", "2026-05-15", "Completada"),
    ("PRJ-004", "Tarea", "Módulo de vacaciones", "COL-004", "2026-07-20", "", "En curso"),
    ("PRJ-004", "Tarea", "Integración con la nómina", "COL-004", "2026-07-31", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Módulo de certificados laborales", "COL-004", "2026-08-25", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Pruebas de seguridad", "COL-004", "2026-09-01", "", "Pendiente"),
    ("PRJ-004", "Hito", "Piloto con Talento Humano", "COL-004", "2026-09-15", "", "Pendiente"),
    ("PRJ-004", "Hito", "Cierre del proyecto", "COL-004", "2026-10-30", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Comité mensual del proyecto — 10/04", "COL-004", "2026-04-10", "2026-04-10", "Completada"),
    ("PRJ-004", "Tarea", "Comité mensual del proyecto — 08/05", "COL-004", "2026-05-08", "2026-05-11", "Completada"),
    ("PRJ-004", "Tarea", "Comité mensual del proyecto — 05/06", "COL-004", "2026-06-05", "2026-06-05", "Completada"),
    ("PRJ-004", "Tarea", "Comité mensual del proyecto — 10/07", "COL-004", "2026-07-10", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Comité mensual del proyecto — 07/08", "COL-004", "2026-08-07", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Comité mensual del proyecto — 04/09", "COL-004", "2026-09-04", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Comité mensual del proyecto — 02/10", "COL-004", "2026-10-02", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Demo quincenal a Talento Humano — 19/06", "COL-004", "2026-06-19", "2026-06-21", "Completada"),
    ("PRJ-004", "Tarea", "Demo quincenal a Talento Humano — 16/07", "COL-004", "2026-07-16", "", "Pendiente"),
    ("PRJ-004", "Tarea", "Demo quincenal a Talento Humano — 30/07", "COL-004", "2026-07-30", "", "Pendiente"),
    # PRJ-005
    ("PRJ-005", "Tarea", "Convocatoria e inscripciones", "COL-001", "2026-02-09", "2026-02-09", "Completada"),
    ("PRJ-005", "Tarea", "Diseño del plan de formación", "COL-001", "2026-02-16", "2026-02-13", "Completada"),
    ("PRJ-005", "Tarea", "Preparación del material", "COL-001", "2026-02-27", "2026-03-03", "Completada"),
    ("PRJ-005", "Hito", "Inicio del primer ciclo", "COL-001", "2026-03-02", "2026-03-02", "Completada"),
    ("PRJ-005", "Tarea", "Talleres del primer ciclo", "COL-001", "2026-04-10", "2026-04-10", "Completada"),
    ("PRJ-005", "Tarea", "Talleres del segundo ciclo", "COL-001", "2026-05-08", "2026-05-12", "Completada"),
    ("PRJ-005", "Tarea", "Evaluación final y certificados", "COL-001", "2026-05-22", "2026-06-01", "Completada"),
    ("PRJ-005", "Hito", "Cierre del programa", "COL-001", "2026-05-29", "2026-06-05", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 06/03", "COL-001", "2026-03-06", "2026-03-06", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 13/03", "COL-001", "2026-03-13", "2026-03-14", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 20/03", "COL-001", "2026-03-20", "2026-03-20", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 27/03", "COL-001", "2026-03-27", "2026-03-27", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 17/04", "COL-001", "2026-04-17", "2026-04-17", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 24/04", "COL-001", "2026-04-24", "2026-04-25", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 01/05", "COL-001", "2026-05-01", "2026-05-01", "Completada"),
    ("PRJ-005", "Tarea", "Sesión semanal de formación — 15/05", "COL-001", "2026-05-15", "2026-05-15", "Completada"),
    # PRJ-006
    ("PRJ-006", "Hito", "Kickoff con la agencia", "COL-006", "2025-09-08", "2025-09-08", "Completada"),
    ("PRJ-006", "Tarea", "Investigación y benchmarking de marca", "COL-006", "2025-09-26", "2025-09-30", "Completada"),
    ("PRJ-006", "Tarea", "Talleres de identidad con directivos", "COL-006", "2025-10-10", "2025-10-10", "Completada"),
    ("PRJ-006", "Tarea", "Propuesta de identidad visual", "COL-006", "2025-10-15", "2025-10-20", "Completada"),
    ("PRJ-006", "Tarea", "Registro de marca ante la superintendencia", "COL-006", "2025-11-21", "2025-12-01", "Completada"),
    ("PRJ-006", "Hito", "Aprobación del manual de marca", "COL-006", "2025-12-10", "2025-12-05", "Completada"),
    ("PRJ-006", "Tarea", "Diseño de aplicaciones de marca", "COL-006", "2025-12-19", "2025-12-22", "Completada"),
    ("PRJ-006", "Tarea", "Producción de material impreso", "COL-006", "2026-01-09", "2026-01-13", "Completada"),
    ("PRJ-006", "Tarea", "Actualización de canales digitales", "COL-006", "2026-01-23", "2026-01-26", "Completada"),
    ("PRJ-006", "Hito", "Lanzamiento interno de la marca", "COL-006", "2026-01-30", "2026-01-28", "Completada"),
    ("PRJ-006", "Hito", "Entrega de archivos finales", "COL-006", "2026-01-30", "2026-01-28", "Completada"),
    ("PRJ-006", "Tarea", "Comité mensual con la agencia — 19/09", "COL-006", "2025-09-19", "2025-09-19", "Completada"),
    ("PRJ-006", "Tarea", "Comité mensual con la agencia — 17/10", "COL-006", "2025-10-17", "2025-10-18", "Completada"),
    ("PRJ-006", "Tarea", "Comité mensual con la agencia — 14/11", "COL-006", "2025-11-14", "2025-11-14", "Completada"),
    ("PRJ-006", "Tarea", "Comité mensual con la agencia — 12/12", "COL-006", "2025-12-12", "2025-12-13", "Completada"),
    ("PRJ-006", "Tarea", "Comité mensual con la agencia — 16/01", "COL-006", "2026-01-16", "2026-01-16", "Completada"),
    # PRJ-007
    ("PRJ-007", "Hito", "Kickoff del proyecto", "COL-007", "2026-05-11", "2026-05-11", "Completada"),
    ("PRJ-007", "Tarea", "Levantamiento de procesos logísticos", "COL-009", "2026-05-29", "2026-06-02", "Completada"),
    ("PRJ-007", "Hito", "Análisis de procesos", "COL-007", "2026-06-15", "2026-06-18", "Completada"),
    ("PRJ-007", "Tarea", "Diseño de la solución", "COL-009", "2026-06-19", "2026-06-24", "Completada"),
    ("PRJ-007", "Tarea", "Preparación de datos maestros", "COL-009", "2026-07-17", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Configuración del módulo de inventarios", "COL-009", "2026-08-01", "", "En curso"),
    ("PRJ-007", "Tarea", "Configuración del módulo de compras", "COL-009", "2026-08-21", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Migración de datos maestros", "COL-009", "2026-09-04", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Integración con facturación", "COL-009", "2026-09-10", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Capacitación a usuarios clave", "COL-007", "2026-09-18", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Pruebas integrales del piloto", "COL-009", "2026-09-25", "", "Pendiente"),
    ("PRJ-007", "Hito", "Piloto en bodega principal", "COL-007", "2026-09-30", "", "Pendiente"),
    ("PRJ-007", "Hito", "Salida en vivo nacional", "COL-007", "2027-01-15", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Comité mensual de dirección — 05/06", "COL-007", "2026-06-05", "2026-06-05", "Completada"),
    ("PRJ-007", "Tarea", "Comité mensual de dirección — 03/07", "COL-007", "2026-07-03", "2026-07-03", "Completada"),
    ("PRJ-007", "Tarea", "Comité mensual de dirección — 07/08", "COL-007", "2026-08-07", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Comité mensual de dirección — 04/09", "COL-007", "2026-09-04", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Comité mensual de dirección — 02/10", "COL-007", "2026-10-02", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Comité mensual de dirección — 06/11", "COL-007", "2026-11-06", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Comité mensual de dirección — 04/12", "COL-007", "2026-12-04", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 12/06", "COL-007", "2026-06-12", "2026-06-12", "Completada"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 26/06", "COL-007", "2026-06-26", "2026-06-27", "Completada"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 10/07", "COL-007", "2026-07-10", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 24/07", "COL-007", "2026-07-24", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 14/08", "COL-007", "2026-08-14", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 28/08", "COL-007", "2026-08-28", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 11/09", "COL-007", "2026-09-11", "", "Pendiente"),
    ("PRJ-007", "Tarea", "Seguimiento quincenal con el implementador — 09/10", "COL-007", "2026-10-09", "", "Pendiente"),
    # PRJ-008
    ("PRJ-008", "Hito", "Kickoff del proyecto", "COL-008", "2026-06-01", "2026-06-01", "Completada"),
    ("PRJ-008", "Tarea", "Inventario de aplicaciones y servicios", "COL-009", "2026-06-08", "2026-06-10", "Completada"),
    ("PRJ-008", "Tarea", "Diseño de la arquitectura híbrida", "COL-008", "2026-06-12", "2026-06-16", "Completada"),
    ("PRJ-008", "Tarea", "Instalación de servidores", "COL-008", "2026-06-20", "2026-06-28", "Completada"),
    ("PRJ-008", "Tarea", "Configuración de almacenamiento", "COL-009", "2026-06-26", "2026-07-01", "Completada"),
    ("PRJ-008", "Tarea", "Migración de servicios no críticos", "COL-009", "2026-07-03", "2026-07-08", "Completada"),
    ("PRJ-008", "Tarea", "Migración de servicios críticos", "COL-008", "2026-07-05", "", "Retrasada"),
    ("PRJ-008", "Tarea", "Pruebas de contingencia", "COL-009", "2026-07-12", "", "Pendiente"),
    ("PRJ-008", "Tarea", "Hardening y seguridad de servidores", "COL-008", "2026-07-24", "", "Pendiente"),
    ("PRJ-008", "Tarea", "Documentación y transferencia a operación", "COL-009", "2026-08-10", "", "Pendiente"),
    ("PRJ-008", "Hito", "Entrega de infraestructura renovada", "COL-008", "2026-08-15", "", "Pendiente"),
    ("PRJ-008", "Tarea", "Comité semanal de infraestructura — 05/06", "COL-008", "2026-06-05", "2026-06-05", "Completada"),
    ("PRJ-008", "Tarea", "Comité semanal de infraestructura — 12/06", "COL-008", "2026-06-12", "2026-06-12", "Completada"),
    ("PRJ-008", "Tarea", "Comité semanal de infraestructura — 19/06", "COL-008", "2026-06-19", "2026-06-20", "Completada"),
    ("PRJ-008", "Tarea", "Comité semanal de infraestructura — 26/06", "COL-008", "2026-06-26", "2026-06-26", "Completada"),
    ("PRJ-008", "Tarea", "Comité semanal de infraestructura — 03/07", "COL-008", "2026-07-03", "2026-07-03", "Completada"),
    ("PRJ-008", "Tarea", "Comité semanal de infraestructura — 10/07", "COL-008", "2026-07-10", "", "Pendiente"),
    ("PRJ-008", "Tarea", "Comité semanal de infraestructura — 17/07", "COL-008", "2026-07-17", "", "Pendiente"),
    # PRJ-009
    ("PRJ-009", "Tarea", "Contratación de la agencia de medios", "COL-006", "2026-09-04", "", "Pendiente"),
    ("PRJ-009", "Tarea", "Plan de campaña y presupuesto de pauta", "COL-006", "2026-09-08", "", "Pendiente"),
    ("PRJ-009", "Tarea", "Definición de audiencias y segmentos", "COL-006", "2026-09-11", "", "Pendiente"),
    ("PRJ-009", "Tarea", "Producción de piezas creativas", "COL-006", "2026-09-18", "", "Pendiente"),
    ("PRJ-009", "Hito", "Aprobación de la campaña por la dirección", "COL-006", "2026-09-25", "", "Pendiente"),
    ("PRJ-009", "Tarea", "Configuración de analítica y medición", "COL-006", "2026-09-29", "", "Pendiente"),
    ("PRJ-009", "Hito", "Lanzamiento de la campaña", "COL-006", "2026-10-01", "", "Pendiente"),
    ("PRJ-009", "Tarea", "Optimización quincenal de la pauta", "COL-006", "2026-10-16", "", "Pendiente"),
    ("PRJ-009", "Tarea", "Campaña especial de Black Friday", "COL-006", "2026-11-27", "", "Pendiente"),
    ("PRJ-009", "Hito", "Cierre y reporte de resultados", "COL-006", "2026-12-15", "", "Pendiente"),
    # PRJ-010
    ("PRJ-010", "Hito", "Kickoff de la auditoría", "COL-002", "2025-11-03", "2025-11-03", "Completada"),
    ("PRJ-010", "Tarea", "Plan de auditoría y muestreo", "COL-002", "2025-11-14", "2025-11-18", "Completada"),
    ("PRJ-010", "Tarea", "Revisión de procesos financieros", "COL-011", "2025-11-28", "2025-12-02", "Completada"),
    ("PRJ-010", "Tarea", "Revisión de procesos operativos", "COL-011", "2025-12-12", "2025-12-19", "Completada"),
    ("PRJ-010", "Tarea", "Entrevistas con líderes de área", "COL-002", "2026-01-09", "2026-01-14", "Completada"),
    ("PRJ-010", "Tarea", "Trabajo de campo en sedes", "COL-011", "2026-01-20", "2026-01-30", "Completada"),
    ("PRJ-010", "Tarea", "Consolidación de hallazgos", "COL-002", "2026-02-06", "2026-02-16", "Completada"),
    ("PRJ-010", "Hito", "Informe final de auditoría", "COL-002", "2026-02-28", "2026-03-10", "Completada"),
    ("PRJ-010", "Tarea", "Presentación a la dirección", "COL-002", "2026-03-06", "2026-03-13", "Completada"),
    ("PRJ-010", "Tarea", "Comité quincenal de auditoría — 21/11", "COL-002", "2025-11-21", "2025-11-21", "Completada"),
    ("PRJ-010", "Tarea", "Comité quincenal de auditoría — 05/12", "COL-002", "2025-12-05", "2025-12-06", "Completada"),
    ("PRJ-010", "Tarea", "Comité quincenal de auditoría — 19/12", "COL-002", "2025-12-19", "2025-12-19", "Completada"),
    ("PRJ-010", "Tarea", "Comité quincenal de auditoría — 16/01", "COL-002", "2026-01-16", "2026-01-17", "Completada"),
    ("PRJ-010", "Tarea", "Comité quincenal de auditoría — 30/01", "COL-002", "2026-01-30", "2026-01-30", "Completada"),
    # PRJ-011
    ("PRJ-011", "Hito", "Kickoff del proyecto", "COL-011", "2026-07-02", "2026-07-02", "Completada"),
    ("PRJ-011", "Tarea", "Contratación del especialista en pagos", "COL-011", "2026-07-09", "2026-07-09", "Completada"),
    ("PRJ-011", "Tarea", "Levantamiento de requisitos", "COL-011", "2026-07-15", "", "En curso"),
    ("PRJ-011", "Tarea", "Mapa de procesos comerciales", "COL-003", "2026-07-24", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Arquitectura de la plataforma", "COL-003", "2026-08-07", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Diseño del catálogo en línea", "COL-003", "2026-08-20", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Configuración del catálogo base", "COL-003", "2026-09-11", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Integración con la pasarela de pagos", "COL-011", "2026-10-02", "", "Pendiente"),
    ("PRJ-011", "Hito", "Demo a clientes piloto", "COL-011", "2026-10-15", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Integración con el CRM", "COL-003", "2026-11-13", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Pruebas con clientes piloto", "COL-011", "2026-12-04", "", "Pendiente"),
    ("PRJ-011", "Hito", "Cierre de la fase 1", "COL-011", "2026-12-18", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Comité mensual del proyecto — 10/07", "COL-011", "2026-07-10", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Comité mensual del proyecto — 07/08", "COL-011", "2026-08-07", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Comité mensual del proyecto — 04/09", "COL-011", "2026-09-04", "", "Pendiente"),
    ("PRJ-011", "Tarea", "Comité mensual del proyecto — 02/10", "COL-011", "2026-10-02", "", "Pendiente"),
    # PRJ-012
    ("PRJ-012", "Hito", "Lanzamiento del programa", "COL-004", "2026-03-16", "2026-03-16", "Completada"),
    ("PRJ-012", "Tarea", "Conformación del comité de bienestar", "COL-004", "2026-03-27", "2026-03-27", "Completada"),
    ("PRJ-012", "Tarea", "Encuesta de clima laboral", "COL-004", "2026-04-01", "2026-04-03", "Completada"),
    ("PRJ-012", "Tarea", "Jornada deportiva", "COL-004", "2026-04-10", "2026-04-10", "Completada"),
    ("PRJ-012", "Tarea", "Actividades de integración por áreas", "COL-004", "2026-04-24", "2026-04-27", "Completada"),
    ("PRJ-012", "Tarea", "Talleres de manejo del estrés", "COL-004", "2026-05-15", "2026-05-19", "Completada"),
    ("PRJ-012", "Tarea", "Feria de bienestar", "COL-004", "2026-06-19", "2026-06-20", "Completada"),
    ("PRJ-012", "Hito", "Cierre del programa de bienestar", "COL-004", "2026-06-30", "2026-06-30", "Completada"),
    ("PRJ-012", "Tarea", "Seguimiento mensual del programa — 03/04", "COL-004", "2026-04-03", "2026-04-03", "Completada"),
    ("PRJ-012", "Tarea", "Seguimiento mensual del programa — 08/05", "COL-004", "2026-05-08", "2026-05-08", "Completada"),
    ("PRJ-012", "Tarea", "Seguimiento mensual del programa — 05/06", "COL-004", "2026-06-05", "2026-06-05", "Completada"),
    ("PRJ-012", "Tarea", "Seguimiento mensual del programa — 26/06", "COL-004", "2026-06-26", "2026-06-26", "Completada"),
]

DATOS_FICTICIOS_TAREAS = [
    {
        "ID_Tarea": f"TAR-{i:03d}", "ID_Proyecto": p,
        "Nombre_Proyecto": _MAPA_NOMBRE_PROY_SEED[p],
        "Tipo": tipo, "Nombre_Tarea": nom,
        "ID_Responsable": resp, "Nombre_Responsable": _MAPA_NOMBRE_COL_SEED[resp],
        "Fecha_Planificada": fp, "Fecha_Real": fr, "Estado": est,
    }
    for i, (p, tipo, nom, resp, fp, fr, est) in enumerate(_TAREAS_BASE, start=1)
]

# --- Módulo 6: Riesgos ---
COLS_RIESGOS = [
    "ID_Riesgo", "ID_Proyecto", "Nombre_Proyecto", "Descripcion_Riesgo",
    "Probabilidad", "Impacto", "Nivel_Riesgo", "Nivel",
    "Plan_Mitigacion", "ID_Responsable", "Nombre_Responsable", "Estado"
]

ESTADOS_RIESGO = ["Abierto", "En mitigación", "Mitigado", "Materializado", "Cerrado"]

# (proyecto, descripción, probabilidad, impacto, plan, responsable, estado)
# Registro de riesgos de muestra: 10 riesgos por proyecto (120 en total), con
# niveles Bajo/Medio/Alto y todos los estados del ciclo de vida. Los proyectos
# finalizados concentran riesgos cerrados/mitigados; los activos y en riesgo,
# riesgos abiertos y materializados, para demostrar la matriz y las alertas.
_RIESGOS_BASE = [
    # PRJ-001 · Implementación Power BI (en curso)
    ("PRJ-001", "Baja adopción de los tableros por parte de la dirección", 2, 3,
     "Sesiones de acompañamiento mensual", "COL-001", "Mitigado"),
    ("PRJ-001", "Calidad deficiente de los datos contables fuente", 4, 4,
     "Reglas de validación y limpieza en la carga de datos", "COL-010", "En mitigación"),
    ("PRJ-001", "Rotación del desarrollador BI principal", 2, 4,
     "Documentación técnica y respaldo cruzado en el equipo", "COL-004", "Abierto"),
    ("PRJ-001", "Cambios de alcance por nuevas solicitudes de la dirección", 4, 3,
     "Comité de cambios quincenal con priorización", "COL-001", "En mitigación"),
    ("PRJ-001", "Retraso en los accesos a las fuentes de datos", 3, 3,
     "Gestión temprana con TI y permisos de solo lectura", "COL-010", "Cerrado"),
    ("PRJ-001", "Licencias insuficientes para los usuarios finales", 2, 2,
     "Inventario de usuarios y compra escalonada", "COL-001", "Abierto"),
    ("PRJ-001", "Indicadores mal definidos por falta de diccionario de datos", 3, 4,
     "Diccionario de indicadores validado con Finanzas", "COL-010", "En mitigación"),
    ("PRJ-001", "Bajo desempeño del modelo con los datos históricos", 3, 3,
     "Optimizar el modelo y crear tablas de resumen", "COL-004", "Abierto"),
    ("PRJ-001", "Dependencia de un proveedor externo para la carga de datos", 2, 3,
     "Cláusulas de soporte y transferencia de conocimiento", "COL-001", "Abierto"),
    ("PRJ-001", "Sobrecarga del equipo por soporte a otros proyectos", 4, 4,
     "Priorización con el comité y apoyo de practicante", "COL-004", "Abierto"),
    # PRJ-002 · Automatización de Facturas (no iniciado)
    ("PRJ-002", "Kickoff pospuesto por el cierre contable de agosto", 3, 2,
     "Agendar el kickoff fuera de las fechas de cierre", "COL-002", "Abierto"),
    ("PRJ-002", "Proveedores que siguen facturando en papel", 4, 3,
     "Plan de digitalización y portal de proveedores", "COL-011", "Abierto"),
    ("PRJ-002", "Resistencia del equipo de cuentas por pagar", 3, 3,
     "Involucrar al equipo desde el diseño del proceso", "COL-002", "Abierto"),
    ("PRJ-002", "Presupuesto insuficiente para la plataforma elegida", 2, 4,
     "Cotizaciones tempranas con tres proveedores", "COL-002", "En mitigación"),
    ("PRJ-002", "Integración compleja con el sistema contable", 3, 4,
     "Prueba de concepto antes de comprar la herramienta", "COL-011", "Abierto"),
    ("PRJ-002", "Cambios normativos de facturación electrónica", 2, 4,
     "Monitoreo normativo trimestral con contabilidad", "COL-011", "Abierto"),
    ("PRJ-002", "Falta de facturas históricas para probar la automatización", 2, 2,
     "Recolectar una muestra de facturas de 6 meses", "COL-002", "Abierto"),
    ("PRJ-002", "Personal clave sin tiempo para el levantamiento", 4, 3,
     "Bloques de agenda acordados con Operaciones", "COL-011", "Abierto"),
    ("PRJ-002", "Errores de lectura automática en facturas escaneadas", 3, 3,
     "Umbral de confianza y revisión manual de excepciones", "COL-002", "Abierto"),
    ("PRJ-002", "Duplicidad de pagos durante la transición", 2, 5,
     "Conciliación diaria durante el periodo de cambio", "COL-011", "Abierto"),
    # PRJ-003 · Migración CRM a la Nube (en riesgo)
    ("PRJ-003", "Retraso del proveedor cloud en la migración final", 4, 4,
     "Plan de contingencia con ambiente paralelo", "COL-003", "En mitigación"),
    ("PRJ-003", "Pérdida de datos históricos durante la migración", 2, 5,
     "Backups triples y validación por muestreo", "COL-003", "Abierto"),
    ("PRJ-003", "Go-live sin completar las pruebas integrales", 5, 4,
     "Congelar el alcance y ejecutar plan de pruebas exprés", "COL-003", "Abierto"),
    ("PRJ-003", "Baja adopción del CRM por el equipo comercial", 4, 3,
     "Talleres prácticos y embajadores por regional", "COL-003", "En mitigación"),
    ("PRJ-003", "Sobrecosto en la consultoría de migración", 4, 4,
     "Renegociar tarifas y cerrar el alcance de la fase 3", "COL-003", "Materializado"),
    ("PRJ-003", "Integración inestable con el sistema de facturación", 3, 4,
     "Pruebas de integración nocturnas automatizadas", "COL-004", "En mitigación"),
    ("PRJ-003", "Clientes duplicados heredados del CRM anterior", 4, 2,
     "Deduplicación asistida antes de migrar", "COL-004", "Mitigado"),
    ("PRJ-003", "Caída del CRM en la primera semana tras el go-live", 3, 5,
     "Soporte reforzado y plan de reversa documentado", "COL-003", "Abierto"),
    ("PRJ-003", "Fuga de datos personales de clientes", 2, 5,
     "Cifrado y control de accesos por rol", "COL-003", "Abierto"),
    ("PRJ-003", "Dependencia de una única consultora experta", 3, 3,
     "Transferencia de conocimiento documentada", "COL-004", "Abierto"),
    # PRJ-004 · Portal de Autoservicio RRHH (en curso)
    ("PRJ-004", "Sobrecarga del desarrollador principal", 4, 4,
     "Redistribuir tareas y sumar apoyo de practicante", "COL-004", "Abierto"),
    ("PRJ-004", "Requisitos ambiguos del área de Talento Humano", 3, 3,
     "Prototipos validados cada quincena", "COL-004", "En mitigación"),
    ("PRJ-004", "Integración con la nómina más compleja de lo previsto", 3, 4,
     "Prueba técnica temprana con el proveedor de nómina", "COL-004", "Abierto"),
    ("PRJ-004", "Datos personales expuestos por permisos mal configurados", 2, 5,
     "Revisión de seguridad y pruebas de acceso por rol", "COL-004", "Abierto"),
    ("PRJ-004", "Baja adopción del portal por los empleados", 3, 3,
     "Campaña interna y guías rápidas de uso", "COL-004", "Abierto"),
    ("PRJ-004", "Cambios normativos laborales que alteran los trámites", 2, 3,
     "Revisión legal semestral del alcance", "COL-004", "Abierto"),
    ("PRJ-004", "Retraso del piloto por vacaciones colectivas", 3, 2,
     "Adelantar el piloto a comienzos de septiembre", "COL-004", "Abierto"),
    ("PRJ-004", "Componente de desarrollo sin soporte del fabricante", 2, 3,
     "Actualizar a la versión con soporte extendido", "COL-004", "Cerrado"),
    ("PRJ-004", "Solicitudes de RRHH por fuera del alcance inicial", 4, 2,
     "Registro de mejoras para una fase 2", "COL-004", "En mitigación"),
    ("PRJ-004", "Pruebas de usuario insuficientes antes del piloto", 3, 4,
     "Reclutar 10 empleados piloto con guion de pruebas", "COL-004", "Abierto"),
    # PRJ-005 · Capacitación en Analítica de Datos (finalizado)
    ("PRJ-005", "Baja asistencia a los talleres presenciales", 3, 3,
     "Sesiones grabadas y horarios alternos", "COL-001", "Mitigado"),
    ("PRJ-005", "Deserción de participantes a mitad del programa", 3, 2,
     "Seguimiento individual y certificados por nivel", "COL-001", "Mitigado"),
    ("PRJ-005", "Instructor principal no disponible en fechas clave", 2, 4,
     "Instructor suplente certificado", "COL-001", "Cerrado"),
    ("PRJ-005", "Contenidos desactualizados frente a las herramientas", 2, 3,
     "Revisión de contenidos previa a cada ciclo", "COL-001", "Cerrado"),
    ("PRJ-005", "Salas y equipos insuficientes para los grupos", 2, 2,
     "Reserva anticipada y modalidad híbrida", "COL-001", "Cerrado"),
    ("PRJ-005", "Sobrecosto por material adicional de formación", 2, 2,
     "Material digital en lugar de impreso", "COL-001", "Mitigado"),
    ("PRJ-005", "Cierre tardío del programa por sesiones reprogramadas", 3, 2,
     "Plan de recuperación de sesiones canceladas", "COL-001", "Materializado"),
    ("PRJ-005", "Conocimiento no aplicado en el puesto de trabajo", 3, 3,
     "Proyectos aplicados por área con seguimiento", "COL-001", "En mitigación"),
    ("PRJ-005", "Evaluaciones sin estandarizar entre grupos", 2, 2,
     "Rúbrica única de evaluación", "COL-001", "Cerrado"),
    ("PRJ-005", "Rotación del personal recién formado", 2, 3,
     "Acuerdos de permanencia con Talento Humano", "COL-001", "Abierto"),
    # PRJ-006 · Rediseño de Marca Corporativa (finalizado)
    ("PRJ-006", "Rechazo de la nueva identidad por la dirección", 2, 4,
     "Validaciones intermedias con el comité directivo", "COL-006", "Cerrado"),
    ("PRJ-006", "Costos de reimpresión del material corporativo", 3, 3,
     "Inventario y transición escalonada del material", "COL-006", "Materializado"),
    ("PRJ-006", "Inconsistencia de marca en los canales digitales", 3, 2,
     "Manual de marca y plantillas oficiales", "COL-006", "Mitigado"),
    ("PRJ-006", "Demoras de la agencia de diseño", 3, 3,
     "Hitos contractuales con penalidades", "COL-006", "Mitigado"),
    ("PRJ-006", "Confusión de los clientes durante la transición", 2, 3,
     "Campaña de comunicación del cambio", "COL-006", "Cerrado"),
    ("PRJ-006", "Conflicto de derechos sobre la tipografía elegida", 1, 4,
     "Verificación legal de licencias de uso", "COL-006", "Cerrado"),
    ("PRJ-006", "Señalética instalada fuera de plazo en las sedes", 2, 2,
     "Cronograma por sedes con proveedor local", "COL-006", "Mitigado"),
    ("PRJ-006", "Pérdida del reconocimiento de la marca anterior", 2, 4,
     "Transición gradual con doble marca por 3 meses", "COL-006", "Cerrado"),
    ("PRJ-006", "Aplicaciones de marca no contempladas (uniformes, flota)", 3, 2,
     "Levantamiento completo de puntos de contacto de marca", "COL-006", "Mitigado"),
    ("PRJ-006", "Desalineación de la marca con la estrategia comercial", 1, 4,
     "Brief validado con Comercial y la Dirección", "COL-006", "Cerrado"),
    # PRJ-007 · Implementación ERP Logístico (en curso)
    ("PRJ-007", "Resistencia al cambio en la operación logística", 4, 3,
     "Plan de gestión del cambio y capacitación temprana", "COL-007", "Abierto"),
    ("PRJ-007", "Desviación del cronograma por complejidad de integraciones", 3, 4,
     "Fases incrementales y buffer de 3 semanas", "COL-007", "Abierto"),
    ("PRJ-007", "Licenciamiento más caro por usuarios adicionales", 3, 4,
     "Auditoría de usuarios y licencias concurrentes", "COL-007", "En mitigación"),
    ("PRJ-007", "Datos de inventario inexactos para la migración", 4, 4,
     "Inventario físico previo a la carga inicial", "COL-009", "Abierto"),
    ("PRJ-007", "Implementador con poca experiencia en logística nacional", 2, 4,
     "Referencias verificadas y acompañamiento del fabricante", "COL-007", "Mitigado"),
    ("PRJ-007", "Paro de operaciones durante el corte al nuevo sistema", 2, 5,
     "Go-live en fin de semana con plan de reversa", "COL-007", "Abierto"),
    ("PRJ-007", "Personal clave de bodega no disponible para pruebas", 3, 3,
     "Turnos de prueba remunerados fuera de horario", "COL-009", "Abierto"),
    ("PRJ-007", "Alcance mal definido en facturación electrónica", 3, 4,
     "Definición contractual con anexo técnico", "COL-007", "En mitigación"),
    ("PRJ-007", "Infraestructura de red insuficiente en bodegas", 3, 3,
     "Diagnóstico de red y redundancia con datos móviles", "COL-009", "En mitigación"),
    ("PRJ-007", "Salida del líder del proyecto hacia otra unidad", 2, 4,
     "Plan de sucesión y documentación de decisiones", "COL-007", "Abierto"),
    # PRJ-008 · Optimización de Infraestructura TI (en riesgo)
    ("PRJ-008", "Sobrecostos por compra de hardware adicional", 4, 3,
     "Renegociar con el proveedor o usar leasing", "COL-008", "Materializado"),
    ("PRJ-008", "Caída de servicios durante la migración", 3, 5,
     "Ventanas de mantenimiento nocturnas", "COL-008", "En mitigación"),
    ("PRJ-008", "Retraso en las entregas de hardware importado", 4, 3,
     "Órdenes anticipadas y proveedor alterno local", "COL-008", "Materializado"),
    ("PRJ-008", "Sobrecarga del equipo TI por incidentes en producción", 5, 4,
     "Congelar cambios no críticos y contratar apoyo externo", "COL-008", "Abierto"),
    ("PRJ-008", "Pérdida de configuraciones de los servidores antiguos", 2, 4,
     "Respaldo completo y documentación previa", "COL-009", "Mitigado"),
    ("PRJ-008", "Incompatibilidad de aplicaciones antiguas", 3, 4,
     "Inventario de aplicaciones y pruebas de compatibilidad", "COL-009", "En mitigación"),
    ("PRJ-008", "Brecha de seguridad durante la coexistencia de ambientes", 2, 5,
     "Segmentación de red y monitoreo reforzado", "COL-008", "Abierto"),
    ("PRJ-008", "Garantías del hardware anterior vencidas antes del corte", 3, 2,
     "Priorizar la migración de los equipos sin garantía", "COL-008", "En mitigación"),
    ("PRJ-008", "Dependencia de un único ingeniero certificado", 3, 4,
     "Certificar a un segundo ingeniero del equipo", "COL-009", "Abierto"),
    ("PRJ-008", "Presupuesto agotado antes de la fase de contingencia", 4, 4,
     "Reserva del 10% aprobada por el comité", "COL-008", "Abierto"),
    # PRJ-009 · Campaña Digital Q4 (no iniciado)
    ("PRJ-009", "Recorte del presupuesto de pauta", 2, 4,
     "Priorizar canales de mejor conversión", "COL-006", "Abierto"),
    ("PRJ-009", "Lanzamiento tardío que pierde la ventana de fin de año", 3, 4,
     "Cronograma inverso desde el Black Friday", "COL-006", "Abierto"),
    ("PRJ-009", "Costos de pauta más altos por la temporada", 4, 3,
     "Reserva anticipada de espacios publicitarios", "COL-006", "Abierto"),
    ("PRJ-009", "Creatividades rechazadas por las plataformas", 2, 2,
     "Revisión previa de las políticas de anuncios", "COL-006", "Abierto"),
    ("PRJ-009", "Campañas agresivas de la competencia en simultáneo", 4, 3,
     "Diferenciación por segmentos nicho", "COL-006", "Abierto"),
    ("PRJ-009", "Página de aterrizaje sin capacidad para picos de tráfico", 2, 4,
     "Prueba de carga antes del lanzamiento", "COL-006", "Abierto"),
    ("PRJ-009", "Datos de segmentación desactualizados", 3, 3,
     "Actualizar audiencias con el CRM migrado", "COL-006", "Abierto"),
    ("PRJ-009", "Aprobaciones tardías de la dirección comercial", 3, 2,
     "Calendario de aprobaciones acordado por adelantado", "COL-006", "Abierto"),
    ("PRJ-009", "Métricas de conversión sin línea base", 2, 3,
     "Definir KPIs y medición desde la primera semana", "COL-006", "Abierto"),
    ("PRJ-009", "Agencia de medios sin disponibilidad en el último trimestre", 2, 4,
     "Contrato firmado antes de septiembre", "COL-006", "En mitigación"),
    # PRJ-010 · Auditoría de Procesos Internos (finalizado)
    ("PRJ-010", "Acceso restringido a la información de los procesos", 3, 3,
     "Acuerdo de confidencialidad y patrocinio directivo", "COL-002", "Cerrado"),
    ("PRJ-010", "Personal que percibe la auditoría como fiscalización", 4, 2,
     "Comunicación del enfoque de mejora, no de castigo", "COL-011", "Mitigado"),
    ("PRJ-010", "Informe final entregado fuera de plazo", 3, 3,
     "Entregas parciales quincenales", "COL-002", "Materializado"),
    ("PRJ-010", "Hallazgos sin plan de acción posterior", 3, 4,
     "Comité de seguimiento de hallazgos", "COL-002", "En mitigación"),
    ("PRJ-010", "Muestra de procesos no representativa", 2, 3,
     "Muestreo estadístico validado", "COL-011", "Cerrado"),
    ("PRJ-010", "Rotación del equipo auditor a mitad del trabajo", 2, 3,
     "Documentación de papeles de trabajo", "COL-011", "Cerrado"),
    ("PRJ-010", "Sesgo por auditar procesos del propio equipo", 2, 4,
     "Asignación cruzada de auditores", "COL-002", "Cerrado"),
    ("PRJ-010", "Duplicidad con la auditoría externa anual", 2, 2,
     "Coordinación de calendarios con el auditor externo", "COL-002", "Cerrado"),
    ("PRJ-010", "Sobrecosto de viáticos por sedes adicionales", 3, 2,
     "Visitas agrupadas por región", "COL-011", "Mitigado"),
    ("PRJ-010", "Información sensible en los papeles de trabajo", 2, 4,
     "Repositorio cifrado con acceso restringido", "COL-002", "Cerrado"),
    # PRJ-011 · Plataforma de Autoservicio Comercial (en curso)
    ("PRJ-011", "Alcance no definido por el área comercial", 3, 3,
     "Workshops de definición de alcance", "COL-011", "En mitigación"),
    ("PRJ-011", "Requisitos regulatorios de la pasarela de pagos", 3, 4,
     "Asesoría especializada en pagos en línea", "COL-011", "Abierto"),
    ("PRJ-011", "Catálogo de productos desactualizado", 4, 3,
     "Responsable único del catálogo y flujo de aprobación", "COL-003", "Abierto"),
    ("PRJ-011", "Clientes B2B renuentes al autoservicio", 4, 4,
     "Programa piloto con clientes aliados", "COL-011", "Abierto"),
    ("PRJ-011", "Integración con el CRM recién migrado", 3, 4,
     "Esperar la estabilización del CRM antes de integrar", "COL-003", "En mitigación"),
    ("PRJ-011", "Fraude en los pedidos en línea", 2, 4,
     "Validación antifraude y límites por cliente", "COL-011", "Abierto"),
    ("PRJ-011", "Tiempos de entrega no comprometidos con Logística", 3, 3,
     "Acuerdo de niveles de servicio con Logística", "COL-011", "Abierto"),
    ("PRJ-011", "Equipo de desarrollo compartido con otros proyectos", 4, 3,
     "Dedicación mínima acordada del 60%", "COL-003", "Abierto"),
    ("PRJ-011", "Exposición de precios y acuerdos comerciales", 2, 5,
     "Perfiles de acceso por cliente y cifrado", "COL-011", "Abierto"),
    ("PRJ-011", "Plataforma elegida sin soporte local", 2, 3,
     "Contrato de soporte con un partner regional", "COL-003", "Abierto"),
    # PRJ-012 · Programa de Bienestar Laboral (finalizado)
    ("PRJ-012", "Baja participación en las actividades", 3, 3,
     "Actividades en horario laboral y por equipos", "COL-004", "Mitigado"),
    ("PRJ-012", "Actividades percibidas como obligatorias", 2, 2,
     "Inscripción voluntaria con incentivos", "COL-004", "Cerrado"),
    ("PRJ-012", "Presupuesto limitado para actividades externas", 3, 2,
     "Alianzas con la caja de compensación", "COL-004", "Mitigado"),
    ("PRJ-012", "Lesiones en las actividades deportivas", 1, 4,
     "Pólizas y calentamiento dirigido", "COL-004", "Cerrado"),
    ("PRJ-012", "Resultados de clima laboral sin planes de acción", 3, 4,
     "Planes de acción por área con seguimiento", "COL-004", "En mitigación"),
    ("PRJ-012", "Dudas sobre la confidencialidad de la encuesta de clima", 2, 4,
     "Encuesta anónima aplicada por un tercero", "COL-004", "Cerrado"),
    ("PRJ-012", "Choque de fechas con cierres operativos", 3, 2,
     "Calendario validado con los líderes de área", "COL-004", "Mitigado"),
    ("PRJ-012", "Proveedores de bienestar sin formalizar", 2, 3,
     "Verificación de proveedores y contratos marco", "COL-004", "Cerrado"),
    ("PRJ-012", "Expectativas de continuidad sin presupuesto aprobado", 3, 3,
     "Propuesta de programa permanente a la dirección", "COL-004", "Abierto"),
    ("PRJ-012", "Desgaste del equipo organizador del programa", 2, 2,
     "Comité rotativo de bienestar", "COL-004", "Cerrado"),
]


def clasificar_nivel_riesgo(nivel):
    """Clasifica el nivel numérico (probabilidad × impacto, 1-25) en Bajo/Medio/Alto."""
    if nivel >= 15:
        return "Alto"
    if nivel >= 8:
        return "Medio"
    return "Bajo"


DATOS_FICTICIOS_RIESGOS = [
    {
        "ID_Riesgo": f"RSG-{i:03d}", "ID_Proyecto": p,
        "Nombre_Proyecto": _MAPA_NOMBRE_PROY_SEED[p],
        "Descripcion_Riesgo": desc, "Probabilidad": prob, "Impacto": imp,
        "Nivel_Riesgo": prob * imp, "Nivel": clasificar_nivel_riesgo(prob * imp),
        "Plan_Mitigacion": plan,
        "ID_Responsable": resp, "Nombre_Responsable": _MAPA_NOMBRE_COL_SEED[resp],
        "Estado": est,
    }
    for i, (p, desc, prob, imp, plan, resp, est) in enumerate(_RIESGOS_BASE, start=1)
]

# --- Registro de alertas (trazabilidad del sistema de alertas tempranas) ---
COLS_ALERTAS = [
    "ID_Alerta", "Fecha_Hora", "Origen", "Nivel", "ID_Referencia",
    "Nombre_Referencia", "ID_Proyecto", "Semana", "Valor_Indicador",
    "Umbral", "Mensaje"
]

# (fecha_hora, origen, nivel, id_ref, nombre_ref, id_proyecto, semana, valor, umbral, mensaje)
_ALERTAS_BASE = [
    ("2026-06-15 18:00", "Monitor de Salud", "Crítico", "COL-004", "Jorge Castaño", "", "2026-06-15", 107.5, 95,
     "Jorge Castaño al 108% de su capacidad semanal (43.0 h de 40 h)."),
    ("2026-06-22 18:00", "Monitor de Salud", "Crítico", "COL-004", "Jorge Castaño", "", "2026-06-22", 112.5, 95,
     "Jorge Castaño al 113% de su capacidad semanal (45.0 h de 40 h)."),
    ("2026-06-29 18:00", "Monitor de Salud", "Crítico", "COL-004", "Jorge Castaño", "", "2026-06-29", 107.5, 95,
     "Jorge Castaño al 108% de su capacidad semanal (43.0 h de 40 h)."),
    ("2026-06-29 18:00", "Monitor de Salud", "Crítico", "COL-003", "Luisa Mendoza", "", "2026-06-29", 97.1, 95,
     "Luisa Mendoza al 97% de su capacidad semanal (34.0 h de 35 h)."),
    ("2026-06-29 18:00", "Monitor de Salud", "Crítico", "COL-008", "Paula Sánchez", "", "2026-06-29", 97.5, 95,
     "Paula Sánchez al 98% de su capacidad semanal (39.0 h de 40 h)."),
    ("2026-07-06 18:00", "Monitor de Salud", "Crítico", "COL-004", "Jorge Castaño", "", "2026-07-06", 110.0, 95,
     "Jorge Castaño al 110% de su capacidad semanal (44.0 h de 40 h)."),
    ("2026-07-06 18:00", "Monitor de Salud", "Crítico", "COL-003", "Luisa Mendoza", "", "2026-07-06", 100.0, 95,
     "Luisa Mendoza al 100% de su capacidad semanal (35.0 h de 35 h)."),
    ("2026-07-06 18:00", "Monitor de Salud", "Crítico", "COL-008", "Paula Sánchez", "", "2026-07-06", 102.5, 95,
     "Paula Sánchez al 103% de su capacidad semanal (41.0 h de 40 h)."),
    ("2026-07-06 18:05", "Monitor de Salud", "Precaución", "COL-009", "Andrés Vargas", "", "2026-07-06", 90.0, 90,
     "Andrés Vargas al 90% de su capacidad semanal (36.0 h de 40 h)."),
    ("2026-07-06 18:05", "Monitor de Salud", "Precaución", "COL-011", "Felipe Muñoz", "", "2026-07-06", 92.5, 90,
     "Felipe Muñoz al 93% de su capacidad semanal (37.0 h de 40 h)."),
    ("2026-06-10 10:30", "Presupuesto", "Crítico", "PRJ-003", "Migración CRM a la Nube", "PRJ-003", "", 46800, 42000,
     "Lo ejecutado ($46,800) supera el presupuesto planificado ($42,000) en 11.4%."),
    ("2026-06-30 09:15", "Presupuesto", "Precaución", "PRJ-008", "Optimización de Infraestructura TI", "PRJ-008", "", 21500, 20000,
     "Lo ejecutado ($21,500) supera el presupuesto planificado ($20,000) en 7.5%."),
]

DATOS_FICTICIOS_ALERTAS = [
    {
        "ID_Alerta": f"ALT-{i:03d}", "Fecha_Hora": fh, "Origen": org, "Nivel": niv,
        "ID_Referencia": idr, "Nombre_Referencia": nr, "ID_Proyecto": idp,
        "Semana": sem, "Valor_Indicador": val, "Umbral": umb, "Mensaje": msg,
    }
    for i, (fh, org, niv, idr, nr, idp, sem, val, umb, msg) in enumerate(_ALERTAS_BASE, start=1)
]


_RUTA_FAVICON = "favicon.png"
st.set_page_config(
    page_title="Portafolio BI · Gestión de proyectos",
    page_icon=_RUTA_FAVICON if os.path.exists(_RUTA_FAVICON) else "📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# TEMA VISUAL: OSCURO (AZUL NOCHE) / CLARO
# =========================================================
# Opciones del tema nativo de Streamlit (controla fondo, widgets y tablas).
TEMAS_STREAMLIT = {
    "Oscuro": {
        "theme.base": "dark",
        "theme.primaryColor": "#4EE28E",
        "theme.backgroundColor": "#101A47",
        "theme.secondaryBackgroundColor": "#182456",
        "theme.textColor": "#F2F6FF",
    },
    "Claro": {
        "theme.base": "light",
        "theme.primaryColor": "#1FA05F",
        "theme.backgroundColor": "#F4F7FE",
        "theme.secondaryBackgroundColor": "#E9EFFB",
        "theme.textColor": "#16255F",
    },
}

# Variables CSS que consumen los estilos personalizados de más abajo.
PALETAS_CSS = {
    "Oscuro": {
        "panel-1": "#1B2A66", "panel-2": "#16214F", "panel-borde": "#2B3A75",
        "form-1": "#1A2557", "form-2": "#141E4A",
        "sidebar-1": "#0E1840", "sidebar-2": "#131F52",
        "texto-fuerte": "#FFFFFF", "texto-suave": "#93A5D6", "texto-etiqueta": "#C7D4F2",
        "acento-etiqueta": "#4EE28E",
        "celeste": "#8AD8F8", "celeste-bg": "rgba(138,216,248,0.16)",
        "verde": "#6FEBA8", "verde-vivo": "#4EE28E", "verde-bg": "rgba(78,226,142,0.16)",
        "rojo": "#FF8FA6", "rojo-vivo": "#FF5470", "rojo-valor": "#FF6B87",
        "rojo-bg": "rgba(255,84,112,0.18)",
        "ambar": "#FFD37A", "ambar-vivo": "#FFC857", "ambar-bg": "rgba(255,200,87,0.16)",
        "oliva": "#E3D96B", "oliva-bg": "rgba(227,217,107,0.16)",
        "neutro": "#AAB9E4", "neutro-bg": "rgba(159,176,222,0.14)",
        "focus-texto": "#F0DCA8", "focus-strong": "#FFD98A",
        "input-borde": "#33468C", "divisor": "#24336E", "pista": "#1C2A63",
        "sombra": "rgba(4,9,34,0.45)", "sombra-suave": "rgba(4,9,34,0.40)",
    },
    "Claro": {
        "panel-1": "#FFFFFF", "panel-2": "#F4F7FD", "panel-borde": "#D8E0F2",
        "form-1": "#FFFFFF", "form-2": "#F6F8FE",
        "sidebar-1": "#FFFFFF", "sidebar-2": "#EDF2FB",
        "texto-fuerte": "#16255F", "texto-suave": "#5B6C9E", "texto-etiqueta": "#3A4A7E",
        "acento-etiqueta": "#0F8A52",
        "celeste": "#1D6FA8", "celeste-bg": "rgba(29,111,168,0.12)",
        "verde": "#0F8A52", "verde-vivo": "#1FA05F", "verde-bg": "rgba(31,160,95,0.14)",
        "rojo": "#C22945", "rojo-vivo": "#E23A5C", "rojo-valor": "#C22945",
        "rojo-bg": "rgba(226,58,92,0.12)",
        "ambar": "#9A6A00", "ambar-vivo": "#E09B2D", "ambar-bg": "rgba(224,155,45,0.16)",
        "oliva": "#837A18", "oliva-bg": "rgba(131,122,24,0.14)",
        "neutro": "#66739B", "neutro-bg": "rgba(102,115,155,0.12)",
        "focus-texto": "#6E5A17", "focus-strong": "#8A6D0F",
        "input-borde": "#C2CFEA", "divisor": "#D8E0F2", "pista": "#E2E9F8",
        "sombra": "rgba(22,37,95,0.12)", "sombra-suave": "rgba(22,37,95,0.08)",
    },
}

if "tema" not in st.session_state:
    # La sesión hereda el tema que esté activo en el servidor para que la
    # paleta CSS y los widgets nativos siempre coincidan.
    st.session_state["tema"] = (
        "Claro" if st._config.get_option("theme.base") == "light" else "Oscuro"
    )


def aplicar_tema(modo):
    for opcion, valor in TEMAS_STREAMLIT[modo].items():
        if st._config.get_option(opcion) != valor:
            st._config.set_option(opcion, valor)


def selector_tema(key):
    """Interruptor de apariencia. Al cambiar, recarga la app con el tema nuevo."""
    claro = st.toggle(
        "Modo claro",
        value=st.session_state["tema"] == "Claro",
        key=key,
        help="Cambia entre la vista oscura y la clara.",
    )
    nuevo = "Claro" if claro else "Oscuro"
    if nuevo != st.session_state["tema"]:
        st.session_state["tema"] = nuevo
        aplicar_tema(nuevo)
        st.rerun()


aplicar_tema(st.session_state["tema"])

# =========================================================
# ESTILOS PERSONALIZADOS (CSS)
# =========================================================
_paleta_activa = PALETAS_CSS[st.session_state["tema"]]
st.markdown(
    "<style>:root{" + "".join(f"--{n}:{v};" for n, v in _paleta_activa.items()) + "}</style>",
    unsafe_allow_html=True,
)
st.markdown("""
<style>
/* Marca de la plataforma */
.logo-row { display: flex; align-items: center; gap: 11px; margin-bottom: 4px; }
.logo-mark {
    flex-shrink: 0;
    width: 38px; height: 38px;
    border-radius: 11px;
    background: linear-gradient(135deg, #2FBF71, #35A7E0);
    color: #071233;
    font-weight: 800; font-size: 0.92rem; letter-spacing: .5px;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 4px 12px rgba(53,167,224,0.35);
}
.logo-title { font-weight: 800; font-size: 1.22rem; color: var(--texto-fuerte); line-height: 1.15; }
.logo-sub { font-size: 0.72rem; color: var(--texto-suave); margin-top: 1px; }

/* Encabezado principal con banda de color */
.banner {
    background: linear-gradient(90deg, #16255F 0%, #1E3C8F 60%, #2853B4 100%);
    border: 1px solid #2E4390;
    padding: 22px 30px;
    border-radius: 12px;
    margin-bottom: 22px;
    box-shadow: 0 10px 30px rgba(4, 9, 34, 0.55);
}
.banner h1 { color: #FFFFFF; margin: 0; font-size: 1.6rem; }
.banner p { color: #A9BCEE; margin: 4px 0 0 0; font-size: 0.95rem; }

/* Tarjetas de métricas */
.metric-card {
    background: linear-gradient(180deg, var(--panel-1) 0%, var(--panel-2) 100%);
    border: 1px solid var(--panel-borde);
    border-left: 5px solid var(--celeste);
    border-radius: 10px;
    padding: 14px 18px;
    box-shadow: 0 4px 16px var(--sombra);
}
.metric-card .label {
    color: var(--acento-etiqueta); font-size: 0.8rem; text-transform: uppercase; letter-spacing: .5px;
    overflow-wrap: normal; word-break: normal; hyphens: none;
}
.metric-card .value {
    color: var(--texto-fuerte); font-size: 1.7rem; font-weight: 700; margin-top: 2px;
    white-space: nowrap; overflow-wrap: normal; word-break: normal;
}
.metric-card.ok { border-left-color: var(--verde-vivo); }
.metric-card.warn { border-left-color: var(--ambar-vivo); }
.metric-card.risk { border-left-color: var(--rojo-vivo); }
.metric-card.risk .value { color: var(--rojo-valor); }

/* Badges de estado */
.badge { padding: 3px 12px; border-radius: 20px; font-size: 0.78rem; font-weight: 600; }
.badge-curso { background: var(--celeste-bg); color: var(--celeste); }
.badge-riesgo { background: var(--rojo-bg); color: var(--rojo); }
.badge-fin { background: var(--verde-bg); color: var(--verde); }
.badge-noini { background: var(--neutro-bg); color: var(--neutro); }
.badge-carga-libre { background: var(--neutro-bg); color: var(--neutro); }
.badge-carga-normal { background: var(--verde-bg); color: var(--verde); }
.badge-carga-alta { background: var(--rojo-bg); color: var(--rojo); }

/* Callout de foco temático (conexión con el Monitor de Salud) */
.focus-callout {
    background: linear-gradient(135deg, rgba(255,200,87,0.12) 0%, rgba(255,200,87,0.05) 100%);
    border: 1px solid rgba(255,200,87,0.35);
    border-left: 5px solid var(--ambar-vivo);
    border-radius: 12px;
    padding: 14px 20px;
    margin-bottom: 20px;
    font-size: 0.9rem;
    line-height: 1.5;
    color: var(--focus-texto);
}
.focus-callout strong { color: var(--focus-strong); }

/* Avatar circular con iniciales para tarjetas de colaboradores */
.avatar-circle {
    width: 44px; height: 44px;
    min-width: 44px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 0.9rem; color: white;
    background: linear-gradient(135deg, #35A7E0, #2853B4);
}
.avatar-circle.alta { background: linear-gradient(135deg, #E8344F, #FF7A6B); }
.avatar-circle.libre { background: linear-gradient(135deg, #46557F, #6B7BA6); }
.avatar-circle.warn { background: linear-gradient(135deg, #D9922B, #FFC857); }

/* Semáforo de saturación (Monitor de Salud) */
.badge-sat-ok { background: var(--verde-bg); color: var(--verde); }
.badge-sat-watch { background: var(--oliva-bg); color: var(--oliva); }
.badge-sat-warn { background: var(--ambar-bg); color: var(--ambar); }
.badge-sat-crit { background: var(--rojo-bg); color: var(--rojo); }
.badge-sat-sub { background: var(--neutro-bg); color: var(--neutro); }
.sat-track {
    background: var(--pista);
    border-radius: 8px;
    height: 10px;
    width: 100%;
    margin-top: 8px;
    overflow: hidden;
}
.sat-fill { height: 10px; border-radius: 8px; }
.sat-fill.ok { background: linear-gradient(90deg, #2FBF71, #63E6A4); }
.sat-fill.watch { background: linear-gradient(90deg, #C8B834, #E8DA5C); }
.sat-fill.warn { background: linear-gradient(90deg, #E09B2D, #FFC857); }
.sat-fill.crit { background: linear-gradient(90deg, #E8344F, #FF7A8C); }
.sat-fill.sub { background: linear-gradient(90deg, #46557F, #8A9BC7); }

/* Botones */
.stButton > button, .stFormSubmitButton > button {
    border-radius: 8px;
    font-weight: 600;
}
.stButton > button[kind="primary"] {
    color: #071233 !important;
    font-weight: 700;
}

/* =====================================================
   FORMULARIO DE CAPTURA — estilo "tarjeta premium"
   ===================================================== */
div[data-testid="stForm"] {
    background: linear-gradient(180deg, var(--form-1) 0%, var(--form-2) 100%);
    border: 1px solid var(--panel-borde);
    border-radius: 18px;
    padding: 30px 34px 22px 34px;
    box-shadow: 0 10px 30px var(--sombra);
}

/* Encabezados de sección dentro del formulario */
.form-section-header {
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 22px 0 16px 0;
    padding-bottom: 10px;
    border-bottom: 2px solid var(--divisor);
}
.form-section-header:first-of-type { margin-top: 2px; }
.form-section-header .icon {
    flex-shrink: 0;
    font-size: 0.82rem;
    font-weight: 800;
    letter-spacing: .5px;
    background: linear-gradient(135deg, #2FBF71, #35A7E0);
    color: white;
    width: 36px; height: 36px;
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 3px 10px rgba(53,167,224,0.35);
}
.form-section-header .title {
    font-weight: 700;
    color: var(--texto-fuerte);
    font-size: 1.02rem;
    line-height: 1.2;
}
.form-section-header .subtitle {
    color: var(--texto-suave);
    font-size: 0.8rem;
    margin-top: 1px;
}

/* Inputs con bordes redondeados y foco resaltado */
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input,
div[data-testid="stDateInput"] input,
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    border-radius: 10px !important;
    border: 1.5px solid var(--input-borde) !important;
    transition: border-color .15s ease, box-shadow .15s ease;
}
div[data-testid="stTextInput"] input:focus,
div[data-testid="stNumberInput"] input:focus,
div[data-testid="stDateInput"] input:focus {
    border-color: var(--verde-vivo) !important;
    box-shadow: 0 0 0 3px var(--verde-bg) !important;
}
div[data-testid="stWidgetLabel"] p {
    font-weight: 600 !important;
    color: var(--texto-etiqueta) !important;
    font-size: 0.86rem !important;
}

/* Botón de envío destacado con degradado */
.stFormSubmitButton > button {
    background: linear-gradient(90deg, #2FBF71, #35A7E0);
    color: #071233;
    border: none;
    padding: 13px 0;
    font-size: 1.02rem;
    font-weight: 700;
    box-shadow: 0 6px 18px rgba(47,191,113,0.30);
    transition: transform .12s ease, box-shadow .12s ease;
}
.stFormSubmitButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 24px rgba(53,167,224,0.40);
    color: #071233;
}

/* Mini-tarjetas de resumen rápido */
.mini-stat {
    background: linear-gradient(180deg, var(--panel-1) 0%, var(--panel-2) 100%);
    border: 1px solid var(--panel-borde);
    border-radius: 14px;
    padding: 12px 10px;
    text-align: center;
    box-shadow: 0 4px 14px var(--sombra-suave);
}
.mini-stat .n {
    font-size: 1.3rem; font-weight: 800; color: var(--texto-fuerte);
    white-space: nowrap; overflow-wrap: normal; word-break: normal;
}
.mini-stat .l {
    font-size: 0.68rem; color: var(--acento-etiqueta); text-transform: uppercase;
    letter-spacing: .3px; margin-top: 3px;
    overflow-wrap: normal; word-break: normal; hyphens: none;
}

/* Encabezado de la tabla de resultados */
.table-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: 8px 0 12px 0;
}
.table-card-header .title { font-weight: 700; font-size: 1.08rem; color: var(--texto-fuerte); }
.count-pill {
    background: var(--celeste-bg); color: var(--celeste); font-weight: 700;
    padding: 4px 14px; border-radius: 20px; font-size: 0.8rem;
}

/* Ambiente general (azul noche o claro según el tema) */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--sidebar-1) 0%, var(--sidebar-2) 100%);
    border-right: 1px solid var(--divisor);
}
h1, h2, h3, h4 { color: var(--texto-fuerte); }
hr { border-color: var(--divisor); }
</style>
""", unsafe_allow_html=True)


# =========================================================
# FUNCIONES DE DATOS
# =========================================================
def leer_hoja(nombre_hoja, columnas):
    if os.path.exists(ARCHIVO_EXCEL):
        try:
            df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=nombre_hoja)
            # Si al Excel le faltan columnas nuevas (versión anterior del
            # archivo o edición manual), las crea vacías en vez de fallar,
            # y reordena todo al formato oficial.
            for col in columnas:
                if col not in df.columns:
                    df[col] = 0 if col.startswith(("Presupuesto", "Avance", "Horas", "Capacidad", "Costo", "Monto", "Beneficio", "Valor", "Umbral")) else ""
            return df[columnas]
        except ValueError:
            return pd.DataFrame(columns=columnas)
    return pd.DataFrame(columns=columnas)


def guardar_hoja(nombre_hoja, df):
    """Guarda una hoja SIN borrar las demás hojas del Excel."""
    if os.path.exists(ARCHIVO_EXCEL):
        with pd.ExcelWriter(
            ARCHIVO_EXCEL, engine="openpyxl",
            mode="a", if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    else:
        df.to_excel(ARCHIVO_EXCEL, sheet_name=nombre_hoja, index=False)


def inicializar_datos():
    hojas_semilla = [
        ("Proyectos", COLS_PROYECTOS, DATOS_FICTICIOS_PROYECTOS),
        ("Colaboradores", COLS_COLABORADORES, DATOS_FICTICIOS_COLABORADORES),
        ("Salud_Equipo", COLS_SALUD, DATOS_FICTICIOS_SALUD),
        ("Presupuesto", COLS_PRESUPUESTO, DATOS_FICTICIOS_PRESUPUESTO),
        ("Tareas_Hitos", COLS_TAREAS, DATOS_FICTICIOS_TAREAS),
        ("Riesgos", COLS_RIESGOS, DATOS_FICTICIOS_RIESGOS),
        ("Alertas", COLS_ALERTAS, DATOS_FICTICIOS_ALERTAS),
    ]
    for nombre, columnas, datos in hojas_semilla:
        if leer_hoja(nombre, columnas).empty:
            guardar_hoja(nombre, pd.DataFrame(datos))


def siguiente_id(df, columna, prefijo):
    """Sugiere el próximo ID disponible con formato PREFIJO-XXX."""
    patron = re.compile(rf"{prefijo}-(\d+)")
    numeros = [int(m.group(1)) for v in df[columna].astype(str) if (m := patron.match(v))]
    return f"{prefijo}-{(max(numeros) + 1 if numeros else 1):03d}"


def parsear_proyectos(cadena):
    """Convierte 'PRJ-001,PRJ-002' en ['PRJ-001', 'PRJ-002'], ignorando vacíos.

    Las celdas vacías de Excel llegan como NaN (no como ""), por lo que se
    valida con pd.isna() antes de intentar separar por comas.
    """
    if pd.isna(cadena) or str(cadena).strip() == "":
        return []
    return [p.strip() for p in str(cadena).split(",") if p.strip()]


def nivel_carga(num_proyectos):
    """Clasifica la carga de trabajo preliminar según cantidad de proyectos asignados.

    Es una vista previa simple (conteo de proyectos); el cálculo real de
    saturación (horas asignadas vs. reales) se hace en el Monitor de Salud.
    """
    if num_proyectos == 0:
        return "Sin asignar", "badge-carga-libre", "libre"
    elif num_proyectos <= 2:
        return "Carga normal", "badge-carga-normal", "normal"
    else:
        return "Carga alta", "badge-carga-alta", "alta"


def lunes_de_la_semana(fecha):
    """Devuelve el lunes de la semana a la que pertenece la fecha dada."""
    return fecha - timedelta(days=fecha.weekday())


def clasificar_saturacion(pct):
    """Clasifica el % de saturación según los umbrales del semáforo.

    Devuelve (etiqueta, clase_badge, clase_barra, clase_avatar).
    Umbrales alineados con la memoria del TFM:
    <60% capacidad disponible · 60-80% saludable · 80-90% seguimiento ·
    90-95% precaución · 95% crítico.
    """
    if pct >= 95:
        return "Crítico", "badge-sat-crit", "crit", "alta"
    if pct >= 90:
        return "Precaución", "badge-sat-warn", "warn", "warn"
    if pct >= 80:
        return "Seguimiento", "badge-sat-watch", "watch", ""
    if pct >= 60:
        return "Saludable", "badge-sat-ok", "ok", ""
    return "Capacidad disponible", "badge-sat-sub", "sub", "libre"


def registrar_alerta(origen, nivel, id_ref, nombre_ref, id_proyecto, semana, valor, umbral, mensaje):
    """Deja registro de una alerta en la hoja «Alertas» (trazabilidad del sistema).

    Este historial alimenta el KPI de alertas activas en Power BI y sirve de
    fuente para los flujos de notificación de Power Automate.
    """
    df_a = leer_hoja("Alertas", COLS_ALERTAS)
    nueva = {
        "ID_Alerta": siguiente_id(df_a, "ID_Alerta", "ALT"),
        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Origen": origen,
        "Nivel": nivel,
        "ID_Referencia": id_ref,
        "Nombre_Referencia": nombre_ref,
        "ID_Proyecto": id_proyecto,
        "Semana": semana,
        "Valor_Indicador": valor,
        "Umbral": umbral,
        "Mensaje": mensaje,
    }
    guardar_hoja("Alertas", pd.concat([df_a, pd.DataFrame([nueva])], ignore_index=True))


def etiqueta_semana(semana_iso):
    """Convierte '2026-07-06' en 'Semana del 06/07/2026'."""
    s = str(semana_iso)[:10]
    return f"Semana del {s[8:10]}/{s[5:7]}/{s[0:4]}"


PERCIBIDA_OPCIONES = ["1 · Muy ligera", "2 · Ligera", "3 · Equilibrada", "4 · Pesada", "5 · Muy pesada"]
PERCIBIDA_EMOJI = {1: "😌", 2: "🙂", 3: "😐", 4: "😰", 5: "🥵"}


def badge_estado(estado):
    clases = {
        "En curso": "badge-curso", "En riesgo": "badge-riesgo",
        "Finalizado": "badge-fin", "No iniciado": "badge-noini",
    }
    return f'<span class="badge {clases.get(estado, "badge-noini")}">{estado}</span>'


# =========================================================
# CARGA DE DATOS
# =========================================================
inicializar_datos()
df_proyectos = leer_hoja("Proyectos", COLS_PROYECTOS)
df_colaboradores = leer_hoja("Colaboradores", COLS_COLABORADORES)
df_salud = leer_hoja("Salud_Equipo", COLS_SALUD)
# Normalización defensiva: si el Excel fue editado a mano, las semanas pueden
# volver como datetime y las horas como texto.
df_salud["Semana"] = df_salud["Semana"].astype(str).str[:10]
for _c in ("Horas_Asignadas", "Horas_Reales", "Capacidad_Horas_Semana"):
    df_salud[_c] = pd.to_numeric(df_salud[_c], errors="coerce").fillna(0)
df_salud["Carga_Percibida"] = (
    pd.to_numeric(df_salud["Carga_Percibida"], errors="coerce").fillna(3).clip(1, 5).astype(int)
)
df_presupuesto = leer_hoja("Presupuesto", COLS_PRESUPUESTO)
df_presupuesto["Monto"] = pd.to_numeric(df_presupuesto["Monto"], errors="coerce").fillna(0)
df_tareas = leer_hoja("Tareas_Hitos", COLS_TAREAS)
df_riesgos = leer_hoja("Riesgos", COLS_RIESGOS)
for _c in ("Probabilidad", "Impacto", "Nivel_Riesgo"):
    df_riesgos[_c] = pd.to_numeric(df_riesgos[_c], errors="coerce").fillna(1).astype(int)


# =========================================================
# ACCESO MULTIUSUARIO: PM O COLABORADOR
# =========================================================
if "rol" not in st.session_state:
    _, col_tema_login = st.columns([5, 1])
    with col_tema_login:
        selector_tema("tema_login")
    st.markdown("""
    <div class="banner" style="display:flex;align-items:center;gap:16px;">
        <div class="logo-mark" style="width:48px;height:48px;font-size:1.1rem;border-radius:13px;">PB</div>
        <div>
            <h1>Portafolio BI · Acceso</h1>
            <p>Selecciona tu perfil para entrar a la plataforma.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_pm, col_colab = st.columns(2)
    with col_pm:
        with st.container(border=True):
            st.markdown("#### Project Manager")
            st.caption(
                "Administra los proyectos, el equipo, el presupuesto detallado, las tareas "
                "y los riesgos (Módulos 1, 2, 4, 5 y 6), y supervisa el Monitor de Salud "
                "de todo el equipo."
            )
            if st.button("Entrar como PM", use_container_width=True, type="primary"):
                st.session_state["rol"] = "PM"
                st.rerun()
    with col_colab:
        with st.container(border=True):
            st.markdown("#### Colaborador")
            st.caption(
                "Reporta tus horas de la semana y cómo sentiste la carga (Módulo 3). "
                "Solo verás y editarás tus propios registros."
            )
            df_activos_login = df_colaboradores[df_colaboradores["Estado"] == "Activo"]
            if df_activos_login.empty:
                st.info("Todavía no hay colaboradores activos. Pide al PM que te registre en el Módulo 2.")
            else:
                opciones_login = [f"{r.ID_Colaborador} · {r.Nombre}" for r in df_activos_login.itertuples()]
                mapa_login = dict(zip(opciones_login, df_activos_login["ID_Colaborador"].astype(str)))
                sel_login = st.selectbox("¿Quién eres?", opciones_login, key="sel_login_colab")
                if st.button("Entrar como colaborador", use_container_width=True):
                    st.session_state["rol"] = "Colaborador"
                    st.session_state["id_colaborador"] = mapa_login[sel_login]
                    st.session_state["nombre_colaborador"] = sel_login.split("·", 1)[1].strip()
                    st.rerun()

    st.caption("Prototipo académico: el acceso identifica a la persona para personalizar la captura; no requiere contraseña.")
    st.stop()

es_pm = st.session_state["rol"] == "PM"
id_usuario = st.session_state.get("id_colaborador", "")
nombre_usuario = st.session_state.get("nombre_colaborador", "")

# Si el colaborador fue eliminado o el archivo se reinició, se cierra la sesión.
if not es_pm and id_usuario not in df_colaboradores["ID_Colaborador"].astype(str).values:
    for _k in ("rol", "id_colaborador", "nombre_colaborador"):
        st.session_state.pop(_k, None)
    st.rerun()

# =========================================================
# NAVEGACIÓN LATERAL
# =========================================================
with st.sidebar:
    st.markdown(
        '<div class="logo-row"><div class="logo-mark">PB</div>'
        '<div><div class="logo-title">Portafolio BI</div>'
        '<div class="logo-sub">Gestión de proyectos y talento humano</div></div></div>',
        unsafe_allow_html=True,
    )
    selector_tema("tema_sidebar")
    if es_pm:
        st.markdown("**Sesión:** Project Manager")
        st.caption("Administras los módulos del portafolio y supervisas la salud del equipo.")
    else:
        st.markdown(f"**Sesión:** {nombre_usuario}")
        st.caption("Tu acceso está enfocado en reportar tus horas y tu carga percibida.")
    if st.button("Cerrar sesión", use_container_width=True):
        for _k in ("rol", "id_colaborador", "nombre_colaborador"):
            st.session_state.pop(_k, None)
        st.rerun()
    st.divider()
    if es_pm:
        modulo = st.radio(
            "Módulos",
            [
                "Inicio / Resumen",
                "1 · Proyectos",
                "2 · Colaboradores",
                "3 · Monitor de Salud",
                "4 · Presupuesto detallado",
                "5 · Tareas e hitos",
                "6 · Riesgos",
            ],
        )
    else:
        modulo = "3 · Monitor de Salud"
        st.markdown("**Módulos**")
        st.markdown("3 · Mi registro semanal")
    st.divider()
    st.caption(f"Fuente de datos: {ARCHIVO_EXCEL}")
    st.caption(
        f"{len(df_proyectos)} proyectos · {len(df_colaboradores)} colaboradores · "
        f"{len(df_tareas)} tareas · {len(df_riesgos)} riesgos"
    )
    if es_pm and os.path.exists(ARCHIVO_EXCEL):
        with open(ARCHIVO_EXCEL, "rb") as _f:
            st.download_button(
                "Descargar base de datos (Excel)",
                _f.read(),
                file_name=ARCHIVO_EXCEL,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="El mismo archivo que alimenta el tablero directivo en Power BI.",
            )
    if es_pm:
        with st.expander("Mantenimiento"):
            if st.button("Reiniciar sistema"):
                if os.path.exists(ARCHIVO_EXCEL):
                    os.remove(ARCHIVO_EXCEL)
                for _k in ("rol", "id_colaborador", "nombre_colaborador"):
                    st.session_state.pop(_k, None)
                st.rerun()


# =========================================================
# MÓDULO: INICIO / RESUMEN
# =========================================================
if modulo.startswith("Inicio"):
    st.markdown("""
    <div class="banner">
        <h1>Portafolio de proyectos</h1>
        <p>Vista general del estado del portafolio y la salud del equipo</p>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("¿Cómo usar la plataforma? · Guía rápida"):
        st.markdown("""
- **Inicio / Resumen** — panorama del portafolio: estado, avance y presupuesto de cada proyecto.
- **1 · Proyectos** — registra y actualiza los proyectos: fechas, presupuesto, avance y estado.
- **2 · Colaboradores** — administra el equipo, sus roles y sus asignaciones a proyectos.
- **3 · Monitor de Salud** — supervisa la saturación semanal del equipo; cada colaborador reporta sus horas y su carga desde su propio acceso.
- **4 · Presupuesto detallado** — define las partidas por rubro y registra los gastos reales contra ellas.
- **5 · Tareas e hitos** — controla el cronograma: qué venció, qué vence pronto y qué se completó.
- **6 · Riesgos** — inventario de riesgos por proyecto con matriz de probabilidad × impacto.

Todo lo que se captura se guarda al instante y alimenta el tablero directivo en Power BI.
Las situaciones que requieren atención (sobrecarga del equipo, sobrecostos, retrasos o riesgos
altos) quedan registradas automáticamente para su seguimiento y notificación.
""")

    total = len(df_proyectos)
    en_curso = int((df_proyectos["Estado_Proyecto"] == "En curso").sum()) if total else 0
    en_riesgo = int((df_proyectos["Estado_Proyecto"] == "En riesgo").sum()) if total else 0
    presupuesto_total = df_proyectos["Presupuesto_Planificado"].sum() if total else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-card"><div class="label">Proyectos totales</div><div class="value">{total}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card ok"><div class="label">En curso</div><div class="value">{en_curso}</div></div>', unsafe_allow_html=True)
    with c3:
        clase = "risk" if en_riesgo > 0 else "ok"
        st.markdown(f'<div class="metric-card {clase}"><div class="label">En riesgo</div><div class="value">{en_riesgo}</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card"><div class="label">Presupuesto total</div><div class="value">${presupuesto_total:,.0f}</div></div>', unsafe_allow_html=True)

    st.write("")
    st.write("#### Estado del portafolio")
    if df_proyectos.empty:
        st.info("Todavía no hay proyectos. Ve al módulo 1 · Proyectos para capturar el primero.")
    else:
        for _, p in df_proyectos.iterrows():
            with st.container(border=True):
                a, b, c = st.columns([3, 2, 2])
                with a:
                    st.markdown(f"**{p['Nombre_Proyecto']}**  \n{p['Cliente_Area']} · Líder: {p['Lider_Proyecto']}")
                with b:
                    st.markdown(badge_estado(p["Estado_Proyecto"]), unsafe_allow_html=True)
                    st.caption(f"Prioridad: {p['Prioridad']}")
                with c:
                    st.progress(int(p["Avance_Pct"]) / 100, text=f"Avance {int(p['Avance_Pct'])}%")


# =========================================================
# MÓDULO 1: PROYECTOS
# =========================================================
elif modulo.startswith("1"):
    st.markdown("""
    <div class="banner">
        <h1>Módulo 1 · Captura de proyectos</h1>
        <p>Registra y administra todos los proyectos de tu portafolio en un solo lugar.</p>
    </div>
    """, unsafe_allow_html=True)

    # ---- Mensajes pendientes de la última acción (creación/edición/borrado) ----
    if "flash_proyecto" in st.session_state:
        for tipo, texto in st.session_state.pop("flash_proyecto"):
            getattr(st, tipo)(texto)

    # ---- Resumen rápido ----
    total_actual = len(df_proyectos)
    en_curso_actual = int((df_proyectos["Estado_Proyecto"] == "En curso").sum()) if total_actual else 0
    en_riesgo_actual = int((df_proyectos["Estado_Proyecto"] == "En riesgo").sum()) if total_actual else 0
    presupuesto_total_actual = df_proyectos["Presupuesto_Planificado"].sum() if total_actual else 0

    s1, s2, s3, s4 = st.columns(4)
    for col, valor, etiqueta in zip(
        (s1, s2, s3, s4),
        (total_actual, en_curso_actual, en_riesgo_actual, f"${presupuesto_total_actual:,.0f}"),
        ("Proyectos registrados", "En curso", "En riesgo", "Presupuesto total"),
    ):
        with col:
            st.markdown(
                f'<div class="mini-stat"><div class="n">{valor}</div><div class="l">{etiqueta}</div></div>',
                unsafe_allow_html=True,
            )

    st.write("")

    with st.form("form_proyecto", clear_on_submit=True):
        st.markdown("""
        <div class="form-section-header">
            <div class="icon">01</div>
            <div>
                <div class="title">Información general</div>
                <div class="subtitle">Identificación y responsables del proyecto</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            id_proj = st.text_input(
                "ID Proyecto *",
                value=siguiente_id(df_proyectos, "ID_Proyecto", "PRJ"),
                help="Se autogenera con el siguiente consecutivo disponible. Puedes cambiarlo si tu organización usa otro formato.",
            )
            cliente = st.text_input("Cliente / Área *", placeholder="Finanzas, Operaciones...")
        with col2:
            nombre = st.text_input("Nombre del proyecto *", placeholder="Implementación Power BI")
            lider = st.text_input("Líder del proyecto *", placeholder="Nombre y apellido")
        col_d1, col_d2 = st.columns([2, 1])
        with col_d1:
            descripcion = st.text_input(
                "Descripción / objetivo del proyecto",
                placeholder="¿Qué busca lograr este proyecto? (opcional pero recomendado)",
            )
        with col_d2:
            tecnologia = st.text_input(
                "Tecnología / categoría",
                placeholder="ERP, Analítica, Formación...",
                help="Permite agrupar y filtrar los proyectos por tipo en los reportes.",
            )

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">02</div>
            <div>
                <div class="title">Cronograma</div>
                <div class="subtitle">Fechas de inicio, cierre planificado y cierre real</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col3, col4 = st.columns(2)
        with col3:
            fecha_inicio = st.date_input("Fecha de inicio", value=date.today())
        with col4:
            fecha_fin = st.date_input(
                "Fecha fin planificada", value=date.today(),
                help="Debe ser igual o posterior a la fecha de inicio.",
            )
        ya_finalizo = st.checkbox(
            "El proyecto ya terminó — registrar la fecha real de cierre",
            help="La fecha real de cierre permite comparar lo planificado contra lo ejecutado (desviación de cronograma).",
        )
        fecha_fin_real = st.date_input("Fecha fin real", value=date.today())

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">03</div>
            <div>
                <div class="title">Presupuesto</div>
                <div class="subtitle">Montos planificados y ejecutados</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col5, col6, col6b = st.columns(3)
        with col5:
            presupuesto = st.number_input("Presupuesto planificado ($)", min_value=0.0, step=500.0)
        with col6:
            ejecutado = st.number_input(
                "Presupuesto ejecutado ($)", min_value=0.0, step=500.0,
                help="Si el proyecto apenas inicia, puedes dejarlo en 0.",
            )
        with col6b:
            beneficio = st.number_input(
                "Beneficio esperado ($)", min_value=0.0, step=500.0,
                help="Valor o ahorro que se espera obtener del proyecto. Permite calcular el retorno de la inversión (ROI) en los reportes.",
            )

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">04</div>
            <div>
                <div class="title">Estado y avance</div>
                <div class="subtitle">Situación actual del proyecto</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        avance = st.slider("Avance (%)", 0, 100, 0)
        col7, col8 = st.columns(2)
        with col7:
            estado = st.selectbox("Estado", ["No iniciado", "En curso", "En riesgo", "Finalizado"])
        with col8:
            prioridad = st.selectbox("Prioridad", ["Alta", "Media", "Baja"])

        st.write("")
        submitted = st.form_submit_button("Guardar proyecto", use_container_width=True)

        if submitted:
            errores = []
            if not id_proj.strip():
                errores.append("El ID del proyecto es obligatorio.")
            elif id_proj.strip() in df_proyectos["ID_Proyecto"].astype(str).values:
                errores.append(f"El ID '{id_proj}' ya existe. Usa uno diferente.")
            if not nombre.strip():
                errores.append("El nombre del proyecto es obligatorio.")
            if not cliente.strip():
                errores.append("El cliente/área es obligatorio.")
            if not lider.strip():
                errores.append("El líder del proyecto es obligatorio.")
            if fecha_fin < fecha_inicio:
                errores.append("La fecha fin no puede ser anterior a la fecha de inicio.")

            if errores:
                for e in errores:
                    st.error(e)
            else:
                nuevo = {
                    "ID_Proyecto": id_proj.strip(),
                    "Nombre_Proyecto": nombre.strip(),
                    "Descripcion": descripcion.strip(),
                    "Cliente_Area": cliente.strip(),
                    "Tecnologia": tecnologia.strip(),
                    "Lider_Proyecto": lider.strip(),
                    "Fecha_Inicio": fecha_inicio.strftime("%Y-%m-%d"),
                    "Fecha_Fin_Planificada": fecha_fin.strftime("%Y-%m-%d"),
                    "Fecha_Fin_Real": fecha_fin_real.strftime("%Y-%m-%d") if ya_finalizo else "",
                    "Presupuesto_Planificado": presupuesto,
                    "Presupuesto_Ejecutado": ejecutado,
                    "Beneficio_Esperado": beneficio,
                    "Avance_Pct": avance,
                    "Estado_Proyecto": estado,
                    "Prioridad": prioridad,
                }
                df_actualizado = pd.concat(
                    [df_proyectos, pd.DataFrame([nuevo])], ignore_index=True
                )
                guardar_hoja("Proyectos", df_actualizado)

                mensajes = [("success", f"Proyecto '{nombre}' guardado correctamente.")]
                if ejecutado > presupuesto and presupuesto > 0:
                    mensajes.append(("warning", "El presupuesto ejecutado supera al planificado. Se guardó igual, pero revisa el dato."))
                if estado == "Finalizado" and avance != 100:
                    mensajes.append(("info", f"Marcaste el proyecto como «Finalizado» pero el avance quedó en {avance}%. Puedes editarlo y ajustarlo a 100% si ya terminó."))
                elif estado == "No iniciado" and avance != 0:
                    mensajes.append(("info", f"El proyecto está «No iniciado» pero el avance quedó en {avance}%. Revisa si el estado debería ser «En curso»."))
                if estado == "Finalizado" and not ya_finalizo:
                    mensajes.append(("info", "El proyecto quedó como «Finalizado» pero no registraste la fecha real de cierre. Puedes agregarla editando el proyecto."))
                st.session_state["flash_proyecto"] = mensajes
                st.rerun()

    # ---- Editar o eliminar un proyecto existente ----
    st.write("")
    with st.expander("Editar o eliminar un proyecto existente", expanded=not df_proyectos.empty):
        if df_proyectos.empty:
            st.caption("Todavía no hay proyectos capturados para editar.")
        else:
            etiquetas = [f"{r.ID_Proyecto} · {r.Nombre_Proyecto}" for r in df_proyectos.itertuples()]
            mapa_id = dict(zip(etiquetas, df_proyectos["ID_Proyecto"].astype(str)))
            etiqueta_sel = st.selectbox("Selecciona un proyecto", etiquetas, key="sel_editar_proyecto")
            id_sel = mapa_id[etiqueta_sel]
            fila = df_proyectos[df_proyectos["ID_Proyecto"].astype(str) == id_sel].iloc[0]

            estados_posibles = ["No iniciado", "En curso", "En riesgo", "Finalizado"]
            prioridades_posibles = ["Alta", "Media", "Baja"]
            try:
                f_inicio_val = pd.to_datetime(fila["Fecha_Inicio"]).date()
            except (ValueError, TypeError):
                f_inicio_val = date.today()
            try:
                f_fin_val = pd.to_datetime(fila["Fecha_Fin_Planificada"]).date()
            except (ValueError, TypeError):
                f_fin_val = date.today()

            fin_real_str = str(fila["Fecha_Fin_Real"])[:10] if pd.notna(fila["Fecha_Fin_Real"]) else ""
            tiene_fin_real = bool(fin_real_str.strip()) and fin_real_str.lower() != "nan"
            try:
                f_fin_real_val = pd.to_datetime(fin_real_str).date() if tiene_fin_real else date.today()
            except (ValueError, TypeError):
                f_fin_real_val = date.today()

            with st.form(f"form_editar_proyecto_{id_sel}"):
                col1, col2 = st.columns(2)
                with col1:
                    e_id = st.text_input("ID Proyecto *", value=str(fila["ID_Proyecto"]), key=f"e_id_{id_sel}")
                    e_cliente = st.text_input("Cliente / Área *", value=str(fila["Cliente_Area"]), key=f"e_cliente_{id_sel}")
                    e_inicio = st.date_input("Fecha de inicio", value=f_inicio_val, key=f"e_inicio_{id_sel}")
                    e_presupuesto = st.number_input(
                        "Presupuesto planificado ($)", min_value=0.0, step=500.0,
                        value=float(fila["Presupuesto_Planificado"] or 0), key=f"e_presupuesto_{id_sel}",
                    )
                with col2:
                    e_nombre = st.text_input("Nombre del proyecto *", value=str(fila["Nombre_Proyecto"]), key=f"e_nombre_{id_sel}")
                    e_lider = st.text_input("Líder del proyecto *", value=str(fila["Lider_Proyecto"]), key=f"e_lider_{id_sel}")
                    e_fin = st.date_input("Fecha fin planificada", value=f_fin_val, key=f"e_fin_{id_sel}")
                    e_ejecutado = st.number_input(
                        "Presupuesto ejecutado ($)", min_value=0.0, step=500.0,
                        value=float(fila["Presupuesto_Ejecutado"] or 0), key=f"e_ejecutado_{id_sel}",
                    )
                e_descripcion = st.text_input(
                    "Descripción / objetivo del proyecto",
                    value=str(fila["Descripcion"]) if pd.notna(fila["Descripcion"]) else "",
                    key=f"e_desc_{id_sel}",
                )
                col_t1, col_t2 = st.columns(2)
                with col_t1:
                    e_tecnologia = st.text_input(
                        "Tecnología / categoría",
                        value=str(fila["Tecnologia"]) if pd.notna(fila["Tecnologia"]) else "",
                        key=f"e_tec_{id_sel}",
                    )
                with col_t2:
                    e_beneficio = st.number_input(
                        "Beneficio esperado ($)", min_value=0.0, step=500.0,
                        value=float(pd.to_numeric(fila["Beneficio_Esperado"], errors="coerce") or 0),
                        key=f"e_benef_{id_sel}",
                    )
                col_fr1, col_fr2 = st.columns(2)
                with col_fr1:
                    e_ya_finalizo = st.checkbox(
                        "El proyecto ya terminó", value=tiene_fin_real, key=f"e_yafin_{id_sel}",
                    )
                with col_fr2:
                    e_fin_real = st.date_input("Fecha fin real", value=f_fin_real_val, key=f"e_finreal_{id_sel}")

                e_avance = st.slider("Avance (%)", 0, 100, int(fila["Avance_Pct"] or 0), key=f"e_avance_{id_sel}")
                col3, col4 = st.columns(2)
                with col3:
                    e_estado = st.selectbox(
                        "Estado", estados_posibles,
                        index=estados_posibles.index(fila["Estado_Proyecto"]) if fila["Estado_Proyecto"] in estados_posibles else 0,
                        key=f"e_estado_{id_sel}",
                    )
                with col4:
                    e_prioridad = st.selectbox(
                        "Prioridad", prioridades_posibles,
                        index=prioridades_posibles.index(fila["Prioridad"]) if fila["Prioridad"] in prioridades_posibles else 0,
                        key=f"e_prioridad_{id_sel}",
                    )

                e_confirmar_borrado = st.checkbox(
                    "Confirmo que deseo eliminar este proyecto de forma permanente",
                    key=f"e_confirmar_{id_sel}",
                )

                st.write("")
                b1, b2 = st.columns(2)
                with b1:
                    actualizar = st.form_submit_button("Guardar cambios", use_container_width=True)
                with b2:
                    eliminar = st.form_submit_button("Eliminar proyecto", use_container_width=True)

                if actualizar:
                    errores_edit = []
                    if not e_id.strip():
                        errores_edit.append("El ID del proyecto es obligatorio.")
                    elif e_id.strip() != id_sel and e_id.strip() in df_proyectos["ID_Proyecto"].astype(str).values:
                        errores_edit.append(f"El ID '{e_id}' ya existe. Usa uno diferente.")
                    if not e_nombre.strip():
                        errores_edit.append("El nombre del proyecto es obligatorio.")
                    if not e_cliente.strip():
                        errores_edit.append("El cliente/área es obligatorio.")
                    if not e_lider.strip():
                        errores_edit.append("El líder del proyecto es obligatorio.")
                    if e_fin < e_inicio:
                        errores_edit.append("La fecha fin no puede ser anterior a la fecha de inicio.")

                    if errores_edit:
                        for e in errores_edit:
                            st.error(e)
                    else:
                        idx_original = fila.name
                        df_proyectos.loc[idx_original, "ID_Proyecto"] = e_id.strip()
                        df_proyectos.loc[idx_original, "Nombre_Proyecto"] = e_nombre.strip()
                        df_proyectos.loc[idx_original, "Descripcion"] = e_descripcion.strip()
                        df_proyectos.loc[idx_original, "Cliente_Area"] = e_cliente.strip()
                        df_proyectos.loc[idx_original, "Tecnologia"] = e_tecnologia.strip()
                        df_proyectos.loc[idx_original, "Beneficio_Esperado"] = e_beneficio
                        df_proyectos.loc[idx_original, "Lider_Proyecto"] = e_lider.strip()
                        df_proyectos.loc[idx_original, "Fecha_Inicio"] = e_inicio.strftime("%Y-%m-%d")
                        df_proyectos.loc[idx_original, "Fecha_Fin_Planificada"] = e_fin.strftime("%Y-%m-%d")
                        df_proyectos.loc[idx_original, "Fecha_Fin_Real"] = e_fin_real.strftime("%Y-%m-%d") if e_ya_finalizo else ""
                        df_proyectos.loc[idx_original, "Presupuesto_Planificado"] = e_presupuesto
                        df_proyectos.loc[idx_original, "Presupuesto_Ejecutado"] = e_ejecutado
                        df_proyectos.loc[idx_original, "Avance_Pct"] = e_avance
                        df_proyectos.loc[idx_original, "Estado_Proyecto"] = e_estado
                        df_proyectos.loc[idx_original, "Prioridad"] = e_prioridad
                        guardar_hoja("Proyectos", df_proyectos)

                        mensajes = [("success", f"Proyecto '{e_nombre}' actualizado correctamente.")]
                        if e_ejecutado > e_presupuesto and e_presupuesto > 0:
                            mensajes.append(("warning", "El presupuesto ejecutado supera al planificado. Se guardó igual, pero revisa el dato."))
                        if e_estado == "Finalizado" and e_avance != 100:
                            mensajes.append(("info", f"Marcaste el proyecto como «Finalizado» pero el avance quedó en {e_avance}%."))
                        elif e_estado == "No iniciado" and e_avance != 0:
                            mensajes.append(("info", f"El proyecto está «No iniciado» pero el avance quedó en {e_avance}%."))
                        st.session_state["flash_proyecto"] = mensajes
                        st.rerun()

                elif eliminar:
                    if not e_confirmar_borrado:
                        st.error("Marca la casilla de confirmación para poder eliminar el proyecto.")
                    else:
                        idx_original = fila.name
                        df_proyectos = df_proyectos.drop(index=idx_original)
                        guardar_hoja("Proyectos", df_proyectos)
                        st.session_state["flash_proyecto"] = [("success", f"Proyecto '{fila['Nombre_Proyecto']}' eliminado correctamente.")]
                        st.rerun()

    st.write("")
    df_mostrar = leer_hoja("Proyectos", COLS_PROYECTOS)
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Proyectos capturados</div>
        <div class="count-pill">{len(df_mostrar)} en total</div>
    </div>
    """, unsafe_allow_html=True)

    if df_mostrar.empty:
        st.info("Todavía no hay proyectos capturados.")
    else:
        buscar = st.text_input(
            "Buscar proyecto",
            placeholder="Filtra por nombre, cliente/área o líder...",
            label_visibility="collapsed",
        )
        df_filtrado = df_mostrar
        if buscar:
            mascara = (
                df_mostrar["Nombre_Proyecto"].astype(str).str.contains(buscar, case=False, na=False)
                | df_mostrar["Cliente_Area"].astype(str).str.contains(buscar, case=False, na=False)
                | df_mostrar["Lider_Proyecto"].astype(str).str.contains(buscar, case=False, na=False)
            )
            df_filtrado = df_mostrar[mascara]

        st.dataframe(
            df_filtrado,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Avance_Pct": st.column_config.ProgressColumn(
                    "Avance", format="%d%%", min_value=0, max_value=100
                ),
                "Presupuesto_Planificado": st.column_config.NumberColumn(
                    "Presup. planificado", format="$%.0f"
                ),
                "Presupuesto_Ejecutado": st.column_config.NumberColumn(
                    "Presup. ejecutado", format="$%.0f"
                ),
                "Fecha_Inicio": st.column_config.TextColumn("Inicio"),
                "Fecha_Fin_Planificada": st.column_config.TextColumn("Fin planificado"),
                "Fecha_Fin_Real": st.column_config.TextColumn("Fin real"),
                "Descripcion": st.column_config.TextColumn("Descripción"),
                "Tecnologia": st.column_config.TextColumn("Tecnología"),
                "Beneficio_Esperado": st.column_config.NumberColumn("Beneficio esperado", format="$%.0f"),
                "Estado_Proyecto": st.column_config.TextColumn("Estado"),
            },
        )
        st.caption(f"Mostrando {len(df_filtrado)} de {len(df_mostrar)} proyecto(s)")


# =========================================================
# MÓDULO 2: COLABORADORES
# =========================================================
elif modulo.startswith("2"):
    st.markdown("""
    <div class="banner">
        <h1>Módulo 2 · Colaboradores</h1>
        <p>Registra a tu equipo, su rol y la disponibilidad semanal de cada persona.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="focus-callout">
        <strong>Por qué vale la pena registrar bien esta información:</strong> la capacidad semanal
        de cada colaborador se compara en el <strong>Monitor de Salud</strong> con las horas que
        realmente dedica a sus proyectos. Así el sistema puede avisarte a tiempo cuando alguien está
        sobrecargado, antes de que afecte su bienestar o el avance del proyecto.
    </div>
    """, unsafe_allow_html=True)

    # ---- Mensajes pendientes de la última acción (creación/edición/borrado) ----
    if "flash_colaborador" in st.session_state:
        for tipo, texto in st.session_state.pop("flash_colaborador"):
            getattr(st, tipo)(texto)

    # ---- Resumen rápido ----
    total_col = len(df_colaboradores)
    activos_col = int((df_colaboradores["Estado"] == "Activo").sum()) if total_col else 0
    capacidad_total_col = (
        df_colaboradores.loc[df_colaboradores["Estado"] == "Activo", "Capacidad_Horas_Semana"].sum()
        if total_col else 0
    )
    alta_carga_col = (
        int(df_colaboradores["Proyectos_Asignados"].apply(lambda x: len(parsear_proyectos(x)) >= 3).sum())
        if total_col else 0
    )

    s1, s2, s3, s4 = st.columns(4)
    for col, valor, etiqueta in zip(
        (s1, s2, s3, s4),
        (total_col, activos_col, f"{capacidad_total_col:,.0f} h", alta_carga_col),
        ("Colaboradores", "Activos", "Capacidad semanal", "Con carga alta"),
    ):
        with col:
            st.markdown(
                f'<div class="mini-stat"><div class="n">{valor}</div><div class="l">{etiqueta}</div></div>',
                unsafe_allow_html=True,
            )

    st.write("")

    opciones_proyectos = [f"{r.ID_Proyecto} · {r.Nombre_Proyecto}" for r in df_proyectos.itertuples()]
    mapa_etiqueta_id_proy = dict(zip(opciones_proyectos, df_proyectos["ID_Proyecto"].astype(str)))
    mapa_id_etiqueta_proy = {v: k for k, v in mapa_etiqueta_id_proy.items()}

    with st.form("form_colaborador", clear_on_submit=True):
        st.markdown("""
        <div class="form-section-header">
            <div class="icon">01</div>
            <div>
                <div class="title">Información personal</div>
                <div class="subtitle">Identificación y contacto del colaborador</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            id_col = st.text_input(
                "ID Colaborador *",
                value=siguiente_id(df_colaboradores, "ID_Colaborador", "COL"),
                help="Se autogenera con el siguiente consecutivo disponible. Puedes cambiarlo si tu organización usa otro formato.",
            )
            email = st.text_input("Correo electrónico", placeholder="nombre@empresa.com")
        with col2:
            nombre_col = st.text_input("Nombre completo *", placeholder="Nombre y apellido")
            fecha_ingreso = st.date_input("Fecha de ingreso", value=date.today())

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">02</div>
            <div>
                <div class="title">Rol y área</div>
                <div class="subtitle">Ubicación del colaborador en la organización</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col3, col4 = st.columns(2)
        with col3:
            rol = st.text_input("Rol / Cargo *", placeholder="Líder de Proyecto, Analista de Datos...")
        with col4:
            area = st.text_input("Área / Departamento *", placeholder="Finanzas, Operaciones...")
        estado_col = st.selectbox("Estado", ["Activo", "Inactivo"])

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">03</div>
            <div>
                <div class="title">Capacidad y asignación</div>
                <div class="subtitle">Disponibilidad semanal y proyectos en los que participa</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col_cap1, col_cap2 = st.columns(2)
        with col_cap1:
            capacidad = st.number_input(
                "Capacidad (horas/semana) *", min_value=0.0, max_value=80.0, value=40.0, step=5.0,
                help="Horas semanales disponibles para trabajar en proyectos. Es la base para calcular la saturación en el Monitor de Salud.",
            )
        with col_cap2:
            costo_hora = st.number_input(
                "Costo por hora ($)", min_value=0.0, max_value=500.0, value=30.0, step=1.0,
                help="Tarifa o costo interno por hora. Permite calcular el costo real de cada proyecto a partir de las horas registradas.",
            )
        proyectos_sel = st.multiselect(
            "Proyectos asignados",
            options=opciones_proyectos,
            help="Puedes dejarlo vacío si el colaborador todavía no tiene proyectos asignados.",
        )

        st.write("")
        submitted_col = st.form_submit_button("Guardar colaborador", use_container_width=True)

        if submitted_col:
            errores_col = []
            if not id_col.strip():
                errores_col.append("El ID del colaborador es obligatorio.")
            elif id_col.strip() in df_colaboradores["ID_Colaborador"].astype(str).values:
                errores_col.append(f"El ID '{id_col}' ya existe. Usa uno diferente.")
            if not nombre_col.strip():
                errores_col.append("El nombre del colaborador es obligatorio.")
            if not rol.strip():
                errores_col.append("El rol/cargo es obligatorio.")
            if not area.strip():
                errores_col.append("El área/departamento es obligatorio.")
            if capacidad <= 0:
                errores_col.append("La capacidad semanal debe ser mayor a 0.")

            if errores_col:
                for e in errores_col:
                    st.error(e)
            else:
                ids_proyectos = [mapa_etiqueta_id_proy[et] for et in proyectos_sel]
                nuevo_col = {
                    "ID_Colaborador": id_col.strip(),
                    "Nombre": nombre_col.strip(),
                    "Rol": rol.strip(),
                    "Area": area.strip(),
                    "Email": email.strip(),
                    "Fecha_Ingreso": fecha_ingreso.strftime("%Y-%m-%d"),
                    "Capacidad_Horas_Semana": capacidad,
                    "Costo_Hora": costo_hora,
                    "Proyectos_Asignados": ",".join(ids_proyectos),
                    "Estado": estado_col,
                }
                df_col_actualizado = pd.concat(
                    [df_colaboradores, pd.DataFrame([nuevo_col])], ignore_index=True
                )
                guardar_hoja("Colaboradores", df_col_actualizado)

                mensajes_col = [("success", f"Colaborador '{nombre_col}' guardado correctamente.")]
                if email.strip() and "@" not in email:
                    mensajes_col.append(("warning", "El correo ingresado no parece válido (falta el '@'). Se guardó igual, pero revisa el dato."))
                if len(ids_proyectos) >= 3:
                    mensajes_col.append(("info", f"«{nombre_col}» quedó asignado a {len(ids_proyectos)} proyectos. Vigila su carga cuando el Monitor de Salud esté disponible."))
                st.session_state["flash_colaborador"] = mensajes_col
                st.rerun()

    # ---- Editar o eliminar un colaborador existente ----
    st.write("")
    with st.expander("Editar o eliminar un colaborador existente", expanded=not df_colaboradores.empty):
        if df_colaboradores.empty:
            st.caption("Todavía no hay colaboradores capturados para editar.")
        else:
            etiquetas_col = [f"{r.ID_Colaborador} · {r.Nombre}" for r in df_colaboradores.itertuples()]
            mapa_id_col = dict(zip(etiquetas_col, df_colaboradores["ID_Colaborador"].astype(str)))
            etiqueta_sel_col = st.selectbox("Selecciona un colaborador", etiquetas_col, key="sel_editar_colaborador")
            id_sel_col = mapa_id_col[etiqueta_sel_col]
            fila_col = df_colaboradores[df_colaboradores["ID_Colaborador"].astype(str) == id_sel_col].iloc[0]

            try:
                f_ingreso_val = pd.to_datetime(fila_col["Fecha_Ingreso"]).date()
            except (ValueError, TypeError):
                f_ingreso_val = date.today()

            ids_actuales = parsear_proyectos(fila_col["Proyectos_Asignados"])
            etiquetas_actuales = [mapa_id_etiqueta_proy[i] for i in ids_actuales if i in mapa_id_etiqueta_proy]

            with st.form(f"form_editar_colaborador_{id_sel_col}"):
                col1, col2 = st.columns(2)
                with col1:
                    e_id_col = st.text_input("ID Colaborador *", value=str(fila_col["ID_Colaborador"]), key=f"ec_id_{id_sel_col}")
                    e_email = st.text_input("Correo electrónico", value=str(fila_col["Email"]), key=f"ec_email_{id_sel_col}")
                    e_rol = st.text_input("Rol / Cargo *", value=str(fila_col["Rol"]), key=f"ec_rol_{id_sel_col}")
                    e_capacidad = st.number_input(
                        "Capacidad (horas/semana) *", min_value=0.0, max_value=80.0, step=5.0,
                        value=float(fila_col["Capacidad_Horas_Semana"] or 0), key=f"ec_capacidad_{id_sel_col}",
                    )
                    e_costo_hora = st.number_input(
                        "Costo por hora ($)", min_value=0.0, max_value=500.0, step=1.0,
                        value=float(pd.to_numeric(fila_col["Costo_Hora"], errors="coerce") or 0), key=f"ec_costo_{id_sel_col}",
                    )
                with col2:
                    e_nombre_col = st.text_input("Nombre completo *", value=str(fila_col["Nombre"]), key=f"ec_nombre_{id_sel_col}")
                    e_ingreso = st.date_input("Fecha de ingreso", value=f_ingreso_val, key=f"ec_ingreso_{id_sel_col}")
                    e_area = st.text_input("Área / Departamento *", value=str(fila_col["Area"]), key=f"ec_area_{id_sel_col}")
                    estados_col_posibles = ["Activo", "Inactivo"]
                    e_estado_col = st.selectbox(
                        "Estado", estados_col_posibles,
                        index=estados_col_posibles.index(fila_col["Estado"]) if fila_col["Estado"] in estados_col_posibles else 0,
                        key=f"ec_estado_{id_sel_col}",
                    )

                e_proyectos_sel = st.multiselect(
                    "Proyectos asignados", options=opciones_proyectos, default=etiquetas_actuales,
                    key=f"ec_proyectos_{id_sel_col}",
                )

                e_confirmar_borrado_col = st.checkbox(
                    "Confirmo que deseo eliminar este colaborador de forma permanente",
                    key=f"ec_confirmar_{id_sel_col}",
                )

                st.write("")
                b1, b2 = st.columns(2)
                with b1:
                    actualizar_col = st.form_submit_button("Guardar cambios", use_container_width=True)
                with b2:
                    eliminar_col = st.form_submit_button("Eliminar colaborador", use_container_width=True)

                if actualizar_col:
                    errores_edit_col = []
                    if not e_id_col.strip():
                        errores_edit_col.append("El ID del colaborador es obligatorio.")
                    elif e_id_col.strip() != id_sel_col and e_id_col.strip() in df_colaboradores["ID_Colaborador"].astype(str).values:
                        errores_edit_col.append(f"El ID '{e_id_col}' ya existe. Usa uno diferente.")
                    if not e_nombre_col.strip():
                        errores_edit_col.append("El nombre del colaborador es obligatorio.")
                    if not e_rol.strip():
                        errores_edit_col.append("El rol/cargo es obligatorio.")
                    if not e_area.strip():
                        errores_edit_col.append("El área/departamento es obligatorio.")
                    if e_capacidad <= 0:
                        errores_edit_col.append("La capacidad semanal debe ser mayor a 0.")

                    if errores_edit_col:
                        for e in errores_edit_col:
                            st.error(e)
                    else:
                        ids_proyectos_edit = [mapa_etiqueta_id_proy[et] for et in e_proyectos_sel]
                        idx_original_col = fila_col.name
                        df_colaboradores.loc[idx_original_col, "ID_Colaborador"] = e_id_col.strip()
                        df_colaboradores.loc[idx_original_col, "Nombre"] = e_nombre_col.strip()
                        df_colaboradores.loc[idx_original_col, "Rol"] = e_rol.strip()
                        df_colaboradores.loc[idx_original_col, "Area"] = e_area.strip()
                        df_colaboradores.loc[idx_original_col, "Email"] = e_email.strip()
                        df_colaboradores.loc[idx_original_col, "Fecha_Ingreso"] = e_ingreso.strftime("%Y-%m-%d")
                        df_colaboradores.loc[idx_original_col, "Capacidad_Horas_Semana"] = e_capacidad
                        df_colaboradores.loc[idx_original_col, "Costo_Hora"] = e_costo_hora
                        df_colaboradores.loc[idx_original_col, "Proyectos_Asignados"] = ",".join(ids_proyectos_edit)
                        df_colaboradores.loc[idx_original_col, "Estado"] = e_estado_col
                        guardar_hoja("Colaboradores", df_colaboradores)

                        mensajes_edit_col = [("success", f"Colaborador '{e_nombre_col}' actualizado correctamente.")]
                        if e_email.strip() and "@" not in e_email:
                            mensajes_edit_col.append(("warning", "El correo ingresado no parece válido (falta el '@')."))
                        if len(ids_proyectos_edit) >= 3:
                            mensajes_edit_col.append(("info", f"«{e_nombre_col}» quedó asignado a {len(ids_proyectos_edit)} proyectos. Vigila su carga en el Monitor de Salud."))
                        st.session_state["flash_colaborador"] = mensajes_edit_col
                        st.rerun()

                elif eliminar_col:
                    if not e_confirmar_borrado_col:
                        st.error("Marca la casilla de confirmación para poder eliminar al colaborador.")
                    else:
                        idx_original_col = fila_col.name
                        df_colaboradores = df_colaboradores.drop(index=idx_original_col)
                        guardar_hoja("Colaboradores", df_colaboradores)
                        st.session_state["flash_colaborador"] = [("success", f"Colaborador '{fila_col['Nombre']}' eliminado correctamente.")]
                        st.rerun()

    # ---- Vista rápida de carga de trabajo (preview del Monitor de Salud) ----
    st.write("")
    df_mostrar_col = leer_hoja("Colaboradores", COLS_COLABORADORES)
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Vista rápida de carga de trabajo</div>
        <div class="count-pill">{len(df_mostrar_col)} en total</div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("Carga preliminar según cantidad de proyectos asignados. El cálculo real de saturación (horas asignadas vs. horas reales) se hará en el Monitor de Salud.")

    if df_mostrar_col.empty:
        st.info("Todavía no hay colaboradores capturados.")
    else:
        mapa_id_nombre_proy = dict(zip(df_proyectos["ID_Proyecto"].astype(str), df_proyectos["Nombre_Proyecto"]))
        for _, fila_vista in df_mostrar_col.iterrows():
            ids_asignados = parsear_proyectos(fila_vista["Proyectos_Asignados"])
            etiqueta_carga, clase_badge, clase_avatar = nivel_carga(len(ids_asignados))
            iniciales = "".join(p[0].upper() for p in str(fila_vista["Nombre"]).split()[:2]) or "?"
            nombres_proyectos = ", ".join(mapa_id_nombre_proy.get(i, i) for i in ids_asignados) if ids_asignados else "Sin proyectos asignados"
            capacidad_fmt = float(fila_vista["Capacidad_Horas_Semana"] or 0)
            with st.container(border=True):
                a, b, c_carga = st.columns([0.6, 3, 2])
                with a:
                    st.markdown(f'<div class="avatar-circle {clase_avatar}">{iniciales}</div>', unsafe_allow_html=True)
                with b:
                    st.markdown(f"**{fila_vista['Nombre']}**  \n{fila_vista['Rol']} · {fila_vista['Area']}")
                    st.caption(f"{nombres_proyectos}")
                with c_carga:
                    st.markdown(f'<span class="badge {clase_badge}">{etiqueta_carga} · {len(ids_asignados)} proy.</span>', unsafe_allow_html=True)
                    st.caption(f"{capacidad_fmt:.0f} h/semana · {fila_vista['Estado']}")

    # ---- Tabla completa de colaboradores ----
    st.write("")
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Colaboradores capturados</div>
        <div class="count-pill">{len(df_mostrar_col)} en total</div>
    </div>
    """, unsafe_allow_html=True)

    if df_mostrar_col.empty:
        st.info("Todavía no hay colaboradores capturados.")
    else:
        buscar_col = st.text_input(
            "Buscar colaborador",
            placeholder="Filtra por nombre, rol o área...",
            label_visibility="collapsed",
            key="buscar_colaborador",
        )
        df_filtrado_col = df_mostrar_col
        if buscar_col:
            mascara_col = (
                df_mostrar_col["Nombre"].astype(str).str.contains(buscar_col, case=False, na=False)
                | df_mostrar_col["Rol"].astype(str).str.contains(buscar_col, case=False, na=False)
                | df_mostrar_col["Area"].astype(str).str.contains(buscar_col, case=False, na=False)
            )
            df_filtrado_col = df_mostrar_col[mascara_col]

        st.dataframe(
            df_filtrado_col,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Capacidad_Horas_Semana": st.column_config.NumberColumn("Capacidad (h/sem)", format="%.0f h"),
                "Costo_Hora": st.column_config.NumberColumn("Costo/hora", format="$%.0f"),
                "Proyectos_Asignados": st.column_config.TextColumn("Proyectos asignados"),
                "Fecha_Ingreso": st.column_config.TextColumn("Ingreso"),
            },
        )
        st.caption(f"Mostrando {len(df_filtrado_col)} de {len(df_mostrar_col)} colaborador(es)")


# =========================================================
# MÓDULO 3: MONITOR DE SALUD — AUTO-REPORTE DEL COLABORADOR
# =========================================================
elif modulo.startswith("3") and not es_pm:
    st.markdown(f"""
    <div class="banner">
        <h1>Módulo 3 · Mi registro semanal</h1>
        <p>Hola, {nombre_usuario}. Reporta tus horas de cada semana y cómo sentiste la carga.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="focus-callout">
        <strong>Por qué reportar cada semana:</strong> tus horas y tu sensación de carga son la base
        para detectar la sobrecarga a tiempo. La saturación compara tus horas reales con tu capacidad:
        ⚪ <strong>&lt;60%</strong> capacidad disponible · 🟢 <strong>60–80%</strong> saludable ·
        🟡 <strong>80–90%</strong> seguimiento · 🟠 <strong>90–95%</strong> precaución ·
        🔴 <strong>95%</strong> crítico.
        Solo tú puedes registrar y editar tus horas; el PM ve el semáforo del equipo para equilibrar la carga.
    </div>
    """, unsafe_allow_html=True)

    # ---- Mensajes pendientes de la última acción ----
    if "flash_salud" in st.session_state:
        for tipo, texto in st.session_state.pop("flash_salud"):
            getattr(st, tipo)(texto)

    # ---- Datos propios ----
    mapa_nombre_proy = dict(zip(df_proyectos["ID_Proyecto"].astype(str), df_proyectos["Nombre_Proyecto"]))
    fila_usuario = df_colaboradores[df_colaboradores["ID_Colaborador"].astype(str) == id_usuario].iloc[0]
    capacidad_mia = float(fila_usuario["Capacidad_Horas_Semana"] or 0)
    ids_mis_proyectos = parsear_proyectos(fila_usuario["Proyectos_Asignados"])
    df_salud_mio = df_salud[df_salud["ID_Colaborador"].astype(str) == id_usuario]
    semanas_mias = sorted(df_salud_mio["Semana"].unique(), reverse=True)

    # ---- Resumen rápido personal (última semana reportada) ----
    if semanas_mias:
        semana_ult = semanas_mias[0]
        df_sem_ult = df_salud_mio[df_salud_mio["Semana"] == semana_ult]
        horas_ult = df_sem_ult["Horas_Reales"].sum()
        sat_ult = horas_ult / capacidad_mia * 100 if capacidad_mia > 0 else 0
        percib_ult = df_sem_ult["Carga_Percibida"].mean()
        emoji_ult = PERCIBIDA_EMOJI.get(int(round(percib_ult)), "😐")
        valor_semana = f"{semana_ult[8:10]}/{semana_ult[5:7]}"
        valor_horas = f"{horas_ult:.1f} h"
        valor_sat = f"{sat_ult:.0f}%" if capacidad_mia > 0 else "—"
    else:
        semana_ult, horas_ult, sat_ult = None, 0, 0
        valor_semana, valor_horas, valor_sat, emoji_ult = "—", "—", "—", "—"

    s1, s2, s3, s4 = st.columns(4)
    for col, valor, etiqueta in zip(
        (s1, s2, s3, s4),
        (valor_semana, valor_horas, valor_sat, emoji_ult),
        ("Última semana reportada", "Mis horas reales", "Mi saturación", "Sensación de carga"),
    ):
        with col:
            st.markdown(
                f'<div class="mini-stat"><div class="n">{valor}</div><div class="l">{etiqueta}</div></div>',
                unsafe_allow_html=True,
            )

    # ---- Mi semáforo (última semana reportada) ----
    st.write("")
    st.markdown("""
    <div class="table-card-header">
        <div class="title">Mi semáforo</div>
    </div>
    """, unsafe_allow_html=True)

    if not semanas_mias:
        st.info("Todavía no has reportado horas. Usa el formulario de abajo para registrar tu primera semana.")
    elif capacidad_mia <= 0:
        st.warning("Tu capacidad semanal está en 0 h, así que no se puede calcular tu saturación. Pide al PM que la actualice en el Módulo 2.")
    else:
        etiqueta_sat, clase_badge, clase_barra, clase_avatar = clasificar_saturacion(sat_ult)
        iniciales = "".join(p[0].upper() for p in str(nombre_usuario).split()[:2]) or "?"
        with st.container(border=True):
            a, b, c = st.columns([0.7, 2.6, 1.7])
            with a:
                st.markdown(f'<div class="avatar-circle {clase_avatar}">{iniciales}</div>', unsafe_allow_html=True)
            with b:
                st.markdown(f"**{nombre_usuario}**  \n{etiqueta_semana(semana_ult)}")
            with c:
                st.markdown(
                    f'<span class="badge {clase_badge}">{etiqueta_sat} · {sat_ult:.0f}%</span>',
                    unsafe_allow_html=True,
                )
            st.markdown(
                f'<div class="sat-track"><div class="sat-fill {clase_barra}" style="width:{min(sat_ult, 100):.0f}%"></div></div>',
                unsafe_allow_html=True,
            )
            st.caption(f"{horas_ult:.1f} h de {capacidad_mia:.0f} h · Sensación de carga: {emoji_ult}")
        if sat_ult >= 95:
            st.error("🔴 Tu última semana quedó en zona crítica. Coméntalo con tu PM para redistribuir la carga.")
        elif sat_ult >= 90:
            st.warning("🟠 Tu última semana quedó en zona de precaución. Vigila tu carga esta semana.")

    # ---- Formulario de captura semanal (solo mis horas) ----
    st.write("")
    opciones_proy_salud = [f"{r.ID_Proyecto} · {r.Nombre_Proyecto}" for r in df_proyectos.itertuples()]
    mapa_etiqueta_id_proy_salud = dict(zip(opciones_proy_salud, df_proyectos["ID_Proyecto"].astype(str)))
    opciones_mis_proy = [e for e, i in mapa_etiqueta_id_proy_salud.items() if i in ids_mis_proyectos]
    if not opciones_mis_proy:
        opciones_mis_proy = opciones_proy_salud

    with st.form("form_salud", clear_on_submit=True):
        st.markdown("""
        <div class="form-section-header">
            <div class="icon">01</div>
            <div>
                <div class="title">Registrar mis horas de la semana</div>
                <div class="subtitle">En qué proyecto trabajaste y en qué semana</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            proyecto_sel = st.selectbox(
                "Proyecto *", opciones_mis_proy,
                help="Se muestran los proyectos que tienes asignados.",
            ) if opciones_mis_proy else None
        with col2:
            fecha_semana = st.date_input(
                "Semana a registrar", value=date.today(),
                help="Puedes elegir cualquier día: el registro se guardará en la semana del lunes correspondiente.",
            )

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">02</div>
            <div>
                <div class="title">Horas de la semana</div>
                <div class="subtitle">Lo planificado vs. lo que realmente trabajaste</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col3, col4 = st.columns(2)
        with col3:
            horas_asignadas = st.number_input(
                "Horas asignadas", min_value=0.0, max_value=80.0, step=1.0,
                help="Horas que se planeó dedicar a este proyecto durante la semana.",
            )
        with col4:
            horas_reales = st.number_input(
                "Horas reales trabajadas", min_value=0.0, max_value=80.0, step=1.0,
                help="Horas efectivamente trabajadas. Este dato alimenta el semáforo de saturación.",
            )

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">03</div>
            <div>
                <div class="title">Tu percepción</div>
                <div class="subtitle">Cómo sentiste la carga esa semana — el dato humano del monitor</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        percibida_sel = st.select_slider(
            "¿Qué tan pesada sentiste la semana?",
            options=PERCIBIDA_OPCIONES, value=PERCIBIDA_OPCIONES[2],
            help="Tu percepción complementa las horas: puedes estar dentro de tu capacidad y aun así sentirte sobrecargado.",
        )
        comentario = st.text_input("Comentario (opcional)", placeholder="Ej: semana de cierre de hito, hubo incidentes...")

        st.write("")
        submitted_salud = st.form_submit_button("Guardar mi registro semanal", use_container_width=True)

        if submitted_salud:
            errores_salud = []
            if not proyecto_sel:
                errores_salud.append("No hay proyectos registrados. Pide al PM que cree al menos uno en el Módulo 1.")
            if horas_asignadas == 0 and horas_reales == 0:
                errores_salud.append("Registra al menos las horas asignadas o las horas reales de la semana.")

            if not errores_salud:
                id_proy_nuevo = mapa_etiqueta_id_proy_salud[proyecto_sel]
                semana_nueva = lunes_de_la_semana(fecha_semana).strftime("%Y-%m-%d")
                duplicado = df_salud[
                    (df_salud["ID_Colaborador"].astype(str) == id_usuario)
                    & (df_salud["ID_Proyecto"].astype(str) == id_proy_nuevo)
                    & (df_salud["Semana"] == semana_nueva)
                ]
                if not duplicado.empty:
                    errores_salud.append(
                        f"Ya reportaste horas en ese proyecto para la {etiqueta_semana(semana_nueva).lower()}. "
                        "Usa la sección «Editar o eliminar» para modificar ese registro."
                    )

            if errores_salud:
                for e in errores_salud:
                    st.error(e)
            else:
                nuevo_registro = {
                    "ID_Registro": siguiente_id(df_salud, "ID_Registro", "SAL"),
                    "Semana": semana_nueva,
                    "ID_Colaborador": id_usuario,
                    "Nombre_Colaborador": nombre_usuario,
                    "ID_Proyecto": id_proy_nuevo,
                    "Nombre_Proyecto": mapa_nombre_proy.get(id_proy_nuevo, id_proy_nuevo),
                    "Horas_Asignadas": horas_asignadas,
                    "Horas_Reales": horas_reales,
                    "Carga_Percibida": int(percibida_sel[0]),
                    "Capacidad_Horas_Semana": capacidad_mia,
                    "Comentario": comentario.strip(),
                }
                df_salud_actualizado = pd.concat([df_salud, pd.DataFrame([nuevo_registro])], ignore_index=True)
                guardar_hoja("Salud_Equipo", df_salud_actualizado)

                mensajes_salud = [("success", "Tu registro semanal quedó guardado. ¡Gracias por reportar a tiempo!")]

                # Alerta temprana: recalcular la saturación de esa semana con el nuevo registro
                if capacidad_mia > 0:
                    total_semana = df_salud_actualizado[
                        (df_salud_actualizado["ID_Colaborador"].astype(str) == id_usuario)
                        & (df_salud_actualizado["Semana"] == semana_nueva)
                    ]["Horas_Reales"].sum()
                    sat_nueva = total_semana / capacidad_mia * 100
                    texto_alerta = f"{nombre_usuario} al {sat_nueva:.0f}% de su capacidad semanal ({total_semana:.1f} h de {capacidad_mia:.0f} h)."
                    if sat_nueva >= 95:
                        mensajes_salud.append(("error", f"🔴 Con este registro quedas al {sat_nueva:.0f}% de tu capacidad esta semana ({total_semana:.1f} h de {capacidad_mia:.0f} h). El PM recibirá la alerta para ayudarte a redistribuir la carga."))
                        registrar_alerta("Monitor de Salud", "Crítico", id_usuario, nombre_usuario, "", semana_nueva, round(sat_nueva, 1), 95, texto_alerta)
                    elif sat_nueva >= 90:
                        mensajes_salud.append(("warning", f"🟠 Con este registro quedas al {sat_nueva:.0f}% de tu capacidad esta semana ({total_semana:.1f} h de {capacidad_mia:.0f} h). Se dejó registro en la hoja de alertas."))
                        registrar_alerta("Monitor de Salud", "Precaución", id_usuario, nombre_usuario, "", semana_nueva, round(sat_nueva, 1), 90, texto_alerta)
                if int(percibida_sel[0]) >= 4:
                    mensajes_salud.append(("info", "Reportaste que la semana se sintió pesada. Coméntalo con tu PM: tu percepción cuenta tanto como las horas."))
                if ids_mis_proyectos and id_proy_nuevo not in ids_mis_proyectos:
                    mensajes_salud.append(("info", f"El proyecto {id_proy_nuevo} no está en tu lista de proyectos asignados. Avisa al PM para que actualice tu asignación en el Módulo 2."))
                st.session_state["flash_salud"] = mensajes_salud
                st.rerun()

    # ---- Editar o eliminar uno de mis registros ----
    st.write("")
    with st.expander("Editar o eliminar uno de mis registros", expanded=False):
        if df_salud_mio.empty:
            st.caption("Todavía no tienes registros de horas para editar.")
        else:
            df_mio_orden = df_salud_mio.sort_values(["Semana", "ID_Registro"], ascending=[False, True])
            etiquetas_reg = [
                f"{r.ID_Registro} · {etiqueta_semana(r.Semana)} · {r.ID_Proyecto}"
                for r in df_mio_orden.itertuples()
            ]
            mapa_id_reg = dict(zip(etiquetas_reg, df_mio_orden["ID_Registro"].astype(str)))
            etiqueta_reg_sel = st.selectbox("Selecciona un registro", etiquetas_reg, key="sel_editar_salud")
            id_reg_sel = mapa_id_reg[etiqueta_reg_sel]
            fila_reg = df_salud[df_salud["ID_Registro"].astype(str) == id_reg_sel].iloc[0]

            etiqueta_proy_actual = next(
                (e for e, i in mapa_etiqueta_id_proy_salud.items() if i == str(fila_reg["ID_Proyecto"])),
                opciones_proy_salud[0] if opciones_proy_salud else None,
            )
            try:
                semana_val = pd.to_datetime(fila_reg["Semana"]).date()
            except (ValueError, TypeError):
                semana_val = date.today()
            percibida_actual = int(fila_reg["Carga_Percibida"] or 3)

            with st.form(f"form_editar_salud_{id_reg_sel}"):
                col1, col2 = st.columns(2)
                with col1:
                    e_proy_sel = st.selectbox(
                        "Proyecto *", opciones_proy_salud,
                        index=opciones_proy_salud.index(etiqueta_proy_actual) if etiqueta_proy_actual in opciones_proy_salud else 0,
                        key=f"es_proy_{id_reg_sel}",
                    )
                    e_asignadas = st.number_input(
                        "Horas asignadas", min_value=0.0, max_value=80.0, step=1.0,
                        value=float(fila_reg["Horas_Asignadas"] or 0), key=f"es_asig_{id_reg_sel}",
                    )
                with col2:
                    e_semana = st.date_input("Semana", value=semana_val, key=f"es_semana_{id_reg_sel}")
                    e_reales = st.number_input(
                        "Horas reales trabajadas", min_value=0.0, max_value=80.0, step=1.0,
                        value=float(fila_reg["Horas_Reales"] or 0), key=f"es_real_{id_reg_sel}",
                    )
                e_percibida = st.select_slider(
                    "¿Qué tan pesada sentiste la semana?", options=PERCIBIDA_OPCIONES,
                    value=PERCIBIDA_OPCIONES[max(1, min(5, percibida_actual)) - 1], key=f"es_percibida_{id_reg_sel}",
                )
                e_comentario = st.text_input(
                    "Comentario (opcional)", value=str(fila_reg["Comentario"]) if pd.notna(fila_reg["Comentario"]) else "",
                    key=f"es_coment_{id_reg_sel}",
                )
                e_confirmar_borrado_sal = st.checkbox(
                    "Confirmo que deseo eliminar este registro de forma permanente",
                    key=f"es_confirmar_{id_reg_sel}",
                )

                st.write("")
                b1, b2 = st.columns(2)
                with b1:
                    actualizar_sal = st.form_submit_button("Guardar cambios", use_container_width=True)
                with b2:
                    eliminar_sal = st.form_submit_button("Eliminar registro", use_container_width=True)

                if actualizar_sal:
                    id_proy_edit = mapa_etiqueta_id_proy_salud[e_proy_sel]
                    semana_edit = lunes_de_la_semana(e_semana).strftime("%Y-%m-%d")
                    errores_edit_sal = []
                    if e_asignadas == 0 and e_reales == 0:
                        errores_edit_sal.append("Registra al menos las horas asignadas o las horas reales.")
                    duplicado_edit = df_salud[
                        (df_salud["ID_Colaborador"].astype(str) == id_usuario)
                        & (df_salud["ID_Proyecto"].astype(str) == id_proy_edit)
                        & (df_salud["Semana"] == semana_edit)
                        & (df_salud["ID_Registro"].astype(str) != id_reg_sel)
                    ]
                    if not duplicado_edit.empty:
                        errores_edit_sal.append("Ya tienes otro registro en ese proyecto para esa semana.")

                    if errores_edit_sal:
                        for e in errores_edit_sal:
                            st.error(e)
                    else:
                        idx_reg = fila_reg.name
                        df_salud.loc[idx_reg, "Semana"] = semana_edit
                        df_salud.loc[idx_reg, "ID_Proyecto"] = id_proy_edit
                        df_salud.loc[idx_reg, "Nombre_Proyecto"] = mapa_nombre_proy.get(id_proy_edit, id_proy_edit)
                        df_salud.loc[idx_reg, "Horas_Asignadas"] = e_asignadas
                        df_salud.loc[idx_reg, "Horas_Reales"] = e_reales
                        df_salud.loc[idx_reg, "Carga_Percibida"] = int(e_percibida[0])
                        df_salud.loc[idx_reg, "Capacidad_Horas_Semana"] = capacidad_mia
                        df_salud.loc[idx_reg, "Comentario"] = e_comentario.strip()
                        guardar_hoja("Salud_Equipo", df_salud)

                        mensajes_edit_sal = [("success", f"Registro {id_reg_sel} actualizado correctamente.")]
                        if capacidad_mia > 0:
                            total_sem_edit = df_salud[
                                (df_salud["ID_Colaborador"].astype(str) == id_usuario)
                                & (df_salud["Semana"] == semana_edit)
                            ]["Horas_Reales"].sum()
                            sat_edit = total_sem_edit / capacidad_mia * 100
                            if sat_edit >= 95:
                                mensajes_edit_sal.append(("error", f"🔴 Con este cambio quedas al {sat_edit:.0f}% de tu capacidad esa semana. Se dejó registro en la hoja de alertas."))
                                registrar_alerta("Monitor de Salud", "Crítico", id_usuario, nombre_usuario, "", semana_edit, round(sat_edit, 1), 95,
                                                 f"{nombre_usuario} al {sat_edit:.0f}% de su capacidad semanal ({total_sem_edit:.1f} h de {capacidad_mia:.0f} h).")
                            elif sat_edit >= 90:
                                mensajes_edit_sal.append(("warning", f"🟠 Con este cambio quedas al {sat_edit:.0f}% de tu capacidad esa semana. Se dejó registro en la hoja de alertas."))
                                registrar_alerta("Monitor de Salud", "Precaución", id_usuario, nombre_usuario, "", semana_edit, round(sat_edit, 1), 90,
                                                 f"{nombre_usuario} al {sat_edit:.0f}% de su capacidad semanal ({total_sem_edit:.1f} h de {capacidad_mia:.0f} h).")
                        st.session_state["flash_salud"] = mensajes_edit_sal
                        st.rerun()

                elif eliminar_sal:
                    if not e_confirmar_borrado_sal:
                        st.error("Marca la casilla de confirmación para poder eliminar el registro.")
                    else:
                        df_salud = df_salud.drop(index=fila_reg.name)
                        guardar_hoja("Salud_Equipo", df_salud)
                        st.session_state["flash_salud"] = [("success", f"Registro {id_reg_sel} eliminado correctamente.")]
                        st.rerun()

    # ---- Mi historial ----
    st.write("")
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Mi historial de registros</div>
        <div class="count-pill">{len(df_salud_mio)} en total</div>
    </div>
    """, unsafe_allow_html=True)

    if df_salud_mio.empty:
        st.info("Todavía no has reportado horas.")
    else:
        f1, f2 = st.columns(2)
        with f1:
            filtro_proy = st.selectbox(
                "Proyecto", ["Todos"] + sorted(df_salud_mio["Nombre_Proyecto"].astype(str).unique()),
                key="filtro_mio_proy",
            )
        with f2:
            filtro_semana = st.selectbox(
                "Semana", ["Todas"] + semanas_mias,
                format_func=lambda s: s if s == "Todas" else etiqueta_semana(s),
                key="filtro_mio_semana",
            )

        df_filtrado_mio = df_salud_mio
        if filtro_proy != "Todos":
            df_filtrado_mio = df_filtrado_mio[df_filtrado_mio["Nombre_Proyecto"].astype(str) == filtro_proy]
        if filtro_semana != "Todas":
            df_filtrado_mio = df_filtrado_mio[df_filtrado_mio["Semana"] == filtro_semana]

        st.dataframe(
            df_filtrado_mio.sort_values(["Semana", "ID_Registro"], ascending=[False, True])[
                ["ID_Registro", "Semana", "Nombre_Proyecto", "Horas_Asignadas", "Horas_Reales", "Carga_Percibida", "Comentario"]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "ID_Registro": st.column_config.TextColumn("ID"),
                "Semana": st.column_config.TextColumn("Semana del"),
                "Nombre_Proyecto": st.column_config.TextColumn("Proyecto"),
                "Horas_Asignadas": st.column_config.NumberColumn("Asignadas", format="%.1f h"),
                "Horas_Reales": st.column_config.NumberColumn("Reales", format="%.1f h"),
                "Carga_Percibida": st.column_config.ProgressColumn(
                    "Carga percibida", format="%d/5", min_value=1, max_value=5
                ),
            },
        )
        st.caption(f"Mostrando {len(df_filtrado_mio)} de {len(df_salud_mio)} registro(s)")


# =========================================================
# MÓDULO 3: MONITOR DE SALUD — SUPERVISIÓN DEL PM
# =========================================================
elif modulo.startswith("3"):
    st.markdown("""
    <div class="banner">
        <h1>Módulo 3 · Monitor de Salud</h1>
        <p>Supervisa la carga real y percibida de todo el equipo, semana a semana.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="focus-callout">
        <strong>Cómo leer este monitor:</strong> la saturación compara las horas que cada persona
        trabajó realmente en la semana con su capacidad disponible.
        ⚪ <strong>&lt;60%</strong> capacidad disponible · 🟢 <strong>60–80%</strong> saludable ·
        🟡 <strong>80–90%</strong> seguimiento · 🟠 <strong>90–95%</strong> precaución ·
        🔴 <strong>95%</strong> crítico.
        Cada colaborador reporta sus propias horas y su carga percibida desde su acceso personal;
        aquí supervisas el panorama completo del equipo y el registro de alertas emitidas.
    </div>
    """, unsafe_allow_html=True)

    # ---- Preparación de datos comunes del módulo ----
    df_col_activos = df_colaboradores[df_colaboradores["Estado"] == "Activo"]
    mapa_nombre_col = dict(zip(df_colaboradores["ID_Colaborador"].astype(str), df_colaboradores["Nombre"]))
    mapa_rol_col = dict(zip(df_colaboradores["ID_Colaborador"].astype(str), df_colaboradores["Rol"]))
    mapa_capacidad_col = {
        str(r.ID_Colaborador): float(r.Capacidad_Horas_Semana or 0)
        for r in df_colaboradores.itertuples()
    }
    semanas_registradas = sorted(df_salud["Semana"].unique(), reverse=True) if not df_salud.empty else []

    def saturacion_semana(id_col, semana):
        """% de saturación de un colaborador en una semana (todas sus horas reales / capacidad)."""
        capacidad = mapa_capacidad_col.get(id_col, 0)
        if capacidad <= 0:
            return None
        horas = df_salud[
            (df_salud["ID_Colaborador"].astype(str) == id_col) & (df_salud["Semana"] == semana)
        ]["Horas_Reales"].sum()
        return horas / capacidad * 100, horas

    # ---- Resumen rápido (última semana con registros) ----
    if semanas_registradas:
        semana_ultima = semanas_registradas[0]
        sats_ultima = []
        for id_c in df_col_activos["ID_Colaborador"].astype(str):
            resultado = saturacion_semana(id_c, semana_ultima)
            if resultado is not None and resultado[1] > 0:
                sats_ultima.append(resultado[0])
        prom_ultima = sum(sats_ultima) / len(sats_ultima) if sats_ultima else 0
        prec_ultima = sum(1 for s in sats_ultima if 90 <= s < 95)
        crit_ultima = sum(1 for s in sats_ultima if s >= 95)
        valor_semana = f"{semana_ultima[8:10]}/{semana_ultima[5:7]}"
    else:
        semana_ultima, prom_ultima, prec_ultima, crit_ultima, valor_semana = None, 0, 0, 0, "—"

    s1, s2, s3, s4 = st.columns(4)
    for col, valor, etiqueta in zip(
        (s1, s2, s3, s4),
        (valor_semana, f"{prom_ultima:.0f}%", prec_ultima, crit_ultima),
        ("Última semana", "Saturación promedio", "En precaución", "En crítico"),
    ):
        with col:
            st.markdown(
                f'<div class="mini-stat"><div class="n">{valor}</div><div class="l">{etiqueta}</div></div>',
                unsafe_allow_html=True,
            )

    # ---- Semáforo del equipo ----
    st.write("")
    st.markdown("""
    <div class="table-card-header">
        <div class="title">Semáforo del equipo</div>
    </div>
    """, unsafe_allow_html=True)

    if not semanas_registradas:
        st.info("Todavía no hay horas registradas. Cada colaborador puede capturar su primera semana desde su acceso.")
    else:
        semana_sel = st.selectbox(
            "Semana a visualizar",
            semanas_registradas,
            format_func=etiqueta_semana,
            key="semana_semaforo",
        )

        # Estado de cada colaborador activo en la semana seleccionada
        estados_semana = []
        for r in df_col_activos.itertuples():
            id_c = str(r.ID_Colaborador)
            resultado = saturacion_semana(id_c, semana_sel)
            if resultado is None or resultado[1] == 0:
                estados_semana.append((id_c, None, 0))
            else:
                estados_semana.append((id_c, resultado[0], resultado[1]))
        # Críticos primero, sin registro al final
        estados_semana.sort(key=lambda x: -1 if x[1] is None else x[1], reverse=True)

        criticos = [(i, s, h) for i, s, h in estados_semana if s is not None and s >= 95]
        precaucion = [(i, s, h) for i, s, h in estados_semana if s is not None and 90 <= s < 95]
        sin_registro = [i for i, s, h in estados_semana if s is None]

        if criticos:
            lineas = [
                f"🔴 **{mapa_nombre_col.get(i, i)}** al **{s:.0f}%** de su capacidad ({h:.1f} h de {mapa_capacidad_col.get(i, 0):.0f} h)"
                for i, s, h in criticos
            ]
            st.error(
                "**Alerta crítica de sobrecarga**\n\n" + "\n\n".join(lineas)
                + "\n\nConsidera redistribuir tareas o ajustar el alcance de sus proyectos."
            )
        if precaucion:
            lineas = [
                f"🟡 **{mapa_nombre_col.get(i, i)}** al **{s:.0f}%** de su capacidad ({h:.1f} h de {mapa_capacidad_col.get(i, 0):.0f} h)"
                for i, s, h in precaucion
            ]
            st.warning("**En zona de precaución**\n\n" + "\n\n".join(lineas))
        if not criticos and not precaucion:
            st.success("Nadie está en zona de precaución ni crítica en esta semana.")

        columnas_grid = st.columns(2)
        pos = 0
        for id_c, sat, horas in estados_semana:
            nombre = mapa_nombre_col.get(id_c, id_c)
            rol = mapa_rol_col.get(id_c, "")
            capacidad = mapa_capacidad_col.get(id_c, 0)
            iniciales = "".join(p[0].upper() for p in str(nombre).split()[:2]) or "?"
            with columnas_grid[pos % 2]:
                with st.container(border=True):
                    a, b, c = st.columns([0.7, 2.6, 1.7])
                    if sat is None:
                        with a:
                            st.markdown(f'<div class="avatar-circle libre">{iniciales}</div>', unsafe_allow_html=True)
                        with b:
                            st.markdown(f"**{nombre}**  \n{rol}")
                        with c:
                            st.markdown('<span class="badge badge-carga-libre">Sin registro</span>', unsafe_allow_html=True)
                        st.markdown('<div class="sat-track"></div>', unsafe_allow_html=True)
                        st.caption(f"0 h de {capacidad:.0f} h disponibles")
                    else:
                        etiqueta_sat, clase_badge, clase_barra, clase_avatar = clasificar_saturacion(sat)
                        percibida_media = df_salud[
                            (df_salud["ID_Colaborador"].astype(str) == id_c) & (df_salud["Semana"] == semana_sel)
                        ]["Carga_Percibida"].mean()
                        emoji = PERCIBIDA_EMOJI.get(int(round(percibida_media)), "😐")
                        with a:
                            st.markdown(f'<div class="avatar-circle {clase_avatar}">{iniciales}</div>', unsafe_allow_html=True)
                        with b:
                            st.markdown(f"**{nombre}**  \n{rol}")
                        with c:
                            st.markdown(
                                f'<span class="badge {clase_badge}">{etiqueta_sat} · {sat:.0f}%</span>',
                                unsafe_allow_html=True,
                            )
                        st.markdown(
                            f'<div class="sat-track"><div class="sat-fill {clase_barra}" style="width:{min(sat, 100):.0f}%"></div></div>',
                            unsafe_allow_html=True,
                        )
                        st.caption(f"{horas:.1f} h de {capacidad:.0f} h · Sensación de carga: {emoji}")
            pos += 1

        if sin_registro:
            nombres_sin = ", ".join(mapa_nombre_col.get(i, i) for i in sin_registro)
            st.caption(f"⚪ Sin horas registradas esta semana: {nombres_sin}")

    # ---- Nota sobre el auto-reporte ----
    st.write("")
    st.info(
        "Las horas y la carga percibida las reporta cada colaborador desde su propio acceso "
        "(perfil «Colaborador» en la pantalla de inicio). Si un registro tiene un error, "
        "pide a la persona que lo corrija desde su sesión."
    )

    # ---- Historial completo ----
    st.write("")
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Historial de registros del equipo</div>
        <div class="count-pill">{len(df_salud)} en total</div>
    </div>
    """, unsafe_allow_html=True)

    if df_salud.empty:
        st.info("Todavía no hay registros de horas.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            filtro_col = st.selectbox(
                "Colaborador", ["Todos"] + sorted(df_salud["Nombre_Colaborador"].astype(str).unique()),
                key="filtro_salud_col",
            )
        with f2:
            filtro_proy = st.selectbox(
                "Proyecto", ["Todos"] + sorted(df_salud["Nombre_Proyecto"].astype(str).unique()),
                key="filtro_salud_proy",
            )
        with f3:
            filtro_semana = st.selectbox(
                "Semana", ["Todas"] + semanas_registradas,
                format_func=lambda s: s if s == "Todas" else etiqueta_semana(s),
                key="filtro_salud_semana",
            )

        df_filtrado_sal = df_salud
        if filtro_col != "Todos":
            df_filtrado_sal = df_filtrado_sal[df_filtrado_sal["Nombre_Colaborador"].astype(str) == filtro_col]
        if filtro_proy != "Todos":
            df_filtrado_sal = df_filtrado_sal[df_filtrado_sal["Nombre_Proyecto"].astype(str) == filtro_proy]
        if filtro_semana != "Todas":
            df_filtrado_sal = df_filtrado_sal[df_filtrado_sal["Semana"] == filtro_semana]

        st.dataframe(
            df_filtrado_sal.sort_values(["Semana", "Nombre_Colaborador"], ascending=[False, True]),
            use_container_width=True,
            hide_index=True,
            column_config={
                "ID_Registro": st.column_config.TextColumn("ID"),
                "Semana": st.column_config.TextColumn("Semana del"),
                "Nombre_Colaborador": st.column_config.TextColumn("Colaborador"),
                "Nombre_Proyecto": st.column_config.TextColumn("Proyecto"),
                "Horas_Asignadas": st.column_config.NumberColumn("Asignadas", format="%.1f h"),
                "Horas_Reales": st.column_config.NumberColumn("Reales", format="%.1f h"),
                "Carga_Percibida": st.column_config.ProgressColumn(
                    "Carga percibida", format="%d/5", min_value=1, max_value=5
                ),
                "Capacidad_Horas_Semana": st.column_config.NumberColumn("Capacidad", format="%.0f h"),
            },
        )
        st.caption(f"Mostrando {len(df_filtrado_sal)} de {len(df_salud)} registro(s)")

    # ---- Registro de alertas emitidas (trazabilidad) ----
    st.write("")
    df_alertas_log = leer_hoja("Alertas", COLS_ALERTAS)
    with st.expander(f"Registro de alertas emitidas ({len(df_alertas_log)} en total)"):
        st.caption(
            "Historial automático de todas las alertas que el sistema ha generado "
            "(saturación, presupuesto y cronograma). Sirve de fuente para las notificaciones "
            "automáticas y para el seguimiento de cada aviso."
        )
        if df_alertas_log.empty:
            st.info("Todavía no se ha emitido ninguna alerta.")
        else:
            st.dataframe(
                df_alertas_log.sort_values("Fecha_Hora", ascending=False),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "ID_Alerta": st.column_config.TextColumn("ID"),
                    "Fecha_Hora": st.column_config.TextColumn("Fecha y hora"),
                    "Nombre_Referencia": st.column_config.TextColumn("Colaborador / Proyecto"),
                    "Valor_Indicador": st.column_config.NumberColumn("Valor", format="%.1f"),
                    "ID_Referencia": None,
                    "ID_Proyecto": None,
                },
            )


# =========================================================
# MÓDULO 4: PRESUPUESTO DETALLADO (PARTIDAS Y GASTOS POR RUBRO)
# =========================================================
elif modulo.startswith("4"):
    st.markdown("""
    <div class="banner">
        <h1>Módulo 4 · Presupuesto detallado</h1>
        <p>Define cuánto se planea gastar en cada rubro y registra los gastos reales para compararlos al instante.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="focus-callout">
        <strong>Cómo funciona el control presupuestal:</strong> primero define las
        <strong>partidas planificadas</strong> de cada proyecto por rubro (Personal, Software,
        Consultoría...). Luego registra cada <strong>gasto real</strong> contra ese rubro:
        la tabla de control te muestra al instante si el rubro va
        🟢 <strong>en rango</strong> (&lt;90% ejecutado) · 🟡 <strong>al límite</strong> (90–100%) ·
        🔴 <strong>excedido</strong> (&gt;100%), y cada exceso queda registrado en la hoja de alertas.
    </div>
    """, unsafe_allow_html=True)

    if "flash_gasto" in st.session_state:
        for tipo, texto in st.session_state.pop("flash_gasto"):
            getattr(st, tipo)(texto)

    # ---- Preparación de datos comunes del módulo ----
    opciones_proy_gasto = [f"{r.ID_Proyecto} · {r.Nombre_Proyecto}" for r in df_proyectos.itertuples()]
    mapa_proy_gasto = dict(zip(opciones_proy_gasto, df_proyectos["ID_Proyecto"].astype(str)))
    mapa_nombre_proy_g = dict(zip(df_proyectos["ID_Proyecto"].astype(str), df_proyectos["Nombre_Proyecto"]))

    df_plan_g = df_presupuesto[df_presupuesto["Tipo"] == "Planificado"]
    df_real_g = df_presupuesto[df_presupuesto["Tipo"] == "Real"]
    total_plan = df_plan_g["Monto"].sum()
    total_real = df_real_g["Monto"].sum()
    pct_ejecucion = total_real / total_plan * 100 if total_plan > 0 else 0

    # Suma plan vs. real por (proyecto, rubro) para el control y las alertas
    if df_presupuesto.empty:
        agg_rubros = pd.DataFrame(columns=["Planificado", "Real"])
    else:
        agg_rubros = df_presupuesto.groupby(["ID_Proyecto", "Categoria", "Tipo"])["Monto"].sum().unstack(fill_value=0)
        for _t in ("Planificado", "Real"):
            if _t not in agg_rubros.columns:
                agg_rubros[_t] = 0

    rubros_excedidos = [
        (idx, fila) for idx, fila in agg_rubros.iterrows()
        if fila["Planificado"] > 0 and fila["Real"] > fila["Planificado"]
    ]
    rubros_al_limite = [
        (idx, fila) for idx, fila in agg_rubros.iterrows()
        if fila["Planificado"] > 0 and 0.9 * fila["Planificado"] <= fila["Real"] <= fila["Planificado"]
    ]
    rubros_sin_partida = [
        (idx, fila) for idx, fila in agg_rubros.iterrows()
        if fila["Planificado"] == 0 and fila["Real"] > 0
    ]

    def estado_rubro(planificado, ejecutado):
        """Semáforo de un rubro según su % de ejecución frente a la partida."""
        if planificado <= 0:
            return "🟠 Sin partida" if ejecutado > 0 else "⚪ Sin movimientos"
        pct = ejecutado / planificado * 100
        if pct > 100:
            return "🔴 Excedido"
        if pct >= 90:
            return "🟡 Al límite"
        if ejecutado == 0:
            return "⚪ Sin gasto aún"
        return "🟢 En rango"

    # ---- Resumen rápido ----
    s1, s2, s3, s4 = st.columns(4)
    for col, valor, etiqueta in zip(
        (s1, s2, s3, s4),
        (f"${total_plan:,.0f}", f"${total_real:,.0f}", f"{pct_ejecucion:.0f}%", len(rubros_excedidos)),
        ("Total planificado", "Total ejecutado", "Ejecución global", "Rubros excedidos"),
    ):
        with col:
            st.markdown(
                f'<div class="mini-stat"><div class="n">{valor}</div><div class="l">{etiqueta}</div></div>',
                unsafe_allow_html=True,
            )

    # ---- Rubros en alerta (panorama de todos los proyectos) ----
    st.write("")
    if not df_presupuesto.empty:
        if rubros_excedidos:
            lineas = [
                f"🔴 **{idx[0]} · {idx[1]}**: ${fila['Real']:,.0f} ejecutados de ${fila['Planificado']:,.0f} planificados (+{(fila['Real'] / fila['Planificado'] - 1) * 100:.0f}%)"
                for idx, fila in rubros_excedidos
            ]
            st.error(
                "**Rubros con la partida excedida**\n\n" + "\n\n".join(lineas)
                + "\n\nRevisa el rubro con el líder del proyecto: amplía la partida o frena el gasto."
            )
        if rubros_al_limite:
            lineas = [
                f"🟡 **{idx[0]} · {idx[1]}**: ${fila['Real']:,.0f} de ${fila['Planificado']:,.0f} ({fila['Real'] / fila['Planificado'] * 100:.0f}% ejecutado)"
                for idx, fila in rubros_al_limite
            ]
            st.warning("**Rubros al límite de su partida (90%)**\n\n" + "\n\n".join(lineas))
        if not rubros_excedidos and not rubros_al_limite:
            st.success("Ningún rubro excede ni está al límite de su partida planificada.")
        if rubros_sin_partida:
            nombres_sp = ", ".join(f"{idx[0]} · {idx[1]}" for idx, _ in rubros_sin_partida)
            st.caption(f"🟠 Gastos sin partida planificada: {nombres_sp}. Define la partida para poder compararlos.")

    # ---- Control plan vs. real por rubro ----
    st.write("")
    st.markdown("""
    <div class="table-card-header">
        <div class="title">Control por rubro del proyecto</div>
    </div>
    """, unsafe_allow_html=True)

    if not opciones_proy_gasto:
        st.info("Todavía no hay proyectos. Crea al menos uno en el Módulo 1 para armar su presupuesto.")
    else:
        proy_control_sel = st.selectbox(
            "Proyecto a revisar", opciones_proy_gasto, key="sel_control_presupuesto",
        )
        id_proy_control = mapa_proy_gasto[proy_control_sel]
        plan_proy = df_plan_g[df_plan_g["ID_Proyecto"].astype(str) == id_proy_control].groupby("Categoria")["Monto"].sum()
        real_proy = df_real_g[df_real_g["ID_Proyecto"].astype(str) == id_proy_control].groupby("Categoria")["Monto"].sum()
        cats_extra = sorted(set(plan_proy.index).union(real_proy.index) - set(CATEGORIAS_GASTO))

        filas_control = []
        for cat in CATEGORIAS_GASTO + cats_extra:
            p = float(plan_proy.get(cat, 0))
            r = float(real_proy.get(cat, 0))
            if p == 0 and r == 0:
                continue
            filas_control.append({
                "Rubro": cat,
                "Planificado": p,
                "Ejecutado": r,
                "Disponible": p - r,
                "Ejecucion_Pct": round(r / p * 100, 1) if p > 0 else None,
                "Estado": estado_rubro(p, r),
            })

        if not filas_control:
            st.info("Este proyecto todavía no tiene partidas ni gastos. Empieza definiendo sus partidas planificadas abajo.")
        else:
            p_tot = sum(f["Planificado"] for f in filas_control)
            r_tot = sum(f["Ejecutado"] for f in filas_control)
            filas_control.append({
                "Rubro": "TOTAL",
                "Planificado": p_tot,
                "Ejecutado": r_tot,
                "Disponible": p_tot - r_tot,
                "Ejecucion_Pct": round(r_tot / p_tot * 100, 1) if p_tot > 0 else None,
                "Estado": estado_rubro(p_tot, r_tot),
            })
            st.dataframe(
                pd.DataFrame(filas_control),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Rubro": st.column_config.TextColumn("Rubro"),
                    "Planificado": st.column_config.NumberColumn("Planificado", format="$%.0f"),
                    "Ejecutado": st.column_config.NumberColumn("Ejecutado", format="$%.0f"),
                    "Disponible": st.column_config.NumberColumn("Disponible", format="$%.0f"),
                    "Ejecucion_Pct": st.column_config.NumberColumn("% ejecución", format="%.0f%%"),
                    "Estado": st.column_config.TextColumn("Estado"),
                },
            )

            # Coherencia con los montos globales capturados en el Módulo 1
            fila_proy_m1 = df_proyectos[df_proyectos["ID_Proyecto"].astype(str) == id_proy_control]
            if not fila_proy_m1.empty:
                presup_m1 = float(pd.to_numeric(fila_proy_m1["Presupuesto_Planificado"], errors="coerce").fillna(0).iloc[0])
                ejec_m1 = float(pd.to_numeric(fila_proy_m1["Presupuesto_Ejecutado"], errors="coerce").fillna(0).iloc[0])
                avisos_coherencia = []
                if presup_m1 > 0 and abs(p_tot - presup_m1) > 1:
                    avisos_coherencia.append(
                        f"las partidas suman **${p_tot:,.0f}** y el presupuesto planificado del Módulo 1 es **${presup_m1:,.0f}**"
                    )
                if abs(r_tot - ejec_m1) > 1:
                    avisos_coherencia.append(
                        f"los gastos reales suman **${r_tot:,.0f}** y el presupuesto ejecutado del Módulo 1 es **${ejec_m1:,.0f}**"
                    )
                if avisos_coherencia:
                    st.info(
                        "Diferencia con el Módulo 1: " + " y ".join(avisos_coherencia)
                        + ". Actualiza el proyecto en el Módulo 1 o ajusta los movimientos para que ambos cuadren."
                    )

    # ---- Captura: partida planificada o gasto real ----
    st.write("")
    tab_partida, tab_gasto = st.tabs(["Definir partida planificada", "Registrar gasto real"])

    with tab_partida:
        with st.form("form_partida", clear_on_submit=True):
            st.markdown("""
            <div class="form-section-header">
                <div class="icon">01</div>
                <div>
                    <div class="title">Nueva partida planificada</div>
                    <div class="subtitle">Cuánto se piensa gastar en un rubro del proyecto</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                p_proyecto = st.selectbox("Proyecto *", opciones_proy_gasto, key="partida_proy") if opciones_proy_gasto else None
                p_categoria = st.selectbox("Rubro / categoría *", CATEGORIAS_GASTO, key="partida_cat")
            with col2:
                p_concepto = st.text_input("Concepto *", placeholder="Licencias anuales, horas del equipo...", key="partida_concepto")
                p_fecha = st.date_input("Fecha de la partida", value=date.today(), key="partida_fecha",
                                        help="Normalmente la fecha en que se aprueba el presupuesto del proyecto.")
            p_monto = st.number_input("Monto planificado ($) *", min_value=0.0, step=100.0, key="partida_monto")

            st.write("")
            submitted_partida = st.form_submit_button("Guardar partida planificada", use_container_width=True)

            if submitted_partida:
                errores_p = []
                if not p_proyecto:
                    errores_p.append("No hay proyectos registrados. Crea al menos uno en el Módulo 1.")
                if not p_concepto.strip():
                    errores_p.append("El concepto de la partida es obligatorio.")
                if p_monto <= 0:
                    errores_p.append("El monto debe ser mayor a 0.")

                if errores_p:
                    for e in errores_p:
                        st.error(e)
                else:
                    id_proy_p = mapa_proy_gasto[p_proyecto]
                    nueva_partida = {
                        "ID_Gasto": siguiente_id(df_presupuesto, "ID_Gasto", "GTO"),
                        "ID_Proyecto": id_proy_p,
                        "Nombre_Proyecto": mapa_nombre_proy_g.get(id_proy_p, id_proy_p),
                        "Fecha": p_fecha.strftime("%Y-%m-%d"),
                        "Categoria": p_categoria,
                        "Concepto": p_concepto.strip(),
                        "Tipo": "Planificado",
                        "Monto": p_monto,
                    }
                    df_presupuesto_act = pd.concat([df_presupuesto, pd.DataFrame([nueva_partida])], ignore_index=True)
                    guardar_hoja("Presupuesto", df_presupuesto_act)

                    mensajes_p = [("success", f"Partida de ${p_monto:,.0f} en «{p_categoria}» guardada para {id_proy_p}.")]
                    plan_rubro_previo = df_plan_g.loc[
                        (df_plan_g["ID_Proyecto"].astype(str) == id_proy_p) & (df_plan_g["Categoria"] == p_categoria), "Monto",
                    ].sum()
                    if plan_rubro_previo > 0:
                        mensajes_p.append(("info", f"El rubro «{p_categoria}» ya tenía partidas por ${plan_rubro_previo:,.0f}; con esta suma ${plan_rubro_previo + p_monto:,.0f} planificados."))
                    plan_total_nuevo = df_presupuesto_act.loc[
                        (df_presupuesto_act["ID_Proyecto"].astype(str) == id_proy_p)
                        & (df_presupuesto_act["Tipo"] == "Planificado"), "Monto",
                    ].sum()
                    presup_m1_p = pd.to_numeric(
                        df_proyectos.loc[df_proyectos["ID_Proyecto"].astype(str) == id_proy_p, "Presupuesto_Planificado"],
                        errors="coerce",
                    ).sum()
                    if presup_m1_p > 0 and plan_total_nuevo > presup_m1_p + 1:
                        mensajes_p.append(("warning", f"Las partidas de {id_proy_p} ya suman ${plan_total_nuevo:,.0f}, por encima del presupuesto planificado del Módulo 1 (${presup_m1_p:,.0f}). Revisa cuál de los dos debe ajustarse."))
                    st.session_state["flash_gasto"] = mensajes_p
                    st.rerun()

    with tab_gasto:
        with st.form("form_gasto", clear_on_submit=True):
            st.markdown("""
            <div class="form-section-header">
                <div class="icon">01</div>
                <div>
                    <div class="title">Nuevo gasto real</div>
                    <div class="subtitle">Qué se compró o contrató, contra qué rubro y por cuánto</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                g_proyecto = st.selectbox("Proyecto *", opciones_proy_gasto, key="gasto_proy") if opciones_proy_gasto else None
                g_categoria = st.selectbox("Rubro / categoría *", CATEGORIAS_GASTO, key="gasto_cat")
            with col2:
                g_concepto = st.text_input("Concepto *", placeholder="Compra de servidores, consultoría fase 1...", key="gasto_concepto")
                g_fecha = st.date_input("Fecha del gasto", value=date.today(), key="gasto_fecha")
            g_monto = st.number_input("Monto ($) *", min_value=0.0, step=100.0, key="gasto_monto")

            st.write("")
            submitted_gasto = st.form_submit_button("Guardar gasto real", use_container_width=True)

            if submitted_gasto:
                errores_g = []
                if not g_proyecto:
                    errores_g.append("No hay proyectos registrados. Crea al menos uno en el Módulo 1.")
                if not g_concepto.strip():
                    errores_g.append("El concepto del gasto es obligatorio.")
                if g_monto <= 0:
                    errores_g.append("El monto debe ser mayor a 0.")

                if errores_g:
                    for e in errores_g:
                        st.error(e)
                else:
                    id_proy_g = mapa_proy_gasto[g_proyecto]
                    nuevo_gasto = {
                        "ID_Gasto": siguiente_id(df_presupuesto, "ID_Gasto", "GTO"),
                        "ID_Proyecto": id_proy_g,
                        "Nombre_Proyecto": mapa_nombre_proy_g.get(id_proy_g, id_proy_g),
                        "Fecha": g_fecha.strftime("%Y-%m-%d"),
                        "Categoria": g_categoria,
                        "Concepto": g_concepto.strip(),
                        "Tipo": "Real",
                        "Monto": g_monto,
                    }
                    df_presupuesto_act = pd.concat([df_presupuesto, pd.DataFrame([nuevo_gasto])], ignore_index=True)
                    guardar_hoja("Presupuesto", df_presupuesto_act)

                    mensajes_g = [("success", f"Gasto de ${g_monto:,.0f} en «{g_categoria}» guardado en {id_proy_g}.")]

                    # Control por rubro: comparar lo ejecutado contra la partida del rubro
                    plan_rubro = df_plan_g.loc[
                        (df_plan_g["ID_Proyecto"].astype(str) == id_proy_g) & (df_plan_g["Categoria"] == g_categoria), "Monto",
                    ].sum()
                    real_rubro = df_presupuesto_act.loc[
                        (df_presupuesto_act["ID_Proyecto"].astype(str) == id_proy_g)
                        & (df_presupuesto_act["Categoria"] == g_categoria)
                        & (df_presupuesto_act["Tipo"] == "Real"), "Monto",
                    ].sum()
                    if plan_rubro <= 0:
                        mensajes_g.append(("info", f"El rubro «{g_categoria}» de {id_proy_g} no tiene partida planificada. Defínela en la pestaña «Definir partida planificada» para poder controlarlo."))
                    elif real_rubro > plan_rubro:
                        pct_rubro = (real_rubro / plan_rubro - 1) * 100
                        nivel_rubro = "Crítico" if pct_rubro > 10 else "Precaución"
                        mensajes_g.append(("warning", f"Con este gasto, el rubro «{g_categoria}» de {id_proy_g} queda en ${real_rubro:,.0f}, por encima de su partida de ${plan_rubro:,.0f} (+{pct_rubro:.1f}%). Se dejó registro en la hoja de alertas."))
                        registrar_alerta("Presupuesto", nivel_rubro, id_proy_g, mapa_nombre_proy_g.get(id_proy_g, id_proy_g), id_proy_g, "",
                                         round(real_rubro, 0), round(plan_rubro, 0),
                                         f"El rubro «{g_categoria}» de {id_proy_g} lleva ${real_rubro:,.0f} ejecutados sobre ${plan_rubro:,.0f} planificados (+{pct_rubro:.1f}%).")
                    elif real_rubro >= 0.9 * plan_rubro:
                        mensajes_g.append(("warning", f"🟡 El rubro «{g_categoria}» de {id_proy_g} llega al {real_rubro / plan_rubro * 100:.0f}% de su partida (${real_rubro:,.0f} de ${plan_rubro:,.0f}). Quedan ${plan_rubro - real_rubro:,.0f} disponibles."))

                    # Control del total del proyecto frente al presupuesto del Módulo 1
                    real_proy_total = df_presupuesto_act.loc[
                        (df_presupuesto_act["ID_Proyecto"].astype(str) == id_proy_g)
                        & (df_presupuesto_act["Tipo"] == "Real"), "Monto",
                    ].sum()
                    presup_proy = pd.to_numeric(
                        df_proyectos.loc[df_proyectos["ID_Proyecto"].astype(str) == id_proy_g, "Presupuesto_Planificado"],
                        errors="coerce",
                    ).sum()
                    if presup_proy > 0 and real_proy_total > presup_proy:
                        pct_desvio = (real_proy_total / presup_proy - 1) * 100
                        nivel_desvio = "Crítico" if pct_desvio > 10 else "Precaución"
                        mensajes_g.append(("warning", f"Con este gasto, lo ejecutado en {id_proy_g} (${real_proy_total:,.0f}) supera su presupuesto planificado (${presup_proy:,.0f}) en {pct_desvio:.1f}%. Se dejó registro en la hoja de alertas."))
                        registrar_alerta("Presupuesto", nivel_desvio, id_proy_g, mapa_nombre_proy_g.get(id_proy_g, id_proy_g), id_proy_g, "",
                                         round(real_proy_total, 0), round(presup_proy, 0),
                                         f"Lo ejecutado (${real_proy_total:,.0f}) supera el presupuesto planificado (${presup_proy:,.0f}) en {pct_desvio:.1f}%.")
                    st.session_state["flash_gasto"] = mensajes_g
                    st.rerun()

    # ---- Editar o eliminar un movimiento ----
    st.write("")
    with st.expander("Editar o eliminar un movimiento existente", expanded=False):
        if df_presupuesto.empty:
            st.caption("Todavía no hay movimientos para editar.")
        else:
            etiquetas_g = [
                f"{r.ID_Gasto} · {r.ID_Proyecto} · {r.Tipo} · {r.Concepto} (${r.Monto:,.0f})"
                for r in df_presupuesto.itertuples()
            ]
            mapa_id_g = dict(zip(etiquetas_g, df_presupuesto["ID_Gasto"].astype(str)))
            etiqueta_g_sel = st.selectbox("Selecciona un movimiento", etiquetas_g, key="sel_editar_gasto")
            id_g_sel = mapa_id_g[etiqueta_g_sel]
            fila_g = df_presupuesto[df_presupuesto["ID_Gasto"].astype(str) == id_g_sel].iloc[0]

            etiqueta_proy_g_actual = next(
                (e for e, i in mapa_proy_gasto.items() if i == str(fila_g["ID_Proyecto"])),
                opciones_proy_gasto[0] if opciones_proy_gasto else None,
            )
            try:
                fecha_g_val = pd.to_datetime(str(fila_g["Fecha"])[:10]).date()
            except (ValueError, TypeError):
                fecha_g_val = date.today()

            with st.form(f"form_editar_gasto_{id_g_sel}"):
                col1, col2 = st.columns(2)
                with col1:
                    eg_proyecto = st.selectbox(
                        "Proyecto *", opciones_proy_gasto,
                        index=opciones_proy_gasto.index(etiqueta_proy_g_actual) if etiqueta_proy_g_actual in opciones_proy_gasto else 0,
                        key=f"eg_proy_{id_g_sel}",
                    )
                    eg_categoria = st.selectbox(
                        "Rubro / categoría *", CATEGORIAS_GASTO,
                        index=CATEGORIAS_GASTO.index(fila_g["Categoria"]) if fila_g["Categoria"] in CATEGORIAS_GASTO else 0,
                        key=f"eg_cat_{id_g_sel}",
                    )
                    eg_tipo = st.selectbox(
                        "Tipo *", ["Planificado", "Real"],
                        index=0 if fila_g["Tipo"] == "Planificado" else 1,
                        key=f"eg_tipo_{id_g_sel}",
                        help="«Planificado»: partida presupuestada. «Real»: gasto ya ejecutado.",
                    )
                with col2:
                    eg_fecha = st.date_input("Fecha del movimiento", value=fecha_g_val, key=f"eg_fecha_{id_g_sel}")
                    eg_concepto = st.text_input("Concepto *", value=str(fila_g["Concepto"]), key=f"eg_concepto_{id_g_sel}")
                    eg_monto = st.number_input(
                        "Monto ($) *", min_value=0.0, step=100.0,
                        value=float(fila_g["Monto"] or 0), key=f"eg_monto_{id_g_sel}",
                    )

                eg_confirmar = st.checkbox(
                    "Confirmo que deseo eliminar este movimiento de forma permanente",
                    key=f"eg_confirmar_{id_g_sel}",
                )
                st.write("")
                b1, b2 = st.columns(2)
                with b1:
                    actualizar_g = st.form_submit_button("Guardar cambios", use_container_width=True)
                with b2:
                    eliminar_g = st.form_submit_button("Eliminar movimiento", use_container_width=True)

                if actualizar_g:
                    if not eg_concepto.strip():
                        st.error("El concepto del movimiento es obligatorio.")
                    elif eg_monto <= 0:
                        st.error("El monto debe ser mayor a 0.")
                    else:
                        id_proy_eg = mapa_proy_gasto[eg_proyecto]
                        idx_g = fila_g.name
                        df_presupuesto.loc[idx_g, "ID_Proyecto"] = id_proy_eg
                        df_presupuesto.loc[idx_g, "Nombre_Proyecto"] = mapa_nombre_proy_g.get(id_proy_eg, id_proy_eg)
                        df_presupuesto.loc[idx_g, "Fecha"] = eg_fecha.strftime("%Y-%m-%d")
                        df_presupuesto.loc[idx_g, "Categoria"] = eg_categoria
                        df_presupuesto.loc[idx_g, "Concepto"] = eg_concepto.strip()
                        df_presupuesto.loc[idx_g, "Tipo"] = eg_tipo
                        df_presupuesto.loc[idx_g, "Monto"] = eg_monto
                        guardar_hoja("Presupuesto", df_presupuesto)

                        mensajes_eg = [("success", f"Movimiento {id_g_sel} actualizado correctamente.")]
                        if eg_tipo == "Real":
                            plan_rubro_e = df_presupuesto.loc[
                                (df_presupuesto["ID_Proyecto"].astype(str) == id_proy_eg)
                                & (df_presupuesto["Categoria"] == eg_categoria)
                                & (df_presupuesto["Tipo"] == "Planificado"), "Monto",
                            ].sum()
                            real_rubro_e = df_presupuesto.loc[
                                (df_presupuesto["ID_Proyecto"].astype(str) == id_proy_eg)
                                & (df_presupuesto["Categoria"] == eg_categoria)
                                & (df_presupuesto["Tipo"] == "Real"), "Monto",
                            ].sum()
                            if plan_rubro_e > 0 and real_rubro_e > plan_rubro_e:
                                pct_rubro_e = (real_rubro_e / plan_rubro_e - 1) * 100
                                mensajes_eg.append(("warning", f"El rubro «{eg_categoria}» de {id_proy_eg} queda en ${real_rubro_e:,.0f}, por encima de su partida de ${plan_rubro_e:,.0f} (+{pct_rubro_e:.1f}%). Se dejó registro en la hoja de alertas."))
                                registrar_alerta("Presupuesto", "Crítico" if pct_rubro_e > 10 else "Precaución",
                                                 id_proy_eg, mapa_nombre_proy_g.get(id_proy_eg, id_proy_eg), id_proy_eg, "",
                                                 round(real_rubro_e, 0), round(plan_rubro_e, 0),
                                                 f"El rubro «{eg_categoria}» de {id_proy_eg} lleva ${real_rubro_e:,.0f} ejecutados sobre ${plan_rubro_e:,.0f} planificados (+{pct_rubro_e:.1f}%).")
                        st.session_state["flash_gasto"] = mensajes_eg
                        st.rerun()

                elif eliminar_g:
                    if not eg_confirmar:
                        st.error("Marca la casilla de confirmación para poder eliminar el movimiento.")
                    else:
                        df_presupuesto = df_presupuesto.drop(index=fila_g.name)
                        guardar_hoja("Presupuesto", df_presupuesto)
                        st.session_state["flash_gasto"] = [("success", f"Movimiento {id_g_sel} eliminado correctamente.")]
                        st.rerun()

    # ---- Historial de movimientos ----
    st.write("")
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Movimientos registrados</div>
        <div class="count-pill">{len(df_presupuesto)} en total</div>
    </div>
    """, unsafe_allow_html=True)

    if df_presupuesto.empty:
        st.info("Todavía no hay movimientos registrados.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            filtro_proy_g = st.selectbox("Proyecto", ["Todos"] + sorted(df_presupuesto["Nombre_Proyecto"].astype(str).unique()), key="filtro_g_proy")
        with f2:
            filtro_tipo_g = st.selectbox("Tipo", ["Todos", "Planificado", "Real"], key="filtro_g_tipo")
        with f3:
            filtro_cat_g = st.selectbox("Categoría", ["Todas"] + CATEGORIAS_GASTO, key="filtro_g_cat")

        df_filtrado_g = df_presupuesto
        if filtro_proy_g != "Todos":
            df_filtrado_g = df_filtrado_g[df_filtrado_g["Nombre_Proyecto"].astype(str) == filtro_proy_g]
        if filtro_tipo_g != "Todos":
            df_filtrado_g = df_filtrado_g[df_filtrado_g["Tipo"] == filtro_tipo_g]
        if filtro_cat_g != "Todas":
            df_filtrado_g = df_filtrado_g[df_filtrado_g["Categoria"] == filtro_cat_g]

        st.dataframe(
            df_filtrado_g.sort_values(["Fecha"], ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "ID_Gasto": st.column_config.TextColumn("ID"),
                "Nombre_Proyecto": st.column_config.TextColumn("Proyecto"),
                "Monto": st.column_config.NumberColumn("Monto", format="$%.0f"),
            },
        )
        suma_plan_filtrada = df_filtrado_g.loc[df_filtrado_g["Tipo"] == "Planificado", "Monto"].sum()
        suma_real_filtrada = df_filtrado_g.loc[df_filtrado_g["Tipo"] == "Real", "Monto"].sum()
        st.caption(
            f"Mostrando {len(df_filtrado_g)} de {len(df_presupuesto)} movimiento(s) · "
            f"Partidas planificadas: ${suma_plan_filtrada:,.0f} · Gastos reales: ${suma_real_filtrada:,.0f}"
        )


# =========================================================
# MÓDULO 5: TAREAS E HITOS (CONTROL DE CRONOGRAMA)
# =========================================================
elif modulo.startswith("5"):
    st.markdown("""
    <div class="banner">
        <h1>Módulo 5 · Tareas e hitos</h1>
        <p>Planifica las tareas y los hitos de cada proyecto, y detecta a tiempo lo vencido y lo que está por vencer.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="focus-callout">
        <strong>Cómo leer el control de cronograma:</strong> cada tarea o hito compara su fecha
        planificada con el día de hoy y con la fecha real de cierre:
        🔴 <strong>vencida</strong> (la fecha pasó y no se ha completado) ·
        🟠 <strong>vence en 7 días</strong> · 🔵 <strong>en plazo</strong> ·
        🟢 <strong>completada a tiempo</strong> · 🟡 <strong>completada con retraso</strong>.
        Los vencimientos quedan registrados en la hoja de alertas para el seguimiento del portafolio.
    </div>
    """, unsafe_allow_html=True)

    if "flash_tarea" in st.session_state:
        for tipo, texto in st.session_state.pop("flash_tarea"):
            getattr(st, tipo)(texto)

    # ---- Preparación: situación de cronograma de cada tarea ----
    hoy = date.today()

    def fecha_o_none(valor):
        """Convierte el valor de una celda de fecha en date, o None si está vacía."""
        f = pd.to_datetime(str(valor)[:10], errors="coerce")
        return None if pd.isna(f) else f.date()

    def situacion_tarea(fecha_plan, fecha_real, estado):
        """Clasifica la situación de una tarea frente a hoy. Devuelve (texto, clave)."""
        plan = fecha_o_none(fecha_plan)
        real = fecha_o_none(fecha_real)
        if estado == "Completada" or real is not None:
            if plan and real and real > plan:
                return f"🟡 Completada con retraso (+{(real - plan).days} d)", "tarde"
            if real is None:
                return "🟢 Completada (sin fecha real)", "ok"
            return "🟢 Completada a tiempo", "ok"
        if plan is not None:
            if plan < hoy:
                return f"🔴 Vencida hace {(hoy - plan).days} d", "vencida"
            dias = (plan - hoy).days
            if dias == 0:
                return "🟠 Vence hoy", "pronto"
            if dias <= 7:
                return f"🟠 Vence en {dias} d", "pronto"
        return "🔵 En plazo", "plazo"

    df_tareas_ctrl = df_tareas.copy()
    if df_tareas_ctrl.empty:
        df_tareas_ctrl["Situacion"] = pd.Series(dtype=str)
        df_tareas_ctrl["_clave"] = pd.Series(dtype=str)
    else:
        _situaciones = [
            situacion_tarea(r["Fecha_Planificada"], r["Fecha_Real"], r["Estado"])
            for _, r in df_tareas_ctrl.iterrows()
        ]
        df_tareas_ctrl["Situacion"] = [s[0] for s in _situaciones]
        df_tareas_ctrl["_clave"] = [s[1] for s in _situaciones]

    completadas_t = int((df_tareas_ctrl["Estado"] == "Completada").sum()) if not df_tareas_ctrl.empty else 0
    vencidas_t = int((df_tareas_ctrl["_clave"] == "vencida").sum()) if not df_tareas_ctrl.empty else 0
    pronto_t = int((df_tareas_ctrl["_clave"] == "pronto").sum()) if not df_tareas_ctrl.empty else 0
    tarde_t = int((df_tareas_ctrl["_clave"] == "tarde").sum()) if not df_tareas_ctrl.empty else 0

    # ---- Resumen rápido ----
    s1, s2, s3, s4 = st.columns(4)
    for col, valor, etiqueta in zip(
        (s1, s2, s3, s4),
        (len(df_tareas_ctrl), completadas_t, vencidas_t, pronto_t),
        ("Tareas e hitos", "Completadas", "Vencidas", "Vencen en 7 días"),
    ):
        with col:
            st.markdown(
                f'<div class="mini-stat"><div class="n">{valor}</div><div class="l">{etiqueta}</div></div>',
                unsafe_allow_html=True,
            )

    # ---- Panorama del cronograma (todos los proyectos) ----
    st.write("")
    if not df_tareas_ctrl.empty:
        df_venc = df_tareas_ctrl[df_tareas_ctrl["_clave"] == "vencida"].sort_values("Fecha_Planificada")
        df_pronto = df_tareas_ctrl[df_tareas_ctrl["_clave"] == "pronto"].sort_values("Fecha_Planificada")

        if not df_venc.empty:
            lineas = [
                f"🔴 **{r.Nombre_Tarea}** ({str(r.Tipo).lower()} de {r.ID_Proyecto}) · {r.Nombre_Responsable} — {str(r.Situacion).replace('🔴 ', '').lower()}"
                for r in df_venc.itertuples()
            ]
            st.error(
                "**Tareas e hitos vencidos sin completar**\n\n" + "\n\n".join(lineas)
                + "\n\nReplanifica la fecha, márcalos como completados o revisa el bloqueo con el responsable."
            )
        if not df_pronto.empty:
            lineas = [
                f"🟠 **{r.Nombre_Tarea}** ({str(r.Tipo).lower()} de {r.ID_Proyecto}) · {r.Nombre_Responsable} — {str(r.Situacion).replace('🟠 ', '').lower()} ({str(r.Fecha_Planificada)[:10]})"
                for r in df_pronto.itertuples()
            ]
            st.warning("**Vencen en los próximos 7 días**\n\n" + "\n\n".join(lineas))
        if df_venc.empty and df_pronto.empty:
            st.success("No hay tareas vencidas ni vencimientos en los próximos 7 días.")
        if tarde_t:
            st.caption(f"🟡 {tarde_t} tarea(s) o hito(s) se completaron después de su fecha planificada — el detalle está en el control por proyecto.")

    # ---- Control de cronograma por proyecto ----
    st.write("")
    st.markdown("""
    <div class="table-card-header">
        <div class="title">Cronograma del proyecto</div>
    </div>
    """, unsafe_allow_html=True)

    opciones_proy_t = [f"{r.ID_Proyecto} · {r.Nombre_Proyecto}" for r in df_proyectos.itertuples()]
    mapa_proy_t = dict(zip(opciones_proy_t, df_proyectos["ID_Proyecto"].astype(str)))
    mapa_nombre_proy_t = dict(zip(df_proyectos["ID_Proyecto"].astype(str), df_proyectos["Nombre_Proyecto"]))

    if not opciones_proy_t:
        st.info("Todavía no hay proyectos. Crea al menos uno en el Módulo 1 para armar su cronograma.")
    else:
        proy_ctrl_sel = st.selectbox("Proyecto a revisar", opciones_proy_t, key="sel_control_tareas")
        id_proy_ctrl = mapa_proy_t[proy_ctrl_sel]
        df_crono = df_tareas_ctrl[df_tareas_ctrl["ID_Proyecto"].astype(str) == id_proy_ctrl].sort_values("Fecha_Planificada")

        if df_crono.empty:
            st.info("Este proyecto todavía no tiene tareas ni hitos. Registra el primero en el formulario de abajo.")
        else:
            st.dataframe(
                df_crono[["Fecha_Planificada", "Tipo", "Nombre_Tarea", "Nombre_Responsable", "Fecha_Real", "Estado", "Situacion"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Fecha_Planificada": st.column_config.TextColumn("Planificada"),
                    "Tipo": st.column_config.TextColumn("Tipo"),
                    "Nombre_Tarea": st.column_config.TextColumn("Tarea / Hito"),
                    "Nombre_Responsable": st.column_config.TextColumn("Responsable"),
                    "Fecha_Real": st.column_config.TextColumn("Real"),
                    "Estado": st.column_config.TextColumn("Estado"),
                    "Situacion": st.column_config.TextColumn("Situación"),
                },
            )
            st.caption(
                "Ordenado por fecha planificada: se lee de arriba hacia abajo como la línea de tiempo del proyecto."
            )

    # ---- Captura ----
    st.write("")
    df_col_activos_t = df_colaboradores[df_colaboradores["Estado"] == "Activo"]
    opciones_resp = [f"{r.ID_Colaborador} · {r.Nombre}" for r in df_col_activos_t.itertuples()]
    mapa_resp = dict(zip(opciones_resp, df_col_activos_t["ID_Colaborador"].astype(str)))
    opciones_resp_todos = [f"{r.ID_Colaborador} · {r.Nombre}" for r in df_colaboradores.itertuples()]
    mapa_resp_todos = dict(zip(opciones_resp_todos, df_colaboradores["ID_Colaborador"].astype(str)))
    mapa_nombre_resp = dict(zip(df_colaboradores["ID_Colaborador"].astype(str), df_colaboradores["Nombre"]))

    with st.form("form_tarea", clear_on_submit=True):
        st.markdown("""
        <div class="form-section-header">
            <div class="icon">01</div>
            <div>
                <div class="title">Qué se va a hacer</div>
                <div class="subtitle">Tarea o hito, proyecto al que pertenece y quién responde por él</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            t_proyecto = st.selectbox("Proyecto *", opciones_proy_t) if opciones_proy_t else None
            t_tipo = st.selectbox(
                "Tipo *", ["Tarea", "Hito"],
                help="«Hito»: momento clave que marca el avance (entrega, go-live). «Tarea»: trabajo a ejecutar.",
            )
        with col2:
            t_nombre = st.text_input("Nombre de la tarea o hito *", placeholder="Entrega del tablero v1, pruebas integrales...")
            t_responsable = st.selectbox("Responsable *", opciones_resp) if opciones_resp else None

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">02</div>
            <div>
                <div class="title">Fechas y estado</div>
                <div class="subtitle">Fecha comprometida vs. fecha en que realmente se completó</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col3, col4 = st.columns(2)
        with col3:
            t_fecha_plan = st.date_input("Fecha planificada", value=date.today())
            t_estado = st.selectbox("Estado *", ESTADOS_TAREA)
        with col4:
            t_completada = st.checkbox(
                "Ya se completó — registrar la fecha real",
                help="La fecha real permite medir la desviación del cronograma.",
            )
            t_fecha_real = st.date_input("Fecha real", value=date.today())

        st.write("")
        submitted_tarea = st.form_submit_button("Guardar tarea o hito", use_container_width=True)

        if submitted_tarea:
            errores_t = []
            if not t_proyecto:
                errores_t.append("No hay proyectos registrados. Crea al menos uno en el Módulo 1.")
            if not t_responsable:
                errores_t.append("No hay colaboradores activos. Registra al menos uno en el Módulo 2.")
            if not t_nombre.strip():
                errores_t.append("El nombre de la tarea o hito es obligatorio.")

            if errores_t:
                for e in errores_t:
                    st.error(e)
            else:
                id_proy_t = mapa_proy_t[t_proyecto]
                id_resp_t = mapa_resp[t_responsable]
                nueva_tarea = {
                    "ID_Tarea": siguiente_id(df_tareas, "ID_Tarea", "TAR"),
                    "ID_Proyecto": id_proy_t,
                    "Nombre_Proyecto": mapa_nombre_proy_t.get(id_proy_t, id_proy_t),
                    "Tipo": t_tipo,
                    "Nombre_Tarea": t_nombre.strip(),
                    "ID_Responsable": id_resp_t,
                    "Nombre_Responsable": mapa_nombre_resp.get(id_resp_t, id_resp_t),
                    "Fecha_Planificada": t_fecha_plan.strftime("%Y-%m-%d"),
                    "Fecha_Real": t_fecha_real.strftime("%Y-%m-%d") if t_completada else "",
                    "Estado": t_estado,
                }
                df_tareas_act = pd.concat([df_tareas, pd.DataFrame([nueva_tarea])], ignore_index=True)
                guardar_hoja("Tareas_Hitos", df_tareas_act)

                mensajes_t = [("success", f"{t_tipo} '{t_nombre}' guardado correctamente.")]
                if t_completada and t_estado != "Completada":
                    mensajes_t.append(("info", "Registraste una fecha real pero el estado no es «Completada». Revisa si el estado debería actualizarse."))
                elif t_estado == "Completada" and not t_completada:
                    mensajes_t.append(("info", "El estado es «Completada» pero no registraste la fecha real. Puedes agregarla editando el registro."))
                if t_completada and t_fecha_real > t_fecha_plan:
                    dias_desvio = (t_fecha_real - t_fecha_plan).days
                    mensajes_t.append(("info", f"Se completó {dias_desvio} día(s) después de lo planificado. La desviación quedará visible en el control de cronograma."))
                if t_estado in ("Pendiente", "En curso") and t_fecha_plan < date.today():
                    dias_vencidos = (date.today() - t_fecha_plan).days
                    mensajes_t.append(("warning", f"La fecha planificada ({t_fecha_plan.strftime('%d/%m/%Y')}) ya pasó hace {dias_vencidos} día(s). Considera marcarla como «Retrasada» o replanificarla. Se dejó registro en la hoja de alertas."))
                    registrar_alerta("Cronograma", "Precaución" if dias_vencidos <= 5 else "Crítico",
                                     id_resp_t, mapa_nombre_resp.get(id_resp_t, id_resp_t), id_proy_t, "",
                                     dias_vencidos, 5,
                                     f"{t_tipo} '{t_nombre.strip()}' de {id_proy_t} lleva {dias_vencidos} día(s) de retraso frente a la fecha planificada.")
                st.session_state["flash_tarea"] = mensajes_t
                st.rerun()

    # ---- Editar o eliminar ----
    st.write("")
    with st.expander("Editar o eliminar una tarea o hito existente", expanded=False):
        if df_tareas.empty:
            st.caption("Todavía no hay tareas o hitos para editar.")
        else:
            etiquetas_t = [
                f"{r.ID_Tarea} · {r.ID_Proyecto} · {r.Nombre_Tarea}"
                for r in df_tareas.itertuples()
            ]
            mapa_id_t = dict(zip(etiquetas_t, df_tareas["ID_Tarea"].astype(str)))
            etiqueta_t_sel = st.selectbox("Selecciona una tarea o hito", etiquetas_t, key="sel_editar_tarea")
            id_t_sel = mapa_id_t[etiqueta_t_sel]
            fila_t = df_tareas[df_tareas["ID_Tarea"].astype(str) == id_t_sel].iloc[0]

            etiqueta_proy_t_act = next(
                (e for e, i in mapa_proy_t.items() if i == str(fila_t["ID_Proyecto"])),
                opciones_proy_t[0] if opciones_proy_t else None,
            )
            etiqueta_resp_act = next(
                (e for e, i in mapa_resp_todos.items() if i == str(fila_t["ID_Responsable"])),
                opciones_resp_todos[0] if opciones_resp_todos else None,
            )
            try:
                fecha_plan_val = pd.to_datetime(str(fila_t["Fecha_Planificada"])[:10]).date()
            except (ValueError, TypeError):
                fecha_plan_val = date.today()
            fecha_real_str = str(fila_t["Fecha_Real"])[:10] if pd.notna(fila_t["Fecha_Real"]) else ""
            tiene_fecha_real = bool(fecha_real_str.strip()) and fecha_real_str.lower() != "nan"
            try:
                fecha_real_val = pd.to_datetime(fecha_real_str).date() if tiene_fecha_real else date.today()
            except (ValueError, TypeError):
                fecha_real_val = date.today()

            with st.form(f"form_editar_tarea_{id_t_sel}"):
                col1, col2 = st.columns(2)
                with col1:
                    et_proyecto = st.selectbox(
                        "Proyecto *", opciones_proy_t,
                        index=opciones_proy_t.index(etiqueta_proy_t_act) if etiqueta_proy_t_act in opciones_proy_t else 0,
                        key=f"et_proy_{id_t_sel}",
                    )
                    et_tipo = st.selectbox(
                        "Tipo *", ["Tarea", "Hito"],
                        index=0 if fila_t["Tipo"] == "Tarea" else 1, key=f"et_tipo_{id_t_sel}",
                    )
                    et_fecha_plan = st.date_input("Fecha planificada", value=fecha_plan_val, key=f"et_fplan_{id_t_sel}")
                with col2:
                    et_nombre = st.text_input("Nombre *", value=str(fila_t["Nombre_Tarea"]), key=f"et_nombre_{id_t_sel}")
                    et_responsable = st.selectbox(
                        "Responsable *", opciones_resp_todos,
                        index=opciones_resp_todos.index(etiqueta_resp_act) if etiqueta_resp_act in opciones_resp_todos else 0,
                        key=f"et_resp_{id_t_sel}",
                    )
                    et_estado = st.selectbox(
                        "Estado *", ESTADOS_TAREA,
                        index=ESTADOS_TAREA.index(fila_t["Estado"]) if fila_t["Estado"] in ESTADOS_TAREA else 0,
                        key=f"et_estado_{id_t_sel}",
                    )
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    et_completada = st.checkbox("Ya se completó", value=tiene_fecha_real, key=f"et_comp_{id_t_sel}")
                with col_f2:
                    et_fecha_real = st.date_input("Fecha real", value=fecha_real_val, key=f"et_freal_{id_t_sel}")

                et_confirmar = st.checkbox(
                    "Confirmo que deseo eliminar esta tarea o hito de forma permanente",
                    key=f"et_confirmar_{id_t_sel}",
                )
                st.write("")
                b1, b2 = st.columns(2)
                with b1:
                    actualizar_t = st.form_submit_button("Guardar cambios", use_container_width=True)
                with b2:
                    eliminar_t = st.form_submit_button("Eliminar", use_container_width=True)

                if actualizar_t:
                    if not et_nombre.strip():
                        st.error("El nombre de la tarea o hito es obligatorio.")
                    else:
                        id_proy_et = mapa_proy_t[et_proyecto]
                        id_resp_et = mapa_resp_todos[et_responsable]
                        idx_t = fila_t.name
                        df_tareas.loc[idx_t, "ID_Proyecto"] = id_proy_et
                        df_tareas.loc[idx_t, "Nombre_Proyecto"] = mapa_nombre_proy_t.get(id_proy_et, id_proy_et)
                        df_tareas.loc[idx_t, "Tipo"] = et_tipo
                        df_tareas.loc[idx_t, "Nombre_Tarea"] = et_nombre.strip()
                        df_tareas.loc[idx_t, "ID_Responsable"] = id_resp_et
                        df_tareas.loc[idx_t, "Nombre_Responsable"] = mapa_nombre_resp.get(id_resp_et, id_resp_et)
                        df_tareas.loc[idx_t, "Fecha_Planificada"] = et_fecha_plan.strftime("%Y-%m-%d")
                        df_tareas.loc[idx_t, "Fecha_Real"] = et_fecha_real.strftime("%Y-%m-%d") if et_completada else ""
                        df_tareas.loc[idx_t, "Estado"] = et_estado
                        guardar_hoja("Tareas_Hitos", df_tareas)

                        mensajes_et = [("success", f"Registro {id_t_sel} actualizado correctamente.")]
                        if et_completada and et_fecha_real > et_fecha_plan:
                            mensajes_et.append(("info", f"Quedó completada {(et_fecha_real - et_fecha_plan).days} día(s) después de lo planificado."))
                        if et_estado in ("Pendiente", "En curso") and et_fecha_plan < date.today():
                            dias_vencidos_e = (date.today() - et_fecha_plan).days
                            mensajes_et.append(("warning", f"La fecha planificada ya pasó hace {dias_vencidos_e} día(s) y el registro sigue «{et_estado}». Se dejó registro en la hoja de alertas."))
                            registrar_alerta("Cronograma", "Precaución" if dias_vencidos_e <= 5 else "Crítico",
                                             id_resp_et, mapa_nombre_resp.get(id_resp_et, id_resp_et), id_proy_et, "",
                                             dias_vencidos_e, 5,
                                             f"{et_tipo} '{et_nombre.strip()}' de {id_proy_et} lleva {dias_vencidos_e} día(s) de retraso frente a la fecha planificada.")
                        st.session_state["flash_tarea"] = mensajes_et
                        st.rerun()

                elif eliminar_t:
                    if not et_confirmar:
                        st.error("Marca la casilla de confirmación para poder eliminar el registro.")
                    else:
                        df_tareas = df_tareas.drop(index=fila_t.name)
                        guardar_hoja("Tareas_Hitos", df_tareas)
                        st.session_state["flash_tarea"] = [("success", f"Registro {id_t_sel} eliminado correctamente.")]
                        st.rerun()

    # ---- Listado ----
    st.write("")
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Tareas e hitos registrados</div>
        <div class="count-pill">{len(df_tareas_ctrl)} en total</div>
    </div>
    """, unsafe_allow_html=True)

    if df_tareas_ctrl.empty:
        st.info("Todavía no hay tareas o hitos registrados.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            filtro_proy_t = st.selectbox("Proyecto", ["Todos"] + sorted(df_tareas_ctrl["Nombre_Proyecto"].astype(str).unique()), key="filtro_t_proy")
        with f2:
            filtro_sit_t = st.selectbox(
                "Situación",
                ["Todas", "🔴 Vencidas", "🟠 Vencen pronto", "🔵 En plazo", "🟢 Completadas a tiempo", "🟡 Completadas con retraso"],
                key="filtro_t_sit",
            )
        with f3:
            filtro_resp_t = st.selectbox("Responsable", ["Todos"] + sorted(df_tareas_ctrl["Nombre_Responsable"].astype(str).unique()), key="filtro_t_resp")

        mapa_filtro_sit = {
            "🔴 Vencidas": "vencida", "🟠 Vencen pronto": "pronto", "🔵 En plazo": "plazo",
            "🟢 Completadas a tiempo": "ok", "🟡 Completadas con retraso": "tarde",
        }
        df_filtrado_t = df_tareas_ctrl
        if filtro_proy_t != "Todos":
            df_filtrado_t = df_filtrado_t[df_filtrado_t["Nombre_Proyecto"].astype(str) == filtro_proy_t]
        if filtro_sit_t != "Todas":
            df_filtrado_t = df_filtrado_t[df_filtrado_t["_clave"] == mapa_filtro_sit[filtro_sit_t]]
        if filtro_resp_t != "Todos":
            df_filtrado_t = df_filtrado_t[df_filtrado_t["Nombre_Responsable"].astype(str) == filtro_resp_t]

        st.dataframe(
            df_filtrado_t.sort_values(["Fecha_Planificada"]).drop(columns=["_clave"]),
            use_container_width=True,
            hide_index=True,
            column_config={
                "ID_Tarea": st.column_config.TextColumn("ID"),
                "Nombre_Proyecto": st.column_config.TextColumn("Proyecto"),
                "Nombre_Tarea": st.column_config.TextColumn("Tarea / Hito"),
                "Nombre_Responsable": st.column_config.TextColumn("Responsable"),
                "Fecha_Planificada": st.column_config.TextColumn("Planificada"),
                "Fecha_Real": st.column_config.TextColumn("Real"),
                "Situacion": st.column_config.TextColumn("Situación"),
                "ID_Proyecto": None,
                "ID_Responsable": None,
            },
        )
        st.caption(f"Mostrando {len(df_filtrado_t)} de {len(df_tareas_ctrl)} registro(s)")


# =========================================================
# MÓDULO 6: RIESGOS (MATRIZ Y CONTROL DE EXPOSICIÓN)
# =========================================================
elif modulo.startswith("6"):
    st.markdown("""
    <div class="banner">
        <h1>Módulo 6 · Riesgos</h1>
        <p>Registra los riesgos de cada proyecto, evalúa su exposición y vigila los críticos antes de que se materialicen.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="focus-callout">
        <strong>Cómo leer el control de riesgos:</strong> el nivel de cada riesgo es
        probabilidad × impacto (1–25): 🟢 <strong>Bajo</strong> (&lt;8) · 🟡 <strong>Medio</strong> (8–14) ·
        🔴 <strong>Alto</strong> (15). Un riesgo está <strong>activo</strong> mientras siga
        «Abierto» o «En mitigación». Los riesgos altos activos y los que se materializan
        quedan registrados en la hoja de alertas.
    </div>
    """, unsafe_allow_html=True)

    if "flash_riesgo" in st.session_state:
        for tipo, texto in st.session_state.pop("flash_riesgo"):
            getattr(st, tipo)(texto)

    # ---- Preparación de datos comunes del módulo ----
    ESTADOS_ACTIVOS_RIESGO = ("Abierto", "En mitigación")
    df_activos_r = df_riesgos[df_riesgos["Estado"].isin(ESTADOS_ACTIVOS_RIESGO)]
    df_altos_activos = df_activos_r[df_activos_r["Nivel"] == "Alto"]
    df_materializados = df_riesgos[df_riesgos["Estado"] == "Materializado"]

    # ---- Resumen rápido ----
    s1, s2, s3, s4 = st.columns(4)
    for col, valor, etiqueta in zip(
        (s1, s2, s3, s4),
        (len(df_riesgos), len(df_activos_r), len(df_altos_activos), len(df_materializados)),
        ("Riesgos", "Activos", "Altos activos", "Materializados"),
    ):
        with col:
            st.markdown(
                f'<div class="mini-stat"><div class="n">{valor}</div><div class="l">{etiqueta}</div></div>',
                unsafe_allow_html=True,
            )

    # ---- Panorama de riesgos críticos (todos los proyectos) ----
    st.write("")
    if not df_riesgos.empty:
        if not df_altos_activos.empty:
            lineas = [
                f"🔴 **{r.ID_Proyecto}** · {r.Descripcion_Riesgo} (P×I = {r.Nivel_Riesgo}) · {r.Nombre_Responsable} · {str(r.Estado).lower()}"
                for r in df_altos_activos.sort_values("Nivel_Riesgo", ascending=False).itertuples()
            ]
            st.error(
                "**Riesgos de nivel Alto sin resolver**\n\n" + "\n\n".join(lineas)
                + "\n\nPrioriza sus planes de mitigación y revísalos en cada comité de proyecto."
            )
        if not df_materializados.empty:
            lineas = [
                f"**{r.ID_Proyecto}** · {r.Descripcion_Riesgo} — plan en curso: {r.Plan_Mitigacion}"
                for r in df_materializados.sort_values("Nivel_Riesgo", ascending=False).itertuples()
            ]
            st.warning("**Riesgos materializados (el problema ya ocurrió)**\n\n" + "\n\n".join(lineas))
        if df_altos_activos.empty and df_materializados.empty:
            st.success("No hay riesgos de nivel Alto activos ni riesgos materializados.")

    # ---- Matriz probabilidad × impacto ----
    st.write("")
    st.markdown("""
    <div class="table-card-header">
        <div class="title">Matriz de riesgos activos (probabilidad × impacto)</div>
    </div>
    """, unsafe_allow_html=True)

    opciones_proy_r = [f"{r.ID_Proyecto} · {r.Nombre_Proyecto}" for r in df_proyectos.itertuples()]
    mapa_proy_r = dict(zip(opciones_proy_r, df_proyectos["ID_Proyecto"].astype(str)))
    mapa_nombre_proy_r = dict(zip(df_proyectos["ID_Proyecto"].astype(str), df_proyectos["Nombre_Proyecto"]))

    OPCION_TODOS = "Todos los proyectos"
    matriz_sel = st.selectbox("Proyecto a revisar", [OPCION_TODOS] + opciones_proy_r, key="sel_matriz_riesgos")
    if matriz_sel == OPCION_TODOS:
        df_matriz = df_activos_r
    else:
        df_matriz = df_activos_r[df_activos_r["ID_Proyecto"].astype(str) == mapa_proy_r[matriz_sel]]

    if df_matriz.empty:
        st.info("No hay riesgos activos (abiertos o en mitigación) para esta selección.")
    else:
        conteo_matriz = df_matriz.groupby(["Probabilidad", "Impacto"]).size()
        etiquetas_imp = ["", "Muy bajo", "Bajo", "Medio", "Alto", "Muy alto"]
        etiquetas_prob = ["Muy baja", "Baja", "Media", "Alta", "Muy alta"]
        filas_html = []
        for imp_ in range(5, 0, -1):
            celdas = []
            for prob_ in range(1, 6):
                nivel_celda = prob_ * imp_
                if nivel_celda >= 15:
                    bg, fg = "var(--rojo-bg)", "var(--rojo)"
                elif nivel_celda >= 8:
                    bg, fg = "var(--ambar-bg)", "var(--ambar)"
                else:
                    bg, fg = "var(--verde-bg)", "var(--verde)"
                n = int(conteo_matriz.get((prob_, imp_), 0))
                contenido = str(n) if n else "–"
                peso = "800" if n else "400"
                celdas.append(
                    f'<td style="background:{bg};color:{fg};text-align:center;padding:12px 0;'
                    f'border-radius:8px;font-weight:{peso};font-size:0.95rem;">{contenido}</td>'
                )
            filas_html.append(
                f'<tr><td style="text-align:right;padding-right:10px;color:var(--texto-suave);'
                f'font-size:0.78rem;font-weight:600;white-space:nowrap;">{imp_} · {etiquetas_imp[imp_]}</td>'
                + "".join(celdas) + "</tr>"
            )
        encabezados_html = "".join(
            f'<td style="text-align:center;color:var(--texto-suave);font-size:0.78rem;font-weight:600;padding:6px 0;">{p_} · {et}</td>'
            for p_, et in zip(range(1, 6), etiquetas_prob)
        )
        st.markdown(
            f'<table style="width:100%;border-collapse:separate;border-spacing:5px;table-layout:fixed;">'
            f'<tr><td style="width:110px;"></td>{encabezados_html}</tr>{"".join(filas_html)}</table>'
            f'<div style="display:flex;justify-content:space-between;color:var(--texto-suave);font-size:0.75rem;margin-top:2px;">'
            f'<span>↑ Impacto</span><span>Probabilidad →</span></div>',
            unsafe_allow_html=True,
        )
        exposicion = int(df_matriz["Nivel_Riesgo"].sum())
        st.caption(
            f"Cada celda cuenta los riesgos activos con esa combinación. "
            f"Exposición activa de la selección: {len(df_matriz)} riesgo(s) · {exposicion} puntos de riesgo (suma de P×I)."
        )

        # ---- Inventario activo priorizado de la selección ----
        st.write("")
        st.markdown("""
        <div class="table-card-header">
            <div class="title">Riesgos activos priorizados</div>
        </div>
        """, unsafe_allow_html=True)
        columnas_inv = ["Nombre_Proyecto", "Descripcion_Riesgo", "Probabilidad", "Impacto", "Nivel_Riesgo", "Nivel", "Plan_Mitigacion", "Nombre_Responsable", "Estado"]
        if matriz_sel != OPCION_TODOS:
            columnas_inv = columnas_inv[1:]
        st.dataframe(
            df_matriz.sort_values("Nivel_Riesgo", ascending=False)[columnas_inv],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Nombre_Proyecto": st.column_config.TextColumn("Proyecto"),
                "Descripcion_Riesgo": st.column_config.TextColumn("Riesgo"),
                "Probabilidad": st.column_config.NumberColumn("P", format="%d"),
                "Impacto": st.column_config.NumberColumn("I", format="%d"),
                "Nivel_Riesgo": st.column_config.NumberColumn("P×I", format="%d"),
                "Plan_Mitigacion": st.column_config.TextColumn("Plan de mitigación"),
                "Nombre_Responsable": st.column_config.TextColumn("Responsable"),
            },
        )
        st.caption("Ordenados de mayor a menor nivel: los primeros son los que exigen atención inmediata.")

    # ---- Captura ----
    st.write("")
    df_col_activos_r = df_colaboradores[df_colaboradores["Estado"] == "Activo"]
    opciones_resp_r = [f"{r.ID_Colaborador} · {r.Nombre}" for r in df_col_activos_r.itertuples()]
    mapa_resp_r = dict(zip(opciones_resp_r, df_col_activos_r["ID_Colaborador"].astype(str)))
    opciones_resp_r_todos = [f"{r.ID_Colaborador} · {r.Nombre}" for r in df_colaboradores.itertuples()]
    mapa_resp_r_todos = dict(zip(opciones_resp_r_todos, df_colaboradores["ID_Colaborador"].astype(str)))
    mapa_nombre_resp_r = dict(zip(df_colaboradores["ID_Colaborador"].astype(str), df_colaboradores["Nombre"]))

    OPCIONES_PROBABILIDAD = ["1 · Muy baja", "2 · Baja", "3 · Media", "4 · Alta", "5 · Muy alta"]
    OPCIONES_IMPACTO = ["1 · Muy bajo", "2 · Bajo", "3 · Medio", "4 · Alto", "5 · Muy alto"]

    with st.form("form_riesgo", clear_on_submit=True):
        st.markdown("""
        <div class="form-section-header">
            <div class="icon">01</div>
            <div>
                <div class="title">Identificación del riesgo</div>
                <div class="subtitle">Qué podría salir mal y en qué proyecto</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        r_proyecto = st.selectbox("Proyecto *", opciones_proy_r) if opciones_proy_r else None
        r_descripcion = st.text_input("Descripción del riesgo *", placeholder="Retraso del proveedor en la entrega, rotación de personal clave...")

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">02</div>
            <div>
                <div class="title">Evaluación</div>
                <div class="subtitle">Qué tan probable es y qué tanto afectaría al proyecto</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            r_probabilidad = st.select_slider(
                "Probabilidad", options=OPCIONES_PROBABILIDAD, value=OPCIONES_PROBABILIDAD[2],
                help="¿Qué tan probable es que ocurra? El nivel del riesgo se calcula automáticamente (probabilidad × impacto).",
            )
        with col2:
            r_impacto = st.select_slider(
                "Impacto", options=OPCIONES_IMPACTO, value=OPCIONES_IMPACTO[2],
                help="Si ocurre, ¿qué tanto afectaría al proyecto?",
            )

        st.markdown("""
        <div class="form-section-header">
            <div class="icon">03</div>
            <div>
                <div class="title">Respuesta al riesgo</div>
                <div class="subtitle">Qué se hará para evitarlo o reducir su efecto, y quién responde</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        r_plan = st.text_input("Plan de mitigación *", placeholder="Plan de contingencia, proveedor alterno, capacitación cruzada...")
        col3, col4 = st.columns(2)
        with col3:
            r_responsable = st.selectbox("Responsable *", opciones_resp_r) if opciones_resp_r else None
        with col4:
            r_estado = st.selectbox("Estado *", ESTADOS_RIESGO)

        st.write("")
        submitted_riesgo = st.form_submit_button("Guardar riesgo", use_container_width=True)

        if submitted_riesgo:
            errores_r = []
            if not r_proyecto:
                errores_r.append("No hay proyectos registrados. Crea al menos uno en el Módulo 1.")
            if not r_responsable:
                errores_r.append("No hay colaboradores activos. Registra al menos uno en el Módulo 2.")
            if not r_descripcion.strip():
                errores_r.append("La descripción del riesgo es obligatoria.")
            if not r_plan.strip():
                errores_r.append("El plan de mitigación es obligatorio: todo riesgo registrado debe tener una respuesta.")

            if errores_r:
                for e in errores_r:
                    st.error(e)
            else:
                prob_val = int(r_probabilidad[0])
                imp_val = int(r_impacto[0])
                nivel_num = prob_val * imp_val
                nivel_txt = clasificar_nivel_riesgo(nivel_num)
                id_proy_r = mapa_proy_r[r_proyecto]
                id_resp_r = mapa_resp_r[r_responsable]
                id_riesgo_nuevo = siguiente_id(df_riesgos, "ID_Riesgo", "RSG")
                nuevo_riesgo = {
                    "ID_Riesgo": id_riesgo_nuevo,
                    "ID_Proyecto": id_proy_r,
                    "Nombre_Proyecto": mapa_nombre_proy_r.get(id_proy_r, id_proy_r),
                    "Descripcion_Riesgo": r_descripcion.strip(),
                    "Probabilidad": prob_val,
                    "Impacto": imp_val,
                    "Nivel_Riesgo": nivel_num,
                    "Nivel": nivel_txt,
                    "Plan_Mitigacion": r_plan.strip(),
                    "ID_Responsable": id_resp_r,
                    "Nombre_Responsable": mapa_nombre_resp_r.get(id_resp_r, id_resp_r),
                    "Estado": r_estado,
                }
                df_riesgos_act = pd.concat([df_riesgos, pd.DataFrame([nuevo_riesgo])], ignore_index=True)
                guardar_hoja("Riesgos", df_riesgos_act)

                mensajes_r = [("success", f"Riesgo guardado con nivel {nivel_txt} ({nivel_num}/25).")]
                if nivel_txt == "Alto" and r_estado in ESTADOS_ACTIVOS_RIESGO:
                    mensajes_r.append(("warning", "Este riesgo quedó en nivel ALTO y sigue activo. Dale prioridad al plan de mitigación. Se dejó registro en la hoja de alertas."))
                    registrar_alerta("Riesgos", "Crítico" if nivel_num >= 20 else "Precaución",
                                     id_riesgo_nuevo, mapa_nombre_proy_r.get(id_proy_r, id_proy_r), id_proy_r, "",
                                     nivel_num, 15,
                                     f"Riesgo de nivel Alto (P×I={nivel_num}) en {id_proy_r}: {r_descripcion.strip()[:120]}")
                if r_estado == "Materializado":
                    mensajes_r.append(("error", "El riesgo se registró como materializado: el problema ya ocurrió. Ejecuta el plan de mitigación y evalúa el impacto en el proyecto. Se dejó registro en la hoja de alertas."))
                    registrar_alerta("Riesgos", "Crítico",
                                     id_riesgo_nuevo, mapa_nombre_proy_r.get(id_proy_r, id_proy_r), id_proy_r, "",
                                     nivel_num, 15,
                                     f"Se materializó el riesgo '{r_descripcion.strip()[:100]}' en {id_proy_r}; ejecutar el plan de mitigación.")
                st.session_state["flash_riesgo"] = mensajes_r
                st.rerun()

    # ---- Editar o eliminar ----
    st.write("")
    with st.expander("Editar o eliminar un riesgo existente", expanded=False):
        if df_riesgos.empty:
            st.caption("Todavía no hay riesgos para editar.")
        else:
            etiquetas_r = [
                f"{r.ID_Riesgo} · {r.ID_Proyecto} · {r.Descripcion_Riesgo[:60]}"
                for r in df_riesgos.itertuples()
            ]
            mapa_id_r = dict(zip(etiquetas_r, df_riesgos["ID_Riesgo"].astype(str)))
            etiqueta_r_sel = st.selectbox("Selecciona un riesgo", etiquetas_r, key="sel_editar_riesgo")
            id_r_sel = mapa_id_r[etiqueta_r_sel]
            fila_r = df_riesgos[df_riesgos["ID_Riesgo"].astype(str) == id_r_sel].iloc[0]

            etiqueta_proy_r_act = next(
                (e for e, i in mapa_proy_r.items() if i == str(fila_r["ID_Proyecto"])),
                opciones_proy_r[0] if opciones_proy_r else None,
            )
            etiqueta_resp_r_act = next(
                (e for e, i in mapa_resp_r_todos.items() if i == str(fila_r["ID_Responsable"])),
                opciones_resp_r_todos[0] if opciones_resp_r_todos else None,
            )
            prob_act = int(fila_r["Probabilidad"] or 3)
            imp_act = int(fila_r["Impacto"] or 3)

            with st.form(f"form_editar_riesgo_{id_r_sel}"):
                er_proyecto = st.selectbox(
                    "Proyecto *", opciones_proy_r,
                    index=opciones_proy_r.index(etiqueta_proy_r_act) if etiqueta_proy_r_act in opciones_proy_r else 0,
                    key=f"er_proy_{id_r_sel}",
                )
                er_descripcion = st.text_input(
                    "Descripción del riesgo *", value=str(fila_r["Descripcion_Riesgo"]), key=f"er_desc_{id_r_sel}",
                )
                col1, col2 = st.columns(2)
                with col1:
                    er_probabilidad = st.select_slider(
                        "Probabilidad", options=OPCIONES_PROBABILIDAD,
                        value=OPCIONES_PROBABILIDAD[max(1, min(5, prob_act)) - 1], key=f"er_prob_{id_r_sel}",
                    )
                with col2:
                    er_impacto = st.select_slider(
                        "Impacto", options=OPCIONES_IMPACTO,
                        value=OPCIONES_IMPACTO[max(1, min(5, imp_act)) - 1], key=f"er_imp_{id_r_sel}",
                    )
                er_plan = st.text_input(
                    "Plan de mitigación *", value=str(fila_r["Plan_Mitigacion"]), key=f"er_plan_{id_r_sel}",
                )
                col3, col4 = st.columns(2)
                with col3:
                    er_responsable = st.selectbox(
                        "Responsable *", opciones_resp_r_todos,
                        index=opciones_resp_r_todos.index(etiqueta_resp_r_act) if etiqueta_resp_r_act in opciones_resp_r_todos else 0,
                        key=f"er_resp_{id_r_sel}",
                    )
                with col4:
                    er_estado = st.selectbox(
                        "Estado *", ESTADOS_RIESGO,
                        index=ESTADOS_RIESGO.index(fila_r["Estado"]) if fila_r["Estado"] in ESTADOS_RIESGO else 0,
                        key=f"er_estado_{id_r_sel}",
                    )

                er_confirmar = st.checkbox(
                    "Confirmo que deseo eliminar este riesgo de forma permanente",
                    key=f"er_confirmar_{id_r_sel}",
                )
                st.write("")
                b1, b2 = st.columns(2)
                with b1:
                    actualizar_r = st.form_submit_button("Guardar cambios", use_container_width=True)
                with b2:
                    eliminar_r = st.form_submit_button("Eliminar riesgo", use_container_width=True)

                if actualizar_r:
                    if not er_descripcion.strip():
                        st.error("La descripción del riesgo es obligatoria.")
                    elif not er_plan.strip():
                        st.error("El plan de mitigación es obligatorio.")
                    else:
                        prob_e = int(er_probabilidad[0])
                        imp_e = int(er_impacto[0])
                        nivel_e = prob_e * imp_e
                        nivel_e_txt = clasificar_nivel_riesgo(nivel_e)
                        estado_anterior = str(fila_r["Estado"])
                        nivel_anterior_txt = clasificar_nivel_riesgo(int(fila_r["Nivel_Riesgo"] or 1))
                        id_proy_er = mapa_proy_r[er_proyecto]
                        id_resp_er = mapa_resp_r_todos[er_responsable]
                        idx_r = fila_r.name
                        df_riesgos.loc[idx_r, "ID_Proyecto"] = id_proy_er
                        df_riesgos.loc[idx_r, "Nombre_Proyecto"] = mapa_nombre_proy_r.get(id_proy_er, id_proy_er)
                        df_riesgos.loc[idx_r, "Descripcion_Riesgo"] = er_descripcion.strip()
                        df_riesgos.loc[idx_r, "Probabilidad"] = prob_e
                        df_riesgos.loc[idx_r, "Impacto"] = imp_e
                        df_riesgos.loc[idx_r, "Nivel_Riesgo"] = nivel_e
                        df_riesgos.loc[idx_r, "Nivel"] = nivel_e_txt
                        df_riesgos.loc[idx_r, "Plan_Mitigacion"] = er_plan.strip()
                        df_riesgos.loc[idx_r, "ID_Responsable"] = id_resp_er
                        df_riesgos.loc[idx_r, "Nombre_Responsable"] = mapa_nombre_resp_r.get(id_resp_er, id_resp_er)
                        df_riesgos.loc[idx_r, "Estado"] = er_estado
                        guardar_hoja("Riesgos", df_riesgos)

                        mensajes_er = [("success", f"Riesgo {id_r_sel} actualizado (nivel {nivel_e_txt}).")]
                        alto_activo_antes = nivel_anterior_txt == "Alto" and estado_anterior in ESTADOS_ACTIVOS_RIESGO
                        alto_activo_ahora = nivel_e_txt == "Alto" and er_estado in ESTADOS_ACTIVOS_RIESGO
                        if alto_activo_ahora and not alto_activo_antes:
                            mensajes_er.append(("warning", "El riesgo quedó en nivel ALTO y sigue activo. Dale prioridad al plan de mitigación. Se dejó registro en la hoja de alertas."))
                            registrar_alerta("Riesgos", "Crítico" if nivel_e >= 20 else "Precaución",
                                             id_r_sel, mapa_nombre_proy_r.get(id_proy_er, id_proy_er), id_proy_er, "",
                                             nivel_e, 15,
                                             f"Riesgo de nivel Alto (P×I={nivel_e}) en {id_proy_er}: {er_descripcion.strip()[:120]}")
                        elif alto_activo_ahora:
                            mensajes_er.append(("warning", "El riesgo sigue en nivel ALTO y activo. Dale prioridad al plan de mitigación."))
                        if er_estado == "Materializado" and estado_anterior != "Materializado":
                            mensajes_er.append(("error", "El riesgo pasó a materializado: el problema ya ocurrió. Ejecuta el plan de mitigación y evalúa el impacto. Se dejó registro en la hoja de alertas."))
                            registrar_alerta("Riesgos", "Crítico",
                                             id_r_sel, mapa_nombre_proy_r.get(id_proy_er, id_proy_er), id_proy_er, "",
                                             nivel_e, 15,
                                             f"Se materializó el riesgo '{er_descripcion.strip()[:100]}' en {id_proy_er}; ejecutar el plan de mitigación.")
                        st.session_state["flash_riesgo"] = mensajes_er
                        st.rerun()

                elif eliminar_r:
                    if not er_confirmar:
                        st.error("Marca la casilla de confirmación para poder eliminar el riesgo.")
                    else:
                        df_riesgos = df_riesgos.drop(index=fila_r.name)
                        guardar_hoja("Riesgos", df_riesgos)
                        st.session_state["flash_riesgo"] = [("success", f"Riesgo {id_r_sel} eliminado correctamente.")]
                        st.rerun()

    # ---- Listado ----
    st.write("")
    st.markdown(f"""
    <div class="table-card-header">
        <div class="title">Riesgos registrados</div>
        <div class="count-pill">{len(df_riesgos)} en total</div>
    </div>
    """, unsafe_allow_html=True)

    if df_riesgos.empty:
        st.info("Todavía no hay riesgos registrados.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            filtro_proy_r = st.selectbox("Proyecto", ["Todos"] + sorted(df_riesgos["Nombre_Proyecto"].astype(str).unique()), key="filtro_r_proy")
        with f2:
            filtro_nivel_r = st.selectbox("Nivel", ["Todos", "Alto", "Medio", "Bajo"], key="filtro_r_nivel")
        with f3:
            filtro_estado_r = st.selectbox("Estado", ["Todos"] + ESTADOS_RIESGO, key="filtro_r_estado")

        df_filtrado_r = df_riesgos
        if filtro_proy_r != "Todos":
            df_filtrado_r = df_filtrado_r[df_filtrado_r["Nombre_Proyecto"].astype(str) == filtro_proy_r]
        if filtro_nivel_r != "Todos":
            df_filtrado_r = df_filtrado_r[df_filtrado_r["Nivel"] == filtro_nivel_r]
        if filtro_estado_r != "Todos":
            df_filtrado_r = df_filtrado_r[df_filtrado_r["Estado"] == filtro_estado_r]

        st.dataframe(
            df_filtrado_r.sort_values("Nivel_Riesgo", ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "ID_Riesgo": st.column_config.TextColumn("ID"),
                "Nombre_Proyecto": st.column_config.TextColumn("Proyecto"),
                "Descripcion_Riesgo": st.column_config.TextColumn("Riesgo"),
                "Probabilidad": st.column_config.ProgressColumn("Probabilidad", format="%d/5", min_value=0, max_value=5),
                "Impacto": st.column_config.ProgressColumn("Impacto", format="%d/5", min_value=0, max_value=5),
                "Nivel_Riesgo": st.column_config.NumberColumn("P×I", format="%d"),
                "Plan_Mitigacion": st.column_config.TextColumn("Plan de mitigación"),
                "Nombre_Responsable": st.column_config.TextColumn("Responsable"),
            },
        )
        st.caption(f"Mostrando {len(df_filtrado_r)} de {len(df_riesgos)} riesgo(s)")


# =========================================================
# ASISTENTE DEL PORTAFOLIO · MOTOR DE CONSULTAS
# =========================================================
NOMBRE_ASISTENTE = "Asistente del portafolio"
EJEMPLOS_ASISTENTE = [
    "¿Qué tareas tenemos pendientes en el proyecto de Power BI?",
    "¿Cuál es el presupuesto ejecutado del CRM y en qué rubro nos estamos pasando?",
    "¿Quién está sobrecargado esta semana?",
    "¿Qué riesgos altos hay abiertos en el portafolio?",
    "¿Qué proyectos van con retraso y de quién dependen?",
]


try:
    # También a nivel de módulo: la ventana necesita sus clases de error.
    import anthropic
except ModuleNotFoundError:
    anthropic = None


def clave_asistente():
    """Lee la clave de la API de los secrets de Streamlit o del entorno."""
    try:
        if "ANTHROPIC_API_KEY" in st.secrets:
            return str(st.secrets["ANTHROPIC_API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY", "").strip()


CLAVE_ASISTENTE = clave_asistente()


def crear_motor_asistente():
    """Prepara el asistente y devuelve la función que responde preguntas.

    Todo vive dentro de esta función a propósito: el script de Streamlit es un
    único espacio de nombres, y nombres genéricos como `texto` o `datos` se
    pisarían con los de los módulos. Devuelve None si no hay clave configurada.
    """
    # El asistente NO recibe la base de datos completa ni inventa cifras: recibe
    # un catálogo de consultas (herramientas) que leen las mismas hojas que
    # alimentan el tablero de Power BI. El modelo interpreta la pregunta, elige
    # la consulta y redacta la respuesta con los datos que esa consulta devuelve.
    MODELO_ASISTENTE = "claude-opus-4-8"
    # Profundidad de razonamiento: "low" responde más rápido, "high" razona más
    # en preguntas comparativas. "medium" es el equilibrio para consultas de datos.
    ESFUERZO_ASISTENTE = "medium"
    MAX_VUELTAS = 6          # tope de consultas encadenadas por pregunta
    MAX_FILAS_RESPUESTA = 60  # tope de filas que se entregan al asistente

    if not CLAVE_ASISTENTE:
        return None
    try:
        import anthropic
    except ModuleNotFoundError:
        return None

    hoy_asis = date.today()

    # -----------------------------------------------------
    # Utilidades de formato y búsqueda
    # -----------------------------------------------------
    def num(valor):
        """Convierte a número tratando las celdas vacías del Excel como 0."""
        v = pd.to_numeric(valor, errors="coerce")
        return 0.0 if pd.isna(v) else float(v)

    def texto(valor):
        """Devuelve el texto de una celda, o un guion si viene vacía."""
        if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
            return "—"
        limpio = str(valor).strip()
        return limpio if limpio and limpio.lower() not in ("nan", "nat", "none") else "—"

    def eur(valor):
        """Formatea un importe al estilo español: 25.000 €."""
        return f"{num(valor):,.0f} €".replace(",", ".")

    def fecha_corta(valor):
        texto = str(valor)[:10]
        return texto if texto and texto.lower() not in ("nan", "nat", "none") else "—"

    def resolver_proyectos(texto):
        """Traduce lo que escribió el usuario ('el ERP') a filas de Proyectos.

        Busca primero por ID exacto y luego por coincidencia parcial del nombre,
        para que el asistente no tenga que conocer los códigos PRJ-XXX.
        """
        if df_proyectos.empty:
            return df_proyectos
        clave = str(texto or "").strip().lower()
        if not clave:
            return df_proyectos
        por_id = df_proyectos[df_proyectos["ID_Proyecto"].astype(str).str.lower() == clave]
        if not por_id.empty:
            return por_id
        return df_proyectos[
            df_proyectos["Nombre_Proyecto"].astype(str).str.lower().str.contains(clave, regex=False)
        ]

    def situacion_cronograma(fecha_plan, fecha_real, estado):
        """Misma clasificación de situación que usa el Módulo 5 (control de cronograma)."""
        def a_fecha(v):
            f = pd.to_datetime(str(v)[:10], errors="coerce")
            return None if pd.isna(f) else f.date()

        plan, real = a_fecha(fecha_plan), a_fecha(fecha_real)
        if estado == "Completada" or real is not None:
            if plan and real and real > plan:
                return f"Completada con retraso (+{(real - plan).days} d)", "tarde"
            return "Completada a tiempo", "ok"
        if plan is not None:
            if plan < hoy_asis:
                return f"Vencida hace {(hoy_asis - plan).days} d", "vencida"
            dias = (plan - hoy_asis).days
            if dias == 0:
                return "Vence hoy", "pronto"
            if dias <= 7:
                return f"Vence en {dias} d", "pronto"
        return "En plazo", "plazo"

    def recortar(lineas, total):
        """Devuelve el bloque de texto avisando si se recortaron filas."""
        cuerpo = "\n".join(lineas)
        if total > len(lineas):
            cuerpo += f"\n… ({total - len(lineas)} fila(s) más no mostradas)"
        return cuerpo

    # -----------------------------------------------------
    # Consultas disponibles para el asistente
    # -----------------------------------------------------
    def h_listar_proyectos():
        if df_proyectos.empty:
            return "No hay proyectos registrados."
        lineas = []
        for r in df_proyectos.itertuples():
            lineas.append(
                f"- {r.ID_Proyecto} · {r.Nombre_Proyecto} | estado: {r.Estado_Proyecto} | "
                f"avance: {num(r.Avance_Pct):.0f}% | prioridad: {texto(r.Prioridad)} | "
                f"líder: {texto(r.Lider_Proyecto)} | área: {texto(r.Cliente_Area)} | "
                f"fin planificado: {fecha_corta(r.Fecha_Fin_Planificada)}"
            )
        return f"{len(df_proyectos)} proyecto(s) en el portafolio:\n" + "\n".join(lineas)

    def h_detalle_proyecto(proyecto):
        coincidencias = resolver_proyectos(proyecto)
        if coincidencias.empty:
            return f"No encontré ningún proyecto que coincida con «{proyecto}»."
        if len(coincidencias) > 1:
            nombres = ", ".join(
                f"{r.ID_Proyecto} ({r.Nombre_Proyecto})" for r in coincidencias.itertuples()
            )
            return f"«{proyecto}» coincide con varios proyectos: {nombres}. Pide al usuario que concrete."

        p = coincidencias.iloc[0]
        pid = str(p["ID_Proyecto"])
        planificado = num(p["Presupuesto_Planificado"])
        ejecutado = num(p["Presupuesto_Ejecutado"])
        pct = (ejecutado / planificado * 100) if planificado else 0

        tareas_p = df_tareas[df_tareas["ID_Proyecto"].astype(str) == pid]
        vencidas = sum(
            1 for _, t in tareas_p.iterrows()
            if situacion_cronograma(t["Fecha_Planificada"], t["Fecha_Real"], t["Estado"])[1] == "vencida"
        )
        pendientes = int((tareas_p["Estado"] != "Completada").sum())
        riesgos_p = df_riesgos[df_riesgos["ID_Proyecto"].astype(str) == pid]
        riesgos_activos = riesgos_p[riesgos_p["Estado"].isin(["Abierto", "En mitigación"])]

        return (
            f"{pid} · {p['Nombre_Proyecto']}\n"
            f"- Descripción: {texto(p['Descripcion'])}\n"
            f"- Estado: {p['Estado_Proyecto']} | avance: {num(p['Avance_Pct']):.0f}% | "
            f"prioridad: {texto(p['Prioridad'])}\n"
            f"- Líder: {texto(p['Lider_Proyecto'])} | área cliente: {texto(p['Cliente_Area'])} | "
            f"tecnología: {texto(p['Tecnologia'])}\n"
            f"- Fechas: inicio {fecha_corta(p['Fecha_Inicio'])}, fin planificado "
            f"{fecha_corta(p['Fecha_Fin_Planificada'])}, fin real {fecha_corta(p['Fecha_Fin_Real'])}\n"
            f"- Presupuesto: planificado {eur(planificado)}, ejecutado {eur(ejecutado)} ({pct:.1f}%)\n"
            f"- Beneficio esperado: {eur(p['Beneficio_Esperado'])}\n"
            f"- Cronograma: {len(tareas_p)} tareas/hitos, {pendientes} sin completar, {vencidas} vencidas\n"
            f"- Riesgos: {len(riesgos_p)} registrados, {len(riesgos_activos)} activos"
        )

    def h_tareas(proyecto=None, situacion=None, responsable=None):
        datos = df_tareas.copy()
        if datos.empty:
            return "No hay tareas ni hitos registrados."
        etiqueta = "todo el portafolio"
        if proyecto:
            coincidencias = resolver_proyectos(proyecto)
            if coincidencias.empty:
                return f"No encontré ningún proyecto que coincida con «{proyecto}»."
            ids = set(coincidencias["ID_Proyecto"].astype(str))
            datos = datos[datos["ID_Proyecto"].astype(str).isin(ids)]
            etiqueta = ", ".join(coincidencias["Nombre_Proyecto"].astype(str))
        if responsable:
            datos = datos[
                datos["Nombre_Responsable"].astype(str).str.lower()
                .str.contains(str(responsable).lower(), regex=False)
            ]

        filas = []
        for _, t in datos.iterrows():
            sit, clave = situacion_cronograma(t["Fecha_Planificada"], t["Fecha_Real"], t["Estado"])
            filas.append((t, sit, clave))

        if situacion == "pendientes":
            filas = [f for f in filas if f[2] in ("vencida", "pronto", "plazo")]
        elif situacion in ("vencida", "pronto", "plazo", "tarde", "ok"):
            filas = [f for f in filas if f[2] == situacion]

        if not filas:
            return f"No hay tareas que cumplan ese criterio en {etiqueta}."

        orden = {"vencida": 0, "pronto": 1, "plazo": 2, "tarde": 3, "ok": 4}
        filas.sort(key=lambda f: (orden.get(f[2], 9), str(f[0]["Fecha_Planificada"])[:10]))
        lineas = [
            f"- [{t['ID_Proyecto']}] {t['Tipo']}: {t['Nombre_Tarea']} | responsable: "
            f"{texto(t['Nombre_Responsable'])} | fecha planificada: "
            f"{fecha_corta(t['Fecha_Planificada'])} | estado: {t['Estado']} | situación: {sit}"
            for t, sit, _ in filas[:MAX_FILAS_RESPUESTA]
        ]
        return f"{len(filas)} tarea(s)/hito(s) en {etiqueta}:\n" + recortar(lineas, len(filas))

    def h_presupuesto(proyecto):
        coincidencias = resolver_proyectos(proyecto)
        if coincidencias.empty:
            return f"No encontré ningún proyecto que coincida con «{proyecto}»."
        if len(coincidencias) > 1:
            nombres = ", ".join(
                f"{r.ID_Proyecto} ({r.Nombre_Proyecto})" for r in coincidencias.itertuples()
            )
            return f"«{proyecto}» coincide con varios proyectos: {nombres}. Pide al usuario que concrete."

        p = coincidencias.iloc[0]
        pid = str(p["ID_Proyecto"])
        mov = df_presupuesto[df_presupuesto["ID_Proyecto"].astype(str) == pid]
        if mov.empty:
            return (
                f"{pid} · {p['Nombre_Proyecto']} no tiene movimientos de presupuesto detallados. "
                f"En la ficha del proyecto figura: planificado {eur(p['Presupuesto_Planificado'])}, "
                f"ejecutado {eur(p['Presupuesto_Ejecutado'])}."
            )

        plan = mov[mov["Tipo"] == "Planificado"].groupby("Categoria")["Monto"].sum()
        real = mov[mov["Tipo"] == "Real"].groupby("Categoria")["Monto"].sum()
        lineas = []
        for cat in sorted(set(plan.index) | set(real.index)):
            pl, re_ = float(plan.get(cat, 0)), float(real.get(cat, 0))
            if pl == 0:
                estado = "gasto sin partida planificada"
            else:
                consumo = re_ / pl * 100
                estado = (
                    f"{consumo:.1f}% consumido"
                    + (" — EXCEDIDO" if consumo > 100 else " — al límite" if consumo >= 90 else "")
                )
            lineas.append(f"- {cat}: planificado {eur(pl)} | real {eur(re_)} | {estado}")

        total_plan, total_real = float(plan.sum()), float(real.sum())
        pct_total = (total_real / total_plan * 100) if total_plan else 0
        return (
            f"Presupuesto de {pid} · {p['Nombre_Proyecto']} ({len(mov)} movimientos)\n"
            f"TOTAL: planificado {eur(total_plan)} | ejecutado {eur(total_real)} | "
            f"{pct_total:.1f}% consumido | disponible {eur(total_plan - total_real)}\n"
            f"Desglose por rubro:\n" + "\n".join(lineas)
        )

    def h_carga_equipo(semana=None, colaborador=None):
        if df_salud.empty:
            return "No hay registros de horas del equipo."
        datos = df_salud.copy()
        if semana:
            datos = datos[datos["Semana"].astype(str).str[:10] == str(semana)[:10]]
            etiqueta = f"la semana del {str(semana)[:10]}"
        else:
            ultima = sorted(datos["Semana"].astype(str).unique())[-1]
            datos = datos[datos["Semana"].astype(str) == ultima]
            etiqueta = f"la última semana registrada ({ultima})"
        if colaborador:
            datos = datos[
                datos["Nombre_Colaborador"].astype(str).str.lower()
                .str.contains(str(colaborador).lower(), regex=False)
            ]
        if datos.empty:
            return f"No hay registros de horas para ese criterio en {etiqueta}."

        agregado = datos.groupby(["ID_Colaborador", "Nombre_Colaborador"]).agg(
            horas=("Horas_Reales", "sum"),
            capacidad=("Capacidad_Horas_Semana", "max"),
            percibida=("Carga_Percibida", "mean"),
            proyectos=("ID_Proyecto", "nunique"),
        ).reset_index()

        lineas = []
        for r in agregado.sort_values("horas", ascending=False).itertuples():
            cap = num(r.capacidad) or 40
            pct = num(r.horas) / cap * 100
            lineas.append(
                f"- {r.Nombre_Colaborador} ({r.ID_Colaborador}): {num(r.horas):.1f} h de {cap:.0f} h "
                f"= {pct:.0f}% de saturación — {clasificar_saturacion(pct)[0]} | "
                f"carga percibida media: {r.percibida:.1f}/5 | proyectos: {r.proyectos}"
            )
        return (
            f"Carga del equipo en {etiqueta} ({len(agregado)} persona(s)).\n"
            "Umbrales del semáforo: <60% capacidad disponible · 60-80% saludable · "
            "80-90% seguimiento · 90-95% precaución · ≥95% crítico.\n" + "\n".join(lineas)
        )

    def h_riesgos(proyecto=None, solo_activos=True, nivel=None):
        datos = df_riesgos.copy()
        if datos.empty:
            return "No hay riesgos registrados."
        etiqueta = "todo el portafolio"
        if proyecto:
            coincidencias = resolver_proyectos(proyecto)
            if coincidencias.empty:
                return f"No encontré ningún proyecto que coincida con «{proyecto}»."
            ids = set(coincidencias["ID_Proyecto"].astype(str))
            datos = datos[datos["ID_Proyecto"].astype(str).isin(ids)]
            etiqueta = ", ".join(coincidencias["Nombre_Proyecto"].astype(str))
        if solo_activos:
            datos = datos[datos["Estado"].isin(["Abierto", "En mitigación"])]
        if nivel:
            datos = datos[datos["Nivel"].astype(str).str.lower() == str(nivel).lower()]
        if datos.empty:
            return f"No hay riesgos que cumplan ese criterio en {etiqueta}."

        datos = datos.sort_values("Nivel_Riesgo", ascending=False)
        lineas = [
            f"- [{r.ID_Proyecto}] {r.Descripcion_Riesgo} | nivel: {r.Nivel} "
            f"(probabilidad {r.Probabilidad}/5 × impacto {r.Impacto}/5 = {r.Nivel_Riesgo}) | "
            f"estado: {r.Estado} | responsable: {texto(r.Nombre_Responsable)} | "
            f"mitigación: {texto(r.Plan_Mitigacion)}"
            for r in datos.head(MAX_FILAS_RESPUESTA).itertuples()
        ]
        exposicion = int(datos["Nivel_Riesgo"].sum())
        return (
            f"{len(datos)} riesgo(s) en {etiqueta} · exposición total (suma P×I): {exposicion}\n"
            + recortar(lineas, len(datos))
        )

    def h_alertas(limite=10, origen=None, nivel=None):
        datos = leer_hoja("Alertas", COLS_ALERTAS)
        if datos.empty:
            return "No hay alertas registradas."
        if origen:
            datos = datos[datos["Origen"].astype(str).str.lower().str.contains(str(origen).lower(), regex=False)]
        if nivel:
            datos = datos[datos["Nivel"].astype(str).str.lower() == str(nivel).lower()]
        if datos.empty:
            return "No hay alertas que cumplan ese criterio."
        datos = datos.sort_values("Fecha_Hora", ascending=False).head(int(limite or 10))
        lineas = [
            f"- {r.Fecha_Hora} | {r.Origen} | {r.Nivel} | {texto(r.Mensaje)}"
            for r in datos.itertuples()
        ]
        return f"Últimas {len(datos)} alerta(s):\n" + "\n".join(lineas)

    CONSULTAS = {
        "listar_proyectos": h_listar_proyectos,
        "detalle_proyecto": h_detalle_proyecto,
        "tareas": h_tareas,
        "presupuesto": h_presupuesto,
        "carga_equipo": h_carga_equipo,
        "riesgos": h_riesgos,
        "alertas": h_alertas,
    }

    HERRAMIENTAS = [
        {
            "name": "listar_proyectos",
            "description": (
                "Devuelve el listado completo de proyectos del portafolio con su estado, "
                "avance, prioridad, líder y fecha de fin planificada. Úsala cuando la pregunta "
                "sea sobre el portafolio en conjunto o para localizar un proyecto por su nombre."
            ),
            "input_schema": {"type": "object", "properties": {}},
        },
        {
            "name": "detalle_proyecto",
            "description": (
                "Ficha completa de un proyecto: fechas, avance, líder, presupuesto planificado y "
                "ejecutado, número de tareas pendientes y vencidas, y riesgos activos. Úsala cuando "
                "pregunten por el estado o la situación general de un proyecto concreto."
            ),
            "input_schema": {
                "type": "object",
                "properties": {
                    "proyecto": {
                        "type": "string",
                        "description": "ID (PRJ-001) o parte del nombre del proyecto.",
                    }
                },
                "required": ["proyecto"],
            },
        },
        {
            "name": "tareas",
            "description": (
                "Lista tareas e hitos con su responsable, fecha planificada, estado y situación "
                "frente a hoy (vencida, por vencer, en plazo, completada). Úsala para cualquier "
                "pregunta sobre qué está pendiente, qué se ha retrasado o qué vence pronto."
            ),
            "input_schema": {
                "type": "object",
                "properties": {
                    "proyecto": {
                        "type": "string",
                        "description": "ID o parte del nombre del proyecto. Omítelo para consultar todo el portafolio.",
                    },
                    "situacion": {
                        "type": "string",
                        "enum": ["pendientes", "vencida", "pronto", "plazo", "tarde", "ok"],
                        "description": (
                            "Filtro de situación: 'pendientes' = todo lo no completado; 'vencida' = pasó "
                            "su fecha sin completarse; 'pronto' = vence en 7 días o menos; 'plazo' = en "
                            "plazo; 'tarde' = completada con retraso; 'ok' = completada a tiempo."
                        ),
                    },
                    "responsable": {
                        "type": "string",
                        "description": "Nombre o parte del nombre de la persona responsable.",
                    },
                },
            },
        },
        {
            "name": "presupuesto",
            "description": (
                "Detalle presupuestario de un proyecto: total planificado frente a ejecutado, "
                "porcentaje consumido, importe disponible y desglose por rubro señalando los "
                "rubros excedidos. Úsala para preguntas de dinero, gasto o desviación."
            ),
            "input_schema": {
                "type": "object",
                "properties": {
                    "proyecto": {
                        "type": "string",
                        "description": "ID (PRJ-001) o parte del nombre del proyecto.",
                    }
                },
                "required": ["proyecto"],
            },
        },
        {
            "name": "carga_equipo",
            "description": (
                "Saturación del equipo en una semana: horas reales frente a capacidad, porcentaje "
                "de saturación con su nivel de semáforo, carga percibida media y número de proyectos "
                "por persona. Úsala para preguntas sobre sobrecarga, disponibilidad o horas."
            ),
            "input_schema": {
                "type": "object",
                "properties": {
                    "semana": {
                        "type": "string",
                        "description": "Semana en formato AAAA-MM-DD (lunes). Omítela para la última semana registrada.",
                    },
                    "colaborador": {
                        "type": "string",
                        "description": "Nombre o parte del nombre de una persona concreta.",
                    },
                },
            },
        },
        {
            "name": "riesgos",
            "description": (
                "Riesgos registrados con su probabilidad, impacto, nivel, estado, responsable y plan "
                "de mitigación, más la exposición total (suma de P×I). Úsala para preguntas sobre "
                "riesgos, amenazas o exposición."
            ),
            "input_schema": {
                "type": "object",
                "properties": {
                    "proyecto": {
                        "type": "string",
                        "description": "ID o parte del nombre del proyecto. Omítelo para todo el portafolio.",
                    },
                    "solo_activos": {
                        "type": "boolean",
                        "description": "Si es true (por defecto) devuelve solo los abiertos y en mitigación.",
                    },
                    "nivel": {
                        "type": "string",
                        "enum": ["Alto", "Medio", "Bajo"],
                        "description": "Filtra por nivel de riesgo.",
                    },
                },
            },
        },
        {
            "name": "alertas",
            "description": (
                "Historial de alertas automáticas del sistema (saturación del equipo, desviaciones "
                "de presupuesto, cronograma y riesgos altos). Úsala cuando pregunten qué avisos hay "
                "o qué requiere atención."
            ),
            "input_schema": {
                "type": "object",
                "properties": {
                    "limite": {"type": "integer", "description": "Número de alertas a devolver (10 por defecto)."},
                    "origen": {
                        "type": "string",
                        "description": "Filtra por origen: Monitor de Salud, Presupuesto, Cronograma o Riesgos.",
                    },
                    "nivel": {
                        "type": "string",
                        "enum": ["Crítico", "Precaución", "Informativa"],
                        "description": "Filtra por nivel de la alerta.",
                    },
                },
            },
        },
    ]

    _catalogo = "\n".join(
        f"- {r.ID_Proyecto}: {r.Nombre_Proyecto} ({r.Estado_Proyecto})"
        for r in df_proyectos.itertuples()
    ) or "(sin proyectos registrados)"

    INSTRUCCIONES = f"""Eres el asistente de una plataforma de gestión de proyectos y salud del equipo.
Respondes al project manager que la utiliza. Hoy es {hoy_asis.strftime('%d/%m/%Y')}.

Cómo trabajas:
- Todas las cifras que des tienen que salir de las consultas disponibles. Nunca inventes ni
  estimes datos, y nunca respondas de memoria: si necesitas un dato, consúltalo.
- Si una consulta no devuelve la información pedida, dilo con claridad en vez de rellenar el hueco.
- Puedes encadenar varias consultas cuando la pregunta lo requiera (por ejemplo, comparar dos proyectos).
- Si la pregunta es ambigua o el nombre del proyecto coincide con varios, pide una aclaración breve.

Cómo respondes:
- En español, directo y breve: primero la respuesta, después el detalle que la sustenta.
- Usa listas cuando enumeres tareas, riesgos o personas, y da los importes en euros.
- Menciona el nombre del proyecto, no solo su código.
- Nada de jerga técnica de la aplicación: no hables de hojas de Excel, columnas ni herramientas.

Proyectos del portafolio:
{_catalogo}

Umbrales que maneja la plataforma:
- Saturación del equipo: por debajo del 60% hay capacidad disponible, 60-80% es saludable,
  80-90% requiere seguimiento, 90-95% es precaución y a partir del 95% es crítico.
- Presupuesto por rubro: por debajo del 90% consumido va bien, del 90 al 100% está al límite
  y por encima del 100% está excedido.
"""

    if "asis_mensajes" not in st.session_state:
        st.session_state["asis_mensajes"] = []

    cliente_asis = anthropic.Anthropic(api_key=CLAVE_ASISTENTE)

    def preguntar_al_asistente(pregunta):
        """Ciclo del agente: consulta -> ejecuta -> vuelve a consultar hasta responder.

        Devuelve (texto_respuesta, consultas_realizadas). Las consultas se ejecutan
        aquí, en la aplicación: el modelo solo pide cuáles quiere y lee el resultado.
        """
        mensajes = st.session_state["asis_mensajes"]
        mensajes.append({"role": "user", "content": pregunta})
        consultas_usadas = []

        for _ in range(MAX_VUELTAS):
            respuesta = cliente_asis.messages.create(
                model=MODELO_ASISTENTE,
                max_tokens=4000,
                system=INSTRUCCIONES,
                thinking={"type": "adaptive"},
                output_config={"effort": ESFUERZO_ASISTENTE},
                tools=HERRAMIENTAS,
                messages=mensajes,
            )
            mensajes.append({"role": "assistant", "content": respuesta.content})

            if respuesta.stop_reason == "refusal":
                return ("No puedo responder a esa petición.", consultas_usadas)

            if respuesta.stop_reason != "tool_use":
                texto = "\n\n".join(b.text for b in respuesta.content if b.type == "text").strip()
                if respuesta.stop_reason == "max_tokens":
                    texto += "\n\n_(respuesta cortada por longitud)_"
                return (texto or "No he podido redactar una respuesta.", consultas_usadas)

            resultados = []
            for bloque in respuesta.content:
                if bloque.type != "tool_use":
                    continue
                consultas_usadas.append((bloque.name, dict(bloque.input)))
                try:
                    salida = CONSULTAS[bloque.name](**bloque.input)
                    fallo = False
                except Exception as exc:  # el error vuelve al modelo para que reaccione
                    salida = f"La consulta falló: {exc}"
                    fallo = True
                resultados.append({
                    "type": "tool_result",
                    "tool_use_id": bloque.id,
                    "content": salida,
                    "is_error": fallo,
                })
            mensajes.append({"role": "user", "content": resultados})

        return ("La consulta se ha alargado demasiado. Prueba a preguntarlo de forma más concreta.",
                consultas_usadas)

    return preguntar_al_asistente


MOTOR_ASISTENTE = crear_motor_asistente()


# =========================================================
# ASISTENTE DEL PORTAFOLIO · VENTANA FLOTANTE
# =========================================================
# Streamlit no trae widgets flotantes: el botón se fija con CSS sobre un
# contenedor con key propia, y la conversación vive en un st.dialog porque un
# desplegable normal se cerraría en cada recarga (y cada pregunta recarga).
if es_pm:
    from streamlit.components.v1 import html as componente_html

    st.markdown("""
    <style>
    /* Se eleva sobre la barra que Streamlit Cloud pinta en esa misma esquina
       ("Manage app" para la dueña, la insignia "Hosted with Streamlit" para
       quien visita): la dibuja el contenedor de la nube, fuera de la app, así
       que la única forma de no quedar tapado es dejarle ese hueco libre. */
    .st-key-boton_asistente {
        position: fixed; right: 26px; bottom: 78px; z-index: 999; width: auto;
    }
    @media (max-width: 640px) {
        .st-key-boton_asistente { right: 12px; bottom: 70px; }
    }
    .st-key-boton_asistente button {
        background: linear-gradient(135deg, var(--verde-vivo), var(--celeste));
        color: #0E1840 !important; font-weight: 700; border: none;
        border-radius: 26px; padding: 11px 22px;
        box-shadow: 0 10px 26px var(--sombra);
    }
    .st-key-boton_asistente button:hover { filter: brightness(1.07); }
    .asis-pista { color: var(--texto-suave); font-size: 0.9rem; margin-bottom: 4px; }
    </style>
    """, unsafe_allow_html=True)

    def limpiar_conversacion():
        st.session_state["asis_mensajes"] = []
        st.session_state.pop("asis_consultas", None)

    def respuestas_redactadas(mensajes):
        """Cuenta las respuestas que llegaron a texto, que son las que se ven.

        Sirve para emparejar cada respuesta del hilo con las consultas que la
        sustentan, sin contar los turnos internos de llamada a las consultas.
        """
        return sum(
            1 for mensaje in mensajes
            if mensaje["role"] == "assistant" and any(
                getattr(b, "type", "") == "text" and getattr(b, "text", "").strip()
                for b in mensaje["content"]
            )
        )

    @st.dialog(NOMBRE_ASISTENTE, width="large")
    def ventana_asistente():
        if MOTOR_ASISTENTE is None:
            st.warning("El asistente todavía no está configurado.")
            st.markdown(
                "Hace falta una clave de acceso al modelo de lenguaje:\n\n"
                "1. Crea una clave en **console.anthropic.com → API keys**.\n"
                "2. En local, guárdala en `.streamlit/secrets.toml` como "
                "`ANTHROPIC_API_KEY = \"sk-ant-...\"`.\n"
                "3. En la versión publicada, pégala en **Settings → Secrets**.\n\n"
                "La clave nunca se guarda en el código ni viaja al repositorio."
            )
            return

        st.caption(
            "Consulto la misma base de datos que alimenta el tablero directivo, "
            "así que las cifras son las reales del portafolio."
        )

        # El hueco se reserva ANTES del campo de escritura para que el hilo
        # quede arriba y el campo abajo, como en cualquier chat. Además permite
        # decidir si se pintan los ejemplos ya sabiendo si hubo pregunta.
        # La altura fija hace que el hilo se desplace DENTRO de la caja en vez
        # de estirar la ventana, y autoscroll la baja sola a la última
        # respuesta para no tener que buscarla a mano.
        zona_chat = st.container(height=430, autoscroll=True, key="hilo_asistente")

        pregunta = st.chat_input("Escribe tu pregunta sobre el portafolio")

        aviso_error = None
        if pregunta:
            marca = len(st.session_state["asis_mensajes"])
            with st.spinner("Consultando los datos del portafolio…"):
                try:
                    _, consultas = MOTOR_ASISTENTE(pregunta)
                except anthropic.AuthenticationError:
                    aviso_error = "La clave de acceso no es válida. Revísala en los secrets de la aplicación."
                except anthropic.RateLimitError:
                    aviso_error = "Demasiadas consultas seguidas. Espera unos segundos y vuelve a preguntar."
                except anthropic.APIConnectionError:
                    aviso_error = "No hay conexión con el servicio del asistente. Revisa tu red e inténtalo de nuevo."
                except anthropic.APIStatusError as exc:
                    aviso_error = f"El servicio del asistente devolvió un error ({exc.status_code}). Inténtalo de nuevo."
            if aviso_error:
                # Se descarta la pregunta fallida para no dejar la conversación
                # a medias (rompería la siguiente llamada).
                del st.session_state["asis_mensajes"][marca:]
            else:
                st.session_state.setdefault("asis_consultas", {})[
                    respuestas_redactadas(st.session_state["asis_mensajes"]) - 1
                ] = consultas

        # Dentro de la caja: las sugerencias solo mientras no hay conversación
        # (al preguntar desaparecen), y después el hilo. Del historial se
        # muestran solo preguntas y respuestas redactadas; los resultados de las
        # consultas quedan detrás, en el registro de trazabilidad.
        with zona_chat:
            if not st.session_state.get("asis_mensajes"):
                st.markdown('<div class="asis-pista">Puedes preguntarme cosas como:</div>',
                            unsafe_allow_html=True)
                for ejemplo in EJEMPLOS_ASISTENTE:
                    st.markdown(f"- {ejemplo}")

            consultas_por_turno = st.session_state.get("asis_consultas", {})
            indice_respuesta = 0
            for mensaje in st.session_state.get("asis_mensajes", []):
                contenido = mensaje["content"]
                if mensaje["role"] == "user":
                    if isinstance(contenido, str):
                        with st.chat_message("user"):
                            st.markdown(contenido)
                    continue
                redactado = "\n\n".join(
                    b.text for b in contenido
                    if getattr(b, "type", "") == "text" and getattr(b, "text", "").strip()
                )
                if not redactado:
                    continue
                with st.chat_message("assistant"):
                    st.markdown(redactado)
                    registro = consultas_por_turno.get(indice_respuesta)
                    if registro:
                        with st.expander("Consultas utilizadas para esta respuesta"):
                            for nombre, argumentos in registro:
                                detalle = ", ".join(f"{k}={v}" for k, v in argumentos.items()) or "sin filtros"
                                st.markdown(f"- `{nombre}` ({detalle})")
                indice_respuesta += 1
            if aviso_error:
                st.error(aviso_error)

        # El autoscroll propio del contenedor no llega a dispararse aquí (el
        # hilo se repinta entero en cada recarga, no se añade poco a poco), así
        # que se baja la caja a mano. Se reintenta porque Streamlit todavía
        # puede estar pintando cuando llega la orden.
        if st.session_state.get("asis_mensajes"):
            # El número de turno tiene que aparecer en el guion: si el contenido
            # no cambia, Streamlit reutiliza el componente y la orden no se
            # vuelve a ejecutar (se quedaría bajando solo en la primera pregunta).
            componente_html(
                f"""
                <script>
                const turno = {len(st.session_state["asis_mensajes"])};
                const bajar = () => {{
                    const caja = window.parent.document.querySelector('.st-key-hilo_asistente');
                    if (caja) caja.scrollTop = caja.scrollHeight;
                }};
                bajar();
                [80, 250, 600].forEach(ms => setTimeout(bajar, ms));
                </script>
                """,
                height=0,
            )

        if st.session_state.get("asis_mensajes"):
            st.button("Nueva conversación", on_click=limpiar_conversacion,
                      key="asis_limpiar", use_container_width=True)

    with st.container(key="boton_asistente"):
        # El diálogo se abre llamando a la función en el flujo del script;
        # hacerlo desde un on_click no lo abre.
        if st.button("Asistente del portafolio", key="asis_abrir",
                     help="Pregunta en lenguaje natural sobre el portafolio"):
            ventana_asistente()
