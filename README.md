# Plataforma de Gestión de Proyectos

Aplicación web de captura de datos para la gestión de una cartera de proyectos:
proyectos, colaboradores, salud del equipo, presupuesto, cronograma y riesgos.
Los datos capturados alimentan un cuadro de mando en Power BI.

Desarrollada con [Streamlit](https://streamlit.io) como parte del Trabajo Fin de
Máster en Business Intelligence (VIU).

## Acceso

La aplicación está desplegada en Streamlit Community Cloud; basta con abrir el
enlace en el navegador (no requiere instalación). En la pantalla de acceso se
elige el rol:

- **Project Manager**: administra todos los módulos y supervisa la salud del equipo.
- **Colaborador**: auto-reporta sus horas y carga de trabajo percibida.

## Datos

`Base_Datos_TFM.xlsx` contiene datos de demostración ficticios (7 hojas:
Proyectos, Colaboradores, Salud_Equipo, Presupuesto, Tareas_Hitos, Riesgos y
Alertas). Si el archivo no existe, la aplicación lo regenera automáticamente
con los datos de muestra.

> Nota: en el despliegue en la nube el almacenamiento es efímero — los cambios
> capturados se conservan durante la sesión de demostración, pero pueden
> restablecerse a los datos de muestra cuando la plataforma reinicia la app.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```
