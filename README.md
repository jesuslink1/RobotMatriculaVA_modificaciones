# 🤖 RobotMatriculaVA_modificaciones

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python)
![Automation](https://img.shields.io/badge/Automation-Selenium-green)
![Data](https://img.shields.io/badge/Data-pandas-yellow)
![Status](https://img.shields.io/badge/Status-Completed-success)

---

## 📌 Descripción

Sistema de **automatización de procesos académicos** enfocado en matrícula, validación de cursos y generación de reportes.

Este proyecto reemplaza tareas manuales mediante scripts en Python, permitiendo procesar grandes volúmenes de datos académicos de forma rápida, consistente y sin errores.

---

## 🎯 Funcionalidades principales

- Validación de mallas académicas  
- Automatización de pre-matrícula  
- Procesos de matrícula masiva  
- Gestión de cambios de plan  
- Manejo de ciclos lectivos  
- Generación automática de reportes  

---

## ⚙️ Stack tecnológico

- Python  
- pandas  
- Selenium  
- Excel / CSV  

---

## 🧠 ¿Qué problema resuelve?

En entornos académicos, muchos procesos se realizan manualmente:

- Validar alumnos uno por uno  
- Revisar cumplimiento de mallas  
- Generar reportes manuales  
- Ejecutar matrícula masiva  

Este sistema automatiza todo ese flujo, reduciendo tiempos y errores humanos.

---

## 📂 Estructura del proyecto

### Carpetas

- Antiguos → Scripts legacy o versiones anteriores  
- Convalidaciones → Validación de cursos entre planes  
- Desmatricula → Procesos de retiro de matrícula  
- En Construccion → Scripts en desarrollo  
- Extensión → Funcionalidades adicionales  
- Lista Alumnos Matrícula Masiva → Procesamiento masivo  
- Reportes → Generación de resultados  

---

## 📜 Scripts clave

### Core
- main_completo.py → Orquestador principal  
- Main_agregacupo.py → Asignación de cupos  

### Validaciones
- Verificador_Malla_Notas.py → Validación académica  
- Verificador_Pre_Matricula.py → Evaluación de aptitud  

### Ciclos
- AgregarCicloLectivo.py → Asignación de ciclos  
- EliminarCicloLectivo.py → Limpieza de registros  

### Gestión académica
- Cambio_Plan.py → Cambio de plan de estudios  

---

## 📊 Flujo del sistema

Excel (input) → Validaciones → Procesos de matrícula → Generación de reportes → Excel / CSV (output)

---

## 📥 Entradas

- PreMatricula.xlsx  
- Cambio_plan.xlsx  
- CodigosVerificarMalla.xlsx  

---

## 📤 Salidas

- resultadoPreMatricula.xlsx  
- resultados_mallas.xlsx  
- reporte.csv  
- reporteCicloLectivo.csv  

---

## 🚀 Impacto

- Reducción del tiempo operativo  
- Eliminación de errores manuales  
- Mejora en toma de decisiones  
- Automatización de procesos repetitivos  

---

## 🧠 Aprendizajes

- Automatización con Python  
- Manejo de datos con pandas  
- Integración con Excel  
- Diseño de scripts reutilizables  
- Resolución de problemas reales  

---

## 👨‍💻 Autor

Jesús Ramos  
Python Developer | Automatización y Análisis de Datos
