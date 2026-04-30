# 📊 Sistema de Evaluación de Alternativas Económicas (SEAE)

Aplicación web desarrollada en **Python con Flask** que permite evaluar y comparar proyectos de inversión utilizando criterios financieros como:

- Valor Presente Neto (VAN)
- Tasa Interna de Retorno (TIR)
- Costo Anual Uniforme Equivalente (CAUE)

---

## 🚀 Descripción

El sistema permite ingresar los flujos de caja de dos proyectos (A y B) y una tasa de interés, para determinar cuál es la mejor alternativa desde el punto de vista financiero.

Incluye:

- Cálculo automático de indicadores
- Comparación entre proyectos
- Generación de reporte en archivo `.txt`
- Visualización gráfica del VAN

---

## 🛠️ Tecnologías utilizadas

- Python 🐍
- Flask 🌐
- Matplotlib 📊
- HTML (renderizado desde Flask)

---

## 📈 Indicadores financieros

### 🔹 Valor Presente Neto (VAN)
Mide el valor actual de los flujos futuros de un proyecto.

- VAN > 0 → Proyecto rentable
- VAN < 0 → Proyecto no recomendable

---

### 🔹 Tasa Interna de Retorno (TIR)
Es la tasa que hace que el VAN sea igual a cero.

- Si TIR > tasa de interés → Aceptar proyecto
- Si TIR < tasa de interés → Rechazar proyecto

---

### 🔹 Costo Anual Uniforme Equivalente (CAUE)
Permite comparar proyectos con diferentes vidas útiles.

---

## ▶️ Cómo ejecutar el proyecto

### 1. Clonar repositorio

```bash
git clone https://github.com/marsharmurg/sistema-seae.git
cd sistema-seae

Crear entorno virtual
python -m venv venv

Activar:

venv\Scripts\activate   # Windows

Instalar dependencias
pip install flask matplotlib
4. Ejecutar aplicación
python app.py

Abrir en navegador
http://localhost:5000

📊 Funcionalidades

✔ Evaluación de dos proyectos
✔ Comparación automática
✔ Generación de reporte descargable
✔ Visualización gráfica

Autores

Grupo de Ingenieria Economica

💡 Notas

Este proyecto fue desarrollado con fines académicos para aplicar conceptos de ingeniería económica en un entorno práctico.