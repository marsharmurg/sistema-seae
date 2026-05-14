SEAE - Sistema de Evaluación de Alternativas Económicas
Descripción
SEAE es una aplicación de escritorio desarrollada en Python que permite evaluar alternativas económicas mediante los métodos de Valor Presente Neto (VPN), Costo Anual Equivalente (CAE) y Tasa Interna de Retorno (TIR). La herramienta está orientada a estudiantes y profesionales de ingeniería económica, ofreciendo una interfaz moderna, persistencia de datos, generación de reportes y visualizaciones gráficas coherentes con el método de evaluación seleccionado.

Características principales
Calculadora financiera universal: resuelve problemas de valor del dinero en el tiempo, anualidades, amortizaciones y conversión de tasas de interés.

Gestión de proyectos: permite crear, duplicar, eliminar y editar múltiples alternativas de inversión, cada una con su propio método (VPN, CAE o TIR).

Cálculos económicos precisos: implementa VPN, CAE, TIR mediante el método de Newton-Raphson y análisis de payback.

Clasificación automática de riesgo (Bajo, Medio, Alto) basada en volatilidad de flujos, margen TIR-TMAR y período de recuperación.

Advertencias inteligentes: detecta flujos insuficientes, tasas extremas o flujos negativos.

Dashboard dinámico: filtro por método, tarjetas con indicadores agregados (aceptados, rechazados, promedios), gráfico de barras de VPN por proyecto y gráfico de pastel de decisiones.

Gráficos especializados: evolución del VPN acumulado, gráfico radar con indicadores normalizados y análisis de sensibilidad del VPN frente a cambios en la tasa de descuento.

Comparación de alternativas: tabla comparativa global y ventana de detalle para dos proyectos.

Generación de reportes: exportación a PDF (con gráficos adaptados al método) y a Excel.

Base de datos SQLite: historial completo de proyectos con búsqueda, eliminación y exportación/importación en formatos JSON y CSV.

Ayuda contextual y glosario de términos financieros.

Interfaz moderna construida con CustomTkinter, modo oscuro y tooltips informativos.

Tecnologías utilizadas
Python 3

CustomTkinter - interfaz gráfica moderna basada en Tkinter

Matplotlib - visualización de datos

NumPy - cálculos numéricos

SQLite3 - persistencia local

FPDF - generación de documentos PDF

OpenPyXL - exportación a Excel

mplcursors (opcional) - tooltips interactivos en gráficos

Instalación
Clonar el repositorio:

bash
git clone https://github.com/tuusuario/seae.git
cd seae
Instalar las dependencias:

bash
pip install -r requirements.txt
Contenido sugerido de requirements.txt:

text
customtkinter
matplotlib
numpy
fpdf
openpyxl
mplcursors
Ejecutar la aplicación:

bash
python seae.py
La base de datos seae_final.db se crea automáticamente en el directorio de trabajo.

Uso
Al iniciar, la aplicación muestra una ventana con pestañas que organizan todas las funcionalidades.

Pestaña Calculadora
Subpestañas: Valor del Dinero, Anualidades, Amortización, Conversión de Tasas y VPN / CAE / TIR.

Permite calcular variables desconocidas dejando el campo correspondiente vacío.

Cada subpestaña permite generar un reporte PDF del cálculo realizado.

Pestaña Proyectos
Aquí se crean, editan y evalúan las alternativas económicas.

Cada proyecto tiene campos para nombre, método (VPN, CAE o TIR), inversión inicial, vida útil, TMAR, valor residual y flujos de caja anuales.

Los flujos se ingresan en entradas individuales con autoguardado diferido.

Botones para calcular, duplicar, limpiar, eliminar, copiar análisis al portapapeles, generar PDF, Excel, abrir auditor de cálculos y visualizar boxplot de flujos.

El resultado muestra el indicador principal resaltado según el método elegido y los secundarios en formato reducido.

Pestaña Dashboard
Filtro por método: Todos, VPN, CAE, TIR.

Tarjetas de resumen: total de proyectos, aceptados, rechazados, riesgo promedio, VPN promedio, TIR promedio, CAE promedio.

Gráfico de barras del VPN de cada proyecto.

Gráfico de pastel con la proporción de aceptación / rechazo.

Tabla con los 5 proyectos más relevantes ordenados por VPN descendente.

Pestaña Gráficos
Evolución del VPN acumulado para todos los proyectos (etiquetas con método).

Gráfico radar comparativo de indicadores normalizados (VPN, TIR, CAE, Payback).

Pestaña Comparación
Tabla completa de todos los proyectos con su decisión de aceptación/rechazo.

Botón para abrir una ventana de comparación detallada entre dos proyectos.

Generación de reporte comparativo global en PDF.

Pestaña Sensibilidad
Permite analizar cómo varía el VPN de un escenario de inversión al modificar la tasa de descuento.

Se define inversión, años, tasa base y flujos. Un slider controla el rango de variación.

El gráfico muestra la curva VPN vs Tasa con zonas de aceptación y rechazo coloreadas.

Pestaña Historial
Tabla con todos los proyectos guardados en la base de datos.

Botones para eliminar registros, exportar a CSV, comparar dos proyectos históricos, importar/exportar en JSON.

Pestaña Glosario
Definiciones estáticas de VPN, TIR, CAE, TMAR, Payback y criterios de riesgo.

Pestaña Ayuda
Preguntas frecuentes con respuestas breves sobre el uso del sistema.

Estructura del proyecto
text
seae/
├── seae.py               # Código fuente principal de la aplicación
├── requirements.txt      # Dependencias necesarias
├── README.md             # Este archivo
└── seae_final.db         # Base de datos SQLite (se genera automáticamente)
