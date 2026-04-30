import matplotlib.pyplot as plt
import os
from flask import Flask, request
from flask import send_file


# Variables globales
resultado_a = {}
resultado_b = {}
mejor_proyecto = ""

app = Flask(__name__)

# Función para calcular VAN
def calcular_valor_presente(tasa, flujos):
    van = 0
    for t in range(len(flujos)):
        van += flujos[t] / (1 + tasa) ** t
    return van

#FuncionCAUE
def calcular_caue(vp, tasa, n):
    return vp * (tasa * (1 + tasa) ** n)/ ((1 + tasa) ** n -1)

# Funcion para calcular TIR
def calcular_tir(flujos, tol=1e-6, max_iter=100):
    low = -0.99
    high = 1.0

    for _ in range(max_iter):
        mid = (low + high) / 2
        vp = calcular_valor_presente(mid, flujos)

        if abs(vp) < tol:
            return mid
        
        if vp > 0:
            low = mid
        else :
            high = mid

    return mid  


# Página principal ( formulario
@app.route("/")
def inicio():
       
       # 🔥 borrar gráfico si existe
    if os.path.exists("grafico.png"):
        os.remove("grafico.png")

       
    return """
        <h1>Sistema de Evaluación Financiera SEAE</h1>
        <p>Ingrese los datos de los proyectos para realizar el análisis financiero.</p>

        <div style="display:flex; gap:50px; align-items:flex-start;">
        <div>
        <form action="/calcular" method="post">

        <h3>Datos Generales</h3>
        <label>Tasa de interés (%):</label><br>
        <input type="text" name="tasa"><br><br>

        <div style="display:flex; gap:50px;">

        <!--PROYECTO A-->
        <div>

        <h3>Proyecto A</h3>

        <label>Flujos A:</label><br>
        <input type="text" name="flujos_a"><br><br>

        <label>Vida del proyecto A:</label><br>
        <input type="text" name="n_a"><br><br>
        </div>

        <!--PROYECTO B -->
        <div>
        <h3>Proyecto B</h3>
        <label>Flujos B:</label><br>
        <input type="text" name="flujos_b"><br><br>

        <label>Vida del proyecto B:</label><br>
        <input type="text" name="n_b"><br><br>
        </div>

        </div>


        <br>
        <button type="submit">Evaluar Proyectos</button>
        </br>
        </form>

        </div>

        <!-- GRAFICO -->
       
        <div>
        <h2>Gráfico generado</h2>
        <img src="/grafico" width="500">
        </div>

       </div> 

        
       
        """
@app.route("/descargar")
def descargar():
    return send_file("reporte.txt", as_attachment=True)

@app.route("/grafico")
def grafico():
    return send_file("grafico.png", mimetype="image/png") 
 
#Procesar datos
@app.route("/calcular" , methods=["POST"])
def calcular():
    tasa = float(request.form["tasa"]) / 100
    
    #Proyecto A
    flujos_a = list(map(float, request.form["flujos_a"].split(",")))
    n_a = int(request.form["n_a"])
    
    #Proyecto B
    flujos_b = list(map(float, request.form["flujos_b"].split(",")))
    n_b = int(request.form["n_b"])

    #Proyecto A
    vp_a = calcular_valor_presente(tasa, flujos_a)
    caue_a = calcular_caue(vp_a, tasa, n_a)
    tir_a = calcular_tir(flujos_a)

    # Proyecto B
    vp_b = calcular_valor_presente(tasa, flujos_b)
    caue_b = calcular_caue(vp_b, tasa, n_b)
    tir_b = calcular_tir(flujos_b)

    #Decisión, evaluando mejor proyecto
    if vp_a > vp_b:
        mejor = "Proyecto A"

    elif vp_b > vp_a:
        mejor = "Proyecto B"

    else:
        mejor = "Ambos proyectos son equivalentes"

     # Descripción profesional
    descripcion = ""

    if vp_a > 0 and vp_b > 0:
        descripcion = "Ambos proyectos son financieramente viables, ya que generan valor positivo."
    elif vp_a < 0 and vp_b < 0:
        descripcion = "Ambos proyectos generan pérdidas, por lo que se recomienda analizar alternativas."
    else:
        descripcion = "Uno de los proyectos es viable y el otro no, por lo que se recomienda seleccionar el que genera valor positivo."

    descripcion += f" En comparación directa, {mejor} presenta mejores indicadores financieros."

    # Valores absolutos para gráfico
    vp_a_abs = abs(vp_a)
    vp_b_abs = abs(vp_b)

    # Crear gráfico con matplotlib
    proyectos = ['Proyecto A', 'Proyecto B']
    valores = [abs(vp_a), abs(vp_b)]

    plt.figure()
    plt.bar(proyectos, valores)
    plt.title("Comparación de Valor Presente (VAN)")
    plt.xlabel("Proyectos")
    plt.ylabel("Valor Presente")

    plt.tight_layout()
    plt.savefig("grafico.png")
    plt.close()  
    
    # Generar reporte TXT
    with open("reporte.txt", "w") as f:
         f.write("SISTEMA DE EVALUACIÓN DE ALTERNATIVAS ECONÓMICAS (SEAE)\n\n")

         f.write("=== PROYECTO A ===\n")
         f.write(f"Valor Presente: {round(vp_a,2)}\n")
         f.write(f"CAUE: {round(caue_a, 2)}\n")
         f.write(f"TIR: {round(tir_a * 100,2)}%\n\n")

         f.write("=== PROYECTO B ===\n")
         f.write(f"Valor Presente: {round(vp_b, 2)}\n")
         f.write(f"CAUE: {round(caue_b, 2)}\n\n")
         f.write(f"TIR: {round(tir_b * 100, 2)}%\n\n")

         f.write("=== DECISIÓN ===\n")
         f.write(f"La mejor opción es: {mejor}\n")


# Respuesta HTML
    return f"""
        <h2>Resultados</h2>

        <h3>Proyecto A</h3>
        <p>Valor Presente: {round(vp_a, 2)}</p>
        <p>CAUE: {round(caue_a, 2)}</p>
        <p>TIR: {round(tir_a * 100, 2)}%</p>

        <h3>Proyecto B</h3>
        <p>Valor Presente: {round(vp_b, 2)}</p>
        <p>CAUE: {round(caue_b, 2)}</p>
        <p>TIR: {round(tir_b * 100, 2)}%</p>

        <h2>Decisión</h2>
        <p>La mejor opción es: <strong>{mejor}</strong></p>

        <h2>Interpretación</h2>
        <p>{descripcion}</p>

        <h2>Comparación Gráfica (Valor Presente)</h2>
        <img src="/grafico" width="500">
        <a href="/descargar">
        <button>Descargar Reporte</button>
        </a>

        <br><br>
        <a href="/">Volver</a>
    """


if __name__== "__main__":

     app.run(debug=True)

