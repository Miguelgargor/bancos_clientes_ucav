### INSTALAR LAS LIBRERÍAS QUE NO ESTÉN INSTALADAS AÚN:
import pip

## Lista de librerías requeridas:
librerias_requeridas = ['flask', 'pandas', 'os', 'rich']

## Verificar si las librerías están instaladas:
for libreria in librerias_requeridas:
    try:
        # Intenta importar la librería:
        __import__(libreria)
    except ImportError:
        # Si la librería no está instalada, agrégala a la lista de instalaciones necesarias:
        pip.main(['install', libreria])
##==================================================================================================================================##


from flask import Flask, request, render_template
import pandas as pd
from FUNCION_RUBEN_BANCOS_CLIENTES import funcion_ruben
import os


app = Flask(__name__)
####################################################################################################################################
@app.route('/')
def index():


    return render_template('index.html')

####################################################################################################################################

@app.route('/procesados', methods=['POST'])   # SI SE CAMBIA LA RUTA TAMBIÉN CAMBIAR EN "index.html" !!
def procesar():
    archivo = request.files['archivo']

    BANCO = request.form['banco']
    Primer_num_documento = request.form['primer_num_documento']
    
    # LLAMAR A LA FUNCIÓN:
    if archivo:
        df_FINAL, mensaje_final, IMPORTE_TOTAL= funcion_ruben(archivo, BANCO, Primer_num_documento)

        # PRIMERAS y ÚLTIMAS FILAS:
        df_FINAL.index= df_FINAL.index+1
        df_primeros= df_FINAL.loc[:,['Fecha registro', 'Fecha de IVA', 'Tipo documento', 'Nº documento', 'Tipo mov.', 'Nº cuenta', 'Nombre de cuenta', 'Importe']].head()
        df_ultimos= df_FINAL.loc[:,['Fecha registro', 'Fecha de IVA', 'Tipo documento', 'Nº documento', 'Tipo mov.', 'Nº cuenta', 'Nombre de cuenta', 'Importe']].tail()

        df_primeras_y_ultimas= pd.concat([df_primeros, df_ultimos])
    
        return render_template('pagina_datos_procesados.html', mensaje=mensaje_final, IMPORTE_TOTAL=IMPORTE_TOTAL, df_primeras_y_ultimas=df_primeras_y_ultimas)
####################################################################################################################################

if __name__ == '__main__':
    app.run(debug=True, port=5001)
####################################################################################################################################