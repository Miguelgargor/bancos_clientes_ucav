import pandas as pd                         # Tratamiento de Datos.
import os                                   # Tratar Directorios (Carpetas).
from rich import print                      # Imprimir texto enriquecido (De colores y negrita).


def funcion_ruben(archivo, BANCO, Primer_num_documento):
    #### 0º) Verificar si la carpeta existe:
    if not os.path.exists('RESULTADOS_BC'):
        os.makedirs('RESULTADOS_BC') # Si no existe, crea la carpeta.
    #===============================================================================================================================================================#
    #===============================================================================================================================================================#

    #### 1º) DEFINIR EL BANCO (Ya sea Santander, BBVA, Caixa...):
    ### A) SANTANDER-> 5545:
    if BANCO=='SANTANDER 5545':

        ## 1- LECTURA DATOS:
        try:
            # Intenta leer el archivo con la extensión ".xls":
            df_BANCO= pd.read_excel(archivo, header=7)
        except FileNotFoundError:
            # Si el archivo ".xls" no se encuentra, intenta leerlo con la extensión ".xlsx":
            df_BANCO= pd.read_excel(archivo, header=7)
        #------------------------------------------------------------------------------------------#
        ## 2- Columna FECHA:
        FECHA= 'Fecha Operación'
        #------------------------------------------------------------------------------------------#
        ## 3- Nº Cuenta (Código del Banco):
        Cod_Banco= 'SANTANDER01'
        #------------------------------------------------------------------------------------------#
        ## 4- Nombre de Cuenta (Nombre del Banco):
        Nombre_Banco= 'SANTANDER 5545'
        #------------------------------------------------------------------------------------------#
        ## 5- NOMBRE PARA GUARDAR EL EXCEL:
        Nombre_GUARDAR= 'BC SANTANDER 5545'
    ####################################################################################################################################

    ### B) SANTANDER-> 0635:
    elif BANCO=='SANTANDER 0635':
        ## 1- LECTURA DATOS:
        try:
            # Intenta leer el archivo con la extensión ".xls":
            df_BANCO= pd.read_excel(archivo, header=7)
        except FileNotFoundError:
            # Si el archivo ".xls" no se encuentra, intenta leerlo con la extensión ".xlsx":
            df_BANCO= pd.read_excel(archivo, header=7)
        #------------------------------------------------------------------------------------------#
        ## 2- Columna FECHA:
        FECHA= 'Fecha Operación'
        #------------------------------------------------------------------------------------------#
        ## 3- Nº Cuenta (Código del Banco):
        Cod_Banco= 'SANTANDER05'
        #------------------------------------------------------------------------------------------#
        ## 4- Nombre de Cuenta (Nombre del Banco):
        Nombre_Banco= 'SANTANDER 0635'
        #------------------------------------------------------------------------------------------#
        ## 5- NOMBRE PARA GUARDAR EL EXCEL:
        Nombre_GUARDAR= 'BC SANTANDER 0635'
    ####################################################################################################################################

    ### C) CAIXABANK-> 9098:
    elif BANCO=='CAIXABANK 9098':
        ## 1- LECTURA DATOS:
        try:
            # Intenta leer el archivo con la extensión ".xls":
            df_BANCO= pd.read_excel(archivo, header=2)
        except FileNotFoundError:
            # Si el archivo ".xls" no se encuentra, intenta leerlo con la extensión ".xlsx":
            df_BANCO= pd.read_excel(archivo, header=2)
        #------------------------------------------------------------------------------------------#
        ## 2- Columna FECHA:
        FECHA= 'Fecha'
        df_BANCO[f'{FECHA}']= df_BANCO[f'{FECHA}'].dt.strftime('%d/%m/%Y')
        #------------------------------------------------------------------------------------------#
        ## 3- Nº Cuenta (Código del Banco):
        Cod_Banco= 'CAIXA01'
        #------------------------------------------------------------------------------------------#
        ## 4- Nombre de Cuenta (Nombre del Banco):
        Nombre_Banco= 'CAIXA 9098'
        #------------------------------------------------------------------------------------------#
        ## 5- NOMBRE PARA GUARDAR EL EXCEL:
        Nombre_GUARDAR= 'BC CAIXA 9098'
    ####################################################################################################################################

    ### D) CAIXABANK-> 0550:
    elif BANCO=='CAIXABANK 0550':
        ## 1- LECTURA DATOS:
        try:
            # Intenta leer el archivo con la extensión ".xls":
            df_BANCO= pd.read_excel(archivo, header=2)
        except FileNotFoundError:
            # Si el archivo ".xls" no se encuentra, intenta leerlo con la extensión ".xlsx":
            df_BANCO= pd.read_excel(archivo, header=2)
        #------------------------------------------------------------------------------------------#
        ## 2- Columna FECHA:
        FECHA= 'Fecha'
        df_BANCO[f'{FECHA}']= df_BANCO[f'{FECHA}'].dt.strftime('%d/%m/%Y')
        #------------------------------------------------------------------------------------------#
        ## 3- Nº Cuenta (Código del Banco):
        Cod_Banco= 'CAIXA02'
        #------------------------------------------------------------------------------------------#
        ## 4- Nombre de Cuenta (Nombre del Banco):
        Nombre_Banco= 'CAIXA 0550'
        #------------------------------------------------------------------------------------------#
        ## 5- NOMBRE PARA GUARDAR EL EXCEL:
        Nombre_GUARDAR= 'BC CAIXA 0550'
    ###################################################################################################################################

    ### E) BBVA:
    elif BANCO=='BBVA':
        ## 1- LECTURA DATOS:
        try:
            # Intenta leer el archivo con la extensión ".xls":
            df_BANCO= pd.read_excel(archivo, header=15)
        except FileNotFoundError:
            # Si el archivo ".xls" no se encuentra, intenta leerlo con la extensión ".xlsx":
            df_BANCO= pd.read_excel(archivo, header=15)

        df_BANCO= df_BANCO.iloc[:,1:] # Quitar la 1ª Columna que no tiene datos.
        #------------------------------------------------------------------------------------------#
        ## 2- Columna FECHA:
        FECHA= 'Fecha Contable'
        #------------------------------------------------------------------------------------------#
        ## 3- Nº Cuenta (Código del Banco):
        Cod_Banco= 'BBVA'
        #------------------------------------------------------------------------------------------#
        ## 4- Nombre de Cuenta (Nombre del Banco):
        Nombre_Banco= 'BBVA'
        #------------------------------------------------------------------------------------------#
        ## 5- NOMBRE PARA GUARDAR EL EXCEL:
        Nombre_GUARDAR= 'BC BBVA'
    ###################################################################################################################################

    ### F) CAJAMAR:
    elif BANCO=='CAJAMAR':
        ## 1- LECTURA DATOS:
        try:
            # Intenta leer el archivo con la extensión ".xls":
            df_BANCO= pd.read_excel(archivo, header=0)
        except FileNotFoundError:
            # Si el archivo ".xls" no se encuentra, intenta leerlo con la extensión ".xlsx":
            df_BANCO= pd.read_excel(archivo, header=0)
        #------------------------------------------------------------------------------------------#
        ## 2- Columna FECHA:
        FECHA= 'Fecha'
        df_BANCO[f'{FECHA}']= df_BANCO[f'{FECHA}'].dt.strftime('%d/%m/%Y')
        #------------------------------------------------------------------------------------------#
        ## 3- Nº Cuenta (Código del Banco):
        Cod_Banco= 'CAJAMAR'
        #------------------------------------------------------------------------------------------#
        ## 4- Nombre de Cuenta (Nombre del Banco):
        Nombre_Banco= 'CAJAMAR 7924'
        #------------------------------------------------------------------------------------------#
        ## 5- NOMBRE PARA GUARDAR EL EXCEL:
        Nombre_GUARDAR= 'BC CAJAMAR'
    #===============================================================================================================================================================#
    #===============================================================================================================================================================#

    #### 2º) CREACIÓN DEL DF RESULTANTE:
    if 'df_BANCO' in locals() or 'df_BANCO' in globals():

        df_BANCO= df_BANCO.sort_values(by=f'{FECHA}')  # Ordenar de más ANTIGUO-> a más RECIENTE.

        ## 1) DATAFRAME PARA BUSINESS CENTRAL:
        df_tratado= pd.DataFrame()
        df_tratado['Importe (DL)']= df_BANCO['Importe']
        df_tratado['Importe']= df_tratado['Importe (DL)']
        df_tratado['Importe debe']= df_tratado['Importe (DL)']
        df_tratado['Fecha registro']= df_BANCO[f'{FECHA}']
        df_tratado['Fecha de IVA']= df_tratado['Fecha registro']
        df_tratado['Tipo documento']= 'Pago'
        #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

        ## 2) Nº DOCUMENTO:
        # Extraer el número al final del Primer_num_documento:
        if Primer_num_documento:
            numero= int(Primer_num_documento[6:])  # Suponiendo que siempre hay 6 caracteres antes del número.

            # Generar los nuevos valores a continuación del primero:
            nuevos_valores= [Primer_num_documento]
            for i in range(1, len(df_tratado)):
                numero += 1
                nuevo_valor= Primer_num_documento[:6] + str(numero)
                nuevos_valores.append(nuevo_valor)

            # Actualizar la columna 'Nº documento' en el DataFrame:
            df_tratado['Nº documento']= nuevos_valores
        else:
            print('[bold red]¡INTRODUCE EL 1º NÚMERO DE DOCUMENTO CORRECTAMENTE![/bold red]')

        #..........................................................................#
        df_tratado['Tipo mov.']= 'Banco'
        df_tratado['Nº cuenta']= Cod_Banco
        df_tratado['Nombre de cuenta']= Nombre_Banco
        df_tratado['Tipo contrapartida']= 'Cuenta'
        df_tratado['Op. triangular']='No'
        df_tratado['Corrección']= 'No'

        df_tratado[['Nº asiento', 'Descripción', 'Importe haber',
            'Cta. contrapartida', 'Liq. por tipo documento', 'Liq. por nº documento', 'Titulacion Código', 'Liq. por id.', 'Grupo contable prod. gen.',
            'Nº mov. de documentos entrantes', 'Grupo contable neg. gen.', 'Tipo de registro gen.', 'Cód. divisa', 'Tipo regis. contrapartida',
            'Gr. contable negocio contrap.', 'Gr. contable producto contrap.', 'Código de fraccionamiento', 'Comentario', 'Tipo de id.',
            'Nombre de empresa correcto', 'CIF/NIF correcto', 'CECO Código', 'Empleados Código', 'Cód. dim. acceso directo 4', 'Interface Code']]=None
        #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

        ## 3) ORDEN DE LAS COLUMNAS (Como lo tiene Rubén):
        df_FINAL= df_tratado[['Fecha registro', 'Fecha de IVA', 'Nº asiento', 'Tipo documento','Nº documento', 'Tipo mov.', 'Nº cuenta', 'Nombre de cuenta','Descripción', 'Importe',
                    'Importe (DL)', 'Importe debe','Importe haber', 'Tipo contrapartida', 'Cta. contrapartida','Liq. por tipo documento', 'Liq. por nº documento',
                    'Titulacion Código','Liq. por id.', 'Grupo contable prod. gen.','Nº mov. de documentos entrantes', 'Grupo contable neg. gen.','Tipo de registro gen.', 'Cód. divisa',
                    'Tipo regis. contrapartida','Op. triangular', 'Gr. contable negocio contrap.','Gr. contable producto contrap.', 'Código de fraccionamiento','Corrección',
                    'Comentario', 'Tipo de id.', 'Nombre de empresa correcto','CIF/NIF correcto', 'CECO Código', 'Empleados Código','Cód. dim. acceso directo 4', 'Interface Code']]
        #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

        ## 4) GUARDAR EL EXCEL:
        df_FINAL.to_excel(f'RESULTADOS_BC/{Nombre_GUARDAR} ({df_FINAL['Fecha registro'].iloc[0].replace('/','-')} a {df_FINAL['Fecha registro'].iloc[-1].replace('/','-')}).xlsx', index=False)
        #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

        ## 5) COMPROBACIÓN DEL Nº DE MOVIMIENTOS:
        num_Filas_BANCO= len(df_BANCO)
        num_Filas_BC= len(df_FINAL)

        if num_Filas_BANCO==num_Filas_BC:
            mensaje_final= f'Se han obtenido {num_Filas_BC} movimientos.'
        else:
            mensaje_final= f'Se han obtenido {num_Filas_BC} movimientos para BC. Sin embargo, en el banco hay {num_Filas_BANCO} movimientos.'
        #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
        
        ## 6) df_FINAL + IMPORTE_TOTAL:
        df_FINAL.reset_index(drop=True, inplace=True)
        IMPORTE_TOTAL= "{:,.2f} €".format(df_FINAL.Importe.sum()).replace(',','_').replace('.',',').replace('_','.')

        
        return df_FINAL, mensaje_final, IMPORTE_TOTAL
