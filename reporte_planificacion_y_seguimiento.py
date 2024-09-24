import BD
import main


def procesa_planificacion_y_seguimiento(tipo_rep, fecha_seleccionada, bd):

    #import Graficos_RP
    import pandas as pd
    import os
    import easygui as eg
    import re
    from datetime import datetime
    import warnings
    warnings.filterwarnings("ignore")


    def abre_cierra (accion, ruta):
        import os
        import subprocess
        ruta = ruta

        if os.path.exists(ruta):
            if accion == "open":
                subprocess.Popen([r"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE", ruta])
            elif accion == "close":
                subprocess.call(['taskkill', '/f', '/im', 'excel.exe'], shell=True)

    #fecha_semana = eg.enterbox("Fecha de semana con el siguiente formato d/m/Y")
    #fecha_semana =  datetime.strptime(fecha_semana, "%d/%m/%Y")
    arch_busqueda_rapida = eg.fileopenbox(title= "Archivo BUSQUEDA RAPIDA")
    arch_horas_planificadas = eg.fileopenbox(title= "Archivo HORAS PLANIFICADAS")

    arch_guardar = eg.filesavebox(title= "Guardar Reporte")

    fieldName = ["""
    Nro de colaboradores
    separados por coma:
    """]
    justificadas = eg.multenterbox("Inconsistencias justificadas", fields=fieldName)[0]

    justificadas = justificadas.split(',')
    justificadas = [str(x) for x in justificadas]
    justificadas = [x.strip() for x in justificadas]

    def procesa_arch_busqueda_rapida(archivo):
        archivo = pd.read_excel(archivo)
        col_sep = archivo["Unnamed: 5"].str.split('/', expand=True)
        archivo = pd.concat([archivo, col_sep], axis=1)

        archivo.drop(["Unnamed: 2", "Unnamed: 5", 1, 2, 4, 3, 6], axis=1, inplace=True)

        archivo = archivo.set_axis(['Nombre',
                                    'NroColab',
                                    'Centro de Costo',
                                    'Puesto',
                                    'Empresa',
                                    'Nochero'],
                                   axis=1
                                   )

        archivo = archivo.iloc[1:, :]

        archivo.reset_index(drop=True, inplace=True)
        archivo['NroColab'] = pd.to_numeric(archivo['NroColab'], downcast='integer', errors='coerce')



        print("Busqueda rapida -*OK*-")
        return archivo
    arch_busqueda_rapida = procesa_arch_busqueda_rapida(arch_busqueda_rapida)
    arch_busqueda_rapida.to_excel("r_brapida.xlsx", index= False)

    def procesa_arch_horas_planificadas(archivo):
        archivo = pd.read_excel(archivo)
        archivo = archivo.iloc[1:, :]
        archivo.reset_index(drop= True, inplace=True)

        archivo = archivo.set_axis(["Nombre Funcionario",
                                    "NroColab",
                                    "Sucursal",
                                    "Hrs Contrato",
                                    "Hrs Semanal",
                                    "Días Consulta",
                                    "Hrs Planificadas",
                                    "Días Planificados",
                                    "Días No Planificados",
                                    "Días Codigos de pago",],
                                    axis=1,
                                    )
        archivo = archivo.iloc[1:, :]
        archivo.reset_index(drop=True, inplace=True)

        archivo['NroColab'] = pd.to_numeric(archivo['NroColab'], downcast='integer', errors='coerce')
        archivo['Hrs Semanal'] = pd.to_numeric(archivo['Hrs Semanal'], downcast='integer', errors='coerce')

        print("Horas planificadas -*OK*-")
        return archivo
    arch_horas_planificadas = procesa_arch_horas_planificadas(arch_horas_planificadas)
    arch_horas_planificadas.to_excel("Buscando_error.xlsx")



    def creando_reporte_planificacion(b_rapida, horas_planificadas):
        archivo = horas_planificadas.copy()
        b_rapida = b_rapida.copy()

        cc_eliminar = ['-',
                       'Comercial',
                       'Administración',
                       'RRHH',
                       'Delivery 532',
                       'Delivery 534',
                       'Auditoria y procesos',
                       'Operaciones',
                       'Bela oficina',
                       'Marketing',
                       'BOTIGA',
                       'Botiga',
                       'Sistemas',
                       'Dirección']

        archivo = archivo[~archivo['Sucursal'].isin(cc_eliminar)]

        archivo.loc[:, 'Diferencia de horarios'] = archivo.loc[:, 'Hrs Semanal'] - archivo.loc[:, 'Hrs Planificadas']

        ## UNION DE LOS DATAFRAMES
        archivo = pd.merge(archivo, b_rapida, how='left', on="NroColab")

        archivo['jornalero'] = 0
        largo = len(archivo['Hrs Contrato'])

        for i in range(largo):
            if archivo['Hrs Contrato'].values[i] * 4 * 4.33 * archivo['Hrs Semanal'].values[i] == 4433.92:
                archivo['jornalero'].values[i] = "J32"
            elif archivo['Hrs Contrato'].values[i] * 3 * 4.33 * archivo['Hrs Semanal'].values[i] == 2494.08:
                archivo['jornalero'].values[i] = "J24"
            elif archivo['Hrs Contrato'].values[i] * 3 * 4.33 * archivo['Hrs Semanal'].values[i] == 1402.92:
                archivo['jornalero'].values[i] = "J18"
            elif archivo['Hrs Contrato'].values[i] * 3 * 4.33 * archivo['Hrs Semanal'].values[i] == 623.52:
                archivo['jornalero'].values[i] = "J12"
            else:
                archivo.loc[i,'jornalero'] = ""

        archivo['Nochero'].replace({"-": "", "NOCHERO": "N45"}, inplace=True)
        colab_tipo = archivo.Nochero + archivo.jornalero
        archivo.insert(2, "Tipo", colab_tipo)
        archivo['Tipo'].replace({"": "C"}, inplace=True)
        archivo['Tipo'] = archivo['Tipo'].astype(str)

        archivo['Puesto'] = archivo['Puesto'].astype(str)
        archivo = archivo.loc[~archivo['Puesto'].str.contains(re.compile('[Ee][Nn][Cc].*'))]
        archivo = archivo.loc[~archivo['Puesto'].str.contains(re.compile('Jefe de Salon OM'))]
        archivo.reset_index(drop=True, inplace=True)

        archivo.drop(['Nombre',
                      'Centro de Costo',
                      'Puesto',
                      'Empresa',
                      'Nochero',
                      'jornalero'],
                     axis=1,
                     inplace=True)

        archivo['Observaciones'] = ""

        ## OK // Falta planificar Hs // Tiene mas Hs planificadas que las contractuales

        for i in range(len(archivo['NroColab'])):
            if archivo.loc[i, 'Diferencia de horarios'] == 0:
                archivo.loc[i, 'Observaciones'] = "OK"
            elif archivo.loc[i, 'Diferencia de horarios'] >= 0:
                archivo.loc[i, 'Observaciones'] = "Falta planificar Hs"
            else:
                archivo.loc[i, 'Observaciones'] = "Tiene mas Hs planificadas que las contractuales"

        ## FALTAN/SOBRAN JORNADAS PLANIFICADAS

        especificaciones = r"C:\Users\bchavat\Desktop\Automat Reportes\r_data\especificaciones_data.xlsx"
        especificaciones = pd.read_excel(especificaciones)
        def compara_valor_real_teorico(centro_de_costo):
            valor_real = round(archivo.loc[i, 'Hrs Contrato'] * archivo.loc[i, 'Hrs Semanal'] / archivo.loc[i, 'Días Planificados'] / archivo.loc[i, 'Días No Planificados'], 2).astype(float)
            if valor_real not in list(round(especificaciones.loc[especificaciones['Centro de Costo'] == centro_de_costo]['valor unico'], 2).astype(float)):
                return True

        for i in range(len(archivo['NroColab'])):
            if archivo.loc[i,'Tipo'] == 'C':
                if re.search(re.compile("FARMASHOP.*"), archivo.loc[i,'Sucursal']):
                    if compara_valor_real_teorico('Farmashop') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("BELA.*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('Bela') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("Ecommerce.*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('Ecommerce 900') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("OM.*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('OM') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("Delivery.*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('Delivery') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("LOG.*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('Logistica') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("P[uU].*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('Logistica') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("SEO.*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('SEO') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                elif re.search(re.compile("Kiehl.*"), archivo.loc[i, 'Sucursal']):
                    if compara_valor_real_teorico('Kiehls') == True:
                        archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
                else:
                    None
            elif archivo.loc[i,'Tipo'] == 'N45':
                valor_real = round(archivo.loc[i, 'Hrs Contrato'] * archivo.loc[i, 'Hrs Semanal'] / archivo.loc[i, 'Días Planificados'] / archivo.loc[i, 'Días No Planificados'], 2).astype(float)
                if valor_real not in list(round(especificaciones.loc[especificaciones['Centro de Costo'] == 'Nochero']['valor unico'], 2).astype(float)):
                    archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
            elif re.search(re.compile("J.*"), archivo.loc[i, 'Tipo']):
                valor_real = round(archivo.loc[i, 'Hrs Contrato'] * archivo.loc[i, 'Hrs Semanal'] / archivo.loc[i, 'Días Planificados'] / archivo.loc[i, 'Días No Planificados'], 2).astype(float)
                if valor_real not in list(round(especificaciones.loc[especificaciones['Centro de Costo'] == 'Jornalero']['valor unico'], 2).astype(float)):
                    archivo.loc[i, 'Observaciones'] = "Faltan/Sobran jornadas planificadas"
            else:
                None

            ## NO TIENE PLANIFICACION
        archivo_null = archivo.isnull()
        for i in range(len(archivo['NroColab'])):
            if archivo_null.loc[i, 'Hrs Planificadas'] == True:
                archivo.loc[i, 'Observaciones'] = "No tiene planificación"

             ## JUSTIFICADAS
        archivo['NroColab'] = archivo['NroColab'].astype(int)
        archivo['NroColab'] = archivo['NroColab'].astype(str)

        print("""
               *****************************
               Evaluación de justificaciones
               *****************************
            """)
        for l in justificadas:
            mask = archivo['NroColab'].str.contains(l)
            if mask.any():
                indice = archivo[mask].index[0]
                archivo.loc[indice, 'Observaciones'] = "Justificada"
                print(f"Se justifica el colaborador nro: {l}")
            else:
                print(f"El colaborador nro {l}, no está en lista")

        archivo['NroColab'] = archivo['NroColab'].astype(float).astype(int)

        print("""
            ************************************
            Fin de evaluación de justificaciones
            ************************************
            """)

        return archivo

    reporte_rp = creando_reporte_planificacion(arch_busqueda_rapida, arch_horas_planificadas)
    reporte_rp.to_excel(f"{arch_guardar}.xlsx", sheet_name= "BASE", index=False)



    ## CREACION DE ANEXOS
    def creacion_anexos(reporte_rp):
        filtros_obs = []
        reporte_rp['NroColab'] = reporte_rp['NroColab'].astype(str)
        for lista, i in zip(['No tiene planificación', 'Faltan/Sobran jornadas planificadas',
                             'Falta planificar Hs', 'Tiene mas Hs planificadas que las contractuales'], [1, 2, 3, 4]):
            obsfil = reporte_rp.loc[:, ['Sucursal', 'NroColab', 'Observaciones', 'Diferencia de horarios']]
            mask = obsfil.loc[:, 'Observaciones'] == lista
            obsfil = obsfil[mask]

            grouped = obsfil.groupby('Sucursal')
            colab_por_sucursal = grouped['NroColab'].apply(', '.join)

            if i == 1 or i == 2:
                colab_count = grouped['NroColab'].count()  ## Cantidad de colaboradores por sucursal.
                obsfil = pd.concat([colab_por_sucursal, colab_count], axis=1)
                obsfil.columns = ['NroColab', 'Recuento']
                obsfil.sort_values(by='Recuento', ascending=False, inplace=True)
            elif i == 3 or i == 4:
                dif_por_sucursal = abs(
                    grouped['Diferencia de horarios'].sum())  ## Cantidad de horas de inconsistencia por sucursal
                obsfil = pd.concat([colab_por_sucursal, dif_por_sucursal], axis=1)
                obsfil.columns = ['NroColab', 'Diferencia de horas']
                obsfil.sort_values(by='Diferencia de horas', ascending=False, inplace=True)

            obsfil.reset_index(inplace=True)

            filtros_obs.append(obsfil)

        inco_filt = reporte_rp.loc[:, "Observaciones"] != 'OK'
        inconsistencias = reporte_rp[inco_filt]

        filtros_obs[2]['Diferencia de horas'] = filtros_obs[2]['Diferencia de horas'].astype(float).round(2)
        filtros_obs[3]['Diferencia de horas'] = filtros_obs[3]['Diferencia de horas'].astype(float).round(2)

        with pd.ExcelWriter(r'C:\Users\bchavat\Desktop\Automat Reportes\Anexo.xlsx') as writer:
            inconsistencias.to_excel(writer, sheet_name='Anexo', index=False)
            filtros_obs[0].to_excel(writer, sheet_name='NoTplan', index=False)
            filtros_obs[1].to_excel(writer, sheet_name='FSJornPlan', index=False)
            filtros_obs[2].to_excel(writer, sheet_name='FaltaPlanHs', index=False)
            filtros_obs[3].to_excel(writer, sheet_name='MasHsPlan', index=False)

        import BD
        if bd is True:
            if tipo_rep == "P":
                BD.insertar_en_base_de_datos(inconsistencias, "Planificacion", fecha_seleccionada)
            elif tipo_rep == "S":
                BD.insertar_en_base_de_datos(inconsistencias, "Seguimiento", fecha_seleccionada)
            else:
                print("Error en seleccion de BD para incrustar datos.")
        else:
            None

        print("Anexo --*OK*--")
        return None
    creacion_anexos(reporte_rp)


    ## GRAFICOS
    def graficos_rp():
        import matplotlib.pyplot as plt
        import re
        import plotly.express as px


        lista = ['.*',
                 'FARMASHOP.*',
                 'BELA.*',
                 'Ecommerce.*',
                 'OM.*',
                 'Delivery.*',
                 'LOG.*',
                 'P[uU].*',
                 'SEO.*',
                 'Kiehl.*']

        nombres = ['Compañía',
                   'Farmashop',
                   'Bela',
                   'Ecommerce 900',
                   'OM',
                   'Delivery',
                   'Log. y Dist.',
                   'Botiga',
                   'SEO',
                   '''Kiehl's''']

        leyenda = reporte_rp['Observaciones'].value_counts().index.tolist()

        for list, nombre in zip(lista, nombres):
            mask = reporte_rp.loc[:, 'Sucursal'].str.contains(re.compile(list))
            reporte_rp_cc = reporte_rp[mask]
            x = reporte_rp_cc.value_counts('Observaciones')
            print(f'    -**- {nombre} -**-')
            print("")
            print(x)
            print("-----------------")

            ep = [0.02]
            for i in range(len(x) - 2):
                i = i + 2
                y = x[i] / abs(x.sum() - x[0])
                resultado = round(y / 10, 3)
                ep.append(resultado)

            if len(x) == 1:
                pull = [0]
            elif len(x) == 2:
                pull = [ep[0], 0.05]
            elif len(x) == 3:
                pull = [ep[0], 0.05, ep[1]]
            elif len(x) == 4:
                pull = [ep[0], 0.05, ep[1], ep[2]]
            elif len(x) == 5:
                pull = [ep[0], 0.05, ep[1], ep[2], ep[3]]
            elif len(x) == 6:
                pull = [ep[0], 0.05, ep[1], ep[2], ep[3], ep[4]]

            fig = plt.figure
            fig = px.pie(values= x,
                         names= x.index.tolist(),
                         title=nombre,
                         color_discrete_sequence=px.colors.qualitative.T10,
                         width=1200,
                         height=900,
                         labels={'size':10})
            fig.update_layout(font=
                                dict(size=14),
                              legend=dict(
                                  font=dict(size=20)
                                        )
                              )
            fig.update_traces(pull= pull)

            fig.write_image(r"C:\Users\bchavat\Desktop\Automat Reportes\graficas_rp\graf_por_centro_de_costo" + f'_{nombre}.png')
    graficos_rp()




    now = datetime.now()
    print(f"""
            Proceso completo, fecha de creación: 
    
            {now.strftime("%Y-%m-%d %H:%M:%S")}
    
          """)

# procesa_planificacion_y_seguimiento("P","16/09/2024", True)
