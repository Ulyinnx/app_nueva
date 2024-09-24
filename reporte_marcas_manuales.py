def procesa_marcas_manuales():
    import pandas as pd
    import os
    import easygui as eg
    import xlrd
    import warnings
    import re
    warnings.filterwarnings("ignore")

    def procesa_reporte_marcas_manuales(centro_de_costo:str):

        #-----------SE PROCESAN LOS ARGUMENTOS DE LA FUNCION-------------

        if centro_de_costo == "FARMASHOP":
            cc_numero = 0
        elif centro_de_costo == "BELA":
            cc_numero = 1
        elif centro_de_costo == "SEO":
            cc_numero = 2
        elif centro_de_costo == "TODO":
            cc_numero = 5
        else:
            print("""
            ************************************************
              No se selecciono un centro de costo correcto
            ************************************************""")
            exit()

        #------------BUSCAR ARCHIVOS LOCALES-------------

        arch_busqueda_rapida = eg.fileopenbox(title= "Archivo BUSQUEDA RAPIDA")
        arch_marcas_manuales = eg.fileopenbox(title= "Archivo MARCAS MANUALES")
        arch_horas_planificadas = eg.fileopenbox(title= "Archivo HORAS PLANIFICADAS")

        #--------PRUEBA DE BOX con INPUT-----------------

        fieldName = ["Cuantos días de consulta son?"]
        q_dias = eg.multenterbox("Días de consulta", fields=fieldName)
        q_dias = int(q_dias[0])

        # ------------------------------------------------
        #MODO DESAROLLADOR - LOS ARCHIVOS DEBEN IR DENTRO DE LA CARPETA AUTOMAT REPORTES.

        # arch_busqueda_rapida = r"C:\Users\bchavat\Desktop\Automat Reportes\Búsqueda rápida.xlsx"
        # arch_marcas_manuales = r"C:\Users\bchavat\Desktop\Automat Reportes\Reporte Marcas Manuales.xls"
        # arch_horas_planificadas = r"C:\Users\bchavat\Desktop\Automat Reportes\Reporte Horas Planificadas.xls"
        # q_dias = 7

        #------------------------------------------------

        def df_marcas_teoricas(dias_de_consulta):
            q_dias = dias_de_consulta
            hs_tipo =                   ["C", "C", "C", "C","C", "C", "C", "C", "C", "C", "C", "N45", "J32", "J24", "J18", "J12"]
            teoricas_hs_semanales =     [20,  24,   25,  29, 30,  32,  34,  35,  39,  40,  44,    0,    32,    18,    15,  12]
            teoricas_marcas_semanales = [10,  12,   20,  22, 20,   8,  22,  20,  22,  20,  22,    20,    16,    12,    10,   6]

            marcas_teoricas = pd.DataFrame({'tipo': hs_tipo,
                                            'hs_semanales': teoricas_hs_semanales,
                                            'marcas_teoricas_semanales': teoricas_marcas_semanales})

            marcas_teoricas['por_dia'] = marcas_teoricas['marcas_teoricas_semanales'] / 7
            marcas_teoricas[f'marc_teo_{q_dias}_dias'] = round((marcas_teoricas['por_dia'] * q_dias),2)
            index_n45 = hs_tipo.index("N45")

            return marcas_teoricas, index_n45

        marcas_teoricas = df_marcas_teoricas(q_dias)[0]
        index_n45 = df_marcas_teoricas(q_dias)[1]
        print('Marcas Teoricas -*OK*-')
        #marcas_teoricas.to_excel("marcas_teo.xlsx")


        def procesa_arch_busqueda_rapida(archivo):
            archivo = pd.read_excel(archivo)
            col_sep = archivo["Unnamed: 5"].str.split('/', expand=True)
            archivo = pd.concat([archivo, col_sep], axis=1)

            archivo.drop(["Unnamed: 2", "Unnamed: 5",1,2,4,3,6], axis= 1, inplace= True)

            archivo = archivo.set_axis(['Nombre',
                                        'NroColab',
                                        'Centro de Costo',
                                        'Puesto',
                                        'Empresa',
                                        'Nochero'],
                                       axis=1
                                        )

            archivo = archivo.iloc[1: , :]

            archivo.reset_index(drop=True, inplace=True)
            archivo['NroColab'] = pd.to_numeric(archivo['NroColab'], downcast='integer', errors='coerce')
            print("Busqueda rapida -*OK*-")
            return archivo

        arch_busqueda_rapida = procesa_arch_busqueda_rapida(arch_busqueda_rapida)
        #arch_busqueda_rapida.to_excel("r_brapida.xlsx")


        def procesa_arch_horas_planificadas(archivo):

            archivo = pd.read_excel(archivo)
            archivo = archivo.iloc[1:, :]

            archivo = archivo.set_axis(["Nombre Funcionario",
                                        "NroColab",
                                        "Sucursal",
                                        "Hrs Contrato",
                                        "hs_semanales",
                                        "Días Consulta",
                                        "Hrs Planificadas",
                                        "Días Planificados",
                                        "Días No Planificados",
                                        "Días Codigos de pago"],
                                       axis=1,
                                       )

            archivo = archivo.iloc[1:, :]

            archivo.drop(["Nombre Funcionario",
                          "Sucursal",
                          "Días Consulta",
                          "Hrs Planificadas",
                          "Días Planificados",
                          "Días No Planificados",
                          "Días Codigos de pago"],
                         axis=1,
                         inplace=True
                         )

            archivo['NroColab'] = pd.to_numeric(archivo['NroColab'], downcast='integer', errors='coerce')
            archivo['hs_semanales'] = pd.to_numeric(archivo['hs_semanales'], downcast='integer', errors='coerce')

            archivo['jornalero'] = 0
            largo = len(archivo['Hrs Contrato'])

            archivo.reset_index(drop=True, inplace=True)
            for i in range(largo):
                if archivo['Hrs Contrato'].values[i] * 4 * 4.33 * archivo['hs_semanales'].values[i] == 4433.92:
                    archivo['jornalero'][i] = "J32"
                elif archivo['Hrs Contrato'].values[i] * 3 * 4.33 * archivo['hs_semanales'].values[i] == 2494.08:
                    archivo['jornalero'][i] = "J24"
                elif archivo['Hrs Contrato'].values[i] * 3 * 4.33 * archivo['hs_semanales'].values[i] == 1402.92:
                    archivo['jornalero'][i] = "J18"
                elif archivo['Hrs Contrato'].values[i] * 3 * 4.33 * archivo['hs_semanales'].values[i] == 623.52:
                    archivo['jornalero'][i] = "J12"
                else:
                    archivo['jornalero'][i] = ""

            print("Horas planificadas -*OK*-")
            return archivo

        arch_horas_planificadas = procesa_arch_horas_planificadas(arch_horas_planificadas)
        #arch_horas_planificadas.to_excel("r_hs_planificadas.xlsx")

        def procesa_arch_marcas_manuales(archivo, cc_numero):
            archivo = pd.read_excel(archivo)
            archivo = archivo.iloc[2:, :]
            archivo = archivo.set_axis(["Nombre",
                                        "NroColab",
                                        "Sucursal",
                                        "Fecha Edición",
                                        "Hora Edición",
                                        "Usuario",
                                        "Fecha Marca",
                                        "Hora Marca"],
                                        axis=1,
                                     )

            mask = archivo["NroColab"] != archivo["Usuario"]
            archivo = archivo[mask]

            archivo.dropna(axis=0, subset=["Fecha Marca"], inplace=True)
            archivo.drop(labels=["Fecha Edición",
                                 "Hora Edición",
                                 "Usuario",
                                 "Hora Marca"],
                         axis=1,
                         inplace=True)
            archivo['NroColab'] = pd.to_numeric(archivo['NroColab'], downcast='integer',
                                                errors='coerce')

            ## ACA FILTTRAMOS LAS COLUMNAS
            if cc_numero == 5:
                archivo = archivo
            else:
                cc_filtro = [r"FARMASHOP\s\d+", r"BELA\s\d+", r"SEO"]
                archivo = archivo.loc[archivo.loc[:, "Sucursal"].str.contains(cc_filtro[cc_numero], regex=True)]

            archivo.reset_index(drop=True, inplace=True)

            print("Marcas manuales -*OK*-")
            return archivo

        arch_marcas_manuales = procesa_arch_marcas_manuales(arch_marcas_manuales, cc_numero)


        def creando_base(arch_marcas_manuales, arch_busqueda_rapida, arch_horas_planificadas):
            print("Creando base de datos...")
            #Concatenando DataFrame
            base_marcas_manuales = pd.merge(arch_marcas_manuales, arch_busqueda_rapida, how='left', on="NroColab")
            #base_marcas_manuales.info()
            base_marcas_manuales = pd.merge(base_marcas_manuales, arch_horas_planificadas, how='left', on="NroColab")
            #base_marcas_manuales.info()
            base_marcas_manuales.drop(labels=["Nombre_y"],
                                      axis=1,
                                      inplace=True)

            print("Base creada -*OK*-")
            return base_marcas_manuales

        base_marcas_manuales = creando_base(arch_marcas_manuales, arch_busqueda_rapida, arch_horas_planificadas)


        def creando_reporte_colab(base_marcas_manuales):
            reporte_colab = base_marcas_manuales.groupby(['NroColab']).agg({'Nombre_x': ['min'],
                                                                            'Sucursal': ['min'],
                                                                            'Puesto': ['min'],
                                                                            'NroColab': ['count'],
                                                                            'hs_semanales': ['min'],
                                                                            'Nochero': ['min'],
                                                                            'jornalero': ['min']
                                                                            })

            reporte_colab.reset_index(drop=False, inplace=True)
            reporte_colab.columns = reporte_colab.columns.droplevel(1)

            reporte_colab = pd.merge(reporte_colab, marcas_teoricas, how='left', on="hs_semanales")

            reporte_colab['Nochero'].replace({"-": "", "NOCHERO": "N45"}, inplace=True)
            reporte_colab['tipo'] = reporte_colab.Nochero + reporte_colab.jornalero
            reporte_colab['tipo'].replace({"": "C"}, inplace=True)

            reporte_colab.drop(['Nochero',
                                'jornalero',
                                "por_dia"],
                               axis= 1,
                               inplace= True)

            reporte_colab = reporte_colab.set_axis(['NroColab',
                                                    'Nombre',
                                                    'Centro de Costo',
                                                    'Puesto',
                                                   f'Marcas manuales en {q_dias} dias',
                                                    'Horas Semanales',
                                                    'Tipo',
                                                    'marcTeo semanal',
                                                   f'MarcTeo {q_dias} dias'],
                                                 axis=1)

            reporte_colab = reporte_colab.loc[
                ~reporte_colab['Puesto'].str.contains(re.compile('[Ee][Nn][Cc].*'))]
            reporte_colab = reporte_colab.loc[
                ~reporte_colab['Puesto'].str.contains(re.compile('Jefe de Salon OM'))]

            reporte_colab.reset_index(drop=True, inplace=True)
            reporte_colab['% marcas manuales/ marcas teóricas'] = 0.0

            largo = len(reporte_colab['% marcas manuales/ marcas teóricas'])

            reporte_colab['Tipo'] = reporte_colab['Tipo'].str.strip()

            for i in range(largo):
                if reporte_colab['Tipo'].values[i] == "N45" and reporte_colab['Horas Semanales'].values[i] == 44 :
                    reporte_colab[f'MarcTeo {q_dias} dias'].values[i] = marcas_teoricas[f'marc_teo_{q_dias}_dias'][index_n45]
                    reporte_colab['% marcas manuales/ marcas teóricas'].values[i] = reporte_colab[f'Marcas manuales en {q_dias} dias'].values[i] / ((20/7) * q_dias)
                elif reporte_colab['Tipo'].values[i] == r"J\d+":
                    if reporte_colab['Horas Semanales'].values[i] == 32:
                        reporte_colab['% marcas manuales/ marcas teóricas'].values[i] = reporte_colab[f'Marcas manuales en {q_dias} dias'].values[i] / ((16/7) * q_dias)
                    elif reporte_colab['Horas Semanales'].values[i] == 24:
                        reporte_colab['% marcas manuales/ marcas teóricas'].values[i] = reporte_colab[f'Marcas manuales en {q_dias} dias'].values[i] / ((12/7) * q_dias)
                    elif reporte_colab['Horas Semanales'].values[i] == 18 or 12:
                        reporte_colab['% marcas manuales/ marcas teóricas'].values[i] = reporte_colab[f'Marcas manuales en {q_dias} dias'].values[i] / ((12/7) * q_dias)
                else:
                    reporte_colab['% marcas manuales/ marcas teóricas'].values[i] = reporte_colab[f'Marcas manuales en {q_dias} dias'].values[i] / reporte_colab[f'MarcTeo {q_dias} dias'].values[i]

            print("Hallando porcentaje de marcas manuales -*OK*-")


            reporte_colab.sort_values(by='% marcas manuales/ marcas teóricas', ascending=False, inplace=True)


            reporte_colab = reporte_colab[['NroColab',
                                            'Nombre',
                                            'Centro de Costo',
                                            'Puesto',
                                            'Horas Semanales',
                                           f'Marcas manuales en {q_dias} dias',
                                           f'MarcTeo {q_dias} dias',
                                            'Tipo',
                                            '% marcas manuales/ marcas teóricas']]

            reporte_colab.reset_index(drop=True, inplace=True)
            return reporte_colab

        reporte_colab = creando_reporte_colab(base_marcas_manuales)
        #reporte_colab.to_excel('reporte_colab.xlsx', index= False)

        print("Reporte Colaborador -*OK*-")


        def creando_reporte_cencost(base_marcas_manuales):
            reporte_cencost = pd.DataFrame()

            reporte_cencost['Centro de Costo'] = base_marcas_manuales['Sucursal']
            reporte_cencost['NroColab'] = base_marcas_manuales['NroColab']
            reporte_cencost.drop_duplicates('NroColab', inplace= True)
            reporte_cencost = reporte_cencost.groupby(['Centro de Costo']).agg({'NroColab': ['count']})

            colab_total = pd.DataFrame()
            colab_total = arch_busqueda_rapida.groupby('Centro de Costo').agg({'NroColab': ['count']})

            reporte_cencost = pd.merge(reporte_cencost, colab_total, on= 'Centro de Costo', how='left')
            reporte_cencost['%Colabordador con al menos una marca manual'] = round(reporte_cencost['NroColab_x'] / reporte_cencost['NroColab_y'], 3)

            reporte_cencost.reset_index(inplace=True)

            reporte_cencost.columns = reporte_cencost.columns.droplevel(1)

            reporte_cencost.sort_values(by= '%Colabordador con al menos una marca manual', ascending=False, inplace=True)

            reporte_cencost.rename(columns={"NroColab_x": "Cantidad de colaboradores con marcas manuales",
                                            "NroColab_y": "Cantidad de colaboradores total"},
                                   inplace=True)

            reporte_cencost.reset_index(inplace=True, drop= True)

            return reporte_cencost

        #reporte_cencost.to_excel('reporte_cencost.xlsx')#, index= False)
        reporte_cencost = creando_reporte_cencost(base_marcas_manuales)

        print("Reporte Centro de Costo -*OK*-")


        #TRANSFORMANDO REPORTE EN EXCEL
        import xlsxwriter
        print("Pasando reportes a Excel...")
        first_writer = pd.ExcelWriter('Reporte_completo.xlsx',
                                      engine='xlsxwriter',
                                      options={'nan_inf_to_errors': True})

        base_marcas_manuales.to_excel(first_writer, sheet_name='Base')

        reporte_colab.to_excel(first_writer, sheet_name='Colaborador', index= False)
        reporte_cencost.to_excel(first_writer, sheet_name='CentroCosto', index= False)

        workbook = first_writer.book
        formato_borde = workbook.add_format({'border': 1})
        limpia_borde = workbook.add_format()
        formato_condicional = {'type': '3_color_scale',
                               'min_color': '#63BE7B',
                               'max_color': '#F8696B'}

        formato_porcentaje = workbook.add_format({'num_format': '0.0%', 'border':1})
        #estilo = workbook.add_format({'bg_color': '#4F81BD',
        #                              'font_color': 'white',
        #                              'align': 'center',
        #                              'valign':'vcenter',
        #                              'text_wrap':True})

        #HOJA COLABORADOR
        worksheet_colab = first_writer.sheets['Colaborador']

        largo_colab = len(reporte_colab)

        for row in range(0, largo_colab):
            valor = reporte_colab['% marcas manuales/ marcas teóricas'][row]
            worksheet_colab.write(f'I{row + 2}', valor, formato_porcentaje)

        worksheet_colab.set_column('A:I', 15, formato_borde)
        worksheet_colab.conditional_format(f'I2:I{largo_colab + 1}', formato_condicional)

        #HOJA CENTRO DE COSTO

        worksheet_cc = first_writer.sheets['CentroCosto']

        largo_cc = len(reporte_cencost)

        for row in range(0, (largo_cc)):
            valor = reporte_cencost['%Colabordador con al menos una marca manual'][row]
            worksheet_cc.write(f'D{row + 2}', valor, formato_porcentaje)

        worksheet_cc.conditional_format(f'D2:D{largo_cc + 1 }', formato_condicional)
        worksheet_cc.set_column(f'A:D', 30, formato_borde)
        #worksheet_cc.set_row(0, 30, estilo)

        workbook.close()
        print('-*Finalizado*-')


    procesa_reporte_marcas_manuales("TODO")

#procesa_marcas_manuales()

