def procesa_horas_extras():
    import numpy as np
    import pandas as pd
    import os
    # import easygui as eg
    import datetime as dt
    from datetime import datetime
    import warnings
    warnings.filterwarnings("ignore")

    start = datetime.now()
    def abre_cierra (accion, ruta):
        import os
        import subprocess
        import time
        import warnings
        warnings.filterwarnings("ignore")

        ruta = ruta
        if os.path.exists(ruta):
            if accion == "open":
                subprocess.Popen([r"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE", ruta])
            elif accion == "close":
                try:
                    subprocess.check_call(['taskkill', '/f', '/im', 'excel.exe'],
                                          shell=True
                                          )
                    time.sleep(2)
                except subprocess.CalledProcessError:
                    print("El proceso de Excel ya se encuentra cerrado.")
        else:
            None

    # arch_busqueda_rapida = eg.fileopenbox(title= "Archivo BUSQUEDA RAPIDA")
    # arch_horas_extras = eg.fileopenbox(title= "Archivo HORAS EXTRAS")
    # arch_horas_planificadas = eg.fileopenbox(title= "Archivo HORAS PLANIFICADAS")
    # arch_horas_trabajadas = eg.fileopenbox(title= "Archivo HORAS TRABAJADAS")
    # arch_horas_empleados = eg.fileopenbox(title= "Archivo HORAS EMPLEADOS")

    directorio = r"C:\Users\bchavat\Desktop\Automat Reportes\ConstruccionExtras"

    arch_busqueda_rapida = pd.read_excel( directorio + "\Búsqueda rápida.xlsx")
    arch_horas_extras = pd.read_excel( directorio + "\Reporte Horas Extras.xls")
    arch_horas_planificadas = pd.read_excel( directorio + "\Reporte Horas Planificadas.xls")

    import pathlib
    file_extension = [".xlsx", ".xls"]

    for ext in file_extension:
        #print(ext)
        path_horas_trabajas = pathlib.Path(directorio + "\Reporte Horas Trabajadas" + ext)
        if path_horas_trabajas.is_file():
            arch_horas_trabajadas = pd.read_excel(path_horas_trabajas)

    arch_horas_empleados = pd.read_excel( directorio + "\Totales Horas Empleados.xls",
                                          header= None
                                          )

    directorio_data = r"C:\Users\bchavat\Desktop\Automat Reportes\r_data"

    arch_cc_info = pd.read_excel(directorio_data + "\centros_de_costo_info.xlsx")
    arch_gc_data = pd.read_excel(directorio_data + "\grupos_de_cargo_data.xlsx")

    def procesa_arch_busqueda_rapida(archivo):
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
        archivo['NroColab'] = pd.to_numeric(archivo['NroColab'], downcast='integer',
                                            errors='coerce')
        print("Busqueda rapida -*OK*-")
        return archivo

    arch_busqueda_rapida = procesa_arch_busqueda_rapida(arch_busqueda_rapida)

    def procesa_horas_trabajadas(archivo):
        name_col = []
        max_col = int(archivo.shape[1])
        for i in range(0, max_col):
            celda = archivo.iloc[1, i]
            name_col.append(celda)

        archivo = archivo.iloc[2:, :]
        archivo = archivo.set_axis(name_col, axis=1)

        archivo[["ID FUNCIONARIO", "SUCURSAL"]] = archivo[["ID FUNCIONARIO", "SUCURSAL"]].astype(int)
        archivo["FECHA"] = pd.to_datetime(archivo["FECHA"]).dt.strftime('%d/%m/%Y')

        archivo.reset_index(drop=True, inplace=True)

        max_row = int(archivo.shape[0])
        for i in range(0, max_row):
            if pd.isna(archivo.iloc[i, 6]):
                archivo.iloc[i, 6] = archivo.iloc[i, 4]

        archivo = archivo.groupby("ID FUNCIONARIO").agg({"HRS CONTRATO": "min",
                                                         "HRS TRABAJADAS": "sum"})
        return archivo

    arch_horas_trabajadas = procesa_horas_trabajadas(arch_horas_trabajadas)

    def procesa_horas_extras(archivo):
        name_col = []
        max_col = int(archivo.shape[1])
        for i in range(0, max_col):
            celda = archivo.iloc[1, i]
            name_col.append(celda)

        archivo = archivo.iloc[2:, :]
        archivo = archivo.set_axis(name_col, axis=1)


        for name, i in zip(name_col, range(0, max_col)):
            if i == 0 or i > 3:
                archivo[name] = pd.to_numeric(archivo[name], downcast='integer',
                                              errors='coerce')


        def rellenar_celda_vacia(columna: int):
            max_row = int(archivo.shape[0])
            for i in range(0, max_row):
                if pd.isna(archivo.iloc[i,columna]):
                    archivo.iloc[i, columna] = archivo.iloc[int(i-1), columna]
                else:
                    None

        rellenar_celda_vacia(0)
        rellenar_celda_vacia(1)
        rellenar_celda_vacia(2)

        archivo.fillna(0, inplace=True)

        no_sum_extras = ['Cod. Funcionario',
                         'Nombre Funcionario',
                         'Sucursal',
                         'Fecha',
                         'Horas por Aprobar',
                         'Hora Extra Rechazada',
                         'Total Horas']
        sum_extras = []
        for name in name_col:
            if name in no_sum_extras:
                None
            else:
                sum_extras.append(name)

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

        archivo['Total extras autorizadas'] = 0
        for i in range(0, len(sum_extras)):
            archivo['Total extras autorizadas'] = archivo['Total extras autorizadas'] + archivo[sum_extras[i]]

        archivo = archivo[archivo['Total extras autorizadas'] != 0]

        puestos = arch_busqueda_rapida[['NroColab', 'Puesto', 'Nochero']]

        archivo = pd.merge(archivo, arch_horas_trabajadas, left_on='Cod. Funcionario',
                           right_on='ID FUNCIONARIO', how='left')
        archivo = pd.merge(archivo, puestos, left_on='Cod. Funcionario',
                           right_on='NroColab', how='left').drop('NroColab', 1)

        archivo = pd.merge(archivo, arch_gc_data, left_on='Puesto',
                           right_on='Puesto', how='left')

        archivo = pd.merge(archivo, arch_cc_info, left_on='Sucursal', right_on='SUCURSAL',
                           how= 'left').drop('SUCURSAL', 1)

        #### CAMBIA DE ORDEN COMO SE MUESTRA LA FECHA>>> DE: d/m/Y a m/d/Y
        #archivo['Fecha'] = pd.to_datetime(archivo['Fecha']).dt.strftime('%d/%m/%Y')

        archivo.drop(['DIRECCION'], axis=1, inplace=True)

        ## HORAS EXTRAS DE ARMADO
        arch_armado = arch_horas_empleados
        name_armados_col = []
        max_col = int(arch_armado.shape[1])
        for i in range(0, max_col):
            celda = arch_armado.iloc[3, i]
            name_armados_col.append(celda)
        arch_armado = arch_armado.set_axis(name_armados_col, axis=1)
        arch_armado = arch_armado.iloc[4:, :]
        arch_armado['N° FUNC.'] = arch_armado['N° FUNC.'].astype(int)
        arch_armado['FECHA'] = pd.to_datetime(arch_armado['FECHA']).dt.strftime('%d/%m/%Y')
        arch_armado.reset_index(drop= True, inplace=True)

        # DISCRIMINACION DE HORAS

        archivo['Extra Indep.'] = 0
        archivo['Extra Nochero'] = 0
        archivo['Extra Armado'] = 0

        max_row = int(archivo.shape[0])
        max_col = int(archivo.shape[1])
        col_ext_autorizada = max_col - 16
        col_hs_armado = max_col - 1
        col_ext_nochero = max_col - 2
        col_ext_indep = max_col - 3
        col_nochero = max_col - 12

        if arch_armado.empty:
            print("No existen horas de armado para el periodo analizado.")
            for i in range(0, max_row):
                if archivo.iloc[i, col_nochero] == "NOCHERO":
                    archivo.iloc[i, col_ext_nochero] = archivo.iloc[i, col_ext_autorizada]
                else:
                    archivo.iloc[i, col_ext_indep] = archivo.iloc[i, col_ext_autorizada]
        else:
            for i in range(0, max_row):
                if  int(archivo.iloc[i, 0]) in arch_armado['N° FUNC.'].values:
                    fecha_extra = archivo.iloc[i, 3]
                    mask_1 = arch_armado['N° FUNC.'] == int(archivo.iloc[i, 0])
                    arch_fil_1 = arch_armado[mask_1]
                    mask_2 = arch_fil_1['FECHA'] == fecha_extra
                    arch_fil = arch_fil_1[mask_2]
                    print(f"""
    **** EVALUACION DE FECHA DE ARMADO de Colab: {int(archivo.iloc[i, 0])} ****
    Fecha extra a evaluar: {fecha_extra}
    Fechas de armado: 
    {arch_fil_1['FECHA']}""")
                    if not arch_fil.empty:
                        archivo.iloc[i, col_hs_armado] = archivo.iloc[i, col_ext_autorizada]
                        print("→ Hay coincidencia ←")
                    else:
                        archivo.iloc[i, col_ext_indep] = archivo.iloc[i, col_ext_autorizada]
                        print("XXX - No coincide - XXX")
                elif archivo.iloc[i, col_nochero] == "NOCHERO":
                    archivo.iloc[i, col_ext_nochero] = archivo.iloc[i, col_ext_autorizada]
                else:
                    archivo.iloc[i, col_ext_indep] = archivo.iloc[i, col_ext_autorizada]
        archivo['Total Extras'] = archivo['Extra Indep.'] + archivo['Extra Armado'] + archivo['Extra Nochero']
        archivo["Fecha"] = pd.to_datetime(archivo["Fecha"]).dt.strftime('%d/%m/%Y')

        return archivo

    arch_horas_extras = procesa_horas_extras(arch_horas_extras)
    arch_horas_extras.to_excel("r_extras_base.xlsx", index= False)
    #-------------

        ## PORCENTAJE DE HORAS EXTRAS POR SUCURSAL ##

    cantidad_hs_trabajadas = pd.merge(arch_horas_trabajadas,arch_busqueda_rapida,
                                      left_on='ID FUNCIONARIO', right_on='NroColab',
                                      how='left')
    cantidad_hs_trabajadas = cantidad_hs_trabajadas.groupby('Centro de Costo').agg({'HRS TRABAJADAS': 'sum'})
    cantidad_hs_trabajadas.reset_index(inplace=True)

    porcentajes_extras = arch_horas_extras
    porcentajes_extras = porcentajes_extras.groupby('Sucursal').agg({'Total Extras': 'sum'})
    porcentajes_extras.reset_index(inplace=True)

    porcentajes_extras = pd.merge(cantidad_hs_trabajadas, porcentajes_extras,
                                  left_on='Centro de Costo',
                                  right_on='Sucursal', how='left'
                                  )
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

    porcentajes_extras = porcentajes_extras[~porcentajes_extras['Centro de Costo'].isin(cc_eliminar)]


    porcentajes_extras = porcentajes_extras.rename(columns={'HRS TRABAJADAS': 'Total Trabajadas'})
    porcentajes_extras = porcentajes_extras[['Centro de Costo','Total Trabajadas', 'Total Extras']]
    porcentajes_extras.fillna(0, inplace=True)
    porcentajes_extras['Porcentaje de Hs extras/Hs trabajadas'] = 0

    max_row = len(porcentajes_extras)
    for i in range(0, max_row):
        porcentajes_extras.iloc[i,3] = porcentajes_extras.iloc[i,2] / porcentajes_extras.iloc[i,1]
        porcentajes_extras.iloc[i,1] = round(porcentajes_extras.iloc[i,1], 2)

    porcentajes_extras.sort_values(by='Porcentaje de Hs extras/Hs trabajadas',
                                   ascending=False, inplace=True)

    sum_trabajadas = sum(porcentajes_extras['Total Trabajadas'])
    sum_extras = sum(porcentajes_extras["Total Extras"])
    porcentajes_extras = porcentajes_extras.append({},ignore_index=True)
    porcentajes_extras.iloc[max_row,0] = "Total general:"
    porcentajes_extras.iloc[max_row,1] = sum_trabajadas
    porcentajes_extras.iloc[max_row,2] = sum_extras
    porcentajes_extras.iloc[max_row,3] = sum_extras / sum_trabajadas

    porcentajes_extras.to_excel("porcentaje_extras.xlsx", index=False)
    #-------------------------

    # # ANEXO BASE
    #
    # sheet_name = f"{nombre} Cod. Pago"
    # workbook.create_sheet(sheet_name)
    # workbook_sheet = workbook[sheet_name]
    # workbook_sheet.sheet_properties.tabColor = color_tab
    #
    # medidas = df_anexo.shape
    #
    # largo_num = medidas[0] + 1
    # ancho_num = medidas[1] + 1
    # ancho_letra = get_column_letter(ancho_num - 1)
    #
    # # Entrada de datos a hoja
    # for row in dataframe_to_rows(df_anexo, index=False, header=True):
    #     workbook_sheet.append(row)
    #
    # # Formateo de hoja general
    # borde_param = Side(style='thin', color="808080")
    # selecciona_celdas(rango=f"A1:{ancho_letra}{largo_num}",
    #                   sheet=workbook_sheet,
    #                   font=Font(name='Calibri Light'),
    #                   alignment=Alignment(horizontal='center', vertical='center'),
    #                   border=Border(left=borde_param, right=borde_param,
    #                                 top=borde_param, bottom=borde_param))
    # # Formateo del header
    # selecciona_celdas(rango=f"A1:{ancho_letra}1",
    #                   sheet=workbook_sheet,
    #                   font=Font(bold=True, color=color_title),
    #                   fill=PatternFill(start_color=color_header, fill_type="solid"))
    #
    # # Ajuste de tamaño de columnas segun celda mas larga
    # for column in workbook_sheet.columns:
    #     max_length = 0
    #     for cell in column:
    #         try:
    #             if len(str(cell.value)) > max_length:
    #                 max_length = len(cell.value)
    #         except:
    #             pass
    #     adjusted_width = (max_length + 3)
    #     workbook_sheet.column_dimensions[cell.column_letter].width = adjusted_width





                     #####################################
                #####              GRAFICOS               #####

    import matplotlib as plt
    import plotly.express as px

    fig = px.pie(values= arch_horas_extras['Total Extras'], names= arch_horas_extras['Gerente'], title="Gerencia",
                         color_discrete_sequence=px.colors.qualitative.T10, width=1200,
                         height=900, labels={'size':10})

    fig.update_layout(font= dict(size=14), legend=dict(font= dict(size= 20)))

    print(arch_horas_extras.groupby(by="Gerente").agg({"Total Extras": "sum"}))

    fig.write_image(r"C:\Users\bchavat\Desktop\Automat Reportes\ConstruccionExtras\graf_prueba.png")


    finish = datetime.now()
    tiempo_transcurrido = finish - start
    print(f"Tiempo de procesamiento: {tiempo_transcurrido}")

# procesa_horas_extras()
