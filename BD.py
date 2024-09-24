def insertar_en_base_de_datos(insert_data, base_name, fecha_seleccionada, ):
    import main
    import sqlite3
    import pandas as pd
    import shutil

    conn = sqlite3.connect(r'C:\Users\bchavat\Desktop\Automat Reportes\r_data\data_base_reportes.db')
    c = conn.cursor()

    c.execute(f'''CREATE TABLE IF NOT EXISTS {base_name}
                 (Nombre_Funcionario text,
                  NroColab integer,
                  Tipo text,
                  Sucursal integer, 
                  Hrs_Contrato integer, 
                  Hrs_Semanal integer, 
                  Días_Consulta integer, 
                  Hrs_Planificadas integer,
                  Días_Planificados integer,
                  Días_No_Planificados integer,
                  Días_Codigos_de_pago integer,
                  Diferencias_de_horarios real,
                  Observaciones text,
                  Fecha date)''')

    df = insert_data
    df["Fecha"] = fecha_seleccionada
    #df["Fecha"] = pd.to_datetime()

    df = df.rename(columns={
        'Nombre Funcionario': 'Nombre_Funcionario',
        'NroColab': 'NroColab',
        'Tipo': 'Tipo',
        'Sucursal': 'Sucursal',
        'Hrs Contrato': 'Hrs_Contrato',
        'Hrs Semanal': 'Hrs_Semanal',
        'Días Consulta': 'Días_Consulta',
        'Hrs Planificadas': 'Hrs_Planificadas',
        'Días Planificados': 'Días_Planificados',
        'Días No Planificados': 'Días_No_Planificados',
        'Días Codigos de pago': 'Días_Codigos_de_pago',
        'Diferencia de horarios': 'Diferencias_de_horarios',
        'Observaciones': 'Observaciones',
        'Fecha': 'Fecha'
    })


    df.to_sql(f'{base_name}', conn, if_exists='append', index=False)

    df_compartir = pd.read_sql_query(f"SELECT * FROM {base_name}", conn)

    # destino = fr"F:\Dimensionamiento\base_planificacion_y_seguimiento\data_base_{base_name}.xlsx"
    # df_compartir.to_excel(destino)

    conn.commit()
    conn.close()
    #
    # base = r"C:\Users\bchavat\Desktop\Automat Reportes\r_data\data_base_reportes.db"
    #


    print(df_compartir)



#
#
# import pandas as pd
#
# columnas = {
#     'Nombre Funcionario': [],
#     'NroColab': [],
#     'Tipo': [],
#     'Sucursal': [],
#     'Hrs Contrato': [],
#     'Hrs Semanal': [],
#     'Días Consulta': [],
#     'Hrs Planificadas': [],
#     'Días Planificados': [],
#     'Días No Planificados': [],
#     'Días Codigos de pago': [],
#     'Diferencia de horarios': [],
#     'Observaciones': [],
#     'Fecha': []
# }
#
# # Crear el DataFrame con las columnas definidas
# df = pd.DataFrame(columns=columnas.keys())
#
# # Agregar algunos valores a cada columna
# df['Nombre Funcionario'] = ['Juan', 'María', 'Carlos']
# df['NroColab'] = [1, 2, 3]
# df['Tipo'] = ['TipoA', 'TipoB', 'TipoC']
# df['Sucursal'] = ['Sucursal1', 'Sucursal2', 'Sucursal3']
# df['Hrs Contrato'] = [40.0, 38.5, 37.0]
# df['Hrs Semanal'] = [38.0, 37.5, 36.0]
# df['Días Consulta'] = [5, 4, 3]
# df['Hrs Planificadas'] = [37.5, 36.5, 35.0]
# df['Días Planificados'] = [4, 3, 2]
# df['Días No Planificados'] = [1, 1, 1]
# df['Días Codigos de pago'] = [2, 1, 3]
# df['Diferencia de horarios'] = [0.5, 1.0, 1.0]
# df['Observaciones'] = ['Observación 1', 'Observación 2', 'Observación 3']
# df['Fecha'] = pd.to_datetime(['2024-01-01', '2024-01-02', '2024-01-03'])
#
# df.to_excel("PRUEBA_1.xlsx")
#
# fecha_prueba = "01/02/2024"
#
# insertar_en_base_de_datos(df, "Planificacion", fecha_prueba)