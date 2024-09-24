# def graficos_rp():
#     import pickle
#
#     with open(r'C:\Users\bchavat\Desktop\Automat Reportes\RP_test', 'rb') as archivo:
#         reporte_rp = pickle.load(archivo)
#
#     import seaborn as sns
#     import matplotlib.pyplot as plt
#     import re
#     import plotly.express as px
#     import pandas as pd
#
#     lista = ['.*',
#              'FARMASHOP.*',
#              'BELA.*',
#              'Ecommerce.*',
#              'OM.*',
#              'Omnicanalidad.*',
#              'LOG.*',
#              'P[uU].*',
#              'SEO.*',
#              'Kiehl.*']
#
#     nombres = ['Compañía',
#                'Farmashop',
#                'Bela',
#                'Ecommerce',
#                'OM',
#                'Omnicanalidad',
#                'Log. y Dist.',
#                'Botiga',
#                'SEO',
#                '''Kiehl's''']
#
#     leyenda = reporte_rp['Observaciones'].value_counts().index.tolist()
#
#     for list, nombre in zip(lista, nombres):
#         mask = reporte_rp.loc[:, 'Sucursal'].str.contains(re.compile(list))
#         reporte_rp_cc = reporte_rp[mask]
#         x = reporte_rp_cc.value_counts('Observaciones')
#         print(f'    -**- {nombre} -**-')
#         print("")
#         print(x)
#         print("-----------------")
#
#         ep = [0.02]
#         for i in range(len(x) - 2):
#             i = i + 2
#             y = x[i] / abs(x.sum() - x[0])
#             resultado = round(y / 10, 3)
#             ep.append(resultado)
#
#         if len(x) == 1:
#             pull = [0]
#         elif len(x) == 2:
#             pull = [ep[0], 0.05]
#         elif len(x) == 3:
#             pull = [ep[0], 0.05, ep[1]]
#         elif len(x) == 4:
#             pull = [ep[0], 0.05, ep[1], ep[2]]
#         elif len(x) == 5:
#             pull = [ep[0], 0.05, ep[1], ep[2], ep[3]]
#         elif len(x) == 6:
#             pull = [ep[0], 0.05, ep[1], ep[2], ep[3], ep[4]]
#
#         fig = plt.figure
#         fig = px.pie(values= x,
#                      names= x.index.tolist(),
#                      title=nombre,
#                      color_discrete_sequence=px.colors.qualitative.T10,
#                      width=1200,
#                      height=900,
#                      labels={'size':10})
#         fig.update_layout(font=
#                             dict(size=14),
#                           legend=dict(
#                               font=dict(size=20)
#                                     )
#                           )
#         fig.update_traces(pull= pull)
#
#         fig.write_image(r"C:\Users\bchavat\Desktop\Automat Reportes\graficas_rp\graf_por_centro_de_costo" + f'_{nombre}.png')
#
# #graficos_rp()