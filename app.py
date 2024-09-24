import flet as ft
from datetime import datetime  # Asegúrate de importar datetime si no está ya importado

def main(page: ft.Page):
    # Configuración de la página
    page.title = "Reportes - Dimensionamiento"
    page.window_width = 800
    page.window_height = 600

    page.theme_mode = ft.ThemeMode.LIGHT  # Establecer modo claro para estilo Cupertino

    # Variables para almacenar archivos seleccionados
    selected_files = {}
    current_button = None

    # Banner de imagen
    banner = ft.Image(src="banner_dim.png", width=800)

    # -----------------------------------------------------------------------------
    # Sección "Base de Datos"
    # -----------------------------------------------------------------------------
    # Comentario: Contenedor para la sección "Base de Datos"
    fecha_text_field = ft.TextField(
        label="Fecha:",
        read_only=True,
    )

    bd_radio_group = ft.Row(
        controls=[
            ft.Container(
                content=ft.Column([
                    ft.Text("Agregar reporte a base de datos:"),
                    ft.Container(
                        content=ft.Row(
                            [
                                ft.Radio(value="Sí", label="Sí"),
                                ft.Radio(value="No", label="No"),
                            ],
                            spacing=10  # Opcional: agrega espacio entre los botones
                        ),
                        padding=10  # Opcional: agrega padding al contenedor
                    )
                ])
            ),
            fecha_text_field
        ]
    )


    base_datos_container = ft.Container(
        content=ft.Column(
            [
                ft.Text("Base de Datos", style=ft.TextThemeStyle.HEADLINE_SMALL),
                ft.Row(
                    [
                        bd_radio_group,
                        fecha_text_field,
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
            ],
        ),
        padding=10,
        border=ft.border.all(1, "black54"),
    )

    # -----------------------------------------------------------------------------
    # Sección "Tipo de Reporte"
    # -----------------------------------------------------------------------------
    # Comentario: Contenedor para la sección "Tipo de Reporte"
    tipo_reporte_radio_group = ft.RadioGroup(
        content=ft.Column(
            [
                ft.Radio(value="Planificacion", label="Planificacion"),
                ft.Radio(value="Seguimiento", label="Seguimiento"),
                ft.Radio(value="Extras", label="Extras"),
                ft.Radio(value="Ausentismos", label="Ausentismos"),
                ft.Radio(value="Marcas Manuales", label="Marcas Manuales"),
            ],
        )
    )

    tipo_reporte_container = ft.Container(
        content=ft.Column(
            [
                ft.Text("Tipo de Reporte", style=ft.TextThemeStyle.HEADLINE_SMALL),
                tipo_reporte_radio_group,
            ],
        ),
        padding=10,
        border=ft.border.all(1, "black54"),
    )

    # -----------------------------------------------------------------------------
    # Sección "Cargar reportes"
    # -----------------------------------------------------------------------------
    # Comentario: Contenedor para la sección "Cargar reportes"
    # Área de carga de reportes con botones en una fila
    # Texto de encabezado
    carga_reportes_text = ft.Text("Cargar reportes:", style=ft.TextThemeStyle.HEADLINE_SMALL)

    # Función para manejar la selección de archivos
    def pick_files(e, button_text):
        nonlocal current_button
        current_button = button_text
        file_picker.pick_files(allow_multiple=False)

    # Botones para cargar reportes
    btn_nomina = ft.ElevatedButton("Nomina", on_click=lambda e: pick_files(e, "Nomina"))
    btn_reporte_planificacion = ft.ElevatedButton(
        "Reporte de planificacion", on_click=lambda e: pick_files(e, "Reporte de planificacion")
    )
    btn_reporte_trabajadas = ft.ElevatedButton(
        "Reporte de horas trabajadas", on_click=lambda e: pick_files(e, "Reporte de horas trabajadas")
    )
    btn_totales_horas_empleados = ft.ElevatedButton(
        "Totales Horas Empleados", on_click=lambda e: pick_files(e, "Totales Horas Empleados")
    )
    btn_reporte_extras = ft.ElevatedButton(
        "Reporte Horas Extras", on_click=lambda e: pick_files(e, "Reporte Horas Extras")
    )

    # Colocar los botones en una fila
    carga_reportes_row = ft.Row(
        [
            btn_nomina,
            btn_reporte_planificacion,
            btn_reporte_trabajadas,
            btn_totales_horas_empleados,
            btn_reporte_extras,
        ],
        scroll=ft.ScrollMode.AUTO,
        wrap=True,
    )

    carga_reportes_container = ft.Container(
        content=ft.Column(
            [
                carga_reportes_text,
                carga_reportes_row,
            ],
        ),
        padding=10,
        border=ft.border.all(1, "black54"),
    )

    # -----------------------------------------------------------------------------
    # Área de texto para salida informativa
    # -----------------------------------------------------------------------------
    output_text_field = ft.TextField(
        value="",
        read_only=True,
        multiline=True,
        expand=True,
        label="Salida Informativa",
    )

    # FilePicker para seleccionar archivos
    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)

    def on_file_picker_result(e):
        if current_button and e.files:
            selected_file = e.files[0]
            selected_files[current_button] = selected_file
            output_text_field.value += f"{current_button} --> cargado con: {selected_file.name}\n"
            page.update()

    file_picker.on_result = on_file_picker_result

    # Botón "Generar Reporte"
    def generar_reporte(e):
        fecha_seleccionada = fecha_text_field.value
        bd_value = ""
        # Obtener el valor seleccionado en bd_radio_group
        for radio in bd_radio_group.controls:
            if isinstance(radio, ft.Radio) and radio.checked:
                bd_value = radio.value
                break

        tipo_reporte = tipo_reporte_radio_group.value

        output_text_field.value += f"Fecha: {fecha_seleccionada}\n"
        output_text_field.value += f"Agregar reporte a base de datos: {bd_value}\n"
        output_text_field.value += f"Tipo de reporte seleccionado: {tipo_reporte}\n"

        # Mostrar archivos seleccionados
        for key, file in selected_files.items():
            output_text_field.value += f"{key}: {file.name}\n"

        output_text_field.value += "Proceso terminado\n"
        page.update()

    btn_generar_reporte = ft.ElevatedButton(
        "Generar Reporte",
        on_click=generar_reporte,
        style=ft.ButtonStyle(
            color=ft.colors.WHITE,
            bgcolor=ft.colors.BLUE,
            text_style=ft.TextStyle(weight=ft.FontWeight.BOLD),
        ),
    )

    # Construcción del layout principal
    main_layout = ft.Column(
        [
            banner,
            # Contenedores separados con comentarios
            base_datos_container,
            tipo_reporte_container,
            carga_reportes_container,
            btn_generar_reporte,
            ft.Divider(height=1, color="black54"),
            output_text_field,
        ],
        scroll=ft.ScrollMode.AUTO,
        expand=True,
    )

    page.add(main_layout)

    # Asigna el FilePicker al botón correspondiente
    page.add(file_picker)

# Ejecuta la aplicación Flet
ft.app(target=main)

