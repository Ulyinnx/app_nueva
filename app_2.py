import flet as ft

def main(page: ft.Page):
    #----------------------------------------------------------------------
    # CONFIGURACION BASICA DE LA PAGINA ###################################
    # ---------------------------------------------------------------------

    # Modo de tema inicial
    page.theme_mode = ft.ThemeMode.DARK
    # Función para alternar el modo de tema -------------------------
    def toggle_theme(e):
        if page.theme_mode == ft.ThemeMode.LIGHT:
            page.theme_mode = ft.ThemeMode.DARK
            theme_icon_button.icon = ft.icons.WB_SUNNY  # Ícono de sol
        else:
            page.theme_mode = ft.ThemeMode.LIGHT
            theme_icon_button.icon = ft.icons.BRIGHTNESS_3  # Ícono de luna
        page.update()

    # Titulo de app --
    page.title = "Panel de gestión - DIMENSIONAMIENTO"

    #----------------------------------------------------------------------
    # ENCABEZADO - HEADER #################################################
    # ---------------------------------------------------------------------

    # Botón para alternar el tema
    theme_icon_button = ft.IconButton(
        icon=ft.icons.WB_SUNNY,
        on_click=toggle_theme,
        tooltip="Alternar tema claro/oscuro",


    )

    header = ft.Container(
        content=ft.Row(
            controls=[
                ft.Text("PANEL DE DIMENSIONAMIENTO", style=ft.TextThemeStyle.TITLE_MEDIUM),
                theme_icon_button
            ],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            height=40,
        ),
        bgcolor=ft.colors.SURFACE_VARIANT,
        expand=True,
        border_radius=10,
        padding=10,
    )


    # Función para mostrar el contenido de cada sección
    def show_content(content):
        container_content.content = ft.Text(content, size=18)
        page.update()

    # Crear el menú como un contenedor
    menu_izq = ft.Container(
        content=ft.Column(
            controls=[
                ft.TextButton(text='Primer reporte',  icon=ft.icons.NAVIGATE_NEXT),
                ft.TextButton(text='Segundo reporte',  icon=ft.icons.NAVIGATE_NEXT),
                ft.TextButton(text='Tercer reporte',  icon=ft.icons.NAVIGATE_NEXT),
                ft.TextButton(text='Cuarto reporte',  icon=ft.icons.NAVIGATE_NEXT),
            ],
        ),
        width=300,
        bgcolor=ft.colors.SURFACE_VARIANT,
        padding=10,
        border_radius=10,
    )





    # # Definición de pestañas
    # tabs = ft.Tabs(
    #     selected_index=0,
    #     animation_duration=300,
    #     height=25,
    #     padding=20,
    #     tabs=[
    #         ft.Tabs(expand=True),
    #         ft.Tab(text="Mi actividad"),
    #         ft.Tab(text="Mis informaciones"),
    #         ft.Tab(text="Mis grupos"),
    #         ft.Tab(text="Preferencias"),
    #     ],
    #     indicator_color=ft.colors.GREY,
    #     on_change=lambda e: show_content(
    #         f"Mostrando: {e.control.tabs[e.control.selected_index].text}",
    #     ),
    # )

    # Contenedor para mostrar el contenido según la sección seleccionada
    container_content = ft.Container(
        content=ft.Text("Mi actividad - Datos iniciales", size=18),
        alignment=ft.alignment.center,
        expand=True,
        padding=20,
        bgcolor=ft.colors.BACKGROUND,
    )

    # Diseño principal que organiza el menú y el contenido
    # layout = ft.Row(
    #     controls=[
    #         menu,
    #         ft.VerticalDivider(width=1),
    #         ft.Column(
    #             controls=[header, tabs, container_content],
    #             expand=True,
    #         ),
    #     ],
    #     expand=True,
    # )

    layout = ft.Column(
         controls=[
             header,
             ft.Row(controls=[
                 menu_izq,
                 # ft.VerticalDivider(width=50),
                  container_content
                 ]
             )
         ])

    # Añadir el layout a la página
    page.add(layout)

# Iniciar la app de Flet
ft.app(target=main)
