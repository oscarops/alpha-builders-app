import os
import tkinter as tk
from tkinter import messagebox, ttk
from tksheet import Sheet

# Intentar importar ReportLab para la generación de reportes en PDF
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


class CalculadoraRendimientoApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Alpha Builders - Control de Rendimiento y Productividad")
        self.geometry("1050x650")
        self.configure(bg="#1e1e2e")

        # Configuración de estilos para Tkinter
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TFrame", background="#1e1e2e")
        self.style.configure(
            "TLabel", background="#1e1e2e", foreground="#cdd6f4", font=("Segoe UI", 10)
        )
        self.style.configure(
            "Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#89b4fa"
        )
        self.style.configure(
            "TButton", font=("Segoe UI", 10, "bold"), padding=6, background="#313244", foreground="#cdd6f4"
        )
        self.style.map(
            "TButton",
            background=[("active", "#45475a")],
            foreground=[("active", "#ffffff")],
        )

        self._crear_interfaz()

    def _crear_interfaz(self):
        # Panel Superior - Encabezado
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", px=20, py=15)

        title_label = ttk.Label(
            header_frame,
            text="ALPHA BUILDERS | Control de Rendimiento de Mano de Obra",
            style="Header.TLabel",
        )
        title_label.pack(side="left")

        # Contenedor para la tabla interactiva (tksheet)
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, px=20, py=10)

        # Encabezados de columna
        headers = [
            "Actividad / Rubro",
            "Unidad",
            "Cant. Ejecutada",
            "Horas-Hombre (HH)",
            "Rend. Real (HH/Unid)",
            "Rend. Teórico",
            "Estado / Diagnóstico",
        ]

        # Datos de ejemplo iniciales
        datos_iniciales = [
            ["Excavación manual", "m3", 12.5, 10.0, "", 0.85, ""],
            ["Hormigón en zapatas", "m3", 8.0, 16.0, "", 1.80, ""],
            ["Mampostería de ladrillo", "m2", 25.0, 30.0, "", 1.10, ""],
            ["Enlucido interior", "m2", 40.0, 28.0, "", 0.75, ""],
        ]

        # Configuración del componente Sheet
        self.sheet = Sheet(
            table_frame,
            data=datos_iniciales,
            headers=headers,
            header_bg="#313244",
            header_fg="#cdd6f4",
            grid_color="#45475a",
            bg="#181825",
            fg="#cdd6f4",
            table_selected_cells_bg="#45475a",
            table_selected_cells_fg="#ffffff",
            show_row_index=True,
        )
        self.sheet.enable_bindings(
            (
                "single_select",
                "row_select",
                "column_width_resize",
                "double_click_column_resize",
                "arrow_keys",
                "row_height_resize",
                "edit_cell",
                "copy",
                "paste",
                "delete",
                "undo",
            )
        )
        self.sheet.pack(fill="both", expand=True)

        # Anchos de columna personalizados
        self.sheet.column_width(0, 220)
        self.sheet.column_width(1, 80)
        self.sheet.column_width(2, 130)
        self.sheet.column_width(3, 140)
        self.sheet.column_width(4, 150)
        self.sheet.column_width(5, 120)
        self.sheet.column_width(6, 160)

        # Panel Inferior - Botones de acción
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", px=20, py=15)

        btn_agregar = ttk.Button(
            btn_frame, text="➕ Agregar Fila", command=self.agregar_fila
        )
        btn_agregar.pack(side="left", padx=5)

        btn_calcular = ttk.Button(
            btn_frame, text="⚡ Calcular Rendimientos", command=self.calcular_rendimientos
        )
        btn_calcular.pack(side="left", padx=5)

        btn_exportar = ttk.Button(
            btn_frame, text="📄 Exportar Reporte PDF", command=self.exportar_pdf
        )
        btn_exportar.pack(side="right", padx=5)

    def agregar_fila(self):
        """Agrega una nueva fila vacía a la tabla."""
        self.sheet.insert_row(["Nueva Actividad", "m2", 0.0, 0.0, "", 1.0, ""])

    def calcular_rendimientos(self):
        """Calcula el rendimiento real y evalúa el desempeño contra el teórico."""
        data = self.sheet.get_sheet_data()

        for i, row in enumerate(data):
            try:
                cantidad = float(row[2]) if row[2] not in ("", None) else 0.0
                hh = float(row[3]) if row[3] not in ("", None) else 0.0
                rend_teorico = float(row[5]) if row[5] not in ("", None) else 0.0

                if cantidad > 0:
                    rend_real = round(hh / cantidad, 3)
                    self.sheet.set_cell_data(i, 4, rend_real)

                    # Diagnóstico
                    if rend_teorico > 0:
                        if rend_real <= rend_teorico:
                            diag = "EFICIENTE ✅"
                        else:
                            diag = "EXCESO DE HH ⚠️"
                    else:
                        diag = "SIN REF."
                    self.sheet.set_cell_data(i, 6, diag)
                else:
                    self.sheet.set_cell_data(i, 4, "N/A")
                    self.sheet.set_cell_data(i, 6, "INCOMPLETO")

            except ValueError:
                self.sheet.set_cell_data(i, 4, "ERROR")
                self.sheet.set_cell_data(i, 6, "DATOS INVÁLIDOS")

        self.sheet.refresh()
        messagebox.showinfo("Cálculo Completado", "Los rendimientos y diagnósticos han sido actualizados.")

    def exportar_pdf(self):
        """Genera un informe PDF con el membrete de Alpha Builders."""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(
                "Librería Faltante",
                "Para exportar en PDF debes instalar ReportLab:\npip install reportlab",
            )
            return

        # Asegurar cálculos antes de exportar
        self.calcular_rendimientos()

        filename = "Reporte_Rendimiento_AlphaBuilders.pdf"
        doc = SimpleDocTemplate(
            filename,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=36,
            bottomMargin=36,
        )
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            "DocTitle",
            parent=styles["Heading1"],
            fontSize=16,
            textColor=colors.HexColor("#1e1e2e"),
            alignment=1,
            spaceAfter=12,
        )

        elements = []
        elements.append(Paragraph("<b>ALPHA BUILDERS</b>", title_style))
        elements.append(Paragraph("Reporte Ejecutivo de Rendimiento de Mano de Obra", styles["Normal"]))
        elements.append(Spacer(1, 15))

        # Encabezados y datos para la tabla PDF
        headers_pdf = [
            "Actividad",
            "Unid.",
            "Cant.",
            "HH",
            "Rend. Real",
            "Rend. Teór.",
            "Estado",
        ]
        tabla_datos = [headers_pdf] + self.sheet.get_sheet_data()

        t = Table(tabla_datos, colWidths=[150, 45, 55, 55, 75, 75, 90])
        t.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#313244")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 9),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#a6adc8")),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 1), (-1, -1), 8),
                ]
            )
        )

        elements.append(t)

        try:
            doc.build(elements)
            messagebox.showinfo(
                "Exportación Exitosa",
                f"El reporte en PDF se ha generado correctamente:\n{os.path.abspath(filename)}",
            )
        except Exception as e:
            messagebox.showerror("Error de Exportación", f"No se pudo crear el PDF: {e}")


if __name__ == "__main__":
    app = CalculadoraRendimientoApp()
    app.mainloop()