from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

OUTPUT = "wifi-guia.pptx"
TEAL = RGBColor(7, 93, 87)
INK = RGBColor(23, 35, 33)
MUTED = RGBColor(91, 108, 101)
PAPER = RGBColor(244, 246, 242)
WHITE = RGBColor(255, 255, 255)
CORAL = RGBColor(207, 84, 58)

slides = [
    ("Wi-Fi, en claro", ["Guía visual de redes inalámbricas", "Conceptos, estándares, componentes y seguridad"]),
    ("¿Qué es Wi-Fi?", ["Tecnología basada en la familia de estándares IEEE 802.11.", "Conecta dispositivos a una red local mediante ondas de radio.", "Wi-Fi no es lo mismo que Internet: la red local puede tener o no acceso a Internet."]),
    ("Estándares y generaciones", ["Wi-Fi 4: IEEE 802.11n, bandas de 2.4 y 5 GHz.", "Wi-Fi 5: IEEE 802.11ac, banda de 5 GHz.", "Wi-Fi 6 / 6E: IEEE 802.11ax; 6E añade 6 GHz donde esté permitido.", "Wi-Fi 7: IEEE 802.11be, mayor capacidad potencial.", "802.16 se asocia con WiMAX; 802.9 es una familia histórica LAN, no Wi-Fi."]),
    ("Cómo viajan los datos", ["Internet y proveedor de servicio", "Módem o terminal óptico (ONT)", "Router: dirige el tráfico entre redes.", "Punto de acceso: conecta dispositivos por radio.", "Adaptador Wi-Fi: recibe la señal en el móvil o computadora.", "La solicitud y la respuesta recorren la ruta en sentidos opuestos."]),
    ("Puntos de acceso", ["Un hotspot ofrece acceso Wi-Fi en un lugar como un aeropuerto, hotel, café o biblioteca.", "Puede requerir una contraseña o iniciar sesión en un portal.", "En redes públicas, confirma el nombre con el establecimiento y desactiva la conexión automática."]),
    ("Seguridad", ["WEP está obsoleto y no debe utilizarse.", "WPA / TKIP también debe evitarse.", "Se recomienda WPA3; WPA2-AES/CCMP es una alternativa segura.", "Usa una contraseña larga y única y cambia la clave de administración del router."]),
    ("Ventajas y límites", ["Ventajas: movilidad, menos cableado y conexión sencilla de varios equipos.", "Límites: cobertura afectada por distancia y obstáculos.", "La interferencia y la cantidad de dispositivos afectan el rendimiento.", "Wi-Fi utiliza batería; la velocidad real no equivale a la cifra teórica del estándar."]),
    ("En resumen", ["El punto de acceso crea el enlace inalámbrico; el router comunica redes.", "La cobertura depende del equipo y del entorno.", "Mantén router y dispositivos actualizados.", "Autores: Daniela Aldheni, Aline Celeste, Cesar Gabriel, Juan Aldrin y Pablo."])
]

presentation = Presentation()
presentation.slide_width = Inches(13.333)
presentation.slide_height = Inches(7.5)

for index, (title, bullets) in enumerate(slides):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = PAPER

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.18), presentation.slide_height)
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = TEAL
    stripe.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.6), Inches(1.05))
    title_frame = title_box.text_frame
    title_frame.clear()
    title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = title
    title_paragraph.font.name = "Aptos Display"
    title_paragraph.font.size = Pt(34 if index else 42)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = TEAL

    divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.82), Inches(1.8), Inches(11.65), Inches(0.035))
    divider.fill.solid()
    divider.fill.fore_color.rgb = CORAL
    divider.line.fill.background()

    body = slide.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(11.2), Inches(4.55))
    frame = body.text_frame
    frame.clear()
    frame.word_wrap = True
    for bullet_index, text in enumerate(bullets):
        paragraph = frame.paragraphs[0] if bullet_index == 0 else frame.add_paragraph()
        paragraph.text = text
        paragraph.level = 0
        paragraph.font.name = "Aptos"
        paragraph.font.size = Pt(24 if index == 0 else 21)
        paragraph.font.color.rgb = INK if bullet_index == 0 else MUTED
        paragraph.space_after = Pt(17)
        if index != 0:
            paragraph.text = "•  " + text

    footer = slide.shapes.add_textbox(Inches(10.9), Inches(7.05), Inches(1.4), Inches(0.2))
    footer_frame = footer.text_frame
    footer_frame.clear()
    footer_paragraph = footer_frame.paragraphs[0]
    footer_paragraph.text = f"{index + 1:02d} / {len(slides):02d}"
    footer_paragraph.alignment = PP_ALIGN.RIGHT
    footer_paragraph.font.name = "Aptos"
    footer_paragraph.font.size = Pt(10)
    footer_paragraph.font.color.rgb = MUTED

presentation.save(OUTPUT)
print(f"Created {OUTPUT} ({len(slides)} slides)")
