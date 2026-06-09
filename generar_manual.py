from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ESPRESSO = RGBColor(0x4A, 0x37, 0x28)
GOLD     = RGBColor(0xC8, 0xA9, 0x6E)
MOCHA    = RGBColor(0x8B, 0x73, 0x55)
SAGE     = RGBColor(0x9C, 0xAF, 0x88)
CREAM    = RGBColor(0xFA, 0xF7, 0xF2)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
DARK     = RGBColor(0x2C, 0x18, 0x10)
GRAY     = RGBColor(0xA0, 0x80, 0x70)


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, color in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if color:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), '6')
            el.set(qn('w:color'), color)
            tcBorders.append(el)
    tcPr.append(tcBorders)


def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18 if level == 1 else 12)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size = Pt(18)
        run.font.color.rgb = ESPRESSO
    elif level == 2:
        run.font.size = Pt(13)
        run.font.color.rgb = MOCHA
    else:
        run.font.size = Pt(11)
        run.font.color.rgb = GOLD
    run.font.name = 'Calibri'
    return p


def body(doc, text, color=None, bold=False, size=10.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = 'Calibri'
    run.bold = bold
    run.font.color.rgb = color or DARK
    return p


def step_table(doc, steps):
    """Tabla numerada de pasos con fondo alternado."""
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    table.columns[0].width = Cm(1.4)
    table.columns[1].width = Cm(14)

    for i, (title, detail) in enumerate(steps, 1):
        row = table.add_row()
        # Numero
        num_cell = row.cells[0]
        set_cell_bg(num_cell, '4A3728')
        np = num_cell.paragraphs[0]
        np.alignment = WD_ALIGN_PARAGRAPH.CENTER
        nr = np.add_run(str(i))
        nr.font.bold = True
        nr.font.size = Pt(13)
        nr.font.color.rgb = GOLD
        nr.font.name = 'Calibri'

        # Contenido
        txt_cell = row.cells[1]
        bg = 'FFFCF8' if i % 2 == 0 else 'FAF7F2'
        set_cell_bg(txt_cell, bg)
        tp = txt_cell.paragraphs[0]
        tr = tp.add_run(title)
        tr.bold = True
        tr.font.size = Pt(10.5)
        tr.font.color.rgb = ESPRESSO
        tr.font.name = 'Calibri'
        if detail:
            tp.add_run('\n')
            dr = tp.add_run(detail)
            dr.font.size = Pt(9.5)
            dr.font.color.rgb = GRAY
            dr.font.name = 'Calibri'

    doc.add_paragraph()


def code_block(doc, code):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Cm(1)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(8)
    run = p.add_run(code)
    run.font.name = 'Courier New'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x4A, 0x37, 0x28)
    # fondo gris claro via shading en el párrafo
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), 'F0EAE0')
    pPr.append(shd)
    return p


def tip_box(doc, text, icon='💡'):
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.rows[0].cells[0]
    set_cell_bg(cell, 'E8F5E0')
    set_cell_border(cell, top='9CAF88', bottom='9CAF88', left='9CAF88', right='9CAF88')
    p = cell.paragraphs[0]
    r = p.add_run(f'{icon}  {text}')
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x3A, 0x5A, 0x2A)
    r.font.name = 'Calibri'
    doc.add_paragraph()


def warning_box(doc, text):
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.rows[0].cells[0]
    set_cell_bg(cell, 'FFF3CD')
    p = cell.paragraphs[0]
    r = p.add_run(f'⚠️  {text}')
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x7A, 0x5A, 0x00)
    r.font.name = 'Calibri'
    doc.add_paragraph()


def divider(doc):
    p = doc.add_paragraph('─' * 72)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    run = p.runs[0]
    run.font.color.rgb = RGBColor(0xE8, 0xDD, 0xD0)
    run.font.size = Pt(8)


# ─────────────────────────────────────────────
doc = Document()

# Márgenes
for section in doc.sections:
    section.top_margin    = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin   = Cm(2.8)
    section.right_margin  = Cm(2.8)

# ── PORTADA ──────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(40)
r = p.add_run('💝 Valente Studio')
r.font.size = Pt(28)
r.font.bold = True
r.font.color.rgb = ESPRESSO
r.font.name = 'Calibri'

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run('Manual: Cómo crear y publicar una aplicación web')
r2.font.size = Pt(14)
r2.font.color.rgb = MOCHA
r2.font.name = 'Calibri'

p3 = doc.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = p3.add_run('Desde cero hasta PWA en iPhone · Junio 2026')
r3.font.size = Pt(10)
r3.font.color.rgb = GRAY
r3.font.name = 'Calibri'
r3.italic = True

doc.add_paragraph()
divider(doc)
doc.add_paragraph()

# ── RESUMEN ───────────────────────────────────
heading(doc, '¿Qué logramos?', 2)
body(doc,
     'Creamos una aplicación web completa (Inventario Wishlist) con diseño elegante, '
     'la separamos en archivos profesionales, la subimos a GitHub y la instalamos '
     'como app nativa en el iPhone — sin App Store ni costo alguno.')

doc.add_paragraph()

# ── SECCIÓN 1 ─────────────────────────────────
heading(doc, '1. Preparar el archivo HTML original')
body(doc, 'El punto de partida fue un único archivo HTML con todo incluido (estilos, lógica y estructura).', color=GRAY)

step_table(doc, [
    ('Tener el archivo .html listo',
     'En este caso: wishlistinventory.html con HTML + CSS + JavaScript todo junto'),
    ('Subir el archivo a Claude Code',
     'Desde la interfaz de Claude Code, usar el ícono de adjunto o arrastrar el archivo al chat'),
    ('Escribir la instrucción',
     '"Sube este archivo a GitHub" — Claude lo lee, lo copia al repositorio y hace el commit'),
])

tip_box(doc, 'Claude Code puede leer archivos HTML, imágenes, PDFs y más. Solo arrástralos al chat.')

# ── SECCIÓN 2 ─────────────────────────────────
heading(doc, '2. Separar el código en archivos profesionales')
body(doc, 'Un buen proyecto separa el código por responsabilidad. Le pedimos a Claude que dividiera el archivo en 4:', color=GRAY)

step_table(doc, [
    ('Escribir la instrucción a Claude',
     '"Crea HTML, CSS, JavaScript y Python por separado"'),
    ('index.html — Estructura',
     'Solo el esqueleto HTML. Enlaza styles.css y script.js con etiquetas <link> y <script src>'),
    ('styles.css — Diseño',
     'Todos los colores, fuentes, animaciones y layout. Variables CSS, tarjetas, modal, filtros'),
    ('script.js — Lógica',
     'CRUD completo, filtros, búsqueda, exportar Excel, drag & drop de imágenes, localStorage'),
    ('app.py — Servidor Python',
     'Servidor HTTP con API REST (GET/POST/PUT/DELETE) para gestionar productos en products.json'),
])

tip_box(doc, 'Separar el código facilita editar cada parte sin tocar las demás. Es la práctica estándar en desarrollo web.')

# ── SECCIÓN 3 ─────────────────────────────────
heading(doc, '3. Subir el proyecto a GitHub')
body(doc, 'GitHub es donde vive el código. Claude manejó todos los comandos git automáticamente.', color=GRAY)

step_table(doc, [
    ('Claude hace el commit y push automáticamente',
     'Detecta los archivos nuevos, los agrega con git add, crea el mensaje y hace git push'),
    ('Rama de trabajo: claude/github-file-upload-why5xd',
     'Cada proyecto puede tener su propia rama. La rama main es la principal/estable'),
    ('Verificar en GitHub',
     'Ir a github.com/rebecanatasharamirezjimenez-design/Valente-studio y confirmar que aparecen los archivos'),
])

body(doc, 'Comandos que ejecutó Claude internamente:', color=GRAY, size=9.5)
code_block(doc,
    'git add index.html styles.css script.js app.py\n'
    'git commit -m "Separar archivos: HTML, CSS, JS, Python"\n'
    'git push -u origin claude/github-file-upload-why5xd'
)

# ── SECCIÓN 4 ─────────────────────────────────
heading(doc, '4. Crear la rama gh-pages para publicar en internet')
body(doc, 'GitHub Pages es un servicio gratuito que convierte tu repositorio en un sitio web público.', color=GRAY)

step_table(doc, [
    ('Claude crea la rama gh-pages automáticamente',
     'Usando la API de GitHub, crea la rama desde la rama de trabajo con todos los archivos'),
    ('Activar GitHub Pages (solo la primera vez)',
     'En GitHub → Settings → Pages → Branch: gh-pages → / (root) → Save'),
    ('Esperar 1-2 minutos',
     'GitHub compila y despliega. Aparece el mensaje "Your site is live at..."'),
    ('Tu URL pública queda lista',
     'https://rebecanatasharamirezjimenez-design.github.io/Valente-studio/'),
])

tip_box(doc, 'GitHub Pages es gratis para repositorios públicos. Ideal para apps sin backend.', icon='🌐')
warning_box(doc, 'Si el repositorio es privado, GitHub Pages requiere plan de pago.')

# ── SECCIÓN 5 ─────────────────────────────────
heading(doc, '5. Instalar como app en el iPhone (PWA)')
body(doc, 'Una PWA (Progressive Web App) se instala desde Safari y funciona como app nativa — sin App Store.', color=GRAY)

step_table(doc, [
    ('Abrir Safari en el iPhone',
     'Importante: debe ser Safari, no Chrome ni otro navegador, para poder instalar'),
    ('Ir a la URL de GitHub Pages',
     'https://rebecanatasharamirezjimenez-design.github.io/Valente-studio/'),
    ('Tocar el ícono de compartir',
     'El cuadro con flecha hacia arriba en la barra inferior de Safari'),
    ('Seleccionar "Añadir a pantalla de inicio"',
     'Buscar la opción en el menú que aparece desde abajo'),
    ('Confirmar el nombre y tocar "Añadir"',
     'Puedes editar el nombre que aparecerá debajo del ícono'),
    ('Abrir desde la pantalla de inicio',
     'Ya funciona como app: pantalla completa, sin barra de Safari, con ícono propio'),
])

tip_box(doc, 'En Android también funciona: Chrome → menú (⋮) → "Añadir a pantalla de inicio"', icon='📱')

# ── SECCIÓN 6 ─────────────────────────────────
heading(doc, '6. Dónde se guardan los datos')

all_rows_data = [
    ('Situación',           'Resultado', True),
    ('Cierras y abres la app', '✅ Datos conservados (localStorage del iPhone)', False),
    ('Reinicias el iPhone',    '✅ Datos conservados', False),
    ('Eliminas el ícono',      '❌ Datos eliminados permanentemente', False),
    ('iOS limpia caché',       '⚠️ Posible pérdida si no usas la app por semanas', False),
]
table = doc.add_table(rows=len(all_rows_data), cols=2)
table.style = 'Table Grid'

for i, (col1, col2, is_header) in enumerate(all_rows_data):
    bg = '4A3728' if is_header else ('FAF7F2' if i % 2 == 0 else 'FFFCF8')
    for j, text in enumerate([col1, col2]):
        cell = table.rows[i].cells[j]
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.font.bold = is_header
        r.font.color.rgb = GOLD if is_header else DARK
        r.font.size = Pt(10 if is_header else 9.5)
        r.font.name = 'Calibri'

doc.add_paragraph()
tip_box(doc, 'Exporta a Excel frecuentemente como respaldo. El botón ya está en la app.', icon='💾')

# ── SECCIÓN 7 ─────────────────────────────────
heading(doc, '7. Para aplicaciones futuras — Guía rápida')
body(doc, 'Estos son los pasos exactos para repetir el proceso con cualquier nueva app:', color=GRAY)

step_table(doc, [
    ('Crear o subir el archivo de diseño a Claude Code',
     'HTML, imagen de mockup, descripción en texto — Claude entiende cualquier formato'),
    ('Pedir la separación de archivos',
     '"Separa esto en HTML, CSS, JavaScript y Python por separado"'),
    ('"Súbelo a GitHub"',
     'Claude hace el commit y push automáticamente a la rama correcta'),
    ('"Crea la rama gh-pages"',
     'Claude la crea con un comando a la API de GitHub'),
    ('Activar GitHub Pages (Settings → Pages → gh-pages)',
     'Solo se hace una vez por repositorio'),
    ('Abrir en Safari → Compartir → Añadir a pantalla de inicio',
     'Listo: tienes tu app instalada en el iPhone'),
])

# ── SECCIÓN 8 ─────────────────────────────────
heading(doc, '8. Próximos pasos opcionales')

mejoras = [
    ('Base de datos en la nube',
     'Conectar a Firebase o Supabase para que los datos sean permanentes y accesibles desde cualquier dispositivo'),
    ('Notificaciones push',
     'Agregar un service worker para recibir alertas aunque la app esté cerrada'),
    ('Modo offline completo',
     'Con un service worker la app funciona sin internet'),
    ('Dominio personalizado',
     'Cambiar la URL por valentestudio.com en vez de github.io'),
    ('Múltiples usuarios',
     'Sistema de login para que varias personas accedan a listas diferentes'),
]

for titulo, detalle in mejoras:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(titulo + ': ')
    r1.bold = True
    r1.font.color.rgb = ESPRESSO
    r1.font.size = Pt(10)
    r1.font.name = 'Calibri'
    r2 = p.add_run(detalle)
    r2.font.color.rgb = GRAY
    r2.font.size = Pt(10)
    r2.font.name = 'Calibri'

doc.add_paragraph()
divider(doc)

# ── PIE ───────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Valente Studio · Creado con Claude Code · Junio 2026')
r.font.size = Pt(8.5)
r.font.color.rgb = GRAY
r.font.name = 'Calibri'
r.italic = True

# ── GUARDAR ───────────────────────────────────
out = '/home/user/Valente-studio/Manual_Valente_Studio.docx'
doc.save(out)
print(f'Documento guardado: {out}')
