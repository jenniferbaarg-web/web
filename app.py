from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)
app.secret_key = 'clave_secreta_para_flash_messages'

# ============================================================
# 1. CONFIGURACIÓN (cargar desde Excel o fallback)
# ============================================================

def load_config():
    try:
        df = pd.read_excel('configurar.xlsx')
        config = df.iloc[0].to_dict()
        
        # Limpiar número de WhatsApp
        if 'tel_whatsapp' in config:
            tel = config['tel_whatsapp']
            if isinstance(tel, float):
                tel = int(tel)
            tel = str(tel)
            tel = ''.join(filter(str.isdigit, tel))
            if tel.startswith('0'):
                tel = tel[1:]
            config['tel_whatsapp'] = tel
        else:
            config['tel_whatsapp'] = '2257401281'
        
        print(f"📱 WhatsApp cargado: {config['tel_whatsapp']}")
        return config
    except Exception as e:
        print(f"❌ Error al cargar configurar.xlsx: {e}")
        return {
            'titulo': 'Jennifer Gauna',
            'subtitulo': 'Estudio Inmobiliario',
            'direccion_cc': 'Calle 2 , Lomas de burro',
            'direccion_su1': 'Gana nº 28, San Cayetano',
            'direccion_suc2': 'Av. Libertador nº 362, San Isidro',
            'tel_cc': '2000-401281',
            'tel_suc1': '2000-401282',
            'tel_suc2': '2000-401283',
            'mail_contacto': 'contacto@jennifer.com',
            'mail_tasaciones': 'tasaciones@jennifer.com',
            'tel_whatsapp': '2000401281',
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'smtp_user': '',
            'smtp_password': ''
        }

config = load_config()

# ============================================================
# 2. INYECTAR CONFIGURACIÓN EN PLANTILLAS
# ============================================================

@app.context_processor
def inject_config():
    return dict(config=config)

# ============================================================
# 3. CARGA DE DATOS DESDE GOOGLE SHEETS
# ============================================================

# --- URLs de Google Sheets ---
URL_PROPIEDADES = "https://docs.google.com/spreadsheets/d/1YH2_Q8fvimzMQgRDDIrBbI-HziJ270I3KLgA8xf1l9g/edit?usp=sharing"
CSV_PROPIEDADES = URL_PROPIEDADES.replace("/edit?usp=sharing", "/export?format=csv")

# URL para emprendimientos (la que me proporcionaste)
URL_EMPRENDIMIENTOS = "https://docs.google.com/spreadsheets/d/1PfSkmk28Wec7LSWhqx1siQqGqd97KaCt/edit?usp=sharing"
CSV_EMPRENDIMIENTOS = URL_EMPRENDIMIENTOS.replace("/edit?usp=sharing", "/export?format=csv")


def load_emprendimientos():
    """Carga datos de emprendimientos desde Google Sheets."""
    try:
        df = pd.read_csv(CSV_EMPRENDIMIENTOS)
        df.columns = df.columns.str.strip()
        
        columnas = ['id_propiedad', 'fecha', 'tipo_op', 'tipo_inm', 'direccion', 'barrio',
                    'precio', 'moneda', 'sup_total', 'sup_cubierta', 'ambientes', 'dormitorios',
                    'banos', 'cocheras', 'antiguedad', 'expensas', 'estado', 'status_com',
                    'detalle', 'link_foto_inicio', 'galeria', 'propietario', 'tel_propietario']
        for col in columnas:
            if col not in df.columns:
                df[col] = ''
        
        columnas_numericas = ['precio', 'sup_total', 'sup_cubierta', 'ambientes', 
                              'dormitorios', 'banos', 'cocheras', 'antiguedad', 'expensas']
        for col in columnas_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        df['id_propiedad'] = df['id_propiedad'].astype(str)
        return df.fillna('')
    except Exception as e:
        print(f"❌ Error al cargar emprendimientos: {e}")
        # Datos de ejemplo (opcional)
        return pd.DataFrame([
            {'id_propiedad': 'E001', 'tipo_op': 'Venta', 'tipo_inm': 'Departamento', 
             'direccion': 'Av. Principal 123', 'barrio': 'Centro', 'precio': 150000, 'moneda': 'U$S',
             'link_foto_inicio': '/static/imagenes_emp/emp001.jpg', 'detalle': 'Excelente oportunidad'},
        ])
def load_propiedades():
    try:
        df = pd.read_csv(CSV_PROPIEDADES)
        df.columns = df.columns.str.strip()
        columnas = ['id_propiedad', 'fecha', 'tipo_op', 'tipo_inm', 'direccion', 'barrio',
                    'precio', 'moneda', 'sup_total', 'sup_cubierta', 'ambientes', 'dormitorios',
                    'banos', 'cocheras', 'antiguedad', 'expensas', 'estado', 'status_com',
                    'link_foto_inicio', 'propietario', 'tel_propietario', 'detalle', 'galeria']
        for col in columnas:
            if col not in df.columns:
                df[col] = ''
        
        # Convertir columnas numéricas a enteros (manejando valores nulos)
        columnas_numericas = ['precio', 'sup_total', 'sup_cubierta', 'ambientes', 
                              'dormitorios', 'banos', 'cocheras', 'antiguedad', 'expensas']
        for col in columnas_numericas:
            if col in df.columns:
                # Reemplazar valores vacíos o no numéricos con 0 y convertir a int
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        # id_propiedad debe ser string para las URLs
        df['id_propiedad'] = df['id_propiedad'].astype(str)
        
        return df.fillna('')
    except Exception as e:
        print(f"❌ Error al cargar propiedades: {e}")
        return pd.DataFrame(columns=['id_propiedad', 'tipo_op', 'tipo_inm', 'direccion', 'barrio',
                                     'precio', 'moneda', 'link_foto_inicio', 'detalle', 'galeria'])

df_propiedades = load_propiedades()
df_emprendimientos = load_emprendimientos()

# ============================================================
# 4. FUNCIÓN PARA ENVIAR CORREOS
# ============================================================

def send_email(destinatario, asunto, mensaje_html, mensaje_plain=None):
    if not destinatario:
        print("❌ Destinatario vacío")
        return False

    if '@' not in destinatario or '.' not in destinatario:
        print(f"❌ '{destinatario}' no es un email válido")
        return False

    if mensaje_plain is None:
        mensaje_plain = "Gracias por tu consulta. Te responderemos a la brevedad."

    if not config.get('smtp_user') or not config.get('smtp_password'):
        print(f"⚠️ SMTP no configurado. Simulando envío a {destinatario}")
        return True

    msg = MIMEMultipart('alternative')
    msg['From'] = config['smtp_user']
    msg['To'] = destinatario
    msg['Subject'] = asunto

    part1 = MIMEText(mensaje_plain, 'plain')
    part2 = MIMEText(mensaje_html, 'html')
    msg.attach(part1)
    msg.attach(part2)

    try:
        with smtplib.SMTP(config['smtp_server'], int(config['smtp_port'])) as server:
            server.starttls()
            server.login(config['smtp_user'], config['smtp_password'])
            server.sendmail(config['smtp_user'], destinatario, msg.as_string())
        print(f"✅ Correo enviado a {destinatario} - Asunto: {asunto}")
        return True
    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")
        return False

# ============================================================
# 5. RUTAS DE LA APLICACIÓN
# ============================================================

@app.route('/')
def index():
    ubicacion = request.args.get('ubicacion', 'Todas')
    operacion = request.args.get('operacion', 'Todas')
    tipo = request.args.get('tipo', 'Todos')

    df_filtrado = df_propiedades.copy()
    if ubicacion != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['barrio'].str.upper() == ubicacion]
    if operacion != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['tipo_op'].str.upper() == operacion]
    if tipo != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['tipo_inm'].str.upper() == tipo]

    ubicaciones = ['Todas'] + sorted(df_propiedades['barrio'].dropna().str.upper().unique().tolist())
    operaciones = ['Todas'] + sorted(df_propiedades['tipo_op'].dropna().str.upper().unique().tolist())
    tipos = ['Todos'] + sorted(df_propiedades['tipo_inm'].dropna().str.upper().unique().tolist())

    return render_template('index.html',
                           propiedades=df_filtrado.to_dict('records'),
                           ubicaciones=ubicaciones,
                           operaciones=operaciones,
                           tipos=tipos,
                           ubicacion_seleccionada=ubicacion,
                           operacion_seleccionada=operacion,
                           tipo_seleccionado=tipo)

@app.route('/emprendimientos')
def emprendimientos():
    emprendimientos = df_emprendimientos.to_dict('records')
    return render_template('emprendimientos.html', emprendimientos=emprendimientos)

@app.route('/propiedad/<id_propiedad>')
def detalle_propiedad(id_propiedad):
    df = df_propiedades.copy()
    propiedad = df[df['id_propiedad'].astype(str) == str(id_propiedad)]
    
    if propiedad.empty:
        return render_template('404.html'), 404
    
    prop = propiedad.iloc[0].to_dict()
    
    # Obtener imágenes de la galería desde la columna 'galeria' (separadas por comas)
    imagenes_galeria = []
    if prop.get('galeria'):
        # Dividir por comas, eliminar espacios en blanco y filtrar vacíos
        imagenes_galeria = [url.strip() for url in prop['galeria'].split(',') if url.strip()]
    
    # Si no hay imágenes en 'galeria', usar link_foto_inicio como fallback
    if not imagenes_galeria and prop.get('link_foto_inicio'):
        imagenes_galeria = [prop['link_foto_inicio']]
    
    print(f"📸 Imágenes en galería: {len(imagenes_galeria)}")
    
    # Propiedades relacionadas (mismo barrio o tipo)
    relacionadas = df[(df['barrio'] == prop['barrio']) | (df['tipo_inm'] == prop['tipo_inm'])].head(4).to_dict('records')
    
    return render_template('detalle.html',
                           propiedad=prop,
                           relacionadas=relacionadas,
                           imagenes_galeria=imagenes_galeria)

@app.route('/tasaciones', methods=['GET', 'POST'])
def tasaciones():
    if request.method == 'POST':
        nombre = request.form.get('nombre')
        email = request.form.get('email')
        telefono = request.form.get('telefono')
        horario = request.form.get('horario')
        direccion = request.form.get('direccion')
        operacion = request.form.get('operacion')
        tipo_propiedad = request.form.get('tipo_propiedad')
        ambientes = request.form.get('ambientes')
        sup_cubierta = request.form.get('sup_cubierta')
        sup_total = request.form.get('sup_total')
        garage = request.form.get('garage')
        amenities = request.form.get('amenities')
        observaciones = request.form.get('observaciones')
        copia = request.form.get('copia') == 'on'

        mensaje_html = f"""
        <h2>Nueva solicitud de tasación</h2>
        <p><strong>Nombre:</strong> {nombre}</p>
        <p><strong>Email:</strong> {email}</p>
        <p><strong>Teléfono:</strong> {telefono}</p>
        <p><strong>Horario de contacto:</strong> {horario}</p>
        <p><strong>Dirección:</strong> {direccion}</p>
        <p><strong>Operación:</strong> {operacion}</p>
        <p><strong>Tipo de propiedad:</strong> {tipo_propiedad}</p>
        <p><strong>Ambientes:</strong> {ambientes}</p>
        <p><strong>Sup. cubierta:</strong> {sup_cubierta} m²</p>
        <p><strong>Sup. total:</strong> {sup_total} m²</p>
        <p><strong>Garage:</strong> {garage}</p>
        <p><strong>Amenities:</strong> {amenities}</p>
        <p><strong>Observaciones:</strong> {observaciones}</p>
        """

        success = send_email(config['mail_tasaciones'], "Nueva solicitud de tasación", mensaje_html)
        if copia and email:
            send_email(email, "Copia de tu solicitud de tasación", mensaje_html, "Gracias por tu solicitud. Te contactaremos pronto.")
        if success:
            flash('¡Tu solicitud de tasación fue enviada con éxito!', 'success')
        else:
            flash('Hubo un error al enviar la solicitud. Intenta nuevamente.', 'danger')
        return redirect(url_for('tasaciones'))
    return render_template('tasaciones.html')

@app.route('/contacto', methods=['GET', 'POST'])
def contacto():
    if request.method == 'POST':
        nombre = request.form.get('nombre')
        email = request.form.get('email')
        telefono = request.form.get('telefono')
        mensaje = request.form.get('mensaje')
        copia = request.form.get('copia') == 'on'

        mensaje_html = f"""
        <h2>Nuevo mensaje de contacto</h2>
        <p><strong>Nombre:</strong> {nombre}</p>
        <p><strong>Email:</strong> {email}</p>
        <p><strong>Teléfono:</strong> {telefono}</p>
        <p><strong>Mensaje:</strong> {mensaje}</p>
        """

        success = send_email(config['mail_contacto'], "Nuevo mensaje de contacto", mensaje_html)
        if copia and email:
            send_email(email, "Copia de tu mensaje de contacto", mensaje_html, "Gracias por contactarnos. Te responderemos pronto.")
        if success:
            flash('¡Tu mensaje fue enviado con éxito!', 'success')
        else:
            flash('Hubo un error al enviar el mensaje. Intenta nuevamente.', 'danger')
        return redirect(url_for('contacto'))
    return render_template('contacto.html')

@app.route('/emprendimiento/<id_propiedad>')
def detalle_emprendimiento(id_propiedad):
    """Página de detalle para un emprendimiento específico."""
    df = df_emprendimientos.copy()
    propiedad = df[df['id_propiedad'].astype(str) == str(id_propiedad)]
    
    if propiedad.empty:
        return render_template('404.html'), 404
    
    prop = propiedad.iloc[0].to_dict()
    
    # Obtener imágenes de la galería desde la columna 'galeria' (separadas por comas)
    imagenes_galeria = []
    if prop.get('galeria'):
        imagenes_galeria = [url.strip() for url in prop['galeria'].split(',') if url.strip()]
    
    # Si no hay imágenes en 'galeria', usar link_foto_inicio como fallback
    if not imagenes_galeria and prop.get('link_foto_inicio'):
        imagenes_galeria = [prop['link_foto_inicio']]
    
    print(f"📸 Imágenes en galería (emprendimiento): {len(imagenes_galeria)}")
    
    # Emprendimientos relacionados (mismo barrio o tipo)
    relacionadas = df[(df['barrio'] == prop['barrio']) | (df['tipo_inm'] == prop['tipo_inm'])].head(4).to_dict('records')
    
    return render_template('detalle.html',
                           propiedad=prop,
                           relacionadas=relacionadas,
                           imagenes_galeria=imagenes_galeria)

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

# ============================================================
# 6. INICIO DE LA APLICACIÓN
# ============================================================

if __name__ == '__main__':
    app.run(debug=True)
