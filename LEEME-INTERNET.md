# Publicar Presentador en internet

Con esto los maestros solo abren un enlace en el navegador, escriben el tema y descargan su PowerPoint. No instalan nada.

Los menús de los sitios web cambian con el tiempo. Si algún nombre no coincide, busca la opción equivalente.

## Paso 1. Subir el código a GitHub

1. Crea una cuenta gratis en github.com.
2. Crea un repositorio nuevo (botón "New"), con el nombre `presentador`. Puede ser privado.
3. Dentro del repositorio, elige "Add file" > "Upload files".
4. Descomprime `presentador.zip` en tu PC y arrastra **el contenido** de la carpeta `presentador` (los archivos `app.py`, `requirements.txt`, `Procfile`, `.gitignore` y la carpeta `templates`). No subas el .zip.
5. Pulsa "Commit changes".

Nunca escribas tu clave de la IA dentro de los archivos. Va en el paso 2.

## Paso 2. Crear el sitio en un servicio de alojamiento

Sirve cualquier servicio que ejecute aplicaciones Python desde GitHub; por ejemplo Render (render.com). Los pasos generales son:

1. Crea una cuenta y conecta tu GitHub.
2. Crea un "Web Service" nuevo y elige el repositorio `presentador`.
3. Configura:
   - Lenguaje / Runtime: **Python 3**
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app --workers 1 --threads 4 --timeout 120`
4. En "Environment" (variables de entorno) agrega:

| Variable | Qué poner |
|---|---|
| `GEMINI_API_KEY` | Tu clave de Google AI Studio (opcional, pero da mejor contenido) |
| `ACCESS_CODES` | Códigos de acceso separados por comas, por ejemplo `maestra-ana-7K2,colegio-sol-9PQ`. Si lo dejas vacío, cualquiera puede usar la app |
| `DAILY_LIMIT` | Presentaciones por código al día (por defecto 20) |
| `CONTACT` | Tu correo o página web. Wikipedia y Wikimedia piden un contacto para usar su servicio |

5. Pulsa crear / desplegar y espera unos minutos. Te darán un enlace público.

## Cómo funciona el control de acceso

- Si pones `ACCESS_CODES`, la página pide un código. Dale un código distinto a cada maestro o escuela; así sabes quién usa la app y puedes quitar un código cambiando la variable.
- Cada código tiene un límite diario (`DAILY_LIMIT`). Esto protege tu cuota de la IA.
- Si no pones códigos, el límite se aplica por dirección IP.
- Los contadores viven en la memoria del servidor: se reinician cuando la app se reinicia o se actualiza.

## Cosas que debes revisar antes de vender

- **Planes gratuitos de alojamiento:** algunos suspenden la app tras un rato sin uso, y la primera visita tarda en cargar. Revisa el plan y precio actuales.
- **IA de Google (Gemini):** el límite gratuito se comparte entre todos tus usuarios. Lee los términos de Google sobre uso comercial y sobre los datos antes de cobrar, y considera un plan de pago si crece el uso.
- **Imágenes:** vienen de Wikimedia Commons, que solo aloja material con licencia libre, pero casi todas exigen dar crédito. La app los incluye en la última diapositiva. Algunas licencias (como CC BY-SA) tienen condiciones extra: léelas si vas a vender las presentaciones como producto.
- **Texto de Wikipedia:** también tiene licencia libre con atribución. La app cita el enlace del artículo en la última diapositiva.

## Qué falta para cobrar de verdad

Esta versión controla el acceso con códigos que tú entregas. Si quieres cuentas de usuario, pagos o suscripciones automáticas, es un paso siguiente aparte.
