# Presentador

Aplicación web local: escribes un tema y genera una presentación `.pptx` que se abre con tu PowerPoint.

## Instalación (Windows)

1. Instala Python 3.10 o superior desde python.org (marca "Add Python to PATH").
2. En esta carpeta, abre una terminal y ejecuta:

```
pip install -r requirements.txt
python app.py
```

3. Abre http://127.0.0.1:5000 en tu navegador.

## IA gratuita (opcional)

Sin configurar nada, la app arma las diapositivas con el texto de Wikipedia.
Para que una IA las redacte mejor:

1. Crea una clave gratis en Google AI Studio (aistudio.google.com).
2. Antes de iniciar la app, en la terminal:

```
set GEMINI_API_KEY=tu_clave_aqui
python app.py
```

Los límites del plan gratuito los define Google y pueden cambiar. El modelo se puede cambiar con la variable `GEMINI_MODEL`.

## Imágenes y notas del orador

- **Imágenes:** se buscan en Wikimedia Commons (gratis). Los créditos de cada foto aparecen en la última diapositiva. Si no hay imagen para alguna diapositiva, la app lo avisa.
- **Notas del orador:** están en el panel de notas de cada diapositiva de PowerPoint (Vista > Notas). Con la clave de Gemini las redacta la IA; sin ella, se toman del texto de Wikipedia.
- Ambas opciones se pueden desmarcar en la pantalla principal.

## Cómo abre PowerPoint

Con la casilla "Abrir automáticamente en PowerPoint" (solo Windows, solo desde tu propia PC),
el archivo se guarda en la carpeta temporal y se abre con el programa asociado a `.pptx`.
Sin la casilla, el navegador descarga el archivo.
