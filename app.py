<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login - PLANI Presentaciones</title>
    <style>
        :root {
            --gold: #D4AF37; /* Color dorado similar al logo */
            --dark: #1a1a1a;
            --light: #f9f9f9;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: var(--light);
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
        }

        .login-card {
            background: white;
            padding: 40px;
            border-radius: 15px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            width: 100%;
            max-width: 400px;
            text-align: center;
        }

        .logo-container {
            margin-bottom: 20px;
        }

        .logo-img {
            max-width: 150px; /* Ajusta el tamaño del logo aquí */
            height: auto;
        }

        h2 {
            color: var(--dark);
            margin: 0 0 5px 0;
            font-size: 24px;
            letter-spacing: 1px;
            font-weight: 700;
        }

        .subtitle {
            color: #666;
            font-size: 14px;
            margin-bottom: 30px;
            text-transform: uppercase;
            letter-spacing: 2px;
        }

        .form-group {
            margin-bottom: 20px;
            text-align: left;
        }

        label {
            display: block;
            margin-bottom: 8px;
            color: #333;
            font-weight: 600;
            font-size: 14px;
        }

        input[type="text"], 
        input[type="password"] {
            width: 100%;
            padding: 12px;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 16px;
            box-sizing: border-box;
            transition: border-color 0.3s;
        }

        input:focus {
            outline: none;
            border-color: var(--gold);
            box-shadow: 0 0 0 3px rgba(212, 175, 55, 0.2);
        }

        .contact-info {
            background-color: #fff8e1; /* Fondo amarillo muy suave */
            border: 1px solid #ffe082;
            color: #5d4037;
            padding: 12px;
            border-radius: 8px;
            font-size: 14px;
            margin-bottom: 20px;
            line-height: 1.5;
        }

        .contact-info strong {
            color: var(--dark);
        }

        button {
            width: 100%;
            padding: 14px;
            background: linear-gradient(135deg, #D4AF37 0%, #B8860B 100%); /* Gradiente dorado */
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
        }

        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(212, 175, 55, 0.4);
        }

        .error-msg {
            color: #e74c3c;
            background-color: #fadbd8;
            padding: 10px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 14px;
        }
    </style>
</head>
<body>

    <div class="login-card">
        <!-- Logo -->
        <div class="logo-container">
            <!-- NOTA: Asegúrate de guardar la imagen del logo como 'logo.png' en la carpeta 'static' -->
           <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/5/51/Google.png/640px-Google.png" alt="Logo Prueba" style="width:150px;">
        </div>

        <!-- Título -->
        <h2>PLANI</h2>
        <div class="subtitle">PRESENTACIONES</div>

        <!-- Mensajes de error (Flash messages) -->
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="error-msg">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <form method="POST">
            <div class="form-group">
                <label for="username">Usuario</label>
                <input type="text" id="username" name="username" placeholder="Ingresa tu usuario" required autofocus>
            </div>

            <div class="form-group">
                <label for="password">Contraseña</label>
                <input type="password" id="password" name="password" placeholder="••••••••" required>
            </div>

            <!-- Información de Contacto -->
            <div class="contact-info">
                Contacta para tu cuenta al<br>
                <strong>TEL: 809-446-6921</strong>
            </div>

            <button type="submit">Iniciar Sesión</button>
        </form>
    </div>

</body>
</html>