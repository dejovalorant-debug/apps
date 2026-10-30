# Sistema de Deudores para Android

Aplicación Android en orientación horizontal, con navegación inferior: Panel, Clientes, Ventas y Pagos. Mantiene el estilo turquesa y una base SQLite local.

## Generar el APK (sin instalar Android Studio en Windows)

1. Crear una cuenta en GitHub si todavía no tenés.
2. Crear un repositorio nuevo y subir **todo el contenido** de esta carpeta.
3. Entrar en la pestaña **Actions** del repositorio.
4. Elegir **Build Android APK** y pulsar **Run workflow** (o esperar el build automático al subir a `main`).
5. Cuando termine, abrir la ejecución y descargar el artefacto **Sistema-Deudores-APK**.
6. Descomprimirlo y pasar el `.apk` al teléfono. Habilitar la instalación desde esa fuente cuando Android lo solicite.

## Datos
La base `sistema_deudores.db` queda dentro del almacenamiento privado de la aplicación. Android la conserva entre aperturas y actualizaciones normales. Desinstalar la app puede borrar esos datos: hacer respaldo antes de desinstalar.

El Excel original se incluye como `assets/datos_iniciales.xlsx` y se importa una sola vez cuando la base todavía está vacía.

## Nota
Este paquete es el proyecto fuente preparado para compilar. El APK se genera en GitHub Actions; no se afirma que el APK ya esté compilado o probado en un dispositivo físico.
