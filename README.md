# JC Control de Solicitudes — GitHub + Streamlit

Repositorio: Dev-FilesJC/PROJECT_DESPACHO_JC

## Archivos

- `app.py` — aplicación Streamlit.
- `CONTROL-SOLICITUDES-ALMACEN.xlsx` — Excel central.
- `requirements.txt` — dependencias.

## Secret de Streamlit Cloud

En la aplicación de Streamlit Cloud, abrir **Settings → Secrets** y agregar:

```toml
GITHUB_TOKEN = "TU_TOKEN_DE_GITHUB"
```

El token debe tener permiso para leer y escribir el contenido del repositorio.

## Flujo

1. La aplicación descarga el Excel actual desde GitHub.
2. El usuario registra, edita o elimina una solicitud.
3. El sistema conserva la hoja `PROGRAMACION_RUTAS`.
4. El XLSX actualizado se publica en GitHub mediante un commit.
5. La aplicación vuelve a cargar el archivo para mostrar el resultado real.

## Importante

No subir nunca el token de GitHub al repositorio ni escribirlo directamente en `app.py`.
