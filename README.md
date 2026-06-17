# Validador de bases de prima y cesantías

Versión optimizada para Streamlit Cloud.

Cambios principales:
- No lee acumulados grandes al momento de subir el archivo.
- Lee acumulados únicamente al presionar **Generar validación de bases**.
- Para TXT/CSV usa lectura por chunks y agrupa por SAP/mes/concepto para reducir memoria.
- Mantiene la lógica: acumulados históricos + histórico de salarios, sin proyección.
- Convierte 31.12.9999 a la fecha de corte/validación digitada por el usuario.
- Reemplaza `use_container_width` por `width="stretch"`.
- Incluye `.streamlit/config.toml` para permitir archivos hasta 1GB.
