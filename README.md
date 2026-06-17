# Validador bases de prima y cesantías - Nómina JMC

Esta app de Streamlit valida las bases de prima y cesantías usando:

- Acumulados históricos de nómina.
- Histórico de salarios por vigencias.
- Master Data / empleados opcional.
- Parametrización opcional de conceptos y áreas de nómina.

No es una herramienta de proyección. No toma un mes proyectado ni suma pagos proyectados.

## Lógica principal

Base calculada = salario histórico promedio por vigencias + promedio variable de acumulados.

Los conceptos de salario fijo `Y010`, `Y011`, `Y020`, `Y050`, `Y051`, `Y090` se calculan desde el histórico salarial para no duplicarlos como variables acumuladas.

## Reglas default por área de nómina

- ZM: días 360.
- ZL: días calendario.
- ZH: días calendario / base 365 para promedio.
- ZP: días calendario / base 365 para promedio.

La app permite editar o cargar esta parametrización.

## Despliegue en Streamlit Cloud

Sube al repositorio:

```text
app.py
requirements.txt
.streamlit/config.toml
```

El archivo `.streamlit/config.toml` sube el límite de carga a 1 GB.
