# Validador de bases de prima y cesantías - Nómina JMC

Herramienta Streamlit para validar bases de prima y cesantías usando:

- Acumulados históricos de nómina.
- Histórico de salarios SAP.
- Master Data opcional.
- Parametrización opcional de conceptos y reglas por área de nómina.

## Corrección clave de esta versión

El histórico de salarios puede venir como reporte SAP TXT con tuberías, por ejemplo:

`|Nº pers.|Número de personal|Desde|Hasta|Importe|Mon.|CC-nómina|Área de nómina|`

La app detecta ese encabezado aunque no esté en la primera fila del archivo.

También aplica esta regla:

`Hasta = 31.12.9999` se reemplaza por la **fecha de corte de validación** digitada por el usuario.

## Lógica principal

Base calculada = salario histórico promedio por vigencias + variables promedio desde acumulados.

Cuando existen varios componentes fijos simultáneos en el histórico salarial, por ejemplo Sueldo Básico + Bono Antigüedad Operación, la app suma los valores activos del tramo, pero cuenta los días una sola vez.

## Reglas por área

- ZM / ADMINISTRATIVOS: días 360.
- ZL / MENSUAL ADMON 365: días calendario.
- ZH / tiempo parcial horas: días calendario/base 365.
- ZP / tiempo parcial días: días calendario/base 365.

## Streamlit Cloud

Incluye `.streamlit/config.toml` para permitir cargas de hasta 1 GB por archivo.
