# Validador bases de prima y cesantías - Nómina JMC

App de Streamlit para validar las bases de prima y cesantías usando únicamente:

- Acumulados históricos de nómina.
- Histórico de salarios por vigencias.
- Master Data / empleados opcional.
- Parametrización opcional de conceptos y áreas de nómina.

No es una herramienta de proyección. No toma mes proyectado ni suma pagos proyectados.

## Corrección incluida

Esta versión lee correctamente reportes planos SAP exportados con tuberías (`|`), por ejemplo:

```text
|Nº pers.|Número de personal|Desde|Hasta|Importe|Mon.|CC-nómina|Área de nómina|
```

También normaliza áreas de nómina cuando vienen como texto largo:

- `MENSUAL ADMON 365` -> `ZL`
- `ADMINISTRATIVOS` -> `ZM`
- `TIEMPO PARCIAL HORA` / `TIEMPOR PARCIAL HORA` -> `ZH`
- `TIEMPO PARCIAL DIA` -> `ZP`

## Lógica principal

```text
Base calculada = salario histórico promedio por vigencias + promedio variable de acumulados
```

El salario histórico incluye los componentes que vengan en el histórico de salarios, por ejemplo sueldo básico y bonos fijos por vigencia.

Para evitar errores, si un empleado tiene sueldo básico y bono fijo en el mismo tramo de fechas, los días de salario se cuentan una sola vez y el valor se suma por componente.

Los conceptos de salario fijo `Y010`, `Y011`, `Y020`, `Y050`, `Y051`, `Y090` no se duplican desde acumulados.

## Reglas default por área de nómina

- `ZM`: días 360.
- `ZL`: días calendario.
- `ZH`: días calendario / base 365 para promedio.
- `ZP`: días calendario / base 365 para promedio.

La app permite editar o cargar esta parametrización.

## Salidas

El Excel generado incluye:

- `Llevar_al_Modelo`
- `Base_Prima`
- `Base_Cesantias`
- `Detalle_Acumulados`
- `Historico_Salarios_Calculo`
- `Conceptos_Usados`
- `Alertas`
- `Log`

## Despliegue en Streamlit Cloud

Sube al repositorio:

```text
app.py
requirements.txt
.streamlit/config.toml
```

El archivo `.streamlit/config.toml` sube el límite de carga a 1 GB.
