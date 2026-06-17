# Validador de bases de prima y cesantías - Nómina JMC

Aplicación Streamlit para generar y validar bases prestacionales de **prima** y **cesantías** usando:

1. Acumulados históricos de nómina.
2. Histórico de salarios por vigencia.
3. Master Data / empleados opcional.
4. Parametrización opcional de conceptos.

## Enfoque del motor

Esta versión **no realiza proyección**. Calcula la base prestacional así:

```text
Base final = Salario histórico promedio + Variables acumuladas promedio
Variables acumuladas promedio = Valor acumulado / días divisor × 30
```

Los conceptos salariales configurados, por ejemplo `Y010`, `Y011`, `Y020`, `Y050`, `Y051`, `Y090`, se usan para validar contra los acumulados SAP, pero el componente salarial principal se toma del histórico de salarios para evitar duplicidad.

## Archivos esperados

### Acumulados de nómina
Debe contener columnas equivalentes a:

- SAP / Nº pers.
- Concepto / CC-nómina
- Valor
- Periodo o fecha de pago
- Texto concepto y cantidad son opcionales

### Histórico de salarios
Debe contener columnas equivalentes a:

- SAP / Nº pers.
- Salario
- Desde / fecha inicio / vigencia desde
- Hasta / fecha fin / vigencia hasta opcional

Si no viene fecha `Hasta`, el motor la infiere con el siguiente cambio salarial del mismo empleado. Si no hay siguiente registro, se asume vigente hasta 31/12/2099.

### Master Data opcional
Permite traer nombre, cédula, CECO, cargo, área de nómina, fecha ingreso y fecha retiro.

### Parametrización opcional
La app permite descargar una plantilla desde el panel lateral.

## Salida Excel

El archivo generado incluye:

- `Llevar_al_Modelo`
- `Base_Prima`
- `Base_Cesantias`
- `Detalle_Acumulados`
- `Historico_Salarios_Calculo`
- `Conceptos_Usados`
- `Alertas`
- `Log`

## Despliegue en Streamlit Cloud

Sube `app.py` y `requirements.txt` a GitHub y crea una app en Streamlit Cloud apuntando a `app.py`.
