[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Plantilla de Excel de Cálculo de Costos para Transitarios: Calculadora de Beneficios y Pérdidas (P&L) por Envío, Asignación de Costos de Contenedores y Seguimiento de Márgenes
<img width="2880" height="1228" alt="image" src="https://github.com/user-attachments/assets/f89c1a0f-f576-43af-b26a-02bae4a47348" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE) [![Platform](https://img.shields.io/badge/Platform-Web%20Browser%20%2B%20Microsoft%20Excel%20(.xlsx)-217346.svg)](#) [![Tool Type](https://img.shields.io/badge/Tool%20Type-Plug--and--Play%20Freight%20Accounting%20Spreadsheet-2251FF.svg)](#)

**Una plantilla de Excel de cálculo de costos para transitarios (`.xlsx`) lista para usar y una calculadora online gratuita de rentabilidad de envíos, diseñada para transitarios, NVOCC, agentes de aduanas y operadores 3PL que necesitan calcular el beneficio neto por envío sin escribir fórmulas.** Siga los **márgenes brutos del conocimiento de embarque (B/L) y de la guía aérea (AWB)**, automatice las **conversiones de divisas (FX) multimoneda (USD, EUR, CNY, GBP)**, divida los **costos de contenedores FCL/LCL** y concilie las **facturas de clientes (AR) contra las cuentas de transportistas marítimos, drayage, despacho de aduanas y transferencias bancarias (AP)** en una hoja de cálculo logística de beneficios y pérdidas ya preconstruida.

**Pruébela gratis en su navegador ahora mismo, o descargue el archivo Excel desbloqueado para pegar sus propios datos de envíos.**

Cada regla de búsqueda, conversión de moneda y fórmula de reparto de costos de contenedor ya está integrada. Pegue sus filas de facturación desde QuickBooks, Xero o su hoja de seguimiento operativo para ver su beneficio exacto por reserva, contenedor y código de cargo de flete en segundos, respaldado por una garantía de devolución de 30 días.

> 🌐 **Pruebe primero la calculadora online** → [Lance la calculadora web gratuita e interactiva de cálculo de costos y P&L de envíos para transitarios (demo en el navegador, sin configuración previa)](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)  
> 📥 **Descargue la plantilla de Excel lista para usar** → [Obtenga la plantilla de Excel desbloqueada de cálculo de costos y rentabilidad de envíos para transitarios (.xlsx con fórmulas de actualización automática)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)  
>
> *Tanto la calculadora web gratuita como el libro de Microsoft Excel descargable funcionan sobre exactamente las mismas tablas contables de fletes preconfiguradas y las mismas fórmulas automatizadas de márgenes.*

---

## Lo que le ayuda a seguir: de los puntos débiles del flete a las soluciones

Esperar los informes contables de fin de mes oculta qué envíos individuales ganaron o perdieron dinero realmente. Esta **hoja de cálculo plug-and-play de cálculo de costos de fletes** corrige los cinco puntos ciegos contables operativos más habituales del flete marítimo, la carga aérea y el forwarding multimodal:

| Tema de búsqueda frecuente y objetivo de seguimiento | El punto débil diario de la contabilidad de fletes (por qué se fuga el margen) | Cómo esta plantilla de Excel automatizada lo corrige y lo sigue |
| :--- | :--- | :--- |
| **Beneficio por envío y P&L a nivel de reserva (margen de `B/L` y `AWB`)** | QuickBooks y Xero agrupan las facturas por mes natural, ocultando los envíos con pérdidas, las cotizaciones con precios demasiado bajos y las rutas comerciales no rentables. | Vincula **ingresos AR de clientes, cuentas AP de proveedores y comisiones bancarias por transferencia** con cada `Booking_ID` individual para que vea en tiempo real el **beneficio neto ($) y el margen bruto (%) por envío**. |
| **Rentabilidad por código de cargo de flete y cargos accesorios** | Conoce el importe total de la factura, pero no puede saber si perdió dinero en **flete marítimo, drayage local, THC, despacho de aduanas, demora o seguro de carga**. | Desglosa ingresos vs. costos de proveedor por **códigos de servicio de flete** granulares (`Ocean Freight`, `THC`, `Origin Inspection`, `Drayage`, `Customs`, `Documentation`) para exponer sobrecostos ocultos. |
| **Asignación de costos compartidos multicontenedor FCL (reparto `TEU` / `FEU`)** | Cuando un conocimiento de embarque madre (MBL) cubre varios contenedores, las tarifas fijas a nivel de reserva (como el despacho de aduanas o la documentación del B/L) se cuentan dos veces o quedan fuera del precio por contenedor. | Separa los **costos directos por contenedor** de las **tarifas compartidas a nivel de reserva**, dividiendo automáticamente los costos compartidos a partes iguales entre todos los contenedores activos de esa reserva. |
| **Conversión de facturas de flete multimoneda (`USD`, `EUR`, `CNY`, `GBP`)** | Pagar a los transportistas marítimos en USD, a los agentes de origen en CNY o EUR y facturar a los clientes en moneda local con tasas calculadas a mano provoca pérdidas por tipo de cambio al cierre de mes. | Convierte cada partida AR y AP en monedas mixtas a su **única moneda base de información** mediante una tabla de búsqueda FX central, marcando de inmediato cualquier código de moneda sin configurar. |
| **Detección de facturas de proveedor faltantes y errores tipográficos en hojas de cálculo** | Un solo error tipográfico en un ID de reserva dentro de una hoja manual con `SUMIFS` o `VLOOKUP` hace que una cuenta de transportista de $2,000 se convierta silenciosamente en `$0`, inflando el beneficio y pagando de más las comisiones de ventas. | Ejecuta una **verificación automática de integridad de datos** (`ERR: UNKNOWN_CURRENCY`, `ERR: INVALID_BOOKING`) que resalta las filas defectuosas a simple vista sin alterar los totales de su panel. |

---

## Cómo calcular el beneficio por envío y ejecutar una auditoría de costos de fletes (tutorial de 4 pasos para principiantes)

No necesita formación contable ni dominio de las fórmulas de Excel para auditar los márgenes de sus envíos. Siga estos cuatro pasos para convertir listas de facturas en bruto y cuentas de transportistas en un exacto **estado de beneficios y pérdidas reserva por reserva**:

### Paso 1: Configure la moneda base, los tipos de cambio y la política de reparto de contenedores
Abra la pestaña `Setup_Params` para establecer las reglas base de su empresa en lenguaje sencillo:
- Elija su **moneda base de información** (por ejemplo `USD`, `EUR` o `GBP`) e introduzca sus **tasas de cambio (FX)** estándar.
- Revise las **categorías de transacción** preasignadas (`Ingresos de clientes`, `COGS directo de proveedor`, `Comisiones bancarias por transferencia` y `Depósitos de clientes`).
- Seleccione su **regla de asignación de costos compartidos** para envíos FCL multicontenedor.

*Pruébelo en vivo:* Vea ahora mismo cómo funciona la tabla de configuración en la [calculadora de navegador gratuita de cálculo de costos de fletes para transitarios](https://hyvoid.github.io/freight-forwarder-job-costing-excel/).

---

### Paso 2: Pegue sus reservas, contenedores y líneas AR/AP de facturas en bruto
Suelte sus IDs de envío en las listas maestras (`Bookings` y `Containers`) y luego copie y pegue sus filas de facturación directamente en la tabla `Raw_Transactions`:
- **Funciona con cualquier fuente de datos:** exportaciones de QuickBooks Online / Desktop, volcados CSV de Xero, registros de facturación de CargoWise o Magaya, facturas de flete de transportistas marítimos, extractos de agentes de aduanas y rastreadores de despacho manuales.
- **Cero arrastre manual de fórmulas:** las tablas de Excel se expanden automáticamente al pegar nuevas filas.

---

### Paso 3: Revise las alertas de error automáticas y consulte su panel de P&L de envíos
En el momento en que pegue sus datos, la hoja de cálculo lo calcula todo en segundo plano:
- Convierte las cuentas de proveedores y las facturas de clientes en moneda extranjera a su moneda base.
- Marca los IDs de reserva mal escritos o las tasas de moneda faltantes en la columna `Data Health Status` para que una cuenta de proveedor sin vincular nunca quede oculta como partida de costo cero.
- Actualiza su **resumen de beneficios y pérdidas por reserva**, su **tabla de márgenes por código de servicio de flete** y su **desglose de costos unitarios por contenedor**.

---

### Paso 4: Descargue la plantilla de Excel desbloqueada para reutilizarla cada mes con un clic
La calculadora gratuita del navegador es ideal para una comprobación rápida en una sola sesión. Cuando necesite guardar sus datos localmente, importar cientos de filas contables mensuales y seguir salidas de buques en curso a lo largo de los ciclos de facturación, pase al libro `.xlsx` descargable.

> 💡 **Por qué los transitarios con poco tiempo descargan el archivo Excel en lugar de construir su propia hoja:**  
> La conciliación de fletes entre monedas, el reparto de costos multicontenedor y la detección de cuentas huérfanas implican una lógica contable compleja bajo el capó. **En esta plantilla de Excel, cada regla de coincidencia, vínculo de búsqueda y fórmula de expansión automática ya está construida y probada para usted.**  
> 
> Ni siquiera necesita hacer clic y arrastrar las esquinas de las celdas para copiar fórmulas hacia abajo. Simplemente descargue el archivo a su ordenador, pegue sus datos de facturación en bruto y obtenga al instante cifras de beneficio listas para auditoría. Durante las temporadas altas de navegación, dedique su tiempo a ganar cotizaciones de cargadores, asegurar espacio en buque y mover carga, mientras una plantilla sencilla de copiar y pegar se encarga de su análisis financiero en segundos.  
> 
> 👉 **[Descargue hoy la plantilla de Excel desbloqueada de cálculo de costos para transitarios (.xlsx) y automatice el P&L de sus envíos](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)**

---

## La lógica de negocio y la metodología: 4 puntos ciegos de beneficio que esta plantilla de Excel resuelve por usted

La mayoría de los transitarios pierde dinero en envíos no porque cotizara mal la tarifa marítima, sino porque las herramientas contables estándar y las hojas de Excel caseras ocultan a dónde va realmente el dinero. Así protege esta plantilla su resultado final, en términos sencillos, aunque no tenga formación contable formal:

### 1. El software contable mensual oculta los envíos con pérdidas (vinculación a nivel de reserva vs. totales mensuales)
**El hecho objetivo:** El software contable estándar como QuickBooks o Xero organiza el dinero por **mes natural**, mientras que el forwarding de fletes ocurre por **reserva, conocimiento de embarque (B/L) y contenedor**. Si factura a un cliente $8,000 en octubre, pero el transportista marítimo y el transportista de drayage le cobran $7,500 en noviembre, octubre parece artificialmente rentable y noviembre parece una pérdida.  
**Su beneficio directo:** Esta plantilla agrupa cada dólar de ingresos, costo de proveedor y comisión bancaria por `Booking_ID`, sin importar el mes en que llegue la factura. Ve el beneficio neto real de cada envío antes de renovar la tarjeta de tarifas de un cliente o pagar las comisiones de ventas.

### 2. Las tarifas fijas de reserva vs. los costos por contenedor distorsionan sus cotizaciones (reparto automático de costos de contenedores FCL)
**El hecho objetivo:** En el flete marítimo y multimodal, algunos costos ocurren una vez por **reserva** (como una tarifa de despacho de aduanas de $150 o un cargo de documentación del B/L), mientras que otros ocurren por **contenedor** (como $450 de manipulación en terminal o drayage local por cada caja 40'HC). En una hoja de cálculo básica, los operadores o multiplican por error la tarifa aduanera de $150 entre 5 contenedores (inflando la estimación de costos en $600 y perdiendo la licitación) u olvidan incluir las tarifas compartidas de la reserva al calcular el beneficio por contenedor.  
**Su beneficio directo:** La plantilla separa los costos específicos de contenedor de los costos compartidos de toda la reserva y reparte las tarifas compartidas a partes iguales entre el número exacto de contenedores de esa reserva:

```text
Single Container Total Cost = Direct Container Cost + (Shared Booking Costs ÷ Total Containers in Booking)

```

Obtiene costos unitarios exactos por contenedor sin hacer divisiones manuales.

### 3. Un solo error tipográfico en una hoja casera falsea silenciosamente un beneficio alto (detección integrada de errores tipográficos y costos faltantes)

**El hecho objetivo:** En una hoja de Excel estándar construida con `SUMIFS`, un solo error tipográfico en el número de referencia de una reserva en una cuenta de drayage de $1,500 hace que Excel ignore esa cuenta por completo y añada `$0` a sus costos. A la inversa, un `VLOOKUP` roto convierte toda su tabla resumen en errores `#N/A`. Vea cómo una sola factura de proveedor sin vincular distorsiona una reserva real:

```text
Reported Booking Revenue:   $8,500
Reported Vendor Costs:      $5,200  <-- ($1,500 drayage bill skipped due to a Booking ID typo)
Apparent Net Profit:        $3,300  <-- (Overstated by $1,500)
Apparent Profit Margin:      38.8%  <-- (Fake margin; actual real margin is only 21.1%)

```

Si confía en ese 38.8%, paga comisiones de más a sus vendedores con su propio bolsillo y sigue cotizando una ruta que apenas llega al punto de equilibrio.

**Su beneficio directo:** Cada fila que pegue pasa por una **verificación automática de integridad de datos**. Si una fila tiene un ID de reserva mal escrito, un código de cargo inválido o una tasa de moneda faltante, la plantilla lo marca con claridad (`ERR: INVALID_BOOKING` o `ERR: UNKNOWN_CURRENCY`) mientras mantiene limpios los cálculos de su resumen. Detecta las cuentas de proveedor faltantes antes de que le cuesten dinero.

### 4. Las monedas extranjeras y los depósitos anticipados de clientes distorsionan el margen bruto real (aislamiento de FX y depósitos a nivel de fila)

**El hecho objetivo:** Los envíos transfronterizos combinan habitualmente flete marítimo en USD, cuentas de agentes de origen en CNY o EUR, comisiones bancarias de transferencias intermedias y depósitos anticipados de clientes. Convertir totales de monedas a mano a fin de mes introduce errores de tipo de cambio, y contar los depósitos anticipados como ingresos ganados hace que un envío parezca rentable antes incluso de pagar las cuentas de los proveedores.

**Su beneficio directo:** La plantilla convierte cada cuenta en moneda extranjera a su moneda base a nivel de fila individual y separa los **depósitos de clientes (efectivo retenido)** de los **ingresos de flete ganados y el COGS**. Ve su margen de beneficio real y su posición de caja efectiva lado a lado.

---

## Hojas de cálculo manuales de fletes y software contable genérico vs. esta plantilla de Excel automatizada de cálculo de costos de envíos

| Factor de comparación operativa | Hojas de Excel manuales y contabilidad genérica (solo QuickBooks / Xero) | Plantilla estandarizada de Excel de cálculo de costos para transitarios (esta herramienta) |
| --- | --- | --- |
| **Visibilidad del beneficio por envío (`MBL` / `HBL` / `AWB`)** | Mezcla todas las facturas mensuales en cuentas del libro mayor, ocultando qué reservas, contenedores o cargadores específicos perdieron dinero. | Genera una **tabla de P&L reserva por reserva** automática que muestra ingresos brutos, COGS de proveedores, comisiones bancarias, beneficio neto ($) y margen (%). |
| **Conversión de facturas multimoneda (`USD`, `EUR`, `CNY`, `GBP`)** | El personal convierte las cuentas de agentes extranjeros con calculadoras de escritorio y tasas inconsistentes, lo que provoca brechas de margen inexplicables en el cierre de mes. | Usa una **tabla centralizada de consulta de tasas FX** para convertir automáticamente cada línea de factura y muestra una advertencia instantánea si falta una tasa de moneda. |
| **Reparto de costos multicontenedor FCL (`20'GP` / `40'HC` / `TEU`)** | Las tarifas de documentación del B/L, el despacho de aduanas y los recargos de reserva se duplican por contenedor o se omiten de los cálculos de costo unitario. | Separa automáticamente las **cuentas directas de contenedor** de los **gastos generales compartidos de la reserva**, repartiendo las tarifas compartidas entre el número exacto de contenedores por reserva. |
| **Margen de cargos accesorios y códigos de cargo (`THC`, `Drayage`, `Customs`)** | Ve la facturación total del envío, pero no puede saber si la **inspección en origen, el drayage terrestre, la demora o el seguro de carga** están filtrando beneficio. | Consolida ingresos vs. costos de proveedor por **código de servicio de flete** para que detecte qué cargos accesorios ganan dinero y cuáles lo pierden. |
| **Mantenimiento de fórmulas y protección contra errores tipográficos (`VLOOKUP` / `SUMIFS`)** | El personal debe arrastrar las fórmulas manualmente en cada nueva fila, y un solo ID de reserva mal escrito hace que los costos de proveedor caigan silenciosamente a cero. | **Tablas autoexpandibles y alertas de error preconstruidas** que se actualizan automáticamente al pegar y ponen en cuarentena los errores tipográficos (`ERR: INVALID_BOOKING`) antes de que corrompan su P&L. |

---

## Para quién es esto: roles logísticos, casos de uso de hojas de cálculo y flujos de trabajo de la plantilla

Este kit de herramientas está diseñado para **transitarios pequeños y medianos, operadores OTI/NVOCC, agentes de aduanas, equipos 3PL y agentes de flete** que quieren visibilidad inmediata del beneficio por envío sin gastar más de $20,000 y seis meses implementando un ERP o TMS empresarial:

* **Propietarios de transitarios y directores gerentes de NVOCC (`Plantilla de Excel de P&L para transitarios`):** Audite el beneficio neto real por envío (después de comisiones bancarias y gastos generales compartidos de la reserva) antes de aprobar tarjetas de tarifas anuales de cargadores o pagar comisiones mensuales a los vendedores.
* **Coordinadores de operaciones junior y despachadores (`Hoja de cálculo plug-and-play de costos de envíos`):** Siga los costos de trabajos FCL y LCL multicontenedor durante las temporadas altas pegando cuentas de proveedores y cargos de clientes en bruto en una plantilla preconstruida, sin escribir fórmulas.
* **Contables logísticos y auxiliares de contabilidad de fletes (`Plantilla de conciliación de facturas AR/AP de fletes`):** Cierre la brecha entre las exportaciones de QuickBooks, Xero o NetSuite y los archivos operativos de envíos, detectando cargos accesorios no facturados a clientes o cuentas de transportistas sin vincular antes del cierre contable mensual.
* **Analistas de precios de flete y gerentes de rutas comerciales (`Calculadora de costos unitarios de contenedores y margen de accesorios`):** Analice si el flete marítimo, la carga aérea, el despacho de aduanas, el seguro de carga o el drayage local impulsan o reducen los márgenes en rutas específicas de origen a destino.
* **CFOs externos y consultores logísticos (`Modelo de auditoría de fugas de margen para 3PL y fletes`):** Suelte las exportaciones de facturación de sus clientes en una hoja estandarizada para descubrir cuentas con precios demasiado bajos, pérdidas cambiarias y tarifas accesorias no recuperadas en una sola tarde.

*Nota: Este modelo de Excel complementa su TMS operativo (como CargoWise, Magaya o GoFreight) y su software de contabilidad fiscal. Actúa como su rápida capa plug-and-play de **rentabilidad de envíos y cálculo de costos por trabajo** cuando necesita respuestas claras de sus datos de facturación de inmediato.*

---

## Escenarios habituales de costos de fletes: cuándo usar esta hoja de cálculo de P&L de envíos

Ya sea que gestione una oficina de forwarding de cinco personas o cientos de contenedores FCL al mes, esta plantilla resuelve los problemas de costos diarios exactos que ralentizan a los equipos logísticos:

* **Cómo calcular el beneficio por envío en Excel sin escribir fórmulas:** Pegue sus líneas de facturas de clientes y cuentas de proveedores en una sola tabla, etiquete cada fila con su `Booking_ID` y deje que la hoja resumen preconstruida calcule automáticamente el beneficio neto y el margen bruto %.
* **Cómo repartir los costos compartidos del conocimiento de embarque (B/L) entre varios contenedores FCL:** Asigne los cargos específicos de contenedor (como drayage o THC) a números de `Container_ID` individuales y deje los cargos de toda la reserva (como el despacho de aduanas o la documentación) a nivel de `Booking_ID`: la plantilla reparte automáticamente los costos compartidos entre sus contenedores 20'GP y 40'HC.
* **Cómo conciliar cuentas retrasadas de transportistas marítimos y drayage contra facturas de clientes:** Mantenga las reservas abiertas en su registro a lo largo de los meses naturales para que los avisos de llegada tardíos, las cuentas de detención/demora o los cargos de división de chasis se deduzcan automáticamente del margen del envío correcto cuando lleguen.
* **Cómo seguir márgenes de flete multimoneda entre USD, EUR y CNY:** Introduzca sus tasas de cambio de referencia una sola vez en `Setup_Params` para normalizar las facturas de agentes de origen en el extranjero, el flete de los transportistas marítimos y la facturación de entrega local en una única moneda de información.
* **Cómo auditar las comisiones sobre el beneficio bruto de brokers de flete y vendedores:** Verifique que cada cuenta de proveedor, tarifa de origen y cargo bancario esté vinculado y contabilizado antes de calcular los pagos de comisiones de ventas sobre el beneficio bruto del envío.

---

## Acerca de Theseus Workshop

**Theseus Workshop** crea modelos financieros de Excel plug-and-play y calculadoras basadas en navegador para operadores que necesitan números claros y fiables hoy mismo, sin el exceso de software empresarial ni programación compleja de hojas de cálculo.

Cada plantilla está diseñada en torno a un objetivo práctico:

> **"Poner los números operativos y financieros exactos en una sola tabla preconstruida para que su próxima decisión de negocio tome segundos en lugar de horas."**

Si quiere dejar de pelearse con hojas `VLOOKUP` rotas y cálculos manuales de monedas, descargue la plantilla completa de Excel a continuación o explore nuestra biblioteca completa de hojas de cálculo operativas y calculadoras de negocio:

* 📥 **Obtenga esta plantilla de costos de fletes:** [Descargue la plantilla de Excel desbloqueada de cálculo de costos y P&L de envíos para transitarios (.xlsx)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)
* 🌐 **Pruebe la calculadora web gratuita:** [Abra la aplicación de navegador interactiva de cálculo de costos de fletes para transitarios](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)
* 🧰 **Explore todos los modelos de Excel listos para usar y kits de herramientas de negocio:** [Visite la tienda oficial de Theseus Workshop en TheseusWorkshop.com](https://www.google.com/search?q=https://www.theseusworkshop.com%3Futm_source%3Dgithub%26utm_medium%3DGitHub%2520README%26utm_campaign%3Dreadme%2520new%2520launch%26utm_content%3Dabout-section-store-link)

---

## Detalles técnicos

<details>
<summary>Para revisores técnicos, profesionales de Excel y colaboradores</summary>

### Arquitectura del libro de trabajo

El libro de trabajo está organizado en torno a cuatro capas funcionales: **reglas y parámetros → datos maestros → entrada de transacciones → cálculo y salida analítica**.

El diseño contiene 11 hojas:

| Capa              | Hoja                   | Rol                                                                             |
| ------------------ | ----------------------- | -------------------------------------------------------------------------------- |
| Reglas y parámetros | `Instructions`          | Instrucciones de uso, reglas de entrada de datos y definiciones de negocio               |
| Reglas y parámetros | `Setup_Params`          | Moneda base, tasas FX, correspondencias de transacciones y política de asignación de costos compartidos |
| Datos maestros        | `Booking_Master`        | Índice de reservas, cliente, estado y número de contenedores                             |
| Datos maestros        | `Container_Master`      | Relación reserva-contenedor 1:N                                            |
| Datos maestros        | `Service_Master`        | Código de servicio estándar, nombre del servicio y categoría                                |
| Entrada de transacciones  | `Raw_Transactions`      | Libro mayor unificado de transacciones financieras más diagnósticos de integridad de datos                |
| Motor de cálculo | `Job_Costing_Engine`    | Agregación a nivel de reserva de ingresos, COGS, comisiones bancarias, beneficio, margen y depósitos  |
| Análisis            | `Booking_PL`            | P&L de una sola reserva y desglose por servicio                                  |
| Análisis            | `Container_Analysis`    | Costos directos a nivel de contenedor y costos compartidos asignados                                |
| Análisis            | `Service_Profitability` | Rentabilidad por servicio a través de las reservas                                           |
| Salida ejecutiva   | `Executive_Dashboard`   | KPI globales, clasificaciones de rentabilidad y advertencias de integridad de datos                    |

El flujo central es:

```text
Setup_Params
      │
      ├──────────────┐
      ▼              ▼
Booking_Master   Service_Master
      │              │
      └──────┬───────┘
             ▼
     Raw_Transactions
             │
     Linkage + FX + Financial
        Treatment Diagnosis
             │
             ▼
    Job_Costing_Engine
             │
    ┌────────┼───────────┐
    ▼        ▼           ▼
Booking_PL  Container_  Service_
            Analysis    Profitability
             │
             └──────┬──────┘
                    ▼
          Executive_Dashboard
```

La arquitectura de origen identifica tres importantes cadenas de búsqueda de múltiples saltos: una cadena de moneda/tratamiento financiero de cuatro saltos, una cadena de costo compartido/desglose de contenedor de cuatro saltos y una cadena de clasificación de servicios de tres saltos. 

La regla de diseño crítica es que **los valores de cálculo y los estados de diagnóstico permanecen separados**. `USD Equivalent` permanece numérico, mientras que `Data Health Status` transporta errores como `ERR: UNKNOWN_CURRENCY`. Esto preserva la compatibilidad con `SUMIFS` al tiempo que impide que las transacciones defectuosas entren en el motor del P&L. 

### Tres trampas que atrapan incluso a operadores de fletes experimentados

#### Trampa 1 — Tratar una tasa FX faltante como un importe válido

**1. Decisión:** Una reserva parece tener un cierto costo y margen.

**2. Defecto oculto:** Una transacción de proveedor está denominada en una moneda que no está configurada en la tabla FX de referencia.

**3. Impacto en la decisión:** Si la tasa faltante se trata incorrectamente como `1.0`, un importe extranjero como `50,000` puede convertirse en `$50,000 USD`, distorsionando radicalmente el resultado de la reserva.

**4. Por qué está mal:** Una tasa de cambio faltante es un fallo de integridad de datos, no una tasa de cambio económica.

**5. Enfoque correcto:** Devolver un valor numérico neutro de `0` en el campo de cálculo e informar por separado `ERR: UNKNOWN_CURRENCY`.

**6. Decisión corregida:** La reserva se marca para completar datos en lugar de recibir un beneficio o una pérdida falsos.

```text
Wrong:
Unknown Currency → FX = 1.0 → USD Equivalent → P&L

Correct:
Unknown Currency → FX = 0
                 → USD Equivalent = 0
                 → Data Health = ERR: UNKNOWN_CURRENCY
                 → P&L excludes invalid transaction
                 → Management fixes source data
```

<details>
<summary>Lógica de la fórmula</summary>

```excel
=MAP(K2:K10000,L2:L10000,
 LAMBDA(curr,m_rate,
   IF(curr="",
      0,
      IF(m_rate>0,
         m_rate,
         XLOOKUP(
           curr,
           Setup_Params!$A$10:$A$20,
           Setup_Params!$B$10:$B$20,
           0
         )
      )
   )
 ))
```

La regla importante es **nunca asignar por defecto el valor 1.0 a una moneda desconocida**. 

</details>

#### Trampa 2 — Leer el P&L de una reserva antes de comprobar la integridad de los datos

**1. Decisión:** Una reserva con ingresos de `$8,500` y COGS de `$5,200` se considera un beneficio de `$3,300`.

**2. Defecto oculto:** Una o más transacciones de origen tienen vínculos de reserva o de servicio inválidos, o una moneda sin resolver.

**3. Impacto en la decisión:** El margen aparente puede estar sobreestimado porque una línea de costo quedó silenciosamente fuera de la agregación.

**4. Por qué está mal:** Un resultado de `SUMIFS` matemáticamente correcto puede seguir siendo económicamente erróneo si las filas de origen están incompletas.

**5. Enfoque correcto:** El motor agrega solo las transacciones cuyo `Data Health Status = "OK"` e informa por separado del número de filas inválidas.

**6. Decisión corregida:** La reserva se trata como **provisional** hasta que se resuelvan las excepciones de datos.

```text
Wrong:
Transactions → SUMIFS → Profit → Decision

Correct:
Transactions → Linkage / FX validation
             → Healthy rows only
             → Profit + Health Warning
             → Decision
```

<details>
<summary>Lógica de la fórmula</summary>

```excel
rev =
SUMIFS(
  Raw_Transactions!O2:O10000,
  Raw_Transactions!B2:B10000,bk,
  Raw_Transactions!N2:N10000,"Revenue",
  Raw_Transactions!P2:P10000,"OK"
)
```

El mismo filtro de integridad se aplica al COGS, las comisiones bancarias y los depósitos. 

</details>

#### Trampa 3 — Asignar el costo compartido a nivel de reserva directamente a cada contenedor

**1. Decisión:** Una reserva de dos contenedores necesita una vista de rentabilidad a nivel de contenedor.

**2. Defecto oculto:** Una línea de COGS a nivel de reserva pertenece al envío en su conjunto, no a un contenedor específico.

**3. Impacto en la decisión:** Asignar el costo compartido completo a ambos contenedores duplica el costo aparente.

**4. Por qué está mal:** La propiedad económica del costo es a nivel de reserva, mientras que el análisis exige una regla de asignación controlada.

**5. Enfoque correcto:** Usar la política configurada `EQUAL_SHARE` y dividir el COGS compartido entre el número de contenedores de la reserva.

**6. Decisión corregida:** Cada contenedor recibe sus costos directos más su parte asignada del costo común, preservando el total a nivel de reserva.

```text
Booking shared COGS = $1,000
Container count     = 2

Allocated shared cost per container
= $1,000 / 2
= $500
```

El libro de trabajo separa explícitamente los costos directos de contenedor de los costos compartidos a nivel de reserva y utiliza un denominador seguro basado en el número de contenedores. 

<details>
<summary>Lógica de la fórmula</summary>

```excel
alloc_cogs =
IF(
  alloc_policy="EQUAL_SHARE",
  direct_cogs + (shared_cogs / cntr_count),
  direct_cogs
)
```

`Setup_Params!B6` controla si el costo compartido se asigna mediante `EQUAL_SHARE` o se deja sin asignar con `NO_ALLOCATE`. 

</details>

### Escenario de ejemplo

Considere una reserva de forwarding con las siguientes transacciones válidas:

| Partida                      | Tipo     | Importe original | Moneda | Equivalente en USD |
| ------------------------- | -------- | --------------: | -------- | -------------: |
| Factura de cliente          | Revenue  |           8,500 | USD      |         $8,500 |
| Costo de proveedor de flete marítimo | COGS     |           3,000 | USD      |         $3,000 |
| Inspección en origen         | COGS     |             900 | USD      |           $900 |
| Seguro de carga           | Revenue  |             400 | USD      |           $400 |
| Comisión bancaria               | Bank Fee |              75 | USD      |            $75 |
| Depósito recuperable       | Deposit  |             500 | USD      |           $500 |

El P&L válido es:

```text
Revenue
= 8,500 + 400
= $8,900

COGS
= 3,000 + 900
= $3,900

Bank Fees
= $75

Net Profit
= $8,900 - $3,900 - $75
= $4,925

Profit Margin
= $4,925 / $8,900
= 55.34%
```

El depósito recuperable de `$500` se registra por separado en lugar de tratarse como ingreso operativo o COGS.

Supongamos ahora que se introduce una transacción de proveedor adicional en una moneda no configurada. El sistema **no** añade silenciosamente la transacción al COGS. Su tasa FX aplicada pasa a ser `0`, su cálculo en USD permanece numérico y el campo de diagnóstico informa `ERR: UNKNOWN_CURRENCY`. La reserva recibe una advertencia de integridad en lugar de presentarse como limpia.

La recomendación operativa cambia en consecuencia:

> **No use el margen del 55.34% como cifra final de rentabilidad hasta que se corrija la transacción sin resolver.**

Esta es la distinción clave entre una calculadora de rentabilidad y un flujo de trabajo controlado de cálculo de costos por trabajo: el resultado va acompañado de evidencia sobre si los datos subyacentes están lo suficientemente completos como para confiar en ellos.

### Referencia de fórmulas

<details>
<summary>Booking Master — Recuento de contenedores</summary>

**Propósito:** Determinar el número de contenedores vinculados a cada reserva y prevenir fallos de división por cero.

```excel
=MAP(
  A2:A10000,
  LAMBDA(
    bk,
    IF(
      bk="",
      "",
      LET(
        cnt,
        COUNTIF(Container_Master!A2:A10000,bk),
        IF(cnt=0,1,cnt)
      )
    )
  )
)
```

Si aún no se ha introducido ningún contenedor, el denominador recurre a `1`. Es una salvaguarda de cálculo, mientras que la ausencia de contenedores debe seguir tratándose operativamente como un problema de calidad de datos. 

</details>

<details>
<summary>Raw Transactions — Tasa FX aplicada</summary>

**Propósito:** Dar prioridad a una tasa FX manual y, en caso contrario, obtener la tasa de moneda configurada.

```excel
=MAP(
  K2:K10000,
  L2:L10000,
  LAMBDA(
    curr,m_rate,
    IF(
      curr="",
      0,
      IF(
        m_rate>0,
        m_rate,
        XLOOKUP(
          curr,
          Setup_Params!$A$10:$A$20,
          Setup_Params!$B$10:$B$20,
          0
        )
      )
    )
  )
)
```

Una moneda desconocida devuelve `0`, no `1.0`. 

</details>

<details>
<summary>Raw Transactions — Tratamiento financiero</summary>

**Propósito:** Asignar los tipos de transacción operativos a las categorías financieras.

```excel
=MAP(
  H2:H10000,
  LAMBDA(
    type,
    IF(
      type="",
      "",
      XLOOKUP(
        type,
        Setup_Params!$E$10:$E$20,
        Setup_Params!$F$10:$F$20,
        "COGS"
      )
    )
  )
)
```

El diseño de origen usa `COGS` como alternativa conservadora para un tipo de transacción no reconocido, reduciendo el riesgo de sobreestimar el beneficio. 

</details>

<details>
<summary>Raw Transactions — Equivalente en USD</summary>

**Propósito:** Convertir el importe original de cada transacción a la moneda base USD del libro de trabajo manteniendo un resultado numérico.

```excel
=MAP(
  J2:J10000,
  M2:M10000,
  LAMBDA(
    amt,rate,
    IF(
      amt="",
      0,
      amt*rate
    )
  )
)
```

Mantener esta columna numérica es importante porque los cálculos `SUMIFS` posteriores nunca deberían recibir texto de diagnóstico como `"INVALID"` en lugar de un número. 

</details>

<details>
<summary>Raw Transactions — Estado de integridad de datos</summary>

**Propósito:** Validar la vinculación de reserva, servicio y FX de forma independiente de las columnas de cálculo financiero.

```excel
=MAP(
  B2:B10000,
  F2:F10000,
  K2:K10000,
  M2:M10000,
  LAMBDA(
    bk,srv,curr,rate,
    IF(
      bk="",
      "",
      IF(
        ISNA(XLOOKUP(
          bk,
          Booking_Master!A2:A10000,
          Booking_Master!A2:A10000
        )),
        "ERR: INVALID_BOOKING",
        IF(
          ISNA(XLOOKUP(
            srv,
            Service_Master!A2:A500,
            Service_Master!A2:A500
          )),
          "ERR: INVALID_SERVICE",
          IF(rate=0,
             "ERR: UNKNOWN_CURRENCY",
             "OK"
          )
        )
      )
    )
  )
)
```

La capa de diagnóstico está intencionadamente separada de los campos numéricos de cálculo. 

</details>

<details>
<summary>Job Costing Engine — P&L a nivel de reserva</summary>

**Propósito:** Generar una lista dinámica de reservas y agregar solo las transacciones sanas.

```excel
=LET(
  bk_list,
  UNIQUE(
    FILTER(
      Booking_Master!A2:A10000,
      Booking_Master!A2:A10000<>""
    )
  ),
  BYROW(
    bk_list,
    LAMBDA(
      bk,
      LET(
        cust,
        XLOOKUP(
          bk,
          Booking_Master!A2:A10000,
          Booking_Master!C2:C10000,
          "Unknown"
        ),
        rev,
        SUMIFS(
          Raw_Transactions!O2:O10000,
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!N2:N10000,"Revenue",
          Raw_Transactions!P2:P10000,"OK"
        ),
        cogs,
        SUMIFS(
          Raw_Transactions!O2:O10000,
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!N2:N10000,"COGS",
          Raw_Transactions!P2:P10000,"OK"
        ),
        fees,
        SUMIFS(
          Raw_Transactions!O2:O10000,
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!N2:N10000,"Bank Fee",
          Raw_Transactions!P2:P10000,"OK"
        ),
        profit,
        rev-cogs-fees,
        margin,
        IF(rev=0,0,profit/rev),
        bad_rows,
        COUNTIFS(
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!P2:P10000,"<>OK"
        ),
        warn_msg,
        IF(
          bad_rows>0,
          "⚠️ "&bad_rows&" TXN DATA ERRORS",
          "CLEAN"
        ),
        HSTACK(
          bk,cust,rev,cogs,fees,
          profit,margin,
          dep,warn_msg
        )
      )
    )
  )
)
```

El control esencial es el criterio repetido `Data Health Status = "OK"`. 

</details>

### Reglas de validación

| Campo / Condición      | Regla                                                            | Comportamiento ante errores                                                                     |
| ---------------------- | --------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `Booking ID`           | Debe existir en `Booking_Master`                                  | `ERR: INVALID_BOOKING`; transacción excluida del P&L                              |
| `Service Code`         | Debe existir en `Service_Master`                                  | `ERR: INVALID_SERVICE`; transacción excluida del P&L                              |
| `Currency`             | Debe tener una tasa FX configurada válida o una tasa FX manual positiva | La tasa FX aplicada pasa a ser `0`; el diagnóstico informa `ERR: UNKNOWN_CURRENCY`                 |
| `Manual FX Rate`       | Un valor positivo anula la tasa FX de referencia                  | Un valor inválido o no positivo recurre a la tasa FX configurada                             |
| `Transaction Type`     | Debe coincidir con la correspondencia de `Setup_Params`                             | Un tipo desconocido recurre conservadoramente a `COGS`                                   |
| `USD Equivalent`       | Debe permanecer numérico                                             | Nunca debe escribirse texto de diagnóstico en el campo de agregación numérico           |
| `Data Health Status`   | Debe ser `OK` para la agregación financiera                          | Las filas no OK se excluyen del P&L de la reserva                                      |
| `Container Count`      | No debe convertirse en cero al usarse como denominador de asignación        | Recurre a `1` para prevenir `#DIV/0!`                                             |
| Asignación de costos compartidos | Controlada por `Setup_Params!B6`                                 | `EQUAL_SHARE` asigna según el número de contenedores; `NO_ALLOCATE` deja solo los costos directos |
| Rentabilidad de la reserva  | Ingresos de cero no deben producir un error de división               | El margen devuelve `0`                                                                 |
| Integridad de datos ejecutiva  | Se cuenta cualquier estado no vacío distinto de `OK`                 | El panel muestra una advertencia en lugar de presentar un estado limpio sin salvedades                      |

La validación cruzada de origen confirma la cobertura de las entradas básicas de reserva, contenedor, servicio y transacción, además de las salidas equivalente en USD, beneficio neto, margen y costo asignado. También valida específicamente la separación entre el texto de diagnóstico y los campos numéricos de `SUMIFS`. 

<!-- Technical Details parent <details> intentionally remains open for Part 2. -->
### Validación cruzada y preparación para el despliegue

La arquitectura del libro de trabajo está diseñada para validar la cadena analítica **antes** de que la dirección confíe en el P&L resultante.

La implementación de origen comprueba explícitamente la cadena de correspondencia de tipos de transacción, la cadena de búsqueda FX, la vinculación con Service Master, la cadena de asignación de costos compartidos y la compatibilidad entre el texto de diagnóstico y los campos de agregación numéricos. 

| Dimensión de validación         | Qué se comprueba                                                                  | Resultado esperado                                          |
| ---------------------------- | -------------------------------------------------------------------------------- | -------------------------------------------------------- |
| **Vinculación de reservas**          | Cada `Booking ID` de transacción existe en `Booking_Master`                        | Las filas inválidas reciben un estado de error                      |
| **Vinculación de servicios**          | Cada `Service Code` de transacción existe en `Service_Master`                      | Las filas inválidas se excluyen del P&L                       |
| **Cobertura FX**              | Cada transacción tiene una tasa FX manual válida o una tasa de referencia configurada | Las monedas desconocidas disparan una advertencia de integridad              |
| **Clasificación financiera** | El tipo de transacción se asigna a Revenue / COGS / Bank Fee / Deposit                     | Los tipos desconocidos usan la alternativa conservadora COGS         |
| **Agregación numérica**      | `USD Equivalent` permanece numérico en todas las condiciones                                    | `SUMIFS` no encuentra texto de diagnóstico              |
| **Agregación condicionada por integridad** | El P&L solo incluye `Data Health Status = "OK"`                                    | Las filas de origen defectuosas no pueden entrar silenciosamente en el beneficio declarado |
| **Asignación de contenedores**     | Los costos compartidos a nivel de reserva usan la política de asignación configurada                  | Ninguna asignación repetida del costo completo entre contenedores       |
| **Protección ante ingresos cero**  | El cálculo del margen maneja ingresos = 0                                    | Sin `#DIV/0!`                                             |
| **Seguimiento ejecutivo**     | Los errores globales de transacción se cuentan por separado                                 | El panel señala problemas de datos de origen sin resolver          |

La matriz de validación cruzada de la implementación establece que los campos de entrada básicos — reserva, contenedor, servicio y datos de transacción — y las salidas principales — equivalente en USD, beneficio neto, margen y costo asignado — están cubiertos por la arquitectura. 

### SOP operativo

El modelo operativo previsto es deliberadamente simple:

```text
1. Establish Master Data
        ↓
2. Record Financial Transactions
        ↓
3. Check Data Health Status
        ↓
4. Review Booking / Service / Container Profitability
        ↓
5. Resolve Exceptions
        ↓
6. Use Clean Results for Management Decisions
```

#### 1. Establecer los datos maestros

Cree la reserva en `Booking_Master`.

Si los contenedores son relevantes, cree sus relaciones reserva-contenedor en `Container_Master`.

Mantenga el catálogo de servicios estandarizado en `Service_Master`.

No comience el análisis financiero tecleando nombres de servicio libres ni identificadores de reserva directamente en las hojas de informes. Las tablas maestras son la capa de referencia para las cadenas de búsqueda posteriores. 

#### 2. Registrar las transacciones financieras

Introduzca ingresos, costos de proveedores, comisiones bancarias, depósitos y otra actividad financiera relevante en `Raw_Transactions`.

Cada transacción debe contener lo que corresponda de:

* `Booking ID`
* `Container ID`, cuando proceda
* `Service Code`
* `Transaction Type`
* `Allocation Level`
* `Original Amount`
* `Currency`
* `Manual FX Rate`, cuando se requiera una tasa específica de la transacción

La tabla de transacciones es la única fuente que debería recibir los registros detallados de eventos financieros. El motor de cálculo no debe editarse manualmente. 

#### 3. Comprobar la integridad de los datos

Revise `Raw_Transactions[Data Health Status]`.

Una fila limpia devuelve:

```text
OK
```

Las excepciones identifican la relación rota, como:

```text
ERR: INVALID_BOOKING
ERR: INVALID_SERVICE
ERR: UNKNOWN_CURRENCY
```

La implementación trata esto deliberadamente como una capa de control separada. Una transacción puede, por tanto, permanecer numéricamente segura y a la vez quedar visiblemente excluida del resultado financiero. 

#### 4. Revisar la rentabilidad

Use:

* `Job_Costing_Engine` para la rentabilidad a nivel de reserva.
* `Booking_PL` para el desglose servicio por servicio de una reserva.
* `Container_Analysis` cuando la reserva contiene varios contenedores.
* `Service_Profitability` para la economía de servicios a través de las reservas.
* `Executive_Dashboard` para el seguimiento a nivel de dirección.

Esto crea una ruta de desglose desde la **rentabilidad global → reserva → servicio → contenedor → transacción**, en lugar de obligar a responder cada pregunta desde un único panel resumen.

### Reglas de mantenimiento de datos

La arquitectura se apoya en datos maestros controlados porque Excel no ofrece el mismo cumplimiento estricto de claves foráneas que una base de datos o una plataforma ERP.

Las reglas de mantenimiento más importantes son:

1. **No sobrescriba los campos generados por fórmulas.**
2. **No teclee nombres de servicio a mano donde debería usarse un código de servicio.**
3. **No introduzca una moneda nueva sin añadir primero su tasa FX a `Setup_Params`.**
4. **No use `1.0` como alternativa para una tasa FX faltante.**
5. **No escriba texto de error en las columnas de agregación numéricas.**
6. **No edite manualmente la salida de `Job_Costing_Engine`.**
7. **Resuelva las filas de transacción no OK antes de tratar el margen de una reserva como definitivo.**
8. **Mantenga la política de asignación de costos compartidos centralizada en `Setup_Params`.**

El diseño de origen prohíbe específicamente codificar parámetros globales de forma rígida dentro de las fórmulas y centraliza la moneda base, las tasas FX, las correspondencias de transacciones y la política de asignación en `Setup_Params`. 

### Solución de problemas

#### `⚠️ X TXN DATA ERRORS` aparece en el panel

**Causas probables:**

* Falta la moneda en la tabla FX.
* El `Booking ID` no existe en `Booking_Master`.
* El `Service Code` no existe en `Service_Master`.
* Una transacción de origen contiene un vínculo inválido.

**Resolución:**

Filtre la columna `Data Health Status` en `Raw_Transactions` para mostrar los valores distintos de `OK`.

Luego resuelva cada problema de datos de origen subyacente en lugar de cambiar manualmente la salida del P&L.

El Executive Dashboard está diseñado para exponer esta condición mediante una tarjeta de integridad global:

```excel
=LET(
    total_errs,
    COUNTIFS(
        Raw_Transactions!P2:P10000,
        "<>OK",
        Raw_Transactions!P2:P10000,
        "<>"
    ),
    IF(
        total_errs=0,
        "✅ ALL DATA HEALTHY",
        "⚠️ ATTENTION: " &
        total_errs &
        " ERRORS IN RAW TRANSACTIONS"
    )
)
```



#### Las fórmulas de matriz dinámica devuelven `#SPILL!`

Una fórmula de desbordamiento requiere que su área de destino permanezca despejada.

Compruebe:

* las celdas directamente debajo de la fórmula;
* las celdas a la derecha de la fórmula;
* los valores introducidos manualmente que ocupan el rango de desbordamiento previsto;
* el formato residual o el contenido copiado que haya creado una obstrucción.

Despeje las celdas que obstruyen y deje que la matriz dinámica se expanda de nuevo.

#### La rentabilidad aparece inesperadamente alta

No cambie inmediatamente la fórmula de beneficio.

Compruebe primero:

```text
Raw Transactions
       ↓
Data Health Status
       ↓
Unknown Currency?
Invalid Booking?
Invalid Service?
       ↓
Only OK rows enter P&L
```

Un margen sospechosamente alto puede deberse a una transacción de costo que falló en la cadena de vinculación y quedó, por tanto, excluida del cálculo.

Esta es exactamente la razón por la que el libro de trabajo distingue **"el cero como estado de cálculo numérico"** de **"el cero como conclusión económica"**. Una tasa FX faltante puede producir `$0.00` en un campo numérico y, al mismo tiempo, producir `ERR: UNKNOWN_CURRENCY` en el campo de diagnóstico. 

#### La rentabilidad del contenedor aparece sobreestimada o subestimada

Compruebe tres cosas por orden:

1. ¿Está seleccionado el `Booking ID` correcto?
2. ¿Está seleccionado el `Container ID` correcto?
3. ¿Usa `Setup_Params!B6` la política de asignación prevista?

Bajo `EQUAL_SHARE`, el COGS compartido a nivel de reserva se divide entre el número de contenedores de la reserva antes de sumarse a los costos directos del contenedor.

Bajo `NO_ALLOCATE`, los costos compartidos a nivel de reserva no se asignan a contenedores individuales.

La distinción importa porque asignar el costo compartido completo a nivel de reserva a cada contenedor duplicaría el mismo costo económico. 

### Límites de la implementación

Este libro de trabajo está diseñado intencionadamente como una **capa de control analítica ligera**, no como un sustituto de una plataforma contable o de un sistema de gestión del transporte.

Es más fuerte cuando:

* los datos operativos y financieros ya existen en hojas de cálculo o exportaciones;
* la unidad de análisis clave es una reserva o un envío;
* la dirección necesita rentabilidad a nivel de trabajo;
* los servicios y cargos deben compararse entre reservas;
* varias monedas necesitan una base de información coherente;
* los costos compartidos requieren una regla de asignación explícita;
* los usuarios necesitan fórmulas transparentes en lugar de automatización opaca.

No está diseñado para proporcionar:

* concurrencia transaccional multiusuario;
* permisos de nivel ERP;
* conciliación bancaria automatizada;
* ingesta automatizada de facturas de transportistas;
* cumplimiento de claves foráneas a nivel de base de datos;
* orquestación de flujos de trabajo empresariales.

El valor del libro de trabajo no reside, por tanto, en el número de hojas que contiene. Es la cadena analítica controlada que conecta **reserva → transacción → clasificación → conversión de moneda → agregación validada → rentabilidad → desglose**.

### Principio de control analítico

La regla de implementación más importante puede resumirse así:

```text
Never let a broken data relationship
look like a valid financial result.
```

Una hoja de cálculo convencional podría permitir:

```text
Missing FX
    ↓
0 USD
    ↓
SUMIFS
    ↓
Higher Profit
```

Este libro de trabajo, en cambio, hace explícito el estado:

```text
Missing FX
    ↓
0 USD Equivalent
    +
ERR: UNKNOWN_CURRENCY
    ↓
Excluded from P&L
    +
Booking Health Warning
    ↓
Correct the source data
    ↓
Recalculate
```

Esa distinción es el fundamento del modelo de fiabilidad del Job Costing Engine. La arquitectura reconoce explícitamente que `SUMIFS` puede agregar con seguridad valores numéricos mientras ignora silenciosamente el texto, de modo que los estados de diagnóstico nunca deben incrustarse directamente dentro de las columnas de cálculo numéricas. 

### Secuencia de despliegue recomendada

Para una nueva implementación del libro de trabajo, la secuencia más segura es:

```text
Phase 1
Setup_Params
    ↓
Define Base Currency
Define FX Rates
Define Transaction Types
Define Financial Treatments
Define Allocation Policy

Phase 2
Booking_Master
Container_Master
Service_Master
    ↓
Establish Reference Data

Phase 3
Raw_Transactions
    ↓
Load Historical / Current Financial Data

Phase 4
Data Health Review
    ↓
Resolve Invalid Bookings
Resolve Invalid Services
Resolve Unknown Currencies

Phase 5
Job_Costing_Engine
    ↓
Validate Revenue / COGS / Fees / Profit

Phase 6
Booking_PL
Container_Analysis
Service_Profitability
Executive_Dashboard
    ↓
Operational Review
```

El resultado es un flujo de trabajo repetible en lugar de un cálculo puntual de hoja de cálculo: establezca los datos de referencia una vez, añada nuevas transacciones operativas, resuelva las excepciones y deje que las capas de cálculo y análisis se actualicen desde el mismo modelo de datos subyacente. 

<!-- Part 2 ends with the Technical Details parent <details> intentionally OPEN. Do not add </details> here. -->
### Notas finales de arquitectura

El libro de trabajo está construido deliberadamente en torno a un **único motor de rentabilidad a nivel de reserva**, en lugar de una colección de calculadoras desconectadas.

La jerarquía analítica es:

```text
Booking
 ├── Customer
 ├── Containers
 │    └── Container-level direct + allocated costs
 └── Transactions
      ├── Revenue
      ├── COGS
      ├── Bank Fees
      └── Deposits
           ↓
      Service-level profitability
           ↓
      Booking-level P&L
           ↓
      Executive-level profitability & data health
```

Esta estructura mantiene la reserva como unidad económica central y permite al mismo tiempo a los usuarios desglosar hacia abajo en servicios y contenedores o hacia arriba en la rentabilidad entre reservas. 

La implementación también reconoce explícitamente tres cadenas de búsqueda de múltiples saltos:

1. **Transacción → Tratamiento financiero → FX → Equivalente en USD → P&L de la reserva**
2. **Reserva → Recuento de contenedores → Política de costos compartidos → Asignación de costos compartidos → Análisis de contenedores**
3. **Código de servicio → Service Master → Categoría de servicio → Rentabilidad del servicio**

Estas cadenas se validan antes de considerar listas para producción las salidas analíticas. 

### Contra lo que protege el motor

El principio de diseño central no es la mera automatización. Es **impedir que una relación de entrada rota se disfrace de conclusión financiera válida**.

| Fallo                                              | Mecanismo de protección                                                           |
| ---------------------------------------------------- | ------------------------------------------------------------------------------ |
| Moneda desconocida                                     | La FX recurre a `0`, mientras que un estado de diagnóstico separado informa de la tasa faltante |
| Tipo de transacción desconocido                                     | La alternativa conservadora `COGS` impide la inflación accidental del beneficio              |
| Reserva inválida                                      | `ERR: INVALID_BOOKING`                                                         |
| Código de servicio inválido                                      | `ERR: INVALID_SERVICE`                                                         |
| Cadena de transacciones rota                             | Las filas no `OK` se excluyen de la agregación del P&L                                |
| Costo compartido duplicado entre contenedores             | Política de asignación explícita y denominador basado en el recuento de contenedores                     |
| Ingresos cero                                         | El cálculo del margen devuelve `0` en lugar de `#DIV/0!`                            |
| Texto de diagnóstico que contamina los cálculos financieros | Las columnas de cálculo numérico y las columnas de diagnóstico permanecen separadas             |

La matriz de alternativas de la implementación rechaza específicamente una alternativa FX de `1.0` porque podría convertir un importe no reconocido en moneda extranjera en un importe en USD implausiblemente grande. 

### Patrón de uso recomendado

Para operaciones continuas, el libro de trabajo está pensado para comportarse como un **modelo operativo actualizable**, no como un informe puntual.

```text
Maintain Setup_Params
        ↓
Maintain Master Data
        ↓
Append New Transactions
        ↓
Resolve Data Health Exceptions
        ↓
Review Booking Profitability
        ↓
Drill Into Service / Container
        ↓
Use Clean Results for Decisions
```

Esto hace que el libro de trabajo sea útil para preguntas recurrentes como:

* ¿Qué reservas son realmente rentables?
* ¿Qué servicios producen los márgenes más fuertes?
* ¿Qué costos de envío están provocando una pérdida inesperada?
* ¿Qué costos a nivel de reserva deben asignarse a los contenedores?
* ¿Son fiables los márgenes declarados, o hay problemas de datos de origen sin resolver que ocultan costos?
* ¿Dónde debería centrarse a continuación la fijación de precios, la negociación con proveedores o la revisión operativa?

El SOP de origen sigue la misma secuencia: establecer los datos maestros, registrar las transacciones financieras, inspeccionar `Data Health Status` y luego usar los paneles de rentabilidad para la revisión. 

### Principio de diseño final

El libro de trabajo no intenta hacer que el análisis financiero parezca más simple ocultando su incertidumbre.

En cambio, hace visible la incertidumbre.

Una reserva con un margen calculado no se trata automáticamente como una reserva digna de confianza.

Una **reserva limpia** significa:

```text
Valid Booking
+ Valid Service
+ Valid Currency / FX
+ Valid Transaction Classification
+ Healthy Transaction Chain
= Profitability Result That Can Be Reviewed
```

Esa distinción es el núcleo de la implementación.

</details>

## Otras herramientas de esta serie

Una colección de herramientas Excel ligeras de apoyo a la decisión que cubren el control operativo, el análisis de rentabilidad, la conciliación de inventarios, la planificación de personal y otros flujos de trabajo de negocio.

* **[rolling-cashflow-forecast](https://github.com/HyVoid/rolling-cashflow-forecast)** — ¿Por qué siempre falta efectivo? — Proyecte la liquidez futura bajo supuestos operativos y financieros cambiantes.
* **[Multi-Currency-Trade-Financial-Analyzer](https://github.com/HyVoid/Multi-Currency-Trade-Financial-Analyzer)** — ¿Qué flujos comerciales son realmente rentables tras los efectos cambiarios? — Mida la rentabilidad comercial multimoneda en operaciones transfronterizas.
* **[RTL_Loan_Sizer_Pricer](https://github.com/HyVoid/RTL_Loan_Sizer_Pricer)** — ¿Qué estructura de financiamiento es sostenible bajo distintos escenarios de prestatario y garantías? — Compare escenarios de préstamo usando supuestos de financiamiento ajustados al riesgo.
* **[Personal-Business-Dual-Track-Budget-Manager](https://github.com/HyVoid/Personal-Business-Dual-Track-Budget-Manager)** — ¿Dónde se mezclan o malasignan el efectivo personal y el del negocio? — Separe los flujos de caja personal y empresariales en pistas presupuestarias con seguimiento independiente.
* **[Multifamily-Value-Add-Acquisition-Model](https://github.com/HyVoid/Multifamily-Value-Add-Acquisition-Model)** — ¿Esta adquisición inmobiliaria sigue creando valor patrimonial tras los supuestos de renovación y operación? — Evalúe los retornos de adquisición multifamiliar bajo financiamiento de renovación y escenarios operativos.
* **[Real-Estate-Valuation-all-in-one-calculator](https://github.com/HyVoid/Real-Estate-Valuation-all-in-one-calculator)** — ¿Cuánto vale una propiedad según los distintos métodos de valoración? — Calcule el valor de la propiedad con múltiples métodos de valoración, incluyendo modelado de impuestos y gastos de capital.
* **[multi-entity-expense-allocation](https://github.com/HyVoid/multi-entity-expense-allocation)** — ¿Se asignan los costos compartidos de forma justa entre entidades, o algunos activos están sobrecargados mientras otros escapan de su exposición real de costos? — Distribuya facturas de proveedores entre múltiples entidades legales y activos operativos usando metodologías de asignación configurables.
* **[investment-pool-nav-excel-template](https://github.com/HyVoid/investment-pool-nav-excel-template)** — ¿Qué se debe realmente a cada participante según su momento de entrada y el valor actual del fondo? — Siga la propiedad de un fondo de inversión de forma continua mediante contabilidad de unidades basada en el NAV.
* **[shopify-cash-flow-forecast-excel-toolkit](https://github.com/HyVoid/shopify-cash-flow-forecast-excel-toolkit)** — ¿Qué compromisos operativos crearán un problema de efectivo en las próximas doce semanas? — Conecte ventas de Shopify, transacciones bancarias, gasto publicitario y compras de inventario en un marco unificado de decisiones de flujo de caja.
* **[Revenue-Management-Decision-Engine](https://github.com/HyVoid/Revenue-Management-Decision-Engine)** — ¿Ha cambiado suficiente información relevante como para justificar un ajuste de precios? — Gobierne las recomendaciones de precios mediante barandillas configurables que eviten una volatilidad innecesaria de ingresos.
* **[amazon-wbr-sqp-analytics-template](https://github.com/HyVoid/amazon-wbr-sqp-analytics-template)** — ¿Qué cambió realmente en el negocio este periodo y qué productos o palabras clave lo impulsaron? — Transforme las exportaciones de Amazon Seller Central SQP y Sellerboard en revisiones empresariales estandarizadas para la dirección.
* **[food-manufacturing-cost-calculator-excel](https://github.com/HyVoid/food-manufacturing-cost-calculator-excel)** — ¿Qué productos son realmente rentables una vez asignados todos los costos de fabricación y cuáles están destruyendo el margen? — Asigne mano de obra de fábrica, servicios, gastos generales y costos de materiales a cada SKU para una verdadera economía unitaria de fabricación.
* **[ecommerce-profit-calculator-excel](https://github.com/HyVoid/ecommerce-profit-calculator-excel)** — ¿Qué pedidos son realmente rentables tras todos los costos ocultos y dónde se fuga el dinero entre canales y monedas? — Reconstruya la rentabilidad real a nivel de pedido conciliando pagos de ventas, costos de envío, comisiones de pago y comisiones de marketplace en operaciones de comercio electrónico multicanal.
* **[construction-job-costing-excel-template](https://github.com/HyVoid/construction-job-costing-excel-template)** — ¿Qué proyectos de contratista y clientes son realmente rentables tras todos los costos del trabajo? — Conecte estimación de leads, ejecución del trabajo y seguimiento de costos reales en un marco unificado de rentabilidad para contratistas.
* **[Flora-Pastry-Coffee-Shop-Operating-Control-System](https://github.com/HyVoid/Flora-Pastry-Coffee-Shop-Operating-Control-System)** — ¿Es la tienda realmente rentable tras todos los costos de ingredientes, mano de obra y operación? — Conecte costeo de recetas, ventas, consumo de inventario y gastos operativos en un marco unificado de rentabilidad y flujo de caja.
* **[cross-entity-logistics-console](https://github.com/HyVoid/cross-entity-logistics-console)** — ¿Qué contenedores y clientes son realmente rentables en las operaciones aduaneras y de flota? — Consolide los datos de despacho aduanero y ejecución de flota en una única vista de rentabilidad a nivel de contenedor.
* **[family-loan-tracker-excel](https://github.com/HyVoid/family-loan-tracker-excel)** — ¿Qué se debe realmente hoy tras años de transferencias e intereses acumulados? — Concilie transferencias de fondos a largo plazo y acumulación de intereses en un saldo pendiente actual.
* **[3-statement-financial-model-excel](https://github.com/HyVoid/3-statement-financial-model-excel)** — ¿Cómo se verán nuestros estados financieros bajo distintos escenarios operativos y de financiamiento? — Proyecte estados financieros integrados vinculando ingresos, gastos, capital de trabajo y deuda en un modelo unificado.
* **[assisted-living-cost-allocation-excel](https://github.com/HyVoid/assisted-living-cost-allocation-excel)** — ¿Cuánto del margen declarado de cada edificio lo genera realmente ese edificio y cuánto la forma en que se asignaron los costos corporativos? — Asigne gastos compartidos entre unidades operativas preservando las transacciones de origen y conciliando los resultados unitarios con los totales de la empresa.
* **[real-estate-development-feasibility-excel](https://github.com/HyVoid/real-estate-development-feasibility-excel)** — ¿A qué precio máximo de adquisición de terreno se convierte este proyecto en una inversión bancable? — Evalúe oportunidades de desarrollo residencial filtrando comparables de terreno, estimando costos de construcción y calculando umbrales de beneficio sobre costo.
* **[uk-residential-costing-software](https://github.com/HyVoid/uk-residential-costing-software)** — ¿Alcanzará realmente este proyecto de construcción su margen objetivo una vez considerados la logística de obra, los factores de desperdicio y los cambios en las tarifas de oficios? — Separe las cantidades físicas del edificio de las tarifas volátiles de oficios y los costos preliminares de obra para escalar los márgenes comerciales a través de presupuestos de licitación y cotizaciones a clientes.
* **[b2b-pricing-landed-cost-excel-model](https://github.com/HyVoid/b2b-pricing-landed-cost-excel-model)** — ¿Qué precios mayoristas y segmentos de clientes preservan realmente el margen una vez costeados por completo el flete, los aranceles, las condiciones de pago y la FX? — Calcule costos landed a nivel de SKU, costo de servicio y cascadas de precio de margen de bolsillo para simular la rentabilidad mayorista transfronteriza.
* **[custom-label-pricing-calculator](https://github.com/HyVoid/custom-label-pricing-calculator)** — ¿Cuánto deberíamos cobrar por un pedido de etiquetas personalizadas en cada tramo de cantidad para seguir siendo competitivos y rentables? — Calcule precios paramétricos de etiquetas modelando los costos por pulgada cuadrada de sustrato, cargos de configuración y descuentos por tramos de cantidad frente a referencias de la competencia.
* **[landed-cost-calculator-retail-deduction-excel](https://github.com/HyVoid/landed-cost-calculator-retail-deduction-excel)** — ¿Cuál es el margen neto real de cada SKU una vez contabilizados los costos landed reales y las deducciones de los minoristas? — Concilie los costos landed estimados del PO contra el flete, los aranceles, la FX y el drayage reales mientras audita las deducciones de remesa de los minoristas para exponer el margen neto a nivel de SKU.
* **[saas-finance-mrr-burn-rate-excel](https://github.com/HyVoid/saas-finance-mrr-burn-rate-excel)** — ¿Podemos asumir los compromisos de contratación, marketing y proveedores de esta semana sin quedarnos sin efectivo antes del cierre contable? — Siga el ritmo semanal de quema de efectivo, las variaciones de la cascada MRR/ARR, las métricas de retención y los costos de nómina frente al presupuesto antes del cierre de fin de mes.
* **[startup-financial-model-excel](https://github.com/HyVoid/startup-financial-model-excel)** — ¿Se quedará sin efectivo nuestro plan de crecimiento antes de que la monetización de las cohortes lo alcance? — Proyecte quema de efectivo plurianual, runway, economía unitaria y dilución de financiación a partir de supuestos de cohortes bottom-up.
* **[shopify-pnl-reconciliation-excel](https://github.com/HyVoid/shopify-pnl-reconciliation-excel)** — ¿Es realmente rentable cada escaparate de Shopify una vez conciliados los costos reales a nivel de SKU, los gastos compartidos y los cargos bancarios? — Concilie las exportaciones de ventas multi-tienda de Shopify contra el COGS a nivel de SKU, el OPEX multimoneda asignado y los extractos bancarios para producir un P&L mensual fiable.
* **[travel-agency-accounting-excel-template](https://github.com/HyVoid/travel-agency-accounting-excel-template)** — ¿Qué reservas de tours generan realmente beneficio neto y flujo de caja positivo tras los costos de proveedores y las cobranzas B2B? — Calcule costos de trabajo por reserva de tour, márgenes brutos, liquidaciones AR/AP y flujo de caja operativo en múltiples monedas.

## Licencia

Este proyecto se publica bajo la **Apache License 2.0**.

Puede usar, modificar, distribuir y construir sobre esta obra conforme a los términos de la Apache License 2.0.
