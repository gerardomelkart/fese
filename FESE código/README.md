FESE 9-1-1

Scripts para procesar mensualmente los formatos FESE del servicio de atención de emergencias 9-1-1, generar los productos de salida, emitir comprobantes por entidad federativa, enviar dichos comprobantes por correo y generar comprobantes de meses anteriores cuando sea necesario.

Estructura del proyecto

fese/
│
├── README.md
│
└── FESE código/
    ├── fese.py
    ├── enviar_comprobantes.py
    ├── generar_comprobantes_pendientes.py
    ├── requirements.txt
    ├── .gitignore
    │
    ├── config/
    │   └── correos_comprobantes.txt
    │
    ├── datos/
    │   ├── .gitkeep
    │   └── feseAAAA-M.rds
    │
    ├── insumos/
    │   └── .gitkeep
    │
    ├── modulos/
    │   ├── init.py
    │   └── comprobantes.py
    │
    ├── plantillas/
    │   ├── COMPROBANTE FESE plantilla.pdf
    │   └── Formato CNIEDT plantilla.xlsx
    │
    └── salidas/
        └── .gitkeep

Scripts principales

fese.py

Es el proceso mensual principal.

Realiza, de forma general, lo siguiente:

1. Calcula automáticamente el periodo a procesar tomando como referencia el mes anterior a la fecha actual.
2. Carga desde datos/ el histórico RDS del mes previo.
3. Lee los formatos Excel colocados en insumos/.
4. Valida y normaliza la información recibida.
5. Incorpora el nuevo mes al histórico.
6. Genera el archivo anual Rep_anualAAAA-MM.xlsx.
7. Genera el nuevo histórico feseAAAA-M.rds.
8. Genera validaciones.json.
9. Llena el formato CNIEDT correspondiente al periodo.
10. Genera un comprobante PDF para cada entidad.
11. Copia los resultados finales a la carpeta FESE de OneDrive.
12. Conserva localmente los históricos necesarios y elimina históricos anteriores que ya no requiere el proceso.
13. Vacía insumos/ únicamente cuando el proceso termina correctamente.

enviar_comprobantes.py

Envía por Outlook los comprobantes del periodo mensual generado por fese.py.

El script:

- Lee los destinatarios desde config/correos_comprobantes.txt.
- Busca los comprobantes correspondientes al último mes cerrado.
- Utiliza la cuenta institucional configurada en el código.
- Puede trabajar en modo BORRADOR o ENVIAR.
- Valida que exista un comprobante para cada entidad antes de comenzar.
- En modo ENVIAR, solicita confirmación antes de mandar correos reales.
- Registra los envíos exitosos para evitar duplicados en futuras ejecuciones.

El registro se guarda dentro de la carpeta de comprobantes del periodo como:

registro_envios_comprobantes.json

generar_comprobantes_pendientes.py

Genera comprobantes de meses anteriores sin modificar el proceso mensual.

El script:

- Busca automáticamente el RDS histórico más reciente en datos/.
- Muestra un menú con las entidades disponibles.
- Permite elegir el año.
- Permite elegir uno o varios meses.
- Usa el mismo diseño y lógica del comprobante mensual.
- Genera los PDF localmente.
- No modifica el RDS.
- No modifica insumos/.
- No envía correos.

Los archivos se generan en:

FESE código/comprobantes_pendientes/ENTIDAD/

modulos/comprobantes.py

Módulo interno utilizado por los otros scripts para construir los comprobantes PDF.

No es necesario ejecutarlo directamente.

---

Requisitos

El proyecto está diseñado para ejecutarse en Windows.

Se requiere:

- Python instalado y disponible mediante el comando py.
- Microsoft Excel de escritorio.
- Microsoft Outlook de escritorio.
- Acceso a la cuenta institucional utilizada para el envío.
- Acceso a la carpeta institucional de OneDrive utilizada por FESE.
- Los paquetes de requirements.txt.

Dependencias Python

Actualmente el proyecto utiliza:

pandas
numpy
openpyxl
pyreadr
unidecode
pywin32
pypdf
reportlab

---

Instalación

1. Clonar el repositorio

git clone https://github.com/gerardomelkart/fese.git
cd fese

2. Entrar a la carpeta del código

cd ".\FESE código"

3. Verificar Python

py --version

4. Instalar las dependencias

py -m pip install --upgrade pip
py -m pip install -r requirements.txt

5. Verificar Microsoft Office

Excel y Outlook deben estar instalados como aplicaciones de escritorio.

Para el envío de correos, Outlook debe tener configurada la cuenta institucional indicada en enviar_comprobantes.py.

6. Preparar el histórico

Los archivos .rds no se guardan en Git.

Por lo tanto, después de clonar el repositorio se debe copiar manualmente a:

FESE código/datos/

el histórico necesario para continuar el proceso.

Ejemplo: para procesar agosto de 2026 debe existir como mínimo:

datos/fese2026-7.rds

El nombre sigue la forma:

feseAAAA-M.rds

7. Revisar la carpeta de destino

Actualmente los scripts principales utilizan como carpeta final:

C:\Users\gerardo.noeller\OneDrive - Secretaría de Seguridad y Protección Ciudadana\Escritorio\FESE

Si el proyecto se instala con otro usuario de Windows o en otro equipo, debe actualizarse CARPETA_DESTINO en:

fese.py
enviar_comprobantes.py

---

Uso mensual normal

1. Reunir los formatos

Reunir los formatos FESE de las 32 entidades federativas correspondientes al periodo.

2. Copiar los formatos a insumos

Colocar todos los Excel dentro de:

FESE código/insumos/

No deben colocarse ahí los archivos históricos RDS.

3. Ejecutar FESE

Desde:

FESE código

ejecutar:

py ".\fese.py"

No es necesario indicar manualmente el mes. El script calcula automáticamente el último mes cerrado.

Por ejemplo, si se ejecuta durante agosto de 2026, procesa julio de 2026.

4. Revisar el resultado

Si termina correctamente, la consola debe finalizar sin Traceback y mostrar el mensaje de proceso terminado junto con los tiempos de ejecución.

Localmente se generan principalmente:

datos/feseAAAA-M.rds
salidas/Rep_anualAAAA-MM.xlsx
salidas/validaciones.json

En la carpeta final de OneDrive se generan o copian, entre otros:

Rep_anualAAAA-MM.xlsx
feseAAAA-M.rds
Formato CNIEDT AAAA-MM.xlsx
Formatos MES AAAA/
Comprobantes MES AAAA/

La carpeta insumos/ se vacía automáticamente al terminar correctamente.

Si el proceso falla antes de completar las salidas, no se debe ejecutar el envío de correos.

---

Envío mensual de comprobantes

El envío de correos es un proceso separado. Ejecutar nuevamente fese.py no manda correos.

1. Revisar destinatarios

Los destinatarios se encuentran en:

config/correos_comprobantes.txt

Formato:

Entidad;Correo
Aguascalientes;correo@dominio.gob.mx
Baja California;correo@dominio.gob.mx
...

Si una entidad cambia de correo, actualizar este archivo antes del envío.

2. Seleccionar modo

Antes de ejecutar el script, revisar esta variable en enviar_comprobantes.py:

MODO = "ENVIAR"  # BORRADOR o ENVIAR

El repositorio puede quedar configurado en cualquiera de los dos modos, por lo que siempre debe revisarse antes de ejecutar.

Usar:

MODO = "BORRADOR"

para realizar pruebas y crear borradores en Outlook sin enviar.

Usar:

MODO = "ENVIAR"

únicamente cuando los comprobantes, destinatarios y periodo ya hayan sido revisados.

3. Ejecutar

py ".\enviar_comprobantes.py"

En modo BORRADOR, los mensajes se guardan como borradores en Outlook.

En modo ENVIAR, el script solicita una confirmación:

Escribe ENVIAR para continuar:

Solo después de escribir exactamente:

ENVIAR

comienza el envío real.

4. Protección contra envíos duplicados

Cada envío correcto se registra en:

Comprobantes MES AAAA/registro_envios_comprobantes.json

Si se vuelve a ejecutar el script para el mismo periodo, una entidad ya registrada se muestra como:

Ya enviado, se omite: Entidad

y no se envía nuevamente.

Si se realizó una prueba real a un correo incorrecto

Si una entidad quedó registrada durante una prueba real, se debe:

1. Abrir registro_envios_comprobantes.json.
2. Localizar únicamente la entrada de la entidad y periodo afectados.
3. Eliminar esa entrada cuidando que el JSON siga siendo válido.
4. Confirmar que config/correos_comprobantes.txt tenga el destinatario correcto.
5. Ejecutar nuevamente enviar_comprobantes.py.

No se debe borrar todo el registro de envíos si las demás entidades ya fueron enviadas correctamente.

---

Generar comprobantes de meses anteriores

Este proceso sirve cuando una entidad solicita uno o varios comprobantes retroactivos.

No es necesario volver a ejecutar el proceso mensual.

Ejecutar:

py ".\generar_comprobantes_pendientes.py"

El script:

1. Carga automáticamente el histórico RDS más reciente.
2. Muestra las entidades disponibles.
3. Solicita la entidad.
4. Solicita el año.
5. Solicita los meses.
6. Solicita confirmación.
7. Genera los comprobantes solicitados.

Ejemplo:

Selecciona entidad: 6

Años disponibles: 2020, 2021, 2022, 2023, 2024, 2025, 2026
Año [2026]:

Meses a generar (ej. 3,4,5,6 o 3-6): 3-6

Entidad: Coahuila
Año: 2026
Meses: Marzo, Abril, Mayo, Junio

Escribe GENERAR para continuar: GENERAR

También pueden indicarse meses separados:

3,5,6

o un rango:

3-6

Los archivos quedan en:

comprobantes_pendientes/ENTIDAD/

Ejemplo:

comprobantes_pendientes/COAHUILA/
├── COMPROBANTE FESE MARZO 2026 - COAHUILA.pdf
├── COMPROBANTE FESE ABRIL 2026 - COAHUILA.pdf
├── COMPROBANTE FESE MAYO 2026 - COAHUILA.pdf
└── COMPROBANTE FESE JUNIO 2026 - COAHUILA.pdf

Estos PDF no se envían automáticamente.

---

Plantillas

El proyecto requiere:

plantillas/COMPROBANTE FESE plantilla.pdf
plantillas/Formato CNIEDT plantilla.xlsx

No cambiar el nombre ni mover estos archivos sin actualizar también las rutas del código.

Si cambia oficialmente alguno de los formatos, primero debe verificarse si la estructura interna sigue siendo compatible con los scripts.

---

Archivos que no se guardan en Git

El .gitignore excluye archivos operativos o generados como:

- RDS históricos.
- Insumos mensuales.
- Excel de salida.
- validaciones.json.
- comprobantes pendientes generados localmente.
- archivos temporales.
- pycache.
- archivos .pyc.

Esto es intencional.

En una instalación nueva, los históricos RDS deben obtenerse del respaldo operativo y copiarse manualmente a datos/.

---

Problemas comunes

py no se reconoce

Python o el Python Launcher no están instalados o no están disponibles en PATH.

Verificar:

py --version

Falta el histórico RDS

Ejemplo de error:

No se encuentra feseAAAA-M.rds

Copiar el histórico correcto a:

datos/

No se encuentra la plantilla

Confirmar que existan:

plantillas/COMPROBANTE FESE plantilla.pdf
plantillas/Formato CNIEDT plantilla.xlsx

Outlook no encuentra la cuenta

El envío requiere que Outlook tenga configurada la cuenta institucional indicada en:

CUENTA_REMITENTE = "admin.bnext.cni@sspc.gob.mx"

Faltan comprobantes para alguna entidad

No enviar.

Primero revisar que fese.py haya generado correctamente los comprobantes de las 32 entidades y que los nombres coincidan con config/correos_comprobantes.txt.

Archivo bloqueado

Si Excel, un RDS o un archivo de salida está abierto o bloqueado por otra aplicación, el proceso puede impedir su reemplazo.

Cerrar el archivo y volver a ejecutar el proceso.

---

Flujo resumido mensual

Reunir los 32 formatos
        ↓
Copiarlos a insumos/
        ↓
py .\fese.py
        ↓
Revisar resultados
        ↓
Revisar comprobantes y correos
        ↓
py .\enviar_comprobantes.py

Para comprobantes anteriores:

py .\generar_comprobantes_pendientes.py
        ↓
Seleccionar entidad
        ↓
Seleccionar año y meses
        ↓
GENERAR
        ↓
comprobantes_pendientes/ENTIDAD/

---

Recomendaciones operativas

- No borrar el RDS del último mes procesado.
- Antes de ejecutar el proceso mensual, verificar que exista el RDS del mes anterior.
- No ejecutar enviar_comprobantes.py hasta revisar los resultados del proceso mensual.
- Para pruebas de correo, utilizar MODO = "BORRADOR".
- Mantener actualizado config/correos_comprobantes.txt.
- No modificar manualmente los RDS.
- No borrar registro_envios_comprobantes.json después de un envío real.
- Conservar respaldos de los resultados finales en la ubicación institucional correspondiente.
