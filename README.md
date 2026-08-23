# Command Lab — generador de comandos administrativos/diagnósticos

## Qué es

Un pequeño conjunto de scripts CLI interactivos que generan, ya listos para copiar y
usar, comandos de administración y diagnóstico habituales en Windows y Linux/macOS
(instalar un MSI, mapear un recurso SMB, resolver un DNS, comprobar un puerto, calcular
un hash, etc.), evitando tener que recordar la sintaxis exacta, el quoting o el orden de
los parámetros de cada herramienta.

Las categorías que organizan el menú están tomadas de una taxonomía de técnicas y
comportamientos previamente clasificada, agrupada y saneada a partir de un dataset de
seguridad más amplio (ver [Procedencia de la taxonomía](#procedencia-de-la-taxonomía)).
Los scripts no reproducen contenido crudo de ese dataset: cada opción genera una
plantilla benigna, administrativa o de diagnóstico.

## Para qué sirve

- Elegir una categoría de técnica/herramienta y ver en 2-3 frases para qué sirve y qué
  necesita.
- Elegir una operación concreta dentro de esa categoría.
- Responder únicamente las preguntas necesarias (host, puerto, ruta, URL...), con una
  explicación breve de qué es cada dato y de dónde suele obtenerse.
- Obtener el comando final, correctamente escapado para el shell de destino
  (Bash, cmd.exe o PowerShell), y copiado al portapapeles cuando el sistema lo permite.
- Volver al menú de categorías en cualquier momento, sin quedar atrapado en una pantalla.

## Cómo funciona

```
seleccionar categoría → elegir operación → introducir parámetros → generar comando → copiar/utilizar el resultado
```

Cada categoría explica primero para qué sirve la herramienta subyacente y qué te va a
pedir. Cada operación pide solo los parámetros que necesita, valida lo básico (puerto en
rango, letra de unidad con formato `Z:`, valores que no rompan el comando) y muestra el
resultado con separadores claros antes de volver al menú.

## Archivos

| Fichero | Qué hace |
|---|---|
| `command-builder-linux.sh` | Generador interactivo de comandos (Bash). 14 categorías, genera el comando y lo copia al portapapeles (`wl-copy`/`xclip`/`xsel`/`pbcopy`, el primero disponible). |
| `command-builder-windows.ps1` | Equivalente en PowerShell. Mismas 14 categorías y misma estructura conceptual; copia con `Set-Clipboard`. |
| `launcher-linux.sh` | Explorador de la taxonomía completa (16 categorías, 1.483 grupos, metadatos del dataset) más una batería de autopruebas locales y benignas (roundtrip Base64, roundtrip de ofuscación, ejecución multilínea, ejecución oculta, una descarga HTTP real pero contra un servidor local en 127.0.0.1). No contacta infraestructura externa ni ejecuta muestras reales. |
| `launcher-windows.ps1` | Mismo explorador y mismas autopruebas, en PowerShell, usando los LOLBins nativos de Windows (msiexec, mshta, rundll32, conhost, pcalua, cscript...) contra artefactos locales inofensivos. |
| `reference/excel_limpio_estructural.xlsx` | Excel estructural saneado del que se derivó la taxonomía (categorías, grupos, metadatos de procedencia). Es solo referencia de auditoría: ningún script lo lee en tiempo de ejecución, los datos ya están embebidos en `launcher-*`. |

`command-builder-*` y `launcher-*` tienen propósitos distintos y no dependen uno del
otro: el primero es una herramienta práctica para obtener un comando concreto; el
segundo es un explorador de la taxonomía completa y un verificador de que el entorno
local puede ejecutar de forma segura cada tipo de comportamiento (útil para validar el
laboratorio antes de usarlo, o simplemente para consultar los metadatos del dataset).

## Linux

Requiere `bash` 4.0 o superior (los scripts usan arrays y `read`; el explorador además
usa expansión a minúsculas). La mayoría de distribuciones Linux ya traen bash 4/5 por
defecto. **En macOS, el bash del sistema (`/bin/bash`) es la versión 3.2** — el
generador de comandos (`command-builder-linux.sh`) funciona igualmente con esa versión,
pero si vas a usar `launcher-linux.sh` instala uno más reciente (`brew install bash`) o
ejecútalo explícitamente con ese binario, p. ej. `/opt/homebrew/bin/bash launcher-linux.sh`.

```sh
chmod +x command-builder-linux.sh launcher-linux.sh   # una sola vez
./command-builder-linux.sh
./launcher-linux.sh
```

Para copiar al portapapeles automáticamente hace falta `wl-copy` (Wayland), `xclip` o
`xsel` (X11) o `pbcopy` (macOS); si no hay ninguno disponible, el comando se muestra
igualmente y solo hay que copiarlo a mano.

## Windows

Requiere PowerShell 5.1 o superior (incluido de serie en Windows 10/11) o PowerShell 7+
(`pwsh`). Si la política de ejecución bloquea scripts locales, ejecuta desde PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\command-builder-windows.ps1
.\launcher-windows.ps1
```

El copiado al portapapeles usa `Set-Clipboard`; si no hay portapapeles disponible (por
ejemplo, una sesión sin entorno gráfico), el script lo indica y muestra el comando igual.

## Categorías

Categorías representadas en `command-builder-*` (14 de las 16 de la taxonomía; las dos
restantes, `00 - Sin contenido` y `ZZ - fila vacía`, no son técnicas y no generan
comando):

| ID | Categoría | Propósito |
|---|---|---|
| 01 | msiexec | Instalar/desinstalar un paquete MSI local de forma silenciosa. |
| 02 | cmd + net use (SMB/WebDAV) | Mapear o desconectar un recurso de red compartido. |
| 03 | conhost --headless | Diagnóstico básico del equipo (usuario, hostname, IP) sin ventana de consola. |
| 04 | mshta | Abrir una aplicación HTML (.hta) local y de confianza. |
| 05 | rundll32 | Accesos directos a paneles de control clásicos de Windows. |
| 06 | wscript / cscript | Ejecutar un script local (VBS/JS) que tú controlas. |
| 07 | pcalua | Abrir aplicaciones integradas de Windows vía el asistente de compatibilidad. |
| 08 | PowerShell (línea de comandos) | Resolver DNS, probar un puerto, calcular un hash, consultar cabeceras HTTP. |
| 09 | cmd.exe | ping, nslookup, whoami, listar una carpeta. |
| 10 | PowerShell embebido | Consultas de una línea: resumen del sistema, procesos por CPU, configuración de red. |
| 11 | shell Unix (bash/ssh/curl) | Comprobaciones equivalentes para un entorno Linux/macOS/WSL (no cmd.exe/PowerShell nativos). |
| 12 | URL / dominio | Cabeceras HTTP, resolución DNS o prueba de puerto sobre un recurso de red. |
| 13 | echo / texto | Imprimir texto o codificarlo en Base64, sin tocar sistemas externos. |
| 99 | otros / diagnóstico | Inventario básico del sistema local (Linux o Windows). |

`launcher-*` añade `00 - Sin contenido` y `ZZ - fila vacía` como entradas informativas
del dataset (sin ejecutor propio) para poder explorar la taxonomía completa.

## Parámetros

Tipos de dato que se piden con más frecuencia:

| Parámetro | Qué es | De dónde suele salir |
|---|---|---|
| Host / dominio | Nombre o dirección del equipo o servicio a comprobar. | El inventario del entorno, DNS interno o la documentación del recurso. |
| IP | Dirección numérica de un equipo. | Igual que el host, cuando no hay nombre DNS. |
| URL | Dirección completa de un recurso web (`https://...`). | El enlace del servicio que quieres comprobar. |
| Puerto | Número de servicio TCP (1-65535). | La documentación del servicio (80/443 web, 22 SSH, 445 SMB...). |
| Ruta | Ubicación de un fichero o carpeta local. | El Explorador de archivos / Finder, o las propiedades del fichero. |
| Recurso compartido | Nombre del share SMB/WebDAV al que te conectas. | La configuración del recurso o quien lo administra. |

Los valores se validan de forma básica (formato de puerto, formato de letra de unidad,
que no contengan una comilla que rompería el comando) antes de generarse el comando.

## Ejemplos de uso

```
$ ./command-builder-linux.sh
...
 9) 09 - cmd.exe
> 9
1) Ping
2) nslookup
...
> 1
Host/IP: 192.0.2.10
============================================================
Comando generado
============================================================

cmd.exe /d /c ping "192.0.2.10"

✓ Copiado automáticamente al portapapeles.
```

```
> 8                     # 08 - PowerShell (línea de comandos)
2) Probar puerto TCP
Host/IP autorizado: app.ejemplo.local
Puerto: 8443
============================================================
Comando generado
============================================================

powershell.exe -NoProfile -Command "Test-NetConnection -ComputerName 'app.ejemplo.local' -Port 8443"
```

## Alcance

Estos scripts son herramientas de **generación y apoyo** para administración,
diagnóstico y laboratorio autorizado. No ejecutan nada por sí mismos contra sistemas de
terceros: construyen el texto del comando y lo dejan listo para que tú decidas cuándo y
dónde ejecutarlo. Úsalos únicamente sobre sistemas y recursos para los que tengas
autorización explícita. `launcher-*` solo ejecuta validaciones locales y benignas
(operaciones en `127.0.0.1` o en ficheros temporales propios); nunca contacta
infraestructura de terceros ni reproduce muestras o cargas reales.

## Procedencia de la taxonomía

Las categorías, grupos y metadatos que organizan estos scripts se derivaron de un
dataset de seguridad clasificado previamente (1.483 grupos / 4.339 registros
clasificados, sobre un total de 4.411 filas de origen), reducido y saneado hasta quedar
en un Excel puramente estructural: solo se conservan categorías, recuentos, fechas
agregadas y presencia/ausencia de características (descarga remota, codificación,
ofuscación, ejecución oculta, persistencia), sin dominios, IPs, URLs ni ningún contenido
ejecutable u orden de consola real. No hace falta conocer el dataset original para usar
ni para entender estos scripts: cada categoría representa un tipo de comportamiento, no
una muestra concreta.
