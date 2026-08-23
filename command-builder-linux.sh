#!/usr/bin/env bash
# Command Builder basado en la taxonomía del dataset.
# Genera plantillas benignas/administrativas y de laboratorio; no reproduce payloads crudos.
set -u

copy_to_clipboard() {
  local text="$1"
  if command -v wl-copy >/dev/null 2>&1; then
    printf '%s' "$text" | wl-copy && return 0
  elif command -v xclip >/dev/null 2>&1; then
    printf '%s' "$text" | xclip -selection clipboard && return 0
  elif command -v xsel >/dev/null 2>&1; then
    printf '%s' "$text" | xsel --clipboard --input && return 0
  elif command -v pbcopy >/dev/null 2>&1; then
    printf '%s' "$text" | pbcopy && return 0
  fi
  return 1
}

pause() {
  printf '\nPulsa Enter para continuar...'
  IFS= read -r _
}

ask_nonempty() {
  local prompt="$1" value
  while true; do
    printf '%s' "$prompt" >&2
    # Sin `|| exit`, un EOF de stdin (pipe cerrado, Ctrl-D) dejaría este
    # bucle girando sin fin y sin pausa, a CPU máxima.
    IFS= read -r value || { printf '\nEntrada finalizada (EOF); saliendo.\n' >&2; exit 1; }
    [[ -n "$value" ]] && { printf '%s' "$value"; return; }
    printf 'El valor no puede estar vacío.\n' >&2
  done
}

ask_port() {
  local p
  while true; do
    p="$(ask_nonempty 'Puerto: ')"
    if [[ "$p" =~ ^[0-9]+$ ]] && (( p >= 1 && p <= 65535 )); then
      printf '%s' "$p"; return
    fi
    printf 'Puerto inválido (1-65535).\n' >&2
  done
}

# Para valores que se insertan dentro de -Command "..." de PowerShell: la
# comilla doble rompería el wrapper, y $ o ` disparan interpolación/subexpresión
# de la comilla doble EXTERNA (p. ej. $(calc)) aunque vayan dentro de comillas
# simples internas — esas no protegen frente al parseo de la comilla externa.
ask_ps_value() {
  local prompt="$1" value
  while true; do
    value="$(ask_nonempty "$prompt")"
    if [[ "$value" != *'"'* && "$value" != *'$'* && "$value" != *'`'* ]]; then
      printf '%s' "$value"; return
    fi
    printf 'El valor no puede contener comillas dobles, $ ni ` (romperían o inyectarían código en el comando PowerShell).\n' >&2
  done
}

ask_drive_letter() {
  local d
  while true; do
    d="$(ask_nonempty 'Letra de unidad (ej. Z:): ')"
    if [[ "$d" =~ ^[A-Za-z]:$ ]]; then
      printf '%s' "$d"; return
    fi
    printf 'Formato inválido; usa una letra seguida de dos puntos, ej. Z:\n' >&2
  done
}

psq() { printf "'%s'" "${1//\'/\'\'}"; }
shq() { printf "'%s'" "${1//\'/\'\\\'\'}"; }
cmdq() {
  local s="$1"
  s="${s//\"/\"\"}"
  printf '"%s"' "$s"
}

show_result() {
  local title="$1" command="$2"
  printf '\n============================================================\n'
  printf '%s\n' "$title"
  printf '============================================================\n\n'
  printf '%s\n' "$command"
  printf '\n'
  if copy_to_clipboard "$command"; then
    printf '✓ Copiado automáticamente al portapapeles.\n'
  else
    printf 'No hay utilidad de portapapeles disponible; copia el comando mostrado.\n'
  fi
  pause
}

category_info() {
  case "$1" in
    01) cat <<'EOF'
Para qué sirve: msiexec es el instalador nativo de Windows para paquetes MSI y se usa para instalar o retirar software empaquetado de forma estándar.
Qué hace aquí: te prepara una instalación o desinstalación silenciosa de un MSI local y de confianza.
Qué te pediré: la ruta completa del archivo .msi; puedes copiarla desde el Explorador de archivos o desde las propiedades del fichero.
EOF
      ;;
    02) cat <<'EOF'
Para qué sirve: net use conecta o desconecta recursos compartidos de red, normalmente SMB o WebDAV, desde Windows.
Qué hace aquí: te permite preparar el mapeo de un recurso autorizado a una letra de unidad o desconectarlo después.
Qué te pediré: servidor/host, nombre o ruta del recurso y letra de unidad; esos datos suelen venir de la configuración del recurso o del administrador del entorno.
EOF
      ;;
    03) cat <<'EOF'
Para qué sirve: conhost --headless aloja procesos de consola de Windows sin mostrar la ventana clásica.
Qué hace aquí: lo usamos para lanzar pequeñas comprobaciones locales de diagnóstico como usuario, hostname o configuración IP.
Qué te pediré: nada adicional en las opciones actuales, porque todas consultan información del propio equipo.
EOF
      ;;
    04) cat <<'EOF'
Para qué sirve: mshta abre aplicaciones HTML de Windows (HTA), normalmente archivos .hta.
Qué hace aquí: genera la apertura de un HTA local y controlado para pruebas de laboratorio o compatibilidad.
Qué te pediré: la ruta completa del archivo .hta; puedes copiarla desde el Explorador de archivos o desde las propiedades del fichero.
EOF
      ;;
    05) cat <<'EOF'
Para qué sirve: rundll32 permite invocar funciones exportadas por DLL de Windows y es una utilidad integrada del sistema.
Qué hace aquí: genera accesos conocidos a paneles y funciones administrativas de Windows, como Opciones de Internet o impresoras.
Qué te pediré: nada en las opciones actuales, porque la DLL y la función concreta ya vienen definidas por la opción que elijas.
EOF
      ;;
    06) cat <<'EOF'
Para qué sirve: wscript y cscript ejecutan scripts de Windows Script Host, como VBS o JavaScript.
Qué hace aquí: prepara la ejecución de un script local que tú controles mediante cscript.
Qué te pediré: la ruta completa del archivo .vbs o .js; normalmente la obtienes copiando la ruta desde el Explorador de archivos.
EOF
      ;;
    07) cat <<'EOF'
Para qué sirve: pcalua forma parte del Program Compatibility Assistant de Windows y puede iniciar aplicaciones locales.
Qué hace aquí: muestra ejemplos sencillos para abrir aplicaciones integradas del sistema a través de ese componente.
Qué te pediré: nada en las opciones actuales, porque las aplicaciones ya están fijadas en cada plantilla.
EOF
      ;;
    08) cat <<'EOF'
Para qué sirve: PowerShell permite administrar Windows y realizar comprobaciones de red, DNS, archivos y HTTP.
Qué hace aquí: genera comandos para resolver nombres, comprobar un puerto TCP, calcular hashes o consultar cabeceras HTTP.
Qué te pediré: según la opción, un dominio/host, una IP, un puerto, una ruta de fichero o una URL; esos datos salen del recurso autorizado que quieras comprobar.
EOF
      ;;
    09) cat <<'EOF'
Para qué sirve: cmd.exe es la consola clásica de Windows y permite ejecutar utilidades básicas del sistema.
Qué hace aquí: genera comprobaciones sencillas como ping, nslookup, whoami o listado de carpetas.
Qué te pediré: según la opción, un host/IP, un dominio o una ruta local; normalmente los obtienes del equipo, servicio o carpeta que estés revisando.
EOF
      ;;
    10) cat <<'EOF'
Para qué sirve: esta sección agrupa pequeñas consultas PowerShell que caben en una sola línea.
Qué hace aquí: genera comandos para ver información del sistema, procesos con más CPU o configuración de red del equipo local.
Qué te pediré: nada en las opciones actuales, porque trabajan directamente sobre el equipo donde ejecutes el comando.
EOF
      ;;
    11) cat <<'EOF'
Para qué sirve: la shell Unix reúne utilidades habituales como curl, dig, ping, sha256sum y ssh para diagnóstico desde Linux o macOS.
Qué hace aquí: genera comprobaciones HTTP, DNS, conectividad, hashes y una prueba de autenticación SSH sin ejecutar órdenes remotas.
Qué te pediré: según la opción, URL, dominio/host, ruta de fichero o host y puerto SSH; esos datos corresponden al recurso autorizado que quieras revisar.
EOF
      ;;
    12) cat <<'EOF'
Para qué sirve: esta categoría agrupa comprobaciones sencillas sobre URLs, dominios y servicios de red.
Qué hace aquí: genera consultas de cabeceras HTTP, resolución DNS o comprobación de disponibilidad de un puerto.
Qué te pediré: una URL, un dominio/host o un host/IP y puerto; la URL es la dirección completa y el puerto identifica el servicio concreto que quieres comprobar.
EOF
      ;;
    13) cat <<'EOF'
Para qué sirve: esta sección trabaja con texto local y transformaciones simples sin acceder a sistemas externos.
Qué hace aquí: genera comandos para imprimir texto o codificarlo en Base64.
Qué te pediré: únicamente el texto que quieras utilizar, que puede ser cualquier cadena de prueba no sensible.
EOF
      ;;
    99) cat <<'EOF'
Para qué sirve: agrupa utilidades básicas de inventario y diagnóstico que no encajan en las categorías anteriores.
Qué hace aquí: genera comandos para consultar información general del propio equipo Linux o Windows.
Qué te pediré: nada en las opciones actuales, porque toda la información se obtiene del sistema local.
EOF
      ;;
  esac
}

category_name() {
  case "$1" in
    01) printf '01 - msiexec (MSI)' ;;
    02) printf '02 - cmd + net use (SMB/WebDAV)' ;;
    03) printf '03 - conhost --headless' ;;
    04) printf '04 - mshta' ;;
    05) printf '05 - rundll32' ;;
    06) printf '06 - wscript / cscript' ;;
    07) printf '07 - pcalua' ;;
    08) printf '08 - PowerShell (línea de comandos)' ;;
    09) printf '09 - cmd.exe' ;;
    10) printf '10 - PowerShell embebido' ;;
    11) printf '11 - shell Unix (bash/ssh/curl)' ;;
    12) printf '12 - URL / dominio' ;;
    13) printf '13 - echo / texto' ;;
    99) printf '99 - otros / diagnóstico' ;;
  esac
}

show_category_info() {
  printf '
============================================================
'
  printf ' INFORMACIÓN DE CATEGORÍAS
'
  printf '============================================================
'
  printf 'Resumen rápido de para qué sirve cada categoría y qué datos puede pedirte.

'
  local id
  for id in 01 02 03 04 05 06 07 08 09 10 11 12 13 99; do
    printf '%s
' "$(category_name "$id")"
    category_info "$id"
    printf '
'
  done
  pause
}

menu() {
  clear 2>/dev/null || true
  printf '============================================================\n'
  printf ' COMMAND BUILDER — laboratorio / administración autorizada\n'
  printf '============================================================\n'
  printf 'Elige por categoría. Después te preguntaré solo los datos necesarios.\n\n'
  printf ' 1) 01 - msiexec (MSI)\n'
  printf ' 2) 02 - cmd + net use (SMB/WebDAV)\n'
  printf ' 3) 03 - conhost --headless\n'
  printf ' 4) 04 - mshta\n'
  printf ' 5) 05 - rundll32\n'
  printf ' 6) 06 - wscript / cscript\n'
  printf ' 7) 07 - pcalua\n'
  printf ' 8) 08 - PowerShell (línea de comandos)\n'
  printf ' 9) 09 - cmd.exe\n'
  printf '10) 10 - PowerShell embebido\n'
  printf '11) 11 - shell Unix (bash/ssh/curl)\n'
  printf '12) 12 - URL / dominio\n'
  printf '13) 13 - echo\n'
  printf '14) 99 - otros / diagnóstico\n'
  printf '15) Información de categorías\n'
  printf ' 0) Salir\n\n'
}

choose_action() {
  local cid="$1" choice target path host port text drive share url
  printf '\n%s\n' "$(category_name "$cid")"
  printf '%s\n' '------------------------------------------------------------'
  category_info "$cid"
  printf '\n'

  case "$cid" in
    01)
      printf '1) Instalar un MSI local\n2) Desinstalar un MSI local\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) path="$(ask_nonempty 'Ruta local del .msi: ')"; show_result "Comando generado" "msiexec.exe /i $(cmdq "$path") /qn /norestart" ;;
        2) path="$(ask_nonempty 'Ruta local del .msi: ')"; show_result "Comando generado" "msiexec.exe /x $(cmdq "$path") /qn /norestart" ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    02)
      printf '1) Mapear recurso SMB\n2) Mapear WebDAV sobre TLS\n3) Desconectar unidad\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1)
          host="$(ask_nonempty 'Servidor/host autorizado: ')"
          share="$(ask_nonempty 'Nombre del recurso compartido: ')"
          drive="$(ask_drive_letter)"
          show_result "Comando generado" "net use $drive $(cmdq "\\\\$host\\$share") /persistent:no"
          ;;
        2)
          host="$(ask_nonempty 'Servidor WebDAV autorizado: ')"
          share="$(ask_nonempty 'Ruta dentro de DavWWWRoot: ')"
          drive="$(ask_drive_letter)"
          show_result "Comando generado" "net use $drive $(cmdq "\\\\${host}@SSL\\DavWWWRoot\\$share") /persistent:no"
          ;;
        3)
          drive="$(ask_drive_letter)"
          show_result "Comando generado" "net use $drive /delete /y"
          ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    03)
      printf '1) Consultar usuario actual\n2) Mostrar hostname\n3) Mostrar configuración IP\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) show_result "Comando generado" 'conhost.exe --headless cmd.exe /d /c whoami' ;;
        2) show_result "Comando generado" 'conhost.exe --headless cmd.exe /d /c hostname' ;;
        3) show_result "Comando generado" 'conhost.exe --headless cmd.exe /d /c ipconfig /all' ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    04)
      printf '1) Abrir un HTA local de confianza\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      if [[ "$choice" == "1" ]]; then
        path="$(ask_nonempty 'Ruta local del .hta: ')"
        show_result "Comando generado" "mshta.exe $(cmdq "$path")"
      else
        printf 'Opción inválida.\n'; pause
      fi
      ;;
    05)
      printf '1) Abrir Opciones de Internet\n2) Abrir impresoras clásicas\n3) Abrir fecha/hora clásica\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) show_result "Comando generado" 'rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl' ;;
        2) show_result "Comando generado" 'rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL PrintersFolder' ;;
        3) show_result "Comando generado" 'rundll32.exe shell32.dll,Control_RunDLL timedate.cpl' ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    06)
      printf '1) Ejecutar VBS local con cscript\n2) Ejecutar JS local con cscript\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1|2) path="$(ask_nonempty 'Ruta local del script que controlas: ')"; show_result "Comando generado" "cscript.exe //nologo $(cmdq "$path")" ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    07)
      printf '1) Abrir Bloc de notas\n2) Abrir Explorador de archivos\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) show_result "Comando generado" 'pcalua.exe -a notepad.exe' ;;
        2) show_result "Comando generado" 'pcalua.exe -a explorer.exe' ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    08)
      printf '1) Resolver DNS\n2) Probar puerto TCP\n3) Calcular SHA256 de fichero\n4) Consultar cabeceras HTTP\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1)
          host="$(ask_ps_value 'Dominio/host: ')"
          show_result "Comando generado" "powershell.exe -NoProfile -Command \"Resolve-DnsName -Name $(psq "$host")\""
          ;;
        2)
          host="$(ask_ps_value 'Host/IP autorizado: ')"; port="$(ask_port)"
          show_result "Comando generado" "powershell.exe -NoProfile -Command \"Test-NetConnection -ComputerName $(psq "$host") -Port $port\""
          ;;
        3)
          path="$(ask_ps_value 'Ruta del fichero: ')"
          show_result "Comando generado" "powershell.exe -NoProfile -Command \"Get-FileHash -Algorithm SHA256 -LiteralPath $(psq "$path")\""
          ;;
        4)
          url="$(ask_ps_value 'URL: ')"
          show_result "Comando generado" "powershell.exe -NoProfile -Command \"Invoke-WebRequest -Method Head -UseBasicParsing -Uri $(psq "$url") | Select-Object StatusCode,Headers\""
          ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    09)
      printf '1) Ping\n2) nslookup\n3) whoami\n4) Listar carpeta\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) host="$(ask_nonempty 'Host/IP: ')"; show_result "Comando generado" "cmd.exe /d /c ping $(cmdq "$host")" ;;
        2) host="$(ask_nonempty 'Dominio/host: ')"; show_result "Comando generado" "cmd.exe /d /c nslookup $(cmdq "$host")" ;;
        3) show_result "Comando generado" 'cmd.exe /d /c whoami /all' ;;
        4) path="$(ask_nonempty 'Ruta de carpeta: ')"; show_result "Comando generado" "cmd.exe /d /c dir $(cmdq "$path")" ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    10)
      printf '1) Resumen de sistema\n2) Top de procesos por CPU\n3) Información de red\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) show_result "Comando generado" 'powershell.exe -NoProfile -Command "$os=Get-CimInstance Win32_OperatingSystem; [pscustomobject]@{Equipo=$env:COMPUTERNAME;Usuario=$env:USERNAME;OS=$os.Caption;Version=$os.Version}"' ;;
        2) show_result "Comando generado" 'powershell.exe -NoProfile -Command "Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 Name,Id,CPU"' ;;
        3) show_result "Comando generado" 'powershell.exe -NoProfile -Command "Get-NetIPConfiguration | Format-List InterfaceAlias,IPv4Address,IPv4DefaultGateway,DNSServer"' ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    11)
      printf '1) Cabeceras HTTP con curl\n2) Resolver DNS\n3) Ping\n4) SHA256 de fichero\n5) Probar autenticación SSH sin ejecutar comandos remotos\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) url="$(ask_nonempty 'URL: ')"; show_result "Comando generado" "curl -fsSIL --max-time 10 $(shq "$url")" ;;
        2) host="$(ask_nonempty 'Dominio/host: ')"; show_result "Comando generado" "dig +short $(shq "$host")" ;;
        3) host="$(ask_nonempty 'Host/IP: ')"; show_result "Comando generado" "ping -c 4 -- $(shq "$host")" ;;
        4) path="$(ask_nonempty 'Ruta del fichero: ')"; show_result "Comando generado" "sha256sum -- $(shq "$path")" ;;
        5)
          host="$(ask_nonempty 'Host SSH autorizado: ')"; port="$(ask_port)"
          show_result "Comando generado" "ssh -v -o BatchMode=yes -o ConnectTimeout=5 -p $port -- $(shq "$host") exit"
          ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    12)
      printf '1) Ver cabeceras HTTP\n2) Resolver DNS\n3) Probar un puerto\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) url="$(ask_nonempty 'URL: ')"; show_result "Comando generado" "curl -fsSIL --max-time 10 $(shq "$url")" ;;
        2) host="$(ask_nonempty 'Dominio/host: ')"; show_result "Comando generado" "nslookup $(shq "$host")" ;;
        3) host="$(ask_nonempty 'Host/IP autorizado: ')"; port="$(ask_port)"; show_result "Comando generado" "nc -vz -w 5 $(shq "$host") $port" ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    13)
      printf '1) Imprimir texto\n2) Codificar texto en Base64\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) text="$(ask_nonempty 'Texto: ')"; show_result "Comando generado" "printf '%s\\n' $(shq "$text")" ;;
        2) text="$(ask_nonempty 'Texto: ')"; show_result "Comando generado" "printf '%s' $(shq "$text") | base64" ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
    99)
      printf '1) Información del sistema Linux\n2) Información del sistema Windows\n'
      printf 'B) Volver al menú de categorías\n'
      printf '> '; read -r choice
      [[ "$choice" =~ ^[bB]$ ]] && return
      case "$choice" in
        1) show_result "Comando generado" 'uname -a; id; ip addr' ;;
        2) show_result "Comando generado" 'systeminfo & whoami /all & ipconfig /all' ;;
        *) printf 'Opción inválida.\n'; pause ;;
      esac
      ;;
  esac
}

while true; do
  menu
  printf '> '
  IFS= read -r sel || { printf '\n'; exit 0; }
  case "$sel" in
    1) choose_action 01 ;;
    2) choose_action 02 ;;
    3) choose_action 03 ;;
    4) choose_action 04 ;;
    5) choose_action 05 ;;
    6) choose_action 06 ;;
    7) choose_action 07 ;;
    8) choose_action 08 ;;
    9) choose_action 09 ;;
    10) choose_action 10 ;;
    11) choose_action 11 ;;
    12) choose_action 12 ;;
    13) choose_action 13 ;;
    14) choose_action 99 ;;
    15) show_category_info ;;
    0|q|Q) exit 0 ;;
    *) printf 'Opción inválida.\n'; pause ;;
  esac
done
