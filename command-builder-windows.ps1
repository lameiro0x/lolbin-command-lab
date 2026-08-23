param()

function Pause-Builder {
    [void](Read-Host "`nPulsa Enter para continuar")
}

function Ask-NonEmpty([string]$Prompt) {
    while ($true) {
        $v = Read-Host $Prompt
        # $null significa stdin cerrado (EOF); sin esta comprobación el bucle
        # gira sin fin. IsNullOrWhiteSpace($null) es true, así que hay que
        # distinguir explícitamente el caso EOF del de "línea vacía".
        if ($null -eq $v) { Write-Host "`nEntrada finalizada (EOF); saliendo."; exit 1 }
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
        Write-Host 'El valor no puede estar vacío.'
    }
}

function Ask-Port {
    while ($true) {
        $p = Ask-NonEmpty 'Puerto'
        $n = 0
        if ([int]::TryParse($p, [ref]$n) -and $n -ge 1 -and $n -le 65535) { return $n }
        Write-Host 'Puerto inválido (1-65535).'
    }
}

function Ask-DriveLetter {
    while ($true) {
        $d = Ask-NonEmpty 'Letra de unidad (ej. Z:)'
        if ($d -match '^[A-Za-z]:$') { return $d }
        Write-Host 'Formato inválido; usa una letra seguida de dos puntos, ej. Z:'
    }
}

# Para valores que se insertan dentro de -Command "..." de PowerShell: la
# comilla doble rompería el wrapper, y $ o ` disparan interpolación/subexpresión
# de la comilla doble EXTERNA (p. ej. $(calc)) aunque vayan dentro de comillas
# simples internas — esas no protegen frente al parseo de la comilla externa.
function Ask-PSValue([string]$Prompt) {
    while ($true) {
        $v = Ask-NonEmpty $Prompt
        if ($v -notmatch '["$`]') { return $v }
        Write-Host 'El valor no puede contener comillas dobles, $ ni ` (romperían o inyectarían código en el comando PowerShell).'
    }
}

function Quote-PS([string]$s) {
    return "'" + $s.Replace("'", "''") + "'"
}

function Quote-Cmd([string]$s) {
    return '"' + $s.Replace('"','""') + '"'
}

function Quote-Sh([string]$s) {
    return "'" + $s.Replace("'", "'\''") + "'"
}

function Show-Result([string]$Command) {
    Write-Host "`n============================================================"
    Write-Host 'COMANDO GENERADO'
    Write-Host "============================================================`n"
    Write-Host $Command
    try {
        Set-Clipboard -Value $Command
        Write-Host "`n✓ Copiado automáticamente al portapapeles."
    } catch {
        Write-Host "`nNo se pudo copiar automáticamente; copia el comando mostrado."
    }
    Pause-Builder
}

function Category-Info([string]$Id) {
    switch ($Id) {
        '01' { @'
Para qué sirve: msiexec es el instalador nativo de Windows para paquetes MSI y se usa para instalar o retirar software empaquetado de forma estándar.
Qué hace aquí: te prepara una instalación o desinstalación silenciosa de un MSI local y de confianza.
Qué te pediré: la ruta completa del archivo .msi; puedes copiarla desde el Explorador de archivos o desde las propiedades del fichero.
'@ }
        '02' { @'
Para qué sirve: net use conecta o desconecta recursos compartidos de red, normalmente SMB o WebDAV, desde Windows.
Qué hace aquí: te permite preparar el mapeo de un recurso autorizado a una letra de unidad o desconectarlo después.
Qué te pediré: servidor/host, nombre o ruta del recurso y letra de unidad; esos datos suelen venir de la configuración del recurso o del administrador del entorno.
'@ }
        '03' { @'
Para qué sirve: conhost --headless aloja procesos de consola de Windows sin mostrar la ventana clásica.
Qué hace aquí: lo usamos para lanzar pequeñas comprobaciones locales de diagnóstico como usuario, hostname o configuración IP.
Qué te pediré: nada adicional en las opciones actuales, porque todas consultan información del propio equipo.
'@ }
        '04' { @'
Para qué sirve: mshta abre aplicaciones HTML de Windows (HTA), normalmente archivos .hta.
Qué hace aquí: genera la apertura de un HTA local y controlado para pruebas de laboratorio o compatibilidad.
Qué te pediré: la ruta completa del archivo .hta; puedes copiarla desde el Explorador de archivos o desde las propiedades del fichero.
'@ }
        '05' { @'
Para qué sirve: rundll32 permite invocar funciones exportadas por DLL de Windows y es una utilidad integrada del sistema.
Qué hace aquí: genera accesos conocidos a paneles y funciones administrativas de Windows, como Opciones de Internet o impresoras.
Qué te pediré: nada en las opciones actuales, porque la DLL y la función concreta ya vienen definidas por la opción que elijas.
'@ }
        '06' { @'
Para qué sirve: wscript y cscript ejecutan scripts de Windows Script Host, como VBS o JavaScript.
Qué hace aquí: prepara la ejecución de un script local que tú controles mediante cscript.
Qué te pediré: la ruta completa del archivo .vbs o .js; normalmente la obtienes copiando la ruta desde el Explorador de archivos.
'@ }
        '07' { @'
Para qué sirve: pcalua forma parte del Program Compatibility Assistant de Windows y puede iniciar aplicaciones locales.
Qué hace aquí: muestra ejemplos sencillos para abrir aplicaciones integradas del sistema a través de ese componente.
Qué te pediré: nada en las opciones actuales, porque las aplicaciones ya están fijadas en cada plantilla.
'@ }
        '08' { @'
Para qué sirve: PowerShell permite administrar Windows y realizar comprobaciones de red, DNS, archivos y HTTP.
Qué hace aquí: genera comandos para resolver nombres, comprobar un puerto TCP, calcular hashes o consultar cabeceras HTTP.
Qué te pediré: según la opción, un dominio/host, una IP, un puerto, una ruta de fichero o una URL; esos datos salen del recurso autorizado que quieras comprobar.
'@ }
        '09' { @'
Para qué sirve: cmd.exe es la consola clásica de Windows y permite ejecutar utilidades básicas del sistema.
Qué hace aquí: genera comprobaciones sencillas como ping, nslookup, whoami o listado de carpetas.
Qué te pediré: según la opción, un host/IP, un dominio o una ruta local; normalmente los obtienes del equipo, servicio o carpeta que estés revisando.
'@ }
        '10' { @'
Para qué sirve: esta sección agrupa pequeñas consultas PowerShell que caben en una sola línea.
Qué hace aquí: genera comandos para ver información del sistema, procesos con más CPU o configuración de red del equipo local.
Qué te pediré: nada en las opciones actuales, porque trabajan directamente sobre el equipo donde ejecutes el comando.
'@ }
        '11' { @'
Para qué sirve: la shell Unix reúne utilidades habituales como curl, dig, ping, sha256sum y ssh para diagnóstico desde Linux o macOS.
Qué hace aquí: genera comprobaciones HTTP, DNS, conectividad, hashes y una prueba de autenticación SSH sin ejecutar órdenes remotas.
Qué te pediré: según la opción, URL, dominio/host, ruta de fichero o host y puerto SSH; esos datos corresponden al recurso autorizado que quieras revisar.
Nota: estos comandos usan sintaxis Unix; ejecútalos en un sistema Linux/macOS, WSL o Git Bash, no directamente en cmd.exe ni en esta consola de PowerShell.
'@ }
        '12' { @'
Para qué sirve: esta categoría agrupa comprobaciones sencillas sobre URLs, dominios y servicios de red.
Qué hace aquí: genera consultas de cabeceras HTTP, resolución DNS o comprobación de disponibilidad de un puerto.
Qué te pediré: una URL, un dominio/host o un host/IP y puerto; la URL es la dirección completa y el puerto identifica el servicio concreto que quieres comprobar.
'@ }
        '13' { @'
Para qué sirve: esta sección trabaja con texto local y transformaciones simples sin acceder a sistemas externos.
Qué hace aquí: genera comandos para imprimir texto o codificarlo en Base64.
Qué te pediré: únicamente el texto que quieras utilizar, que puede ser cualquier cadena de prueba no sensible.
'@ }
        '99' { @'
Para qué sirve: agrupa utilidades básicas de inventario y diagnóstico que no encajan en las categorías anteriores.
Qué hace aquí: genera comandos para consultar información general del propio equipo Linux o Windows.
Qué te pediré: nada en las opciones actuales, porque toda la información se obtiene del sistema local.
'@ }
    }
}

function Category-Name([string]$Id) {
    switch ($Id) {
        '01' { '01 - msiexec (MSI)' }
        '02' { '02 - cmd + net use (SMB/WebDAV)' }
        '03' { '03 - conhost --headless' }
        '04' { '04 - mshta' }
        '05' { '05 - rundll32' }
        '06' { '06 - wscript / cscript' }
        '07' { '07 - pcalua' }
        '08' { '08 - PowerShell (línea de comandos)' }
        '09' { '09 - cmd.exe' }
        '10' { '10 - PowerShell embebido' }
        '11' { '11 - shell Unix (bash/ssh/curl)' }
        '12' { '12 - URL / dominio' }
        '13' { '13 - echo / texto' }
        '99' { '99 - otros / diagnóstico' }
    }
}

function Show-CategoryInfo {
    Write-Host "`n============================================================"
    Write-Host ' INFORMACIÓN DE CATEGORÍAS'
    Write-Host '============================================================'
    Write-Host 'Resumen rápido de para qué sirve cada categoría y qué datos puede pedirte.'
    Write-Host
    foreach ($id in @('01','02','03','04','05','06','07','08','09','10','11','12','13','99')) {
        Write-Host (Category-Name $id)
        Write-Host (Category-Info $id)
        Write-Host
    }
    Pause-Builder
}

function Show-Menu {
    Clear-Host
    Write-Host '============================================================'
    Write-Host ' COMMAND BUILDER — laboratorio / administración autorizada'
    Write-Host '============================================================'
    Write-Host 'Elige por categoría. Después te preguntaré solo los datos necesarios.'
    Write-Host
    Write-Host ' 1) 01 - msiexec (MSI)'
    Write-Host ' 2) 02 - cmd + net use (SMB/WebDAV)'
    Write-Host ' 3) 03 - conhost --headless'
    Write-Host ' 4) 04 - mshta'
    Write-Host ' 5) 05 - rundll32'
    Write-Host ' 6) 06 - wscript / cscript'
    Write-Host ' 7) 07 - pcalua'
    Write-Host ' 8) 08 - PowerShell (línea de comandos)'
    Write-Host ' 9) 09 - cmd.exe'
    Write-Host '10) 10 - PowerShell embebido'
    Write-Host '11) 11 - shell Unix (bash/ssh/curl)'
    Write-Host '12) 12 - URL / dominio'
    Write-Host '13) 13 - echo'
    Write-Host '14) 99 - otros / diagnóstico'
    Write-Host '15) Información de categorías'
    Write-Host ' 0) Salir'
    Write-Host
}

function Build-Category([string]$Id) {
    Write-Host
    Write-Host (Category-Name $Id)
    Write-Host '------------------------------------------------------------'
    Write-Host (Category-Info $Id)
    Write-Host

    switch ($Id) {
        '01' {
            Write-Host '1) Instalar un MSI local'
            Write-Host '2) Desinstalar un MSI local'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            if ($c -eq '1') { $path = Ask-NonEmpty 'Ruta local del .msi'; Show-Result ("msiexec.exe /i {0} /qn /norestart" -f (Quote-Cmd $path)) }
            elseif ($c -eq '2') { $path = Ask-NonEmpty 'Ruta local del .msi'; Show-Result ("msiexec.exe /x {0} /qn /norestart" -f (Quote-Cmd $path)) }
            else { Write-Host 'Opción inválida.'; Pause-Builder }
        }
        '02' {
            Write-Host '1) Mapear recurso SMB'
            Write-Host '2) Mapear WebDAV sobre TLS'
            Write-Host '3) Desconectar unidad'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                '1' {
                    $hostn = Ask-NonEmpty 'Servidor/host autorizado'
                    $share = Ask-NonEmpty 'Nombre del recurso compartido'
                    $drive = Ask-DriveLetter
                    Show-Result ("net use {0} {1} /persistent:no" -f $drive,(Quote-Cmd ("\\{0}\{1}" -f $hostn,$share)))
                }
                '2' {
                    $hostn = Ask-NonEmpty 'Servidor WebDAV autorizado'
                    $share = Ask-NonEmpty 'Ruta dentro de DavWWWRoot'
                    $drive = Ask-DriveLetter
                    Show-Result ("net use {0} {1} /persistent:no" -f $drive,(Quote-Cmd ("\\{0}@SSL\DavWWWRoot\{1}" -f $hostn,$share)))
                }
                '3' {
                    $drive = Ask-DriveLetter
                    Show-Result ("net use {0} /delete /y" -f $drive)
                }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '03' {
            Write-Host '1) Consultar usuario actual'
            Write-Host '2) Mostrar hostname'
            Write-Host '3) Mostrar configuración IP'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                '1' { Show-Result 'conhost.exe --headless cmd.exe /d /c whoami' }
                '2' { Show-Result 'conhost.exe --headless cmd.exe /d /c hostname' }
                '3' { Show-Result 'conhost.exe --headless cmd.exe /d /c ipconfig /all' }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '04' {
            Write-Host '1) Abrir un HTA local de confianza'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            if ($c -eq '1') {
                $path = Ask-NonEmpty 'Ruta local del .hta'
                Show-Result ("mshta.exe {0}" -f (Quote-Cmd $path))
            } else { Write-Host 'Opción inválida.'; Pause-Builder }
        }
        '05' {
            Write-Host '1) Abrir Opciones de Internet'
            Write-Host '2) Abrir impresoras clásicas'
            Write-Host '3) Abrir fecha/hora clásica'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                '1' { Show-Result 'rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl' }
                '2' { Show-Result 'rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL PrintersFolder' }
                '3' { Show-Result 'rundll32.exe shell32.dll,Control_RunDLL timedate.cpl' }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '06' {
            Write-Host '1) Ejecutar VBS local con cscript'
            Write-Host '2) Ejecutar JS local con cscript'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            if ($c -in @('1','2')) {
                $path = Ask-NonEmpty 'Ruta local del script que controlas'
                Show-Result ("cscript.exe //nologo {0}" -f (Quote-Cmd $path))
            } else { Write-Host 'Opción inválida.'; Pause-Builder }
        }
        '07' {
            Write-Host '1) Abrir Bloc de notas'
            Write-Host '2) Abrir Explorador de archivos'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            if ($c -eq '1') { Show-Result 'pcalua.exe -a notepad.exe' }
            elseif ($c -eq '2') { Show-Result 'pcalua.exe -a explorer.exe' }
            else { Write-Host 'Opción inválida.'; Pause-Builder }
        }
        '08' {
            Write-Host '1) Resolver DNS'
            Write-Host '2) Probar puerto TCP'
            Write-Host '3) Calcular SHA256 de fichero'
            Write-Host '4) Consultar cabeceras HTTP'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                '1' {
                    $hostn = Ask-PSValue 'Dominio/host'
                    Show-Result ("powershell.exe -NoProfile -Command `"Resolve-DnsName -Name {0}`"" -f (Quote-PS $hostn))
                }
                '2' {
                    $hostn = Ask-PSValue 'Host/IP autorizado'
                    $port = Ask-Port
                    Show-Result ("powershell.exe -NoProfile -Command `"Test-NetConnection -ComputerName {0} -Port {1}`"" -f (Quote-PS $hostn),$port)
                }
                '3' {
                    $path = Ask-PSValue 'Ruta del fichero'
                    Show-Result ("powershell.exe -NoProfile -Command `"Get-FileHash -Algorithm SHA256 -LiteralPath {0}`"" -f (Quote-PS $path))
                }
                '4' {
                    $url = Ask-PSValue 'URL'
                    Show-Result ("powershell.exe -NoProfile -Command `"Invoke-WebRequest -Method Head -UseBasicParsing -Uri {0} | Select-Object StatusCode,Headers`"" -f (Quote-PS $url))
                }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '09' {
            Write-Host '1) Ping'
            Write-Host '2) nslookup'
            Write-Host '3) whoami'
            Write-Host '4) Listar carpeta'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                '1' { $h=Ask-NonEmpty 'Host/IP'; Show-Result ("cmd.exe /d /c ping {0}" -f (Quote-Cmd $h)) }
                '2' { $h=Ask-NonEmpty 'Dominio/host'; Show-Result ("cmd.exe /d /c nslookup {0}" -f (Quote-Cmd $h)) }
                '3' { Show-Result 'cmd.exe /d /c whoami /all' }
                '4' { $p=Ask-NonEmpty 'Ruta de carpeta'; Show-Result ("cmd.exe /d /c dir {0}" -f (Quote-Cmd $p)) }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '10' {
            Write-Host '1) Resumen de sistema'
            Write-Host '2) Top de procesos por CPU'
            Write-Host '3) Información de red'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                '1' { Show-Result 'powershell.exe -NoProfile -Command "$os=Get-CimInstance Win32_OperatingSystem; [pscustomobject]@{Equipo=$env:COMPUTERNAME;Usuario=$env:USERNAME;OS=$os.Caption;Version=$os.Version}"' }
                '2' { Show-Result 'powershell.exe -NoProfile -Command "Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 Name,Id,CPU"' }
                '3' { Show-Result 'powershell.exe -NoProfile -Command "Get-NetIPConfiguration | Format-List InterfaceAlias,IPv4Address,IPv4DefaultGateway,DNSServer"' }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '11' {
            Write-Host '1) Cabeceras HTTP con curl'
            Write-Host '2) Resolver DNS'
            Write-Host '3) Ping'
            Write-Host '4) SHA256 de fichero'
            Write-Host '5) Probar autenticación SSH sin ejecutar comandos remotos'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                '1' { $u=Ask-NonEmpty 'URL'; Show-Result ("curl -fsSIL --max-time 10 {0}" -f (Quote-Sh $u)) }
                '2' { $h=Ask-NonEmpty 'Dominio/host'; Show-Result ("dig +short {0}" -f (Quote-Sh $h)) }
                '3' { $h=Ask-NonEmpty 'Host/IP'; Show-Result ("ping -c 4 -- {0}" -f (Quote-Sh $h)) }
                '4' { $p=Ask-NonEmpty 'Ruta del fichero'; Show-Result ("sha256sum -- {0}" -f (Quote-Sh $p)) }
                '5' {
                    $h=Ask-NonEmpty 'Host SSH autorizado'
                    $port=Ask-Port
                    Show-Result ("ssh -v -o BatchMode=yes -o ConnectTimeout=5 -p {0} -- {1} exit" -f $port,(Quote-Sh $h))
                }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '12' {
            Write-Host '1) Ver cabeceras HTTP'
            Write-Host '2) Resolver DNS'
            Write-Host '3) Probar un puerto'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            switch ($c) {
                # curl.exe y nslookup.exe son nativos en Windows 10+: se ejecutan
                # aquí mismo, así que usan comillas cmd (Quote-Cmd), no bash.
                '1' { $u=Ask-NonEmpty 'URL'; Show-Result ("curl.exe -fsSIL --max-time 10 {0}" -f (Quote-Cmd $u)) }
                '2' { $h=Ask-NonEmpty 'Dominio/host'; Show-Result ("nslookup {0}" -f (Quote-Cmd $h)) }
                '3' {
                    # nc no es nativo de Windows; el equivalente nativo es Test-NetConnection.
                    $h=Ask-PSValue 'Host/IP autorizado'; $port=Ask-Port
                    Show-Result ("powershell.exe -NoProfile -Command `"Test-NetConnection -ComputerName {0} -Port {1}`"" -f (Quote-PS $h),$port)
                }
                default { Write-Host 'Opción inválida.'; Pause-Builder }
            }
        }
        '13' {
            Write-Host '1) Imprimir texto'
            Write-Host '2) Codificar texto en Base64'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            # printf/base64 no son nativos de Windows: se usa PowerShell nativo.
            if ($c -eq '1') {
                $t = Ask-PSValue 'Texto'
                Show-Result ("powershell.exe -NoProfile -Command `"Write-Output {0}`"" -f (Quote-PS $t))
            } elseif ($c -eq '2') {
                $t = Ask-PSValue 'Texto'
                Show-Result ("powershell.exe -NoProfile -Command `"[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes({0}))`"" -f (Quote-PS $t))
            } else { Write-Host 'Opción inválida.'; Pause-Builder }
        }
        '99' {
            Write-Host '1) Información del sistema Linux'
            Write-Host '2) Información del sistema Windows'
            Write-Host 'B) Volver al menú de categorías'
            $c = Read-Host '>'
            if ($c -match '^[bB]$') { return }
            if ($c -eq '1') { Show-Result 'uname -a; id; ip addr' }
            elseif ($c -eq '2') { Show-Result 'systeminfo & whoami /all & ipconfig /all' }
            else { Write-Host 'Opción inválida.'; Pause-Builder }
        }
    }
}

while ($true) {
    Show-Menu
    $s = Read-Host '>'
    # $null = EOF (stdin cerrado); sin esto el bucle giraría sin fin.
    if ($null -eq $s) { break }
    switch ($s) {
        '1'  { Build-Category '01' }
        '2'  { Build-Category '02' }
        '3'  { Build-Category '03' }
        '4'  { Build-Category '04' }
        '5'  { Build-Category '05' }
        '6'  { Build-Category '06' }
        '7'  { Build-Category '07' }
        '8'  { Build-Category '08' }
        '9'  { Build-Category '09' }
        '10' { Build-Category '10' }
        '11' { Build-Category '11' }
        '12' { Build-Category '12' }
        '13' { Build-Category '13' }
        '14' { Build-Category '99' }
        '15' { Show-CategoryInfo }
        { $_ -in @('0','q','Q') } { break }
        default { Write-Host 'Opción inválida.'; Pause-Builder }
    }
}
