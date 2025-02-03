<#
	.SYNOPSIS
        Script de instalación de extensiones para Visual Studio Code desde archivos .vsix.

	.DESCRIPTION
        Este script permite instalar múltiples extensiones de Visual Studio Code desde archivos `.vsix` ubicados en el mismo directorio que el script.
        Tiene un una Funcion de Logs para la buena gestion

	.PARAMETER -SystemInstall
		Esto pone el script en modo SystemInstall, configurado para ver por pantalla

    .PARAMETER -ExtensionPath
		Podemos configurarlo para instalar solo una extension

	.EXAMPLE
		Ejecuta con privilegios de administrador o implementación de SCCM
        
        Debe ejecutar el script con uno de los siguientes argumentos válidos:
        -ExtensionPath <ruta_del_plugin>   : Para instalar una extensión específica desde una ruta.
        -SystemInstall                     : Para instalar todas las extensiones de la carpeta Plugins.

		powershell.exe -executionpolicy bypass -file "Instalar.ps1" -SystemInstall
        powershell.exe -executionpolicy bypass -file "Instalar.ps1" -ExtensionPath "C:\ruta\plugin.vsix"

	.NOTES
		===========================================================================
		Created By:		Carlos Campos
		Created Date:	02/02/2025, 19:40 PM
		Version:		1.6
		File:			Instalar.ps1
		Copyright (c)2025 Campos
		===========================================================================

	.LICENSE
		Por la presente, se otorga el permiso, de forma gratuita, a cualquier persona que obtenga una copia
		de este software y archivos de documentación asociados (el software), para tratar
		en el software sin restricción, incluidos los derechos de los derechos
		para usar copiar, modificar, fusionar, publicar, distribuir sublicense y /o vender
		copias del software y para permitir a las personas a quienes es el software
		proporcionado para hacerlo, sujeto a las siguientes condiciones:

		El aviso de derechos de autor anterior y este aviso de permiso se incluirán en todos
		copias o porciones sustanciales del software.

		EL SOFTWARE SE PROPORCIONA TAL CUAL, SIN GARANTÍA DE NINGÚN TIPO, EXPRESA O
		IMPLÍCITA, INCLUYENDO PERO SIN LIMITARSE A LAS GARANTÍAS DE COMERCIABILIDAD
		ADECUACIÓN PARA UN PROPÓSITO PARTICULAR Y NO INFRACCIÓN. EN NINGÚN CASO LOS
		AUTORES O TITULARES DE LOS DERECHOS DE AUTOR SERÁN RESPONSABLES DE NINGUNA RECLAMACIÓN, DAÑOS U OTROS
		RESPONSABILIDAD, YA SEA EN UNA ACCIÓN CONTRACTUAL, AGRAVIO O DE OTRO MODO, QUE SURJA DE
		DE O EN CONEXIÓN CON EL SOFTWARE O EL USO U OTROS TRATOS EN EL tSOFTWARE.
#>
[CmdletBinding()]
param (
    [switch]$SystemInstall,
    [string]$ExtensionPath,
    [ValidateNotNullOrEmpty()][string]$PSConsoleTitle = "Instalar Plugins Visual Code",
    [ValidateNotNullOrEmpty()][string]$logFile = "00 InstalarPluginsVC.log",
    [ValidateNotNullOrEmpty()][string]$logPath = "$env:TEMP\Software"
)

function Set-ConsoleTitle {
    <#
        .SYNOPSIS
            Establece el título de la consola PowerShell.
    #>        
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][String]$ConsoleTitle
    )  
    $host.ui.RawUI.WindowTitle = $ConsoleTitle
}

function Write-Log {
    <#
    .SYNOPSIS
        Sistema de Logs
    
    .DESCRIPTION
        Con esta funcion podemos hacer un sistema de logs muy intuitivos para SCCM o uso personal
    
    .PARAMETER Message
        Write-Log -Message "Ponemos un mensaje"
    
    .PARAMETER LogType
        -LogType "INFO"  #Default
        -LogType "WARNING"
        -LogType "ERROR"
    
    .EXAMPLE
        Write-Log -Message "Ponemos un mensaje" -LogType "ERROR"
        Write-Log "Ponemos un mensaje"
    
    .NOTES
        Primero hay que definir 2 parametros:
        $logFile = "ScriptLog.log"    # Nombre del Fichero
        $logPath = "$env:TEMP"  # Ruta del Fichero en la carpeta %TEMP%
        $logPath = "$env:WINDIR\Logs\Software" Ruta del Fichero en la carpeta Windows\Logs
    #>

    param (
        [string]$Message,
        [string]$LogType = "INFO"
    )

    try {
        # Verifica si el directorio de logs existe, sino lo crea
        if (-Not (Test-Path -Path $logPath)) {
            New-Item -Path $logPath -ItemType Directory -Force
        }
        
        $logFilePath = "$logPath\$logFile"
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "$timestamp [$LogType] $Message"
        Add-Content -Path $logFilePath -Value $logMessage -Force
    }
    catch {
        Write-Log "Error al iniciar el log o crear la carpeta de logs: $_" -LogType "ERROR"
    }
}

Write-Log "Inicio del script $PSConsoleTitle." 

function Get-RelativePath {
    <#
        .SYNOPSIS
            Devuelve la ubicación del script en ejecución.

        .EXAMPLE
            $pRuta = Get-RelativePath
    #>
    [CmdletBinding()][OutputType([string])]
    param ()
    return (Split-Path $SCRIPT:MyInvocation.MyCommand.Path -Parent) + "\"
}
    
function Install-VSCodeExtension {
    <#
    .SYNOPSIS
        Instala todas las extensiones de VSCode desde los archivos .vsix en la carpeta 'Plugins' dentro del directorio del script o puede instalar solo uno.
    .DESCRIPTION
        Esta función busca todos los archivos `.vsix` en la carpeta 'Plugins' dentro del directorio donde se encuentra el script y luego instala cada una de las extensiones en VSCode.
        Si alguna extensión ya está instalada, se omite, y si la instalación falla, se informa del error.
        Si se proporciona un archivo específico como argumento, instala solo esa extensión.
    .NOTES
        Debe ejecutar el script con uno de los siguientes argumentos válidos:
        -ExtensionPath <ruta_del_plugin>   : Para instalar una extensión específica desde una ruta.
        -SystemInstall                     : Para instalar todas las extensiones de la carpeta Plugins.

        Ejemplo: Install-VSCodeExtension -ExtensionPath 'C:\ruta\plugin.vsix'
        Ejemplo: Install-VSCodeExtension -SystemInstall

        ===========================================================================
		Created By:		Carlos Campos
		Created Date:	02/02/2025, 19:40 PM
		Version:		1.8
		Function:		Install-VSCodeExtension
		Copyright (c)2025 Campos
		===========================================================================

    #>
    [CmdletBinding()]
    param(
        [string]$ExtensionPath = $null
    )

    $InstallLocation = "$env:USERPROFILE\.vscode\extensions"
    $Success = $true

    if ($SystemInstall) {
        # Obtener la ruta del script y agregar la carpeta 'Plugins'
        $scriptPath = Join-Path -Path (Get-RelativePath) -ChildPath "Plugins"

        # Buscar todos los archivos .vsix en la carpeta 'Plugins'
        $vsixFiles = Get-ChildItem -Path $scriptPath -Filter *.vsix

        if ($vsixFiles.Count -eq 0) {
            Write-Host "No se encontraron archivos .vsix en la carpeta del script." -ForegroundColor Red
            Write-Log "No se encontraron archivos .vsix en la carpeta del script." -LogType "ERROR"
            return $false
        }

        # Obtener la lista de extensiones instaladas con versiones
        $installedExtensions = code --list-extensions --show-versions | ForEach-Object {
            if ($_ -match '^(.+?)@(.+)$') {
                [PSCustomObject]@{
                    Name    = $matches[1]
                    Version = $matches[2]
                }
            }
        }

        $maxExtensionLength = ($vsixFiles | ForEach-Object { $_.Name.Length }) | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

        # Función para comparar versiones
        function Compare-Versions ($version1, $version2) {
            $v1 = $version1 -split '\.'
            $v2 = $version2 -split '\.'
            for ($i = 0; $i -lt [math]::Max($v1.Count, $v2.Count); $i++) {
                $num1 = if ($i -lt $v1.Count) { [int]$v1[$i] } else { 0 }
                $num2 = if ($i -lt $v2.Count) { [int]$v2[$i] } else { 0 }
                if ($num1 -lt $num2) { return -1 }
                if ($num1 -gt $num2) { return 1 }
            }
            return 0
        }

        # Iterar sobre cada archivo .vsix y verificar su instalación
        foreach ($file in $vsixFiles) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            # Extraer nombre y versión del archivo vsix
            if ($baseName -match '^(.*)-(\d+\.\d+\.\d+)$') {
                $extensionName = $matches[1]
                $fileVersion = $matches[2]

                # Crear un espacio en blanco para alinear las extensiones
                $spacePaddingLength = $maxExtensionLength - $baseName.Length
                if ($spacePaddingLength -lt 0) { $spacePaddingLength = 0 }
                $spacePadding = " " * $spacePaddingLength

                Write-Host "Instalando la Extension $extensionName ($fileVersion)...$spacePadding`t" -NoNewline
                Write-Log "Instalando la Extension $extensionName ($fileVersion)..."

                # Verificar si la extensión ya está instalada
                $installedExtension = $installedExtensions | Where-Object { $_.Name -eq $extensionName }

                if ($installedExtension) {
                    $installedVersion = $installedExtension.Version
                    $comparison = Compare-Versions $installedVersion $fileVersion

                    if ($comparison -eq 0) {
                        Write-Host "✅ Ya instalada" -ForegroundColor Cyan
                        Write-Log "La extensión $extensionName ya está instalada con la versión ($fileVersion)."
                        continue
                    }
                    elseif ($comparison -gt 0) {
                        Write-Host "⚠️ Version Instalada $installedVersion" -ForegroundColor Yellow
                        Write-Log "La versión instalada de $extensionName ($installedVersion) es más reciente que la disponible ($fileVersion), no se instala."
                        continue
                    }
                    else {
                        Write-Host "🔄 Versión desactualizada ($installedVersion), se actualizará." -ForegroundColor Magenta
                        Write-Log "La extensión $extensionName tiene una versión antigua ($installedVersion), se procederá a actualizar a $fileVersion."
                    }
                }

                & code --install-extension "`"$($file.FullName)`"" > $null 2>&1
                $exitCode = $LASTEXITCODE

                Start-Sleep -Seconds 1

                # Verificar el código de salida
                if ($exitCode -eq 0) {
                    Write-Host "✅ OK" -ForegroundColor Green
                    Write-Log "Se ha instalado la extensión: $extensionName ($fileVersion)."
                    $Success = $true
                }
                else {
                    Write-Host "❌ ERROR" -ForegroundColor Red
                    Write-Log "Error al instalar la extensión $extensionName ($fileVersion)." -LogType "ERROR"
                    $Success = $false
                }
            }
            else {
                Write-Host "❌ El archivo '$($file.Name)' no tiene el formato esperado 'nombre-version.vsix'." -ForegroundColor Red
                Write-Log "Formato incorrecto en '$($file.Name)'." -LogType "ERROR"
            }
        }

        if ($Success) {
            Write-Host "$spacePadding`t✅ Todas las extensiones se instalaron correctamente." -ForegroundColor Green
            Write-Log "Todas las extensiones se instalaron correctamente." -LogType "WARNING"
        }
        else {
            Write-Host "$spacePadding`t❌ Hubo errores en la instalación de algunas extensiones." -ForegroundColor Red
            Write-Log "Errores en la instalación de algunas extensiones." -LogType "ERROR"
        }

        return $Success

    }    
    if ($ExtensionPath) {
        # Modo instalación de una sola extensión
        if (-not (Test-Path -Path $ExtensionPath)) {
            Write-Host "El archivo $ExtensionPath no existe." -ForegroundColor Red
            Write-Log "El archivo $ExtensionPath no existe." -LogType "ERROR"
            return $false
        }

        $ExtensionName = [System.IO.Path]::GetFileNameWithoutExtension($ExtensionPath)
        Write-Host "Instalando la extensión $ExtensionName...`t" -NoNewline
        Write-Log "Instalando la extensión $ExtensionName..."

        if (Test-Path "$InstallLocation\$ExtensionName*") {
            Write-Host "Ya instalada" -ForegroundColor Cyan
            Write-Log "Ya estaba instalada la extension $ExtensionName"
            return $true
        }

        #Write-Host "Instalando la extensión desde: $ExtensionPath" -ForegroundColor Cyan
        & code --install-extension "`"$ExtensionPath`"" > $null 2>&1
        $exitCode = $LASTEXITCODE

        Start-Sleep -Seconds 1

        if ($exitCode -eq 0 -and (Test-Path "$InstallLocation\$ExtensionName*")) {
            Write-Host "OK" -ForegroundColor Green
            Write-Log "Se ha instalado la Extension: $ExtensionName"
            $Success = $true
        }
        else {
            Write-Host "ERROR" -ForegroundColor Red
            Write-Log "Error al instalar la extension ${ExtensionName}: $_" -LogType "ERROR"
            $Success = $false
        }
        return $Success
    }

    return $Success
}


Set-ConsoleTitle -ConsoleTitle $PSConsoleTitle
Clear-Host
$Success = $true
# Comprobación adicional si $SystemInstall.IsPresent está definido
If ($SystemInstall) {
    Write-Host "Instalando Extensiones:" -ForegroundColor Magenta
    $Status = Install-VSCodeExtension
    If ($Status = $false) {
        $Success = $false
    }
} 
If ($ExtensionPath) {
    $Status = Install-VSCodeExtension -ExtensionPath "$ExtensionPath"
    If ($Status -eq $false) { 
        $Success = $false
    }
}
If (-Not $SystemInstall -and -Not $ExtensionPath) {
    Write-Host "Debe ejecutar el script con uno de los siguientes argumentos válidos:" -ForegroundColor Magenta
    Write-Host "-ExtensionPath <ruta_del_plugin>   : Para instalar una extensión específica desde una ruta." -ForegroundColor Cyan
    Write-Host "-SystemInstall                     : Para instalar todas las extensiones de la carpeta Plugins." -ForegroundColor Cyan
    Write-Host "Ejemplo: .\Instalar.ps1 -ExtensionPath 'C:\ruta\plugin.vsix'" -ForegroundColor Green
    Write-Host "Ejemplo: .\Instalar.ps1 -SystemInstall" -ForegroundColor Green
}
# Si alguna instalación falló, salir con código de error
If (-Not $Success) {
    Exit 1
}
#Write-Host "Instalación completada con éxito." -ForegroundColor Green
Write-Log "Instalación completada con exito." -LogType "WARNING"
exit 0
    