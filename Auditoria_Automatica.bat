@echo off
:: Script para auditoría automática de equipos
setlocal enabledelayedexpansion

echo ========================================================
echo         AUDITORIA AUTOMATICA DE SISTEMAS
echo ========================================================
echo.

:: Definir la ubicación del archivo de auditoría
set "archivo_auditoria=%~dp0auditoria.csv"

:: Verificar si el archivo ya existe, si no, agregar encabezados
echo Verificando archivo de auditoria...
if not exist "%archivo_auditoria%" (
    echo Usuario,Nombre,Procesador,RAM,S.O.,H.D.,IP,TIPO > "%archivo_auditoria%"
    echo [INFO] Archivo de auditoria creado con encabezados.
) else (
    echo [INFO] Usando archivo de auditoria existente.
)

echo.
echo [INFO] Iniciando recopilacion de informacion del sistema...

:: Variables para almacenar la información
set "username=%USERNAME%"
set "computername=Desconocido"
set "processor=Desconocido"
set "ram=Desconocido"
set "os=Desconocido"
set "harddisk=Desconocido"
set "ips_and_types=No IP"
set "tipo=PC"

:: Obtener nombre del equipo
for /f "tokens=*" %%a in ('hostname') do (
    set "computername=%%a"
)

:: Obtener procesador
for /f "tokens=*" %%a in ('powershell -command "(Get-CimInstance Win32_Processor).Name"') do (
    set "processor=%%a"
)

:: Obtener RAM
for /f "tokens=*" %%a in ('powershell -command "[math]::Ceiling((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB)"') do (
    set "ram=%%a GB"
)

:: Obtener Sistema Operativo
for /f "tokens=*" %%a in ('powershell -command "(Get-WmiObject Win32_OperatingSystem).Caption"') do (
    set "os=%%a"
)

:: Obtener tamaño de todos los discos duros (suma todos los discos físicos)
for /f "tokens=*" %%a in ('powershell -command "$total = 0; Get-Disk | ForEach-Object { $total += [math]::Ceiling($_.Size / 1GB) }; $total"') do (
    set "harddisk=%%a GB"
)

:: Crear un script PowerShell para obtener IPs y sus tipos
echo $result = @() > "%TEMP%\get_ip_info.ps1"
echo Get-NetIPAddress -AddressFamily IPv4 ^| Where-Object { $_.IPAddress -notmatch '^169\.254\.' -and $_.IPAddress -ne '127.0.0.1' } ^| ForEach-Object { >> "%TEMP%\get_ip_info.ps1"
echo     $ipType = if ($_.PrefixOrigin -eq 'Dhcp') { 'DHCP' } else { 'Estatica' } >> "%TEMP%\get_ip_info.ps1"
echo     $result += "$($_.IPAddress) [$ipType]" >> "%TEMP%\get_ip_info.ps1"
echo } >> "%TEMP%\get_ip_info.ps1"
echo if ($result.Count -eq 0) { Write-Host "No IP" } else { Write-Host ($result -join ", ") } >> "%TEMP%\get_ip_info.ps1"

:: Ejecutar el script PowerShell en modo seguro
for /f "tokens=*" %%a in ('powershell -ExecutionPolicy Bypass -File "%TEMP%\get_ip_info.ps1"') do (
    set "ips_and_types=%%a"
)

:: Mostrar información recopilada
echo.
echo === INFORMACIóN DEL SISTEMA ===
echo Usuario: %username%
echo Equipo: %computername%
echo Procesador: %processor%
echo RAM: %ram%
echo Sistema Operativo: %os%
echo Disco Duro: %harddisk%
echo IPs y Tipos: %ips_and_types%
echo =============================
echo.

:: Agregar línea al archivo CSV
echo %username%,%computername%,%processor%,%ram%,%os%,%harddisk%,%ips_and_types%,%tipo% >> "%archivo_auditoria%"
echo [INFO] Datos guardados correctamente en %archivo_auditoria%
echo.

echo === PROCESO DE AUDITORIA COMPLETADO ===
echo.
echo Los datos han sido exportados al archivo CSV.
echo Ruta del archivo: %archivo_auditoria%
echo.

:: Exportar a Excel si existe el script
if exist "%~dp0ExportarExcelAgroal.ps1" (
    echo [INFO] Exportando a Excel...
    powershell -ExecutionPolicy Bypass -File "%~dp0ExportarExcelAgroal.ps1" -csvPath "%archivo_auditoria%"
  echo.
    
) else (
    :: Simplemente abrir el archivo con el programa predeterminado
    echo [INFO] Abriendo archivo CSV...
    start "" "%archivo_auditoria%"
    
    echo.
    echo Si Excel no se abre automáticamente, puedes encontrar el archivo CSV en:
    echo %archivo_auditoria%
    echo.
    echo [INFO] El programa se cerrará automáticamente en 2 segundos...
)

:: Limpiar archivo temporal
if exist "%TEMP%\get_ip_info.ps1" del "%TEMP%\get_ip_info.ps1"

echo [INFO] El programa se cerrará automáticamente en 2 segundos...
    timeout /t 2 /nobreak > nul
