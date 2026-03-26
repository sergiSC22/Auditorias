#  Auditoría Automática de Sistemas

Script de auditoría automática para inventario de equipos 
en entornos empresariales. Recopila información del hardware 
y software de cada equipo y genera un informe en Excel 
con formato profesional.

# ¿Qué hace?

- Recopila automáticamente: usuario, nombre del equipo, 
  procesador, RAM, sistema operativo, disco duro e IPs
- Detecta si la IP es estática o DHCP
- Genera un CSV acumulativo (un equipo por fila)
- Exporta a Excel con formato corporativo:
  cabeceras en verde, filas alternas, bordes y autofit

# Tecnologías

- Batch Script (.bat)
- PowerShell (integrado)
- Excel COM Automation

# Caso de uso real

Desarrollado y usado en entorno empresarial real para 
automatizar el inventario de equipos, eliminando el 
proceso manual de recogida de datos.

## Cómo usarlo

1. Ejecutar `auditoria_automatica.bat` como administrador
2. El script genera `auditoria.csv` en la misma carpeta
3. El archivo `ExportarExcel.ps1`, abre automáticamente 
   el Excel formateado
