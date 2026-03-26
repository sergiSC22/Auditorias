# Script para exportar datos del CSV a Excel con formato específico

param(
    [string]$csvPath
)

try {
    # Verificar si existe el archivo CSV
    if (-not (Test-Path $csvPath)) {
        Write-Host "Error: No se encontró el archivo CSV en la ruta $csvPath" -ForegroundColor Red
        exit 1
    }

    # Crear objeto Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    
    # Crear un nuevo libro de Excel
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "Agroal"
    
    # Cambiar el título de la ventana de Excel
    $excel.Caption = "Auditoria de empresas"
    # Cambiar las propiedades del libro
    $workbook.Title = "Auditoria de empresas"
    
    # Leer datos del CSV
    $csvData = Import-Csv $csvPath
    
    # Definir exactamente las columnas que queremos mostrar (sin CORREO y sin TIPO)
    $columnasAMostrar = @("Usuario", "Nombre", "Procesador", "RAM", "S.O.", "H.D.", "IP")
    
    # Escribir encabezados en Excel
    for ($i = 0; $i -lt $columnasAMostrar.Count; $i++) {
        $worksheet.Cells.Item(1, $i + 1) = $columnasAMostrar[$i]
    }
    
    # Formatear encabezados
    $headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $columnasAMostrar.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15 # Color de fondo gris claro
    
    # Escribir datos (solo las columnas especificadas)
    $row = 2
    foreach ($record in $csvData) {
        for ($col = 0; $col -lt $columnasAMostrar.Count; $col++) {
            $nombreColumna = $columnasAMostrar[$col]
            $worksheet.Cells.Item($row, $col + 1) = $record.$nombreColumna
        }
        $row++
    }
    

    # Ajustar ancho de columnas automáticamente según el contenido
    $usedRange = $worksheet.UsedRange
    $usedRange.EntireColumn.AutoFit()
    
    # Añadir un poco de espacio extra (20%) para mejor legibilidad
    for ($col = 1; $col -le $columnasAMostrar.Count; $col++) {
        $currentWidth = $worksheet.Columns.Item($col).ColumnWidth
        $worksheet.Columns.Item($col).ColumnWidth = $currentWidth * 1.2
    }
    
    # Aplicar el estilo de tabla verde (estilo 14 de la captura de pantalla)
    
    # 1. Formato de encabezados - verde oscuro con texto blanco
    $headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $columnasAMostrar.Count))
    $headerRange.Interior.Color = 5287936  # Verde oscuro del estilo de tabla
    $headerRange.Font.Color = 16777215  # Blanco
    $headerRange.Font.Bold = $true
    
    # 2. Bordes de la tabla completa - línea gris fina
    $tableRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item($row-1, $columnasAMostrar.Count))
    
    # Configuración de todos los bordes
    $allBorders = 7 # xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlInsideHorizontal, xlInsideVertical
    for ($border = 7; $border -le 12; $border++) {
        $tableRange.Borders.Item($border).LineStyle = 1  # Línea sólida
        $tableRange.Borders.Item($border).Weight = 2     # Peso fino
        $tableRange.Borders.Item($border).Color = 10066329  # Color gris
    }
    
    # 3. Filas alternas con color verde claro
    for ($i = 2; $i -lt $row; $i += 2) {
        $rowRange = $worksheet.Range($worksheet.Cells.Item($i, 1), $worksheet.Cells.Item($i, $columnasAMostrar.Count))
        $rowRange.Interior.Color = 14348274  # Verde muy claro (exactamente el de la imagen)
    }
    
    # 4. Alineación de texto
    $tableRange.HorizontalAlignment = -4131  # xlLeft (alineado a la izquierda como en la imagen)
    $tableRange.VerticalAlignment = -4107    # xlCenter (centrado verticalmente)
    
    # 5. Pequeño espacio interno para las celdas (padding)
    $tableRange.WrapText = $false
    $tableRange.ShrinkToFit = $false
    $tableRange.IndentLevel = 1
    
    # Guardar el archivo Excel
    $excelPath = [System.IO.Path]::ChangeExtension($csvPath, ".xlsx")
    $workbook.SaveAs($excelPath)
    
    Write-Host "Archivo Excel guardado correctamente en: $excelPath" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}