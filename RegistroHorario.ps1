# Script: resumenAsistencias.ps1
# Autor: Fernando Cisneros Chavez (verdevenus23@gmail.com)
# Fecha: 21 de agosto de 2025
# Licencia: MIT

Add-Type -AssemblyName System.Windows.Forms

# Crear la ventana
$form = New-Object System.Windows.Forms.Form
$form.Text = "Resumen de Asistencias"
$form.Size = New-Object System.Drawing.Size(320,210)
$form.StartPosition = "CenterScreen"

# Función para pedir entrada de texto
function Pedir-Input($mensaje, $titulo) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    return [Microsoft.VisualBasic.Interaction]::InputBox($mensaje, $titulo, "")
}

# Checkbox modo depuracion
$chkDebug = New-Object System.Windows.Forms.CheckBox
$chkDebug.Text = "Mostrar consola"
$chkDebug.Size = New-Object System.Drawing.Size(260,20)
$chkDebug.Location = New-Object System.Drawing.Point(30,130)
$chkDebug.Font = New-Object System.Drawing.Font("Segoe UI", 8)

# Boton Asistencias (generarResumen.py)
$btnAsistencias = New-Object System.Windows.Forms.Button
$btnAsistencias.Text = "Reporte de asistencias"
$btnAsistencias.Size = New-Object System.Drawing.Size(100,35)
$btnAsistencias.Location = New-Object System.Drawing.Point(30,40)
$btnAsistencias.Add_Click({
    $hoy  = Get-Date -Format "yyyy-MM-dd"
    $ayer = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")

    $fechaInicio = Pedir-Input "Ingresa la fecha de inicio ( ej. ayer: $ayer )" "Fecha Inicio"
    if (-not $fechaInicio) { $fechaInicio = $ayer }

    $fechaFin = Pedir-Input "Ingresa la fecha de fin ( ej. hoy: $hoy )" "Fecha Fin"
    if (-not $fechaFin) { $fechaFin = $hoy }

    $asistenciasPath = Join-Path -Path (Get-Location) -ChildPath "scripts\generarResumen.py"

    # Deshabilitar botón mientras corre para evitar doble ejecución
    $btnAsistencias.Enabled = $false
    $btnAsistencias.Text    = "Ejecutando..."

    if ($chkDebug.Checked) {
        # python.exe + consola visible — esperamos a que termine para leer el código de salida
        $proc = Start-Process "python" `
            -ArgumentList "`"$asistenciasPath`"", $fechaInicio, $fechaFin, "--debug" `
            -PassThru -Wait
    } else {
        # pythonw.exe sin consola — también esperamos
        $proc = Start-Process "pythonw" `
            -ArgumentList "`"$asistenciasPath`"", $fechaInicio, $fechaFin `
            -PassThru -Wait
    }

    # Restaurar botón
    $btnAsistencias.Enabled = $true
    $btnAsistencias.Text    = "Reporte de asistencias"

    # Python ya muestra sus propios diálogos para todos los errores conocidos (códigos 1–4).
    # PowerShell solo interviene si el proceso no pudo arrancar (código negativo),
    # lo que indica que python/pythonw no se encontró o el archivo .py no existe.
    if ($null -eq $proc -or $proc.ExitCode -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se pudo iniciar el proceso.`n`n¿Está instalado Python y en el PATH?`nVerifica también que exista:`n$asistenciasPath",
            "Error al iniciar",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

$btnProcesar = New-Object System.Windows.Forms.Button
$btnProcesar.Text = "Procesar"
$btnProcesar.Size = New-Object System.Drawing.Size(100,35)
$btnProcesar.Location = New-Object System.Drawing.Point(150,40)
$btnProcesar.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    $openFileDialog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx"
    $openFileDialog.Title = "Selecciona un archivo Excel"
    $openFileDialog.CheckFileExists = $true

    $dialogResult = $openFileDialog.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $openFileDialog.FileName
        Write-Host "Archivo seleccionado: $filePath"
    } else {
        Write-Host "No se seleccionó ningún archivo."
        return
    }

    $procesarPath = Join-Path -Path (Get-Location) -ChildPath "scripts\ResumirReporte.py"

    # pythonw ejecuta sin abrir consola — los dialogos se muestran desde el script
    Start-Process "pythonw" -ArgumentList "`"$procesarPath`"", "`"$filePath`""
})

$btnSalir = New-Object System.Windows.Forms.Button
$btnSalir.Text = "Salir"
$btnSalir.Size = New-Object System.Drawing.Size(120,30)
$btnSalir.Location = New-Object System.Drawing.Point(80,100)
$btnSalir.Add_Click({
    $form.Close()
})

$iconBase64 = @"
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
"@ -replace '\s',''
$iconBytes  = [Convert]::FromBase64String($iconBase64)
$iconStream = New-Object IO.MemoryStream(,$iconBytes)
$form.Icon  = New-Object System.Drawing.Icon($iconStream)

# Agregar botones y mostrar ventana
$form.Controls.Add($chkDebug)
$form.Controls.Add($btnAsistencias)
$form.Controls.Add($btnProcesar)
$form.Controls.Add($btnSalir)

$form.Add_FormClosed({[System.Environment]::Exit(0)})
[void]$form.ShowDialog()