# PowerShell script to test save and read
$cs = "Server=(localdb)\MSSQLLocalDB;Database=ColorimetriaDB;Trusted_Connection=True;Connect Timeout=5;"
Write-Host "--- PRUEBA DE GUARDADO Y LECTURA (V4) ---" -ForegroundColor Blue

try {
    $conn = New-Object System.Data.SqlClient.SqlConnection($cs)
    $conn.Open()
    Write-Host "[1] Conexión establecida con (localdb)\MSSQLLocalDB" -ForegroundColor Green

    # 1. Insert Cabecera
    Write-Host "[2] Insertando cabecera de prueba..." -NoNewline
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "INSERT INTO tbl_analisis_cabecera (ShadeName, LotNo, FechaRegistro, Status_TL84) OUTPUT INSERTED.Id_Lote VALUES ('TEST_SHADE', 'TEST_LOT', GETDATE(), 'PASS')"
    $idLote = $cmd.ExecuteScalar()
    Write-Host " OK (ID: $idLote)" -ForegroundColor Green

    # 2. Insert Detalle
    Write-Host "[3] Insertando detalle de prueba..." -NoNewline
    $cmd.CommandText = "INSERT INTO tbl_analisis_detalle (Id_Lote, DyeCode, DyeName, Concentration_Original, Proportion_Original) VALUES ($idLote, 'D001', 'Dye Test Red', 0.12345, 100.0)"
    $cmd.ExecuteNonQuery() | Out-Null
    Write-Host " OK" -ForegroundColor Green

    # 3. Read back (Join)
    Write-Host "[4] Recuperando datos (JOIN)..." -NoNewline
    $cmd.CommandText = "SELECT c.ShadeName, d.DyeName FROM tbl_analisis_cabecera c INNER JOIN tbl_analisis_detalle d ON c.Id_Lote = d.Id_Lote WHERE c.Id_Lote = $idLote"
    $reader = $cmd.ExecuteReader()
    if ($reader.Read()) {
        Write-Host " OK" -ForegroundColor Green
        Write-Host "  -> Shade: $($reader['ShadeName'])" -ForegroundColor Gray
        Write-Host "  -> Dye:   $($reader['DyeName'])" -ForegroundColor Gray
    }
    $reader.Close()

    # 4. Cleanup (Optional, but I'll leave it for the user to see in the table)
    # Write-Host "[5] Limpiando datos de prueba..." -NoNewline
    # $cmd.CommandText = "DELETE FROM tbl_analisis_cabecera WHERE Id_Lote = $idLote"
    # $cmd.ExecuteNonQuery() | Out-Null
    # Write-Host " OK" -ForegroundColor Cyan

    $conn.Close()
    Write-Host "------------------------------------"
    Write-Host "¡VALIDACIÓN EXITOSA!" -ForegroundColor Green
} catch {
    Write-Host "[FALLO]" -ForegroundColor Red
    Write-Host " Error: $($_.Exception.Message)" -ForegroundColor Gray
}
