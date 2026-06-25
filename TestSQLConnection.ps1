# PowerShell script to test SQL Server connectivity
$instances = @(".\SQLEXPRESS", ".", "(localdb)\MSSQLLocalDB")

Write-Host "--- VALIDACION DE CONEXION SQL ---" -ForegroundColor Blue

foreach ($inst in $instances) {
    $cs = "Server=$inst;Database=master;Trusted_Connection=True;Connect Timeout=2;"
    Write-Host "Instancia: $inst ... " -NoNewline
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($cs)
        $conn.Open()
        Write-Host "[CONECTADO]" -ForegroundColor Green
        
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT name FROM master.dbo.sysdatabases WHERE name = 'ColorimetriaDB'"
        $db = $cmd.ExecuteScalar()
        
        if ($db) {
            Write-Host " -> OK: 'ColorimetriaDB' existe. " -ForegroundColor Cyan
        } else {
            Write-Host " -> ADVERTENCIA: 'ColorimetriaDB' no existe. " -ForegroundColor Yellow
        }
        $conn.Close()
    } catch {
        Write-Host "[FALLO]" -ForegroundColor Red
        # Write-Host " Error: $($_.Exception.Message)"
    }
}
