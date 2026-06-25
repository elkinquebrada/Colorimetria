# PowerShell script to apply SQL structure
$sqlFile = "c:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\Color\LogicDocs\v4_database_structure.sql"
$instances = @(".\SQLEXPRESS", ".", "(localdb)\MSSQLLocalDB")
$sqlContent = Get-Content -Raw $sqlFile

Write-Host "--- APLICANDO ESTRUCTURA SQL V4 ---" -ForegroundColor Blue

foreach ($inst in $instances) {
    $cs = "Server=$inst;Database=master;Trusted_Connection=True;Connect Timeout=2;"
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($cs)
        $conn.Open()
        Write-Host "Conectado a $inst. Creando Base de Datos..." -ForegroundColor Green
        
        # Split GO to execute chunks
        $chunks = $sqlContent -split "(?m)^\s*GO\s*"
        foreach ($chunk in $chunks) {
            if ($chunk.Trim()) {
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = $chunk
                $cmd.ExecuteNonQuery() | Out-Null
            }
        }
        
        Write-Host "Base de Datos y Tablas creadas exitosamente en $inst." -ForegroundColor Cyan
        $conn.Close()
        exit 0
    } catch {
        # Silent try other instances
    }
}
Write-Error "No se pudo conectar a ninguna instancia de SQL Server."
exit 1
