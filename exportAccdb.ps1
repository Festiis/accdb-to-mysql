<#
.SYNOPSIS
    PowerShell script to export an Access database to separate SQL files per table.

.DESCRIPTION
    Exports each Access table into its own SQL file.
    - <TableName>.sql for table structure + data
    - relationships.sql for all foreign key constraints

    Avoids building one huge SQL string in memory by streaming directly to disk.

.AUTHOR
    Festim Nuredini

.VERSION
    2.0

.LASTMODIFIED
    Date: 2025-09-29
    Changes: Export per table, streaming inserts, better performance on large datasets.
#>

Function SqlQueryForTable {
  param(
      [string]$tableName,
      [bool]$exportData,
      [bool]$escapeBackslashes,
      [string]$outputPath
  )

  $dbAutoIncrField = 16
  $dbFixedField = 1

  $table = $accessDatabase.TableDefs.Item($tableName)

  $primaryKey = ""
  foreach ($index in $table.Indexes) {
    if ($index.Name -eq "PrimaryKey") {
      $primaryKey = $index.Fields[0].Name
      break
    }
  }

  $fields = @()
  foreach ($field in $table.Fields) {
    $fieldType = switch ($field.Type) {
      1  { "BIT(1) DEFAULT b'0'" }
      {$_ -in 3, 4} { "INT(11) DEFAULT '0'" }
      {$_ -in 5, 6, 7} { "DOUBLE DEFAULT '0'" }
      8 {
        if ($field.Name -eq "Tidpunkt" -or $field.Name -eq "FaerdigTid") { "DATE NULL DEFAULT NULL" }
        else { "DATETIME NULL DEFAULT NULL" }
      }
      10 {
        if ($field.Properties["AllowZeroLength"].Value -eq $true -and $field.Name -ne $primaryKey) {
          "VARCHAR($($field.Size)) NULL DEFAULT NULL"
        } else { "VARCHAR($($field.Size)) NOT NULL" }
      }
      12 { "LONGTEXT" }
      default { "VARCHAR(255)" }
    }

    if ($field.Properties["Attributes"].Value -eq ($dbAutoIncrField + $dbFixedField)) {
      $fieldType = "INT(11) NOT NULL AUTO_INCREMENT"
    }

    if ($field.Name -eq $primaryKey) {
      $fieldType += " PRIMARY KEY"
    }

    $fields += "``$($field.Name)`` $fieldType"
  }

  # Write CREATE TABLE
  "CREATE TABLE $tableName (" + [string]::Join(", ", $fields) + ");" | Out-File -FilePath $outputPath -Encoding UTF8

  if ($exportData -and !($tableName -in @("tblChangeLog","tblTransLog"))) {
    $data = $accessDatabase.OpenRecordset("SELECT * FROM $tableName")

    if (!($null -eq $data -or $data.EOF)) {
      $fieldNames = ($table.Fields | ForEach-Object { $_.Name })
      "INSERT INTO $tableName (" + [string]::Join(", ", $fieldNames) + ") VALUES" | Out-File -FilePath $outputPath -Append -Encoding UTF8

      $firstRow = $true
      while (!$data.EOF) {
        $values = @()
        foreach ($field in $table.Fields) {
          $value = $data.Fields.Item($field.Name).Value
          $value = switch ($data.Fields.Item($field.Name).Type) {
            1 { if ($null -eq $value -or "" -eq $value) { "False" } else { $value } }
            {$_ -in 3,4} { if ($null -eq $value -or "" -eq $value) { "0" } else { $value } }
            {$_ -in 5,6,7} { if ($null -eq $value -or "" -eq $value) { "0" } else { "$value".Replace(",",".") } }
            8 { if ($null -eq $value -or "" -eq $value) { "NULL" } elseif ($field.Name -in @("Tidpunkt","FaerdigTid")) { "'" + $value.ToString('yyyy-MM-dd') + "'" } else { "'" + $value.ToString('yyyy-MM-dd HH:mm:ss') + "'" } }
            {$_ -in 10,12} { if ($null -eq $value) { "NULL" } elseif ($escapeBackslashes) { "'" + "$value".Replace("\","\\") + "'" } else { "'$value'" } }
            default { if ($null -eq $value) { "NULL" } else { "'$value'" } }
          }
          $values += $value
        }
        $rowSql = "(" + [string]::Join(", ", $values) + ")"
        if ($firstRow) { $firstRow = $false; $rowSql | Out-File -FilePath $outputPath -Append -Encoding UTF8 }
        else { ("," + $rowSql) | Out-File -FilePath $outputPath -Append -Encoding UTF8 }
        $data.MoveNext()
      }
      ";" | Out-File -FilePath $outputPath -Append -Encoding UTF8
    }
    $data.Close()
  }
}

Function GenerateRelationshipSql {
  param(
      [string]$primaryTableName,
      [string]$primaryFieldName,
      [string]$foreignTableName,
      [string]$foreignFieldName
  )
  return "ALTER TABLE $foreignTableName ADD FOREIGN KEY (``$foreignFieldName``) REFERENCES $primaryTableName(``$primaryFieldName``);"
}

# --- Main script ---
$scriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$accessDatabasePath = Get-ChildItem -Path $scriptDirectory -Filter *.accdb | Select-Object -ExpandProperty FullName
if (-not $accessDatabasePath) { Write-Host "No .accdb file found in the script directory."; exit }

$exportData = (Read-Host "Do you want to export data to SQL files? (yes/no)") -match "^(y|yes)$"
$escapeBackslashes = $false
if ($exportData) { $escapeBackslashes = (Read-Host "Do you want to escape backslashes? (yes (MySQL)/no)") -match "^(y|yes)$" }

$accessPassword = Read-Host -Prompt "Enter the password for the Access database" -AsSecureString
$accessPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($accessPassword))

try {
  $accessDBEngine = New-Object -ComObject "DAO.DBEngine.120"
  $accessDatabase = $accessDBEngine.OpenDatabase($accessDatabasePath, 0, $false, "MS Access;PWD=$accessPassword")
}
catch { Write-Host "Failed to open the Access database: $_.Exception.Message"; exit }

$exportFolderPath = Join-Path -Path $scriptDirectory -ChildPath "mysql"
if (-not (Test-Path $exportFolderPath)) { New-Item -ItemType Directory -Path $exportFolderPath | Out-Null }
else { Remove-Item -Path $exportFolderPath\* -Recurse -Force }

$totalTables = $accessDatabase.TableDefs.Count
$currentTable = 0

foreach ($table in $accessDatabase.TableDefs) {
  $tableName = $table.Name
  $currentTable++
  Write-Progress -Activity "Processing Tables" -Status "Processing $tableName ($currentTable of $totalTables)" -PercentComplete (($currentTable/$totalTables)*100)

  if ($table.Attributes -band 0x01 -or $tableName.ToLower().StartsWith("msys") -or $tableName.ToLower().StartsWith("~") -or $tableName -eq "FL") { continue }

  $outputPath = Join-Path -Path $exportFolderPath -ChildPath "$tableName.sql"
  SqlQueryForTable -tableName $tableName -exportData $exportData -escapeBackslashes $escapeBackslashes -outputPath $outputPath
}

# Export relationships
$relPath = Join-Path -Path $exportFolderPath -ChildPath "relationships.sql"
foreach ($relation in $accessDatabase.Relations) {
  if ($relation.Table.ToLower().StartsWith("msys") -or $relation.ForeignTable.ToLower().StartsWith("msys") -or $relation.Table -eq $relation.ForeignTable) { continue }
  GenerateRelationshipSql -primaryTableName $relation.Table -primaryFieldName $relation.Fields[0].Name -foreignTableName $relation.ForeignTable -foreignFieldName $relation.Fields[0].ForeignName |
    Out-File -FilePath $relPath -Append -Encoding UTF8
}

$accessDatabase.Close()
Write-Host "Export completed. Each table has its own .sql file in $exportFolderPath."
