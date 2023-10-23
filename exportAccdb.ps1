<#
.SYNOPSIS
    PowerShell script to export an Access database to SQL format.

.DESCRIPTION
    This script is designed to facilitate the conversion of an Access database to SQL format.
    It generates SQL queries for table creation, data insertion, and relationships, allowing
    for easy migration of the database structure and content to other SQL-based systems.

.AUTHOR
    Festim Nuredini

.VERSION
    1.0

.LASTMODIFIED
    Date: 2023-10-01
    Changes: Better handling of values and data types.

.NOTES
    - Ensure you have PowerShell installed on your system.
    - Modify the script if necessary to match your Access database.
    
#>

# Function to generate SQL query for an Access table including data insertion
Function SqlQueryForTable {
  param(
      [string]$tableName,
      [bool]$exportData
  )

  # AutoNumber attribute
  $dbAutoIncrField = 16
  # Fixed attribute
  $dbFixedField = 1

  # Get the Access table
  $table = $accessDatabase.TableDefs.Item($tableName)

  $primaryKey = ""
  foreach ($index in $table.Indexes) {
    if ($index.Name -eq "PrimaryKey") {
        $primaryKey = $index.Fields[0].Name
        break
    }
  }

  # Get the field names and types
  $fields = @()
  foreach ($field in $table.Fields) {
    $fieldType = switch ($field.Type) {
      1  { "BIT(1) DEFAULT b'0'" }    # dbBoolean
      {$_ -in 3, 4}  { "INT(11) DEFAULT '0'" }    # dbInteger, dbLong
      {$_ -in 5, 6, 7}  { "DOUBLE DEFAULT '0'" }    # dbCurrency, dbSingle, dbDouble
      8  { 
        # this is specific for Axami database
        if ($field.Name -eq "Tidpunkt" -or $field.Name -eq "FaerdigTid") {
          "DATE NULL DEFAULT NULL"
        }
        else {
          "DATETIME NULL DEFAULT NULL"
        }
      }    # dbdate
      10 { 
        if ($field.Properties["AllowZeroLength"].Value -eq $true) {
          "VARCHAR($($fiel.Size)) NULL DEFAULT NULL"
        } else {
          "VARCHAR($($fiel.Size)) NOT NULL"
        } 
      }   # dbText
      12 { "LONGTEXT" }       # dbMemo (Long Text)
      default { "VARCHAR(255)" }  # Default to VARCHAR(255) for unknown types
    }

    # handle ms access auto increment ID fields
    if ($field.Properties["Attributes"].Value -eq ($dbAutoIncrField + $dbFixedField)) {
        $fieldType = "INT(11) NOT NULL AUTO_INCREMENT"
    }
    

    # handle PRIMARY KEY
    if ($field.Name -eq $primaryKey) {
      $fieldType += " PRIMARY KEY"
    }

    $fields += "``$($field.Name)`` $fieldType"
  }

  # Create the SQL CREATE TABLE statement
  $sqlQuery = "CREATE TABLE $tableName (" + [string]::Join(", ", $fields) + ");"

  if ($exportData) {
    if (!($tableName -eq "tblChangeLog" -or $tableName -eq "tblTransLog")) {
      $data = $accessDatabase.OpenRecordset("SELECT * FROM $tableName")
    
      # Check if the recordset is null or empty
      if (!($null -eq $data -or $data.EOF)) {
        # Create the SQL INSERT INTO statements
        $sqlQuery += "`nINSERT INTO $tableName (" + [string]::Join(", ", ($table.Fields | ForEach-Object { $_.Name })) + ")"
        $sqlQuery += "`nVALUES"
  
        while (!$data.EOF) {
            $values = @()
            foreach ($field in $table.Fields) {
              $value = $data.Fields.Item($field.Name).Value
              $value = switch ($data.Fields.Item($field.Name).Type) {  
                1  {
                  if ($null -eq $value -or "" -eq $value) {
                    "False"
                  } else {
                    $value
                  }
                } # dbBoolean
                {$_ -in 3, 4}  {
                  if ($null -eq $value -or "" -eq $value) {
                    "0"
                  } else {
                    $value
                  }
                } # dbInteger, dbLong 
                {$_ -in 5, 6, 7}  {
                  if ($null -eq $value -or "" -eq $value) {
                    "0"
                  } else {
                    "$value".Replace(",", ".")
                  }
                } # dbCurrency, dbSingle, dbDouble
                8  { 
                  if ($null -eq $value -or "" -eq $value) {
                    "NULL"
                  } elseif ($field.Name -eq "Tidpunkt" -or $field.Name -eq "FaerdigTid") {
                    "'" + $value.ToString('yyyy-MM-dd') + "'"  # Format the date
                  } else {
                    "'" + $value.ToString('yyyy-MM-dd HH:mm:ss') + "'"  # Format the date
                  }
                } # dbdate
                {$_ -in 10, 12}  { 
                  if ($null -eq $value) {
                    "NULL"
                  } else {
                    "'$value'"
                  }
                } # dbText, dbMemo (Long Text)
                default { "'$value'" }  # Default to VARCHAR(255) for unknown types
              }
              
              $values += $value
            }
            $sqlQuery += "`n(" + [string]::Join(", ", $values) + "),"
            $data.MoveNext()
        }
  
        $sqlQuery = $sqlQuery.TrimEnd(",")
        $sqlQuery += ";"
      }
  
      $data.Close()
    }
  }

  return $sqlQuery
}

# Function to generate SQL for relationships
Function GenerateRelationshipSql {
  param(
      [string]$primaryTableName,
      [string]$primaryFieldName,
      [string]$foreignTableName,
      [string]$foreignFieldName
  )

  return "ALTER TABLE $foreignTableName ADD FOREIGN KEY (``$foreignFieldName``) REFERENCES $primaryTableName(``$primaryFieldName``);"
}

# Get the directory of the script
$scriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# Find the Access database file (assumed to be the only .accdb file in the script directory)
$accessDatabasePath = Get-ChildItem -Path $scriptDirectory -Filter *.accdb | Select-Object -ExpandProperty FullName

# Check if exactly one .accdb file was found
if (-not $accessDatabasePath) {
  Write-Host "No .accdb file found in the script directory."
  exit
}

# Prompt the user if they want to export data to SQL files
$exportData = Read-Host -Prompt "Do you want to export data to SQL files? (yes/no)"
$exportData = $exportData -eq "yes" -or $exportData -eq "y"

# Prompt user for the database password
$accessPassword = Read-Host -Prompt "Enter the password for the Access database" 

# Open the Access database using the temporary SQL script
try {
  $accessDBEngine = New-Object -ComObject "DAO.DBEngine.120"
  $accessDatabase = $accessDBEngine.OpenDatabase($accessDatabasePath, 0, $false, "MS Access;PWD=$accessPassword")
}
catch {
  Write-Host "Failed to open the Access database: $_.Exception.Message"
  exit
}

# Create the 'export' subfolder if it doesn't exist
$exportFolderPath = Join-Path -Path $scriptDirectory -ChildPath "mysql"
if (-not (Test-Path $exportFolderPath)) {
    New-Item -ItemType Directory -Path $exportFolderPath | Out-Null
} else {
  # Remove all files from the export folder
  Remove-Item -Path $exportFolderPath\* -Recurse -Force
}


# Create an array to store all SQL statements
$sqlStatements = @()

# Progress
$totalTables = $accessDatabase.TableDefs.Count
$currentTable = 0

foreach ($table in $accessDatabase.TableDefs) {
  $tableName = $table.Name
  # Update progress
  $currentTable++
  $progressMessage = "Processing table $tableName ($currentTable of $totalTables)"
  Write-Progress -Activity "Processing Tables" -Status $progressMessage -PercentComplete (($currentTable / $totalTables) * 100)

  if ($table.Attributes -band 0x01 -or $table.Name.ToLower().StartsWith("msys") -or $table.Name.ToLower().StartsWith("~") -or $table.Name -eq "FL") {
    # Ignore system tables
    continue
  }

  
  $sqlQuery = SqlQueryForTable -tableName $tableName -exportData $exportData

  # Add the table creation SQL to the array
  $sqlStatements += $sqlQuery
}

# Loop through relationships and generate SQL
foreach ($relation in $accessDatabase.Relations) {
  if ($relation.Table.ToLower().StartsWith("msys") -or $relation.Table.ToLower().StartsWith("~")) {
    # Ignore system tables
    continue
  }

  if ($relation.ForeignTable.ToLower().StartsWith("msys") -or $relation.ForeignTable.ToLower().StartsWith("~")) {
    # Ignore system tables
    continue
  }

  if ($relation.Table -eq $relation.ForeignTable) {
    #not possible in mysql
    continue
  }

  $relationshipSql = GenerateRelationshipSql `
      -primaryTableName $relation.Table `
      -primaryFieldName $relation.Fields[0].Name `
      -foreignTableName $relation.ForeignTable `
      -foreignFieldName $relation.Fields[0].ForeignName

  # Add the relationship SQL to the array
  $sqlStatements += $relationshipSql
}

# Combine all SQL statements into a single string
$combinedSql = $sqlStatements -join "`n`n"

# Save all SQL to a single file
$outputFilePath = Join-Path -Path $exportFolderPath -ChildPath "database.sql"
Set-Content -Path $outputFilePath -Value $combinedSql

# Close the Access database
$accessDatabase.Close()

# Display export completion message
Write-Host "Export completed successfully. SQL queries are saved to $exportFolderPath."

# Clean up the temporary SQL script
#Remove-Item $tempSqlScriptPath
