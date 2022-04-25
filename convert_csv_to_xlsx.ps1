$filename = $_.BaseName

#Défini l'emplacement et le séparateur       
$csv = $_.FullName #Emplacement du fichier source
$xlsx = "C:\Users\jopetit\Desktop\Rapport VEEAM\Excel\$filename.xlsx"# Nomme le fichier Excel comme le CSV

$delimiter = ";" #séparateur utilisé dans le fichier CSV

# Création d'un classeur Excel avec une feuille vide
$excel = New-Object -ComObject excel.application
$workbook = $excel.Workbooks.Add(1)   
$worksheet = $workbook.worksheets.Item(1)

# Crée la commande QueryTables.Add et reformate les données
$TxtConnector = ("TEXT;" + $csv)      
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$query.TextFilePlatform = 65001

# Execute & supprime la requête d'import   
$query.Refresh()
$query.Delete()

# Sauvegarde et ferme le classeur en xlsx.  
$Workbook.SaveAs($xlsx,51)
$excel.Quit()