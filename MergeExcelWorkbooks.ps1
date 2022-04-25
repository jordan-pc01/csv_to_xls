$varMaDate = get-date -format "yyyyMMdd"
$sourceFolderPath = "C:\Users\jopetit\Desktop\Rapport VEEAM\Excel"
$finalFileName = "SVR_Veeam_Check_Backup_Status_$varMaDate"
$outSource ="C:\Users\jopetit\Desktop\Rapport VEEAM\Backup_files\"

$OutputFilePath = "$outSource$finalFileName.xlsx"

$XLfiles = Get-ChildItem $sourceFolderPath -Filter *.xlsx

foreach ($XLfile in $XLfiles) {

    <# Hints,
    - If there is only 1 sheet or if you want to import data from the 1st sheet (ordinal, i.e index 0) in the Excel file the below would work
    - Else please use the 'WorkSheetName' parameter to specify the sheet name to import from
    - Use the NoHeader switch of Import-Excel if the source Excel sheets do not contain a header; HeaderName parameter can be used in combination with NoHeader to specify a custom header name
    #>
    Import-Excel $XLfile.FullName | Export-Excel $OutputFilePath -WorksheetName $XLfile.BaseName    

}
Write-Output "Le fichier $finalFileName est dans le répertoire $outSource "