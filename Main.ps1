function Release-Ref ($ref) {
([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
[System.__ComObject]$ref) -gt 0)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
}
# -----------------------------------------------------
$objExcel = new-object -comobject excel.application 
$objExcel.Visible = $True 
$objWorkbook = $objExcel.Workbooks.Open("C:\Users\william.silva\Documents\teste.xlsx")
$objWorksheet = $objWorkbook.Worksheets.Item(1)
$intRow = 1
Do {
        $name = $objWorksheet.Cells.Item($intRow, 1).Value()
        $uri = $objWorksheet.Cells.Item($intRow, 2).Value()
        if ($uri){
        curl $uri -OutFile C:\Users\$env:USERNAME\Downloads\Res\$name.pdf -ErrorAction SilentlyContinue
        }
        $intRow++
}
While ($objWorksheet.Cells.Item($intRow,1).Value() -ne $null)
$objExcel.Quit()
$a = Release-Ref($objWorksheet)
$a = Release-Ref($objWorkbook)
$a = Release-Ref($objExcel)
