function Release-Ref ($ref) {
([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
[System.__ComObject]$ref) -gt 0)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
}
# -----------------------------------------------------
$objExcel = new-object -comobject excel.application 

# Visualização da planilha / Caso Não queira ver só trocar de $True Para $False
$objExcel.Visible = $True 

# Selecionado Via path(Caminho) a planilha de Base
$objWorkbook = $objExcel.Workbooks.Open("C:\Users\william.silva\Downloads\IMPORTAÇÃO.xlsx")

# Selecionando a Primeira Aba da Planilha
$objWorksheet = $objWorkbook.Worksheets.Item(1)

# Selecionando a Segunda Linha / Por Qual linha Começar o Loop
$intRow = 2
Do {
        # Pegando o valor da primeira coluna Como Nome do arquivo
        $name = $objWorksheet.Cells.Item($intRow, 1).Value()
        
        # Pegando o valor da segunda coluna Como URL de download
        $uri = $objWorksheet.Cells.Item($intRow, 2).Value()
        
        # Se tiver valor (URL) Fazer requisição Curl
        if ($uri){
        Write-Output ($name)
        
        # Criar a pasta "\Downloads\Res" ou Trocar o caminho nessa linha de destino do arquivo baixado. 
        curl $uri -OutFile "C:\Users\$($env:USERNAME)\Downloads\Res\$($name)-1.pdf" -ErrorAction SilentlyContinue
        }
        $intRow++
}
While ($objWorksheet.Cells.Item($intRow,1).Value() -ne $null)
$objExcel.Quit()
$a = Release-Ref($objWorksheet)
$a = Release-Ref($objWorkbook)
$a = Release-Ref($objExcel)
