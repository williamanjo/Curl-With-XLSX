# Curl-With-XLSX

Ferramenta Desenvolvida em PowerShell Para o Download de Arquivos listados em uma planilha.<br>

# Como Utilizar 

 - Planilha deve conter **Nome** do Arquivo e **URL** na mesma Linha em colunas **Diferêntes** , mas na mesma **ABA**.
 - Configurações a serem Feitas no Algoritmo:
   - Na linha [#14](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L14) Referenciar A **Planilha** de Base.
   - Na linha [#17](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L17) Indicar Qual **Aba** se encontra a Base.
   - Na linha [#20](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L20) Indicar em Qual **linha** Começar.
   - Na linha [#23](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L23) Indicar em Qual **Coluna** se encontra o **Nome** do Arquivo.
   - Na linha [#26](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L26) Indicar em Qual **Coluna** se encontra a **URL** de Download.
   - Na linha [#33](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L33) Indicar no Parâmetro **-OutFile** o **tipo** de arquivo / **Local de destino** do Download.
