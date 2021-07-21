# Curl-With-XLSX - PowerShell 

Ferramenta Desenvolvida em PowerShell Para o Download de Arquivos listados em uma planilha.<br>

## Como Utilizar 

 - Planilha deve conter **Nome** do Arquivo e **URL** na mesma Linha em colunas **Diferêntes** , mas na mesma **ABA**.
 - Configurações a serem Feitas no Algoritmo:
   - Na linha [#14](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L14) Referenciar A **Planilha** de Base.
   - Na linha [#17](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L17) Indicar Qual **Aba** se encontra a Base.
   - Na linha [#20](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L20) Indicar em Qual **linha** Começar.
   - Na linha [#23](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L23) Indicar em Qual **Coluna** se encontra o **Nome** do Arquivo.
   - Na linha [#26](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L26) Indicar em Qual **Coluna** se encontra a **URL** de Download.
   - Na linha [#33](https://github.com/williamanjo/Curl-With-XLSX/blob/8422fb934479a8d7d850871f5224768930af7575/Main.ps1#L33) Indicar no Parâmetro **-OutFile** o **tipo** de arquivo / **Local de destino** do Download.

# Curl-With-XLSX - C#

Ferramenta Desenvolvida em C# Para a maior Facilidade de Utilização.<br>

## Formulário Como utilizar
![image](https://user-images.githubusercontent.com/69880957/126486521-5730ee35-c96a-4ec6-95d2-531c2fef4e89.png)
<br>
 - Selecionar A planilha.
 - Selecionar A Sheet.
 - Se tem Cabeçalho ou não na planilha.
 - Identificação é o numero da coluna com o nome do arquivo.
 - URL é o numero da coluna que está a URL do Download.
 - Selecionar Local de Descarga dos arquivos.
 - Tipo de extenção Do arquivo (Padrão: .pdf ).
 - Se tem Versão do arquivo Eg.(arquivo-"1".pdf) ,Se não colocar nada o arquivo vai vir sem o traço no final (arquivo.pdf).

# Exemplo :
## Planilha

![image](https://user-images.githubusercontent.com/69880957/126492680-3ff380a3-b7d4-437c-a83d-af5840288c9a.png)

## Configuração

![image](https://user-images.githubusercontent.com/69880957/126492753-71b69296-5ad5-4416-9261-9dcdec635fc3.png)

## Resultado

![image](https://user-images.githubusercontent.com/69880957/126492993-3fb13bad-e31e-425b-b0fc-8708058dee1f.png)



