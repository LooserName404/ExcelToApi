# ExcelToApi
Um projeto que exemplifica a utilização do ASP.NET com planilhas no formato xlsx.

## Utilizando
Para utilizar o projeto, altere a variável `Directory` no arquivo `Data/Spreadsheet.cs` para utilizar o diretório correto em seu computador.
Após isso, basta inicializar o projeto pelo Visual Studio ou utilizar o comando `dotnet run`.

## Informações
A API usa do ASP.NET MVC API para servir os dados em formato JSON.
A biblioteca para realizar as operações nas planilhas é a `DocumentFormat.OpenXml`.
