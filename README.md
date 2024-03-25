# Este módulo é responsável por importar os dados de um arquivo Excel para um banco de dados SQL Server.
    Ele recebe como parâmetro um array de caminhos de arquivos Excel e um objeto de configuração do SQL Server.
    Exemplo que consumo do modulo importExcelToSql:

```js
    const importExcelToSql = require('../utils/importExcelToSql.js');
    importExcelToSql(configSQL, [pathFile1, pathFile2, ...]).then((result) => {
        if (result.success) {
            console.log('Os dados foram importados com sucesso.');
        } else {
            console.error(`Ocorreu um erro ao importar os dados: ${result.error}`);
        }
    });
```
    O objeto de configuração do SQL Server deve conter as seguintes propriedades:
```js
    {
        user
        password
        server
        database
        options{
            trustServerCertificate: true,
            connectionTimeout: 60000,
        }
    }
```
    O array de caminhos de arquivos Excel deve conter os caminhos dos arquivos que serão importados.
    O módulo retorna um objeto com as seguintes propriedades:
```js
    {
        sucess: boolean,
        error: string,
        importProcess: {
            flow: string[],
            error: string
        },
        readExcel: {
            flow: string[],
            error: string
        },
        createTable: {
            flow: string[],
            error: string
        },
        insertIntoDB: {
            flow: string[],
            error: string
        }
    }
```
    A propriedade **sucess** indica se o processo de importação foi bem sucedido.
    A propriedade error contém a mensagem de erro, caso ocorra algum erro.
    As demais propriedades contém um array de strings com o fluxo dos processos ou a mensagem de erro correspondente.

    
    Ficheiro utilizado como referencia para importar dados de ficheiros excel utilizando a função "\api\utils\importExcelToSql.js" a estrutura a ser utilizada é separada pelas
    sheets que se desejam importar. Cada sheet tem um array de objetos que contem as informações necessárias para importar os dados para a base de dados.
    A estrutura é a seguinte:
```js
        {
            name: "nome da tabela na base de dados",
            worksheetIndex: 0, //index da sheet no ficheiro excel que será utilizado caso o nome da sheet não seja encontrado.
            worksheetName: "nome da sheet",
            deleteTable: true, //se a tabela deve ser apagada antes de inserir os dados. Se false, os dados são adicionados à tabela existente.
            filePath: 'caminho do ficheiro', //caminho do ficheiro excel em geral ficam na pasta "//WSSPADAF001/Drops/Uploads/..pasta do projeto/..nome do ficheiro.xlsx"
            headerRow: 0, //linha onde se encontram os nomes das colunas (começa do zero)
            headers:
                [
                    {
                        name: "nome da coluna na base de dados",
                        columnType: "tipo de dados da coluna (SQL Type)",
                        excelColumnName: "nome da coluna no excel",
                        excelColumnIndex: 0, //index da coluna no excel começa do zero. É utilizado caso o nome da coluna não seja encontrado.
                    },
                    ... //Manter o mesmo padrão para todas as colunas. 
                ]
        }
```
    Para importar mais de uma sheet do mesmo ficheiro, basta criar um novo objeto com as informações necessárias mantendo o mesmo filePath.
    Somente serão importadas as colunas que estiverem presentes no array de headers. Caso a coluna não seja encontrada, será ignorada.
