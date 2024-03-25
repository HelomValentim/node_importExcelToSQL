/*
    Ficheiro utilizado como referencia para importar dados de ficheiros excel utilizando a função "\api\utils\importExcelToSql.js" a estrutura a ser utilizada é separada pelas
    sheets que se desejam importar. Cada sheet tem um array de objetos que contem as informações necessárias para importar os dados para a base de dados.
    A estrutura é a seguinte:
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

    Para importar mais de uma sheet do mesmo ficheiro, basta criar um novo objeto com as informações necessárias mantendo o mesmo filePath.
    Somente serão importadas as colunas que estiverem presentes no array de headers. Caso a coluna não seja encontrada, será ignorada.
*/

const customers_list = [
    {
        name: "customers_list",
        worksheetIndex: 0,
        worksheetName: "Sheet1",
        deleteTable: true,
        filePath: 'C:/Users/Usuario/Desktop/Excel_clients.xlsx',
        headerRow: 0,
        headers:
            [
                {
                    name: "name_client",
                    columnType: "NVARCHAR(150)",
                    excelColumnName: "Name",
                    excelColumnIndex: 0,
                },
                {
                    name: "phone_number",
                    columnType: "NVARCHAR(15)",
                    excelColumnName: "Contact",
                    excelColumnIndex: 1,
                },
                {
                    name: "address_client",
                    columnType: "NVARCHAR(250)",
                    excelColumnName: "Address",
                    excelColumnIndex: 2,
                },
                {
                    name: "vip_client",
                    columnType: "bit", //OR boolean
                    excelColumnName: "Important Client",
                    excelColumnIndex: 3,
                }
            ]
    }
]

const customers_list_blocked = [
    {
        name: "customers_list_blocked",
        worksheetIndex: 1,
        worksheetName: "blocked customers",
        deleteTable: true,
        filePath: '//WSSPADAF001/Drops/Uploads/ControloMaster/dadost.xlsb',
        headerRow: 0,
        headers:
            [
                {
                    name: "name_client",
                    columnType: "NVARCHAR(150)",
                    excelColumnName: "Name",
                    excelColumnIndex: 0,
                },
                {
                    name: "phone_number",
                    columnType: "NVARCHAR(15)",
                    excelColumnName: "Contact",
                    excelColumnIndex: 1,
                },
                {
                    name: "address_client",
                    columnType: "NVARCHAR(250)",
                    excelColumnName: "Address",
                    excelColumnIndex: 2,
                },
                {
                    name: "vip_client",
                    columnType: "bit", //OR boolean
                    excelColumnName: "Important Client",
                    excelColumnIndex: 3,
                },
                {
                    name: "blocking_date",
                    columnType: "datetime",
                    excelColumnName: "Blocking Date",
                    excelColumnIndex: 4,
                }

            ]
    }
]

module.exports = {
    customers_list: customers_list,
    customers_list_blocked: customers_list_blocked
}