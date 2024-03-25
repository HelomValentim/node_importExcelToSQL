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
