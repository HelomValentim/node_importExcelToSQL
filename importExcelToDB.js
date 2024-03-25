/*
    Este módulo é responsável por importar os dados de um arquivo Excel para um banco de dados SQL Server.
    Ele recebe como parâmetro um array de caminhos de arquivos Excel e um objeto de configuração do SQL Server.
    Exemplo que consumo do modulo importExcelToSql:
    const importExcelToSql = require('../utils/importExcelToSql.js');
    importExcelToSql(configSQL, [pathFile1, pathFile2, ...]).then((result) => {
        if (result.success) {
            console.log('Os dados foram importados com sucesso.');
        } else {
            console.error(`Ocorreu um erro ao importar os dados: ${result.error}`);
        }
    });
    O objeto de configuração do SQL Server deve conter as seguintes propriedades:
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
    O array de caminhos de arquivos Excel deve conter os caminhos dos arquivos que serão importados.
    O módulo retorna um objeto com as seguintes propriedades:
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
    A propriedade sucess indica se o processo de importação foi bem sucedido.
    A propriedade error contém a mensagem de erro, caso ocorra algum erro.
    As demais propriedades contém um array de strings com o fluxo dos processos ou a mensagem de erro correspondente.

*/


const xlsx = require('xlsx');
const sql = require('mssql');
const importExcelReferences = require("/src/importExcelReferences.js");
const sqlConfig = require("/src/sqlConfig.js");

const logs = {
    success: false,
    error: '',
    importProcess: {
        flow: [],
        error: ''
    },
    readExcel: {
        flow: [],
        error: ''
    },
    createTable: {
        flow: [],
        error: ''
    },
    insertDataTable: {
        flow: [],
        error: ''
    },
};


async function readExcel(pathFile, sheets) {
    // Ler o arquivo Excel
    let workbook = xlsx.readFile(pathFile, { cellNF: true });

    if (!workbook) {
        logs.readExcel.error = `Não foi possível ler o arquivo ${pathFile}, verifique se o caminho está correto e se o arquivo é um arquivo válido.`;
        return;
    }

    if (workbook.SheetNames.length > 1) {
        logs.readExcel.flow.push(`Foram encontradas mais de uma planilha no arquivo ${pathFile}. Apenas as planilhas ${sheets} serão importadas.`);
        workbook.SheetNames = workbook.SheetNames.filter(sheetName => sheets.includes(sheetName));
    }
    // Objeto para armazenar os fileData de todas as planilhas
    const fileData = {};
    let configSheet;

    // Iterar sobre cada planilha
    workbook.SheetNames.forEach(sheetName => {
        logs.readExcel.flow.push(`Iniciando a leitura da planilha ${sheetName} do arquivo ${pathFile}.`);
        // Obter os fileData da planilha atual
        const worksheet = workbook.Sheets[sheetName];
        configSheet = findDataSheet(sheetName, pathFile);

        //Verifica se existe a referencia para a planilha, se não existir tenta buscar pelo index
        if (!configSheet) {
            logs.readExcel.flow.push(`Não foi encontrada a planilha "${sheetName}". Buscando referencia pelo index...`);
            configSheet = findDataSheetIndex(pathFile);

            if (!configSheet) {
                logs.readExcel.error = `Não foi encontrada nenhuma referência para o arquivo com path "${pathFile}".`;
                return;
            }
        }

        // Converter os fileData da planilha para JSON
        const sheetData = xlsx.utils.sheet_to_json(worksheet, {
            raw: false,
            header: 1,
            defval: null,
        });

        let xlsHeader = sheetData[parseInt(configSheet['headerRow'])];
        let tempObj;
        let nameColumnsFind = true;
        let headersName = [];

        for (let i = 0; i < configSheet.headers.length; i++) headersName.push(configSheet['headers'][i]['name']);

        //Verificar se todas as colunas da referencia existe no excel com o nome correcto
        for (let i = 0; i < configSheet.headers.length; i++) {
            if (xlsHeader.indexOf(configSheet.headers[i].excelColumnName) === -1) {
                nameColumnsFind = false;
                logs.readExcel.flow.push(`A coluna "${configSheet.headers[i].excelColumnName}" não foi encontrada na planilha "${sheetName}" verificando a quantidade de colunas...`);
                if (xlsHeader.length === headersName.length) {
                    logs.readExcel.flow.push(`Existe a mesma quantidade de colunas na planilha "${sheetName}" que na referencia, alterando o nome das colunas pelo index...`);
                    sheetData[parseInt(configSheet['headerRow'])] = headersName
                    break;
                } else {
                    logs.readExcel.error = `A quantidade de colunas na planilha "${sheetName}" não é igual a quantidade de colunas na referencia.`;
                    return false;
                }
            }
        }

        if (nameColumnsFind) {
            logs.readExcel.flow.push(`Todas as colunas foram encontradas na planilha "${sheetName}" alterando nome das colunas pela referencia.`);
            sheetData[parseInt(configSheet['headerRow'])] = sheetData[parseInt(configSheet['headerRow'])].map(value => value = configSheet.headers.find(header => header.excelColumnName === value).name);
        }

        sheetData.forEach((row) => {
            tempObj = {};
            tempObj.isFilled = true;
            configSheet.headers.forEach((header) => {
                tempObj[header.name] = row[xlsHeader.indexOf(header.excelColumnName)];

                if (tempObj[header.name] == null) {
                    tempObj.isFilled = false;
                }
            });
        });
        fileData[sheetName] = sheetData;
        logs.readExcel.flow.push(`Os dados da planilha "${sheetName}" foram lidos com sucesso.`);
    });
    // Retornar o objeto de fileData completo
    logs.readExcel.flow.push(`Todos os dados do ficheiro ${pathFile} foram lidos com sucesso.`);
    return fileData;
}

async function createTable(pool, configSheet) {
    const transaction = new sql.Transaction(pool);
    try {
        const table = configSheet['name'];
        logs.createTable.flow.push(`Iniciando a criação da tabela "${table}".`);

        // Criar uma transação
        await transaction.begin();

        const columns = configSheet.headers.map(coluna => `[${coluna.name}] ${coluna.columnType}`).join(',');
        logs.createTable.flow.push(`Colunas definidas para a tabela: ${columns}.`);

        // Criar a tabela
        await transaction.request().query(`DROP TABLE IF EXISTS ${table}; CREATE TABLE ${table} (${columns})`);

        // Commit da transação
        await transaction.commit();
        logs.createTable.flow.push(`Tabela "${table}" criada com sucesso.`);
        await sql.close();
    } catch (error) {
        // Rollback da transação em caso de erro
        logs.createTable.error = `Ocorreu um erro ao criar a tabela "${table}": ${error}`;
        await transaction.rollback();
        throw error;
    }
}

async function insertDataTable(sqlConfig, dataTable, configSheet) {
    try {
        logs.insertDataTable.flow.push(`Iniciando a inserção dos dados na tabela ${configSheet['name']}.`);
        // Cria uma nova instância de Pool usando a configuração
        const pool = new sql.ConnectionPool(sqlConfig);
        logs.insertDataTable.flow.push(`Abrindo a conexão com o banco de dados.`);

        // Abre a conexão
        await pool.connect()
            .catch(err => {
                // console.error('Erro ao abrir a conexão com o banco de fileData:', err);
                logs.insertDataTable.error = `Erro ao abrir a conexão com o banco de dados: ${err}`;
            });

        if (!configSheet) { return; }
        const tableName = configSheet['name'];
        const deleteTable = configSheet['deleteTable']
        const headers = dataTable[parseInt(configSheet['headerRow'])];

        for (i = 0; i < parseInt(configSheet['headerRow']) + 1; i++) dataTable.shift(); // Remove a primeira linha do array, que contém os nomes das colunas

        logs.insertDataTable.flow.push(`Verificando se a tabela "${tableName}" existe na base de dados.`);
        const tableExists = await pool.request().query(`SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'${tableName}'`);
        if (tableExists.recordset.length === 0) {
            logs.insertDataTable.flow.push(`A tabela "${tableName}" não existe. Criando...`);
            await createTable(pool, configSheet);
        } else {
            if (deleteTable) {
                logs.insertDataTable.flow.push(`A tabela "${tableName}" existe e o parametro de apagar a tabela está ativada. Apagando dados da tabela...`);
                await pool.request().query(`TRUNCATE TABLE ${tableName}`);
            } else {
                logs.insertDataTable.flow.push(`A tabela "${tableName}" existe e o parametro de apagar a tabela está desativada.`);
            }
        }

        const table = new sql.Table(tableName); // or temporary table, e.g. #temptable
        table.create = true;
        // Adicionando as colunas à
        for (const column of headers) {
            table.columns.add(column, sql.NVarChar(sql.MAX), { nullable: true });
        }

        // Adicionar os registros à table
        for (const data of dataTable) {
            table.rows.add(...Object.values(data));
        }
        logs.insertDataTable.flow.push(`Fazendo o bulk dos dados.`);

        await pool.request().bulk(table);
        logs.insertDataTable.flow.push(`Os dados foram inseridos com sucesso na tabela ${tableName}.`);
        await pool.close();

    } catch (error) {
        logs.insertDataTable.error = `Ocorreu um erro ao inserir os dados na tabela ${tableName}: ${error}`;
        throw error;
    }
}

function findDataSheet(sheetName, pathFile) {
    const allReferences = Object.values(importExcelReferences);
    for (const reference of allReferences) {
        const fileData = reference.find(dado => dado.filePath === pathFile && dado.worksheetName === sheetName);
        if (fileData) {
            return fileData;
        }
    }
    return null; // Retorna null se não encontrar correspondência
}

function findDataSheetIndex(pathFile) {
    const allReferences = Object.values(importExcelReferences);
    let sheetsReferences = [];
    for (let i in allReferences) {
        const fileData = allReferences[i].find(dado => dado.filePath === pathFile);
        if (fileData) {
            sheetsReferences.push(allReferences[i]);
        }
    }

    switch (sheetsReferences.length) {
        case 0:
            logs.importProcess.error = `Não foi encontrada nenhuma referência para o arquivo com path "${pathFile}".`;
            return null;
        case 1:
            return sheetsReferences[0][0];
        default:
            logs.importProcess.error = `Foram encontradas mais de uma referência para o arquivo com path "${pathFile}".`;
            return null;
    }
}

async function findSheetsReference(pathFile) {
    const allReferences = Object.values(importExcelReferences);
    const fileData = [];
    for (const reference of allReferences) {
        const dado = reference.find(dado => dado.filePath === pathFile);
        if (dado) {
            fileData.push(dado.worksheetName);
        }
    }
    if (fileData) {
        return fileData
    } else {
        return null;
    }
}

async function importExcelToSql(sqlConfig, pathFiles = []) {
    try {
        logs.importProcess.flow.push(`Iniciando o processo de importação.`);
        logs.importProcess.flow.push(`Foram informados ${pathFiles.length} arquivos para importação.`);
        for (const pathFile of pathFiles) {
            logs.importProcess.flow.push(`Iniciando o processo de importação do arquivo ${pathFile}.`);

            // Ler o arquivo Excel
            logs.importProcess.flow.push(`Buscando as sheets para o ficheiro path ${pathFile} no ficheiro de referência.`);
            const sheets = await findSheetsReference(pathFile);
            logs.importProcess.flow.push(`As sheets encontradas para o ficheiro são: ${sheets}.`);
            const fileData = await readExcel(pathFile, sheets);
            if (!fileData) {
                logs.importProcess.error = `Ocorreu um erro ao ler o arquivo ${pathFile}.`;
                return logs;
            }

            // Iterar sobre cada planilha
            for (const sheetName in fileData) {
                let configSheet = findDataSheet(sheetName, pathFile);

                //Verifica se existe a referencia para a planilha, se não existir tenta buscar pelo index
                if (!configSheet) {
                    logs.importProcess.flow.push(`Não foi encontrada a planilha "${sheetName}". Buscando referencia pelo index...`);
                    configSheet = findDataSheetIndex(pathFile);

                    if (!configSheet) {
                        logs.importProcess.error = `Não foi encontrada nenhuma referência para o arquivo com path "${pathFile}".`;
                        return;
                    }
                }
                logs.importProcess.flow.push(`Verificando se a planilha "${sheetName}" existe na referencia, ou se existe apenas uma planilha no arquivo.`);
                if (sheets.includes(sheetName) || sheets.length === 1) {
                    if (sheets.includes(sheetName)) {
                        logs.importProcess.flow.push(`A planilha "${sheetName}" existe na referencia.`);
                    } else {
                        logs.importProcess.flow.push(`A planilha "${sheetName}" não existe na referencia, mas o ficheiro tem apenas uma planilha no arquivo.`);
                    }
                    const dataSheet = fileData[sheetName];
                    if (dataSheet.length > 0) {
                        // Inserir os fileData na table
                        await insertDataTable(sqlConfig, dataSheet, configSheet);
                    }
                }
            }
            // return true;

        }
    } catch (error) {
        logs.error = `Ocorreu um erro ao importar os dados: ${error}`;
        return logs;
    }
    if (logs.readExcel.error || logs.createTable.error || logs.insertDataTable.error) {
        throw new Error(`${logs.readExcel.error ? logs.readExcel.error : logs.createTable.error ? logs.createTable.error : logs.insertDataTable.error}`);
    }
    logs.success = true;
    logs.importProcess.flow.push(`Todos os dados foram inseridos com sucesso.`);
    return logs;
}

// Exemplo de uso
const pathFile = 'C:/Users/Usuario/Desktop/Excel_clients.xlsx';
importExcelToSql(sqlConfig, [pathFile]).then(resp =>
    console.log(resp)
);

// module.exports = {
//     importExcelToSql
// }
