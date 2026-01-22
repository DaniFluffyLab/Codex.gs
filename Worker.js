/** 
 * Este projeto busca intermediar as comunicações entre o JS e o Google Sheets
 * para otimizar a leitura e escrita do Sheets como uma database. O objetivo é
 * utilizar ua sintaxe análoga ao do processamento dos Maps do JS, mas com 
 * comandos de iniciar e encerrar uma transação.
 */

class CodexWorker {

    constructor(sheetId, tableName, keyColumnName, options) {


        // FASE 0 - VALIDAÇÃO DE DEPENDÊNCIAS E DEFINIÇÃO DE VARS GLOBAIS

        if (typeof Sheets === 'undefined') {
            throw new Error(
                `[Codex] The "Google Sheets API" Advanced API is not enabled. ` +
                `To use Codex library, you need to activate it with identifier "Sheets".` +
                `Documentation: https://developers.google.com/apps-script/guides/services/advanced`
            );
        };


        /**
         * O identificador único (ID) da planilha de origem.
         * @type {string}
         * @private
         */
        this._sheetID = sheetId;


        /**
         * A instância da planilha (arquivo) do Google Sheets.
         * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
         * @private
         */
        this._sheet;                                                                        // Declara variável
        try { this._sheet = SpreadsheetApp.openById(this._sheetID) }                        // Carrega planilha
        catch (e) { throw Error(`${this._log} - Failed to load spreadsheet: ${e.stack}`) }; // Retorna algum erro


        /**
         * O nome da aba (página) dentro da planilha que será manipulada.
         * @type {string}
         * @private
         */
        this._tableName = tableName;


        /**
         * A aba específica dentro da planilha que servirá como tabela.
         * @type {GoogleAppsScript.Spreadsheet.Sheet}
         * @private
         */
        this._table;                                                                        // Declara variável
        try { this._table = this._sheet.getSheetByName(this._tableName) }                   // Carrega página
        catch (e) { throw Error(`${this._log} - Failed to load sheet/tab: ${e.stack}`) };   // Retorna outros erros


        /**
         * O identificador numérico único (GID) da aba dentro da planilha.
         * @type {number}
         * @private
         */
        this._tableID = this._table.getSheetId();


        /**
         * O nome do cabeçalho da coluna utilizada como chave primária.
         * @type {string}
         * @private
         */
        this._keyColumnName = keyColumnName;


        /**
         * Objeto de configurações da instância do Worker.
         * @type {{
         * mode: ("minimal"|"full"),
         * columns: string[]
         * }}
         * @private
         */
        this._options = {
            mode: "minimal",
            columns: [],
            ...options
        };


        /**
         * Conjunto contendo todas as chaves (IDs) identificadas na planilha.
         * Utilizado para verificar a existência de registros sem necessariamente carregar seus dados.
         * @type {Set<string>}
         * @private
         */
        this._allkeys = new Set();


        /**
         * Conjunto contendo as chaves (IDs) cujos dados completos já foram carregados em memória.
         * @type {Set<string>}
         * @private
         */
        this._loadedkeys = new Set();


        /**
         * Mapa contendo os registros carregados da planilha.
         * Associa cada identificador único ao seu respectivo objeto de dados de linha.
         * @type {Map<string, Object>}
         * @private
         */
        this._data = new Map();


        /**
         * Mapa que rastreia o status de sincronização das chaves alteradas na transação atual.
         * Associa a key do registro ao seu estado pendente para o próximo commit.
         * @type {Map<string, ("new"|"modified"|"deleted")>}
         * @private
         */
        this._keyStatus = new Map();


        /**
         * Prefixo identificador utilizado em mensagens de log e erros da instância.
         * @type {string}
         * @private
         */
        this._log = `[Codex] SheetID:"${this._sheetID}"; Table: "${this._tableName}"`;


        // Carrega dados de índices de colunas
        let columnIndexes = this._getColumnIndexes()



        // Obtém dados baseados no modo de operação

        switch (this._options.mode) {

            case "minimal":

                let keys_colIdx;                                                                    // Cria var para guardar índice
                try { keys_colIdx = columnIndexes.get(this._keyColumnName) }                        // Procura pelo nome
                catch (e) { throw Error(`${this._log} - Error locating keyColumn: ${e.stack}`) }    // Retorna outros erros
                if (keys_colIdx == undefined) throw Error(`${this._log} - keyColumn not found`)     // Se não achar coluna, lança erro

                try {
                    let lastRow = this._table.getLastRow()      // Obtém última linha
                    if (lastRow >= 2) {                         // Se planilha não está vazia

                        // Efetua request na API
                        let keys_rawValues = Sheets.Spreadsheets.Values.batchGetByDataFilter(
                            {
                                dataFilters: [{
                                    gridRange: {
                                        sheetId: this._tableID,
                                        startRowIndex: 1,
                                        startColumnIndex: keys_colIdx,
                                        endRowIndex: lastRow,
                                        endColumnIndex: keys_colIdx + 1
                                    }
                                }],
                                majorDimension: "COLUMNS",
                                valueRenderOption: "UNFORMATTED_VALUE",
                                dateTimeRenderOption: "FORMATTED_STRING"
                            },
                            this._sheetID,
                        ).valueRanges[0].valueRange.values[0]


                        for (let k of keys_rawValues) {                             // Para cada key
                            if (k == "" || k == null || k == undefined) continue    // Ignora keys vazias
                            this._allkeys.add(String(k).trim())                     // Adiciona Key ao Set mestre
                        }
                    }
                }
                catch (e) { throw Error(`${this._log} - Error to get keys: ${e.stack}`) }  // Retorna outros erros
                break;


            case "full":

                try { this._fetchNewData(true, columnIndexes) }                                 // Requisita dados
                catch (e) { throw Error(`${this._log} - Error to get values: ${e.stack}`) }     // Retorna erros
                break;
        }



    }

    /**
     * Mapeia os nomes das colunas da planilha para seus respectivos índices numéricos (0-based).
     * Realiza a leitura do cabeçalho (linha 1) e valida as colunas solicitadas nas configurações, 
     * garantindo que a coluna de chave primária esteja sempre presente no mapeamento.
     * * @returns {Map<string, number>} Um Map onde a chave é o nome da coluna (header) e o valor é o seu índice físico (0-based).
     * @throws {Error} Lança erro se a planilha não contiver colunas ou se uma coluna solicitada nas configurações não existir.
     * @private
     */
    _getColumnIndexes() {

        let columnIndexes;                                              // Cria var para índices das colunas
        let lastColumn = this._table.getLastColumn()                    // Obtém última coluna
        if (lastColumn == 0) throw Error(`No columns found`)            // Lança erro se sem colunas
        let columnArray = this._table.getRange(1, 1, 1, lastColumn)     // Seleciona cabeçalho
            .getValues()[0]                                             // Obtém dados
        columnIndexes = new Map(columnArray.map((v, i) => [v, i]))      // Insere dados dos índices no Map

        // Caso hajam parâmetros sobre quais colunas obter
        if (this._options.columns.length != 0) {
            let hasInvalid = this._options.columns.some(item => !columnIndexes.has(item));          // Verifica a validade das colunas
            if (hasInvalid) throw Error(`One or more requested columns do not exist.`);             // Lança erro se inválido 
            let filteredColumnIndexes = new Map()                                                   // Cria Map para colunas filtradas
            filteredColumnIndexes.set(this._keyColumnName, columnIndexes.get(this._keyColumnName))  // Garante coluna de key
            for (let columnName of this._options.columns) {                                         // Para cada coluna requisitada:
                filteredColumnIndexes.set(columnName, columnIndexes.get(columnName))                    // Armazena seu valor no Map de filtradas
            }
            columnIndexes = filteredColumnIndexes                                                   // Atualiza var de retorno
        }

        // Retorna array de índices
        return columnIndexes
    }

    /**
     * Requisita novos dados da planilha via API avançada e realiza o pivoteamento para o cache interno.
     * Suporta busca por colunas completas (modo COLUMNS) ou por linhas específicas (modo ROWS).
     * * @param {number[]|boolean} requestedRows - Array com índices de linhas (1-based) para busca seletiva, 
     * ou true para realizar a requisição de colunas completas.
     * @param {Map<string, number>} [columnIndexes] - Mapa contendo os nomes das colunas e seus respectivos 
     * índices. Caso omitido, utiliza o mapeamento padrão da instância.
     * @throws {Error} Se o parâmetro requestedRows não for um array nem o valor booleano true.
     * @private
     */
    _fetchNewData(requestedRows, columnIndexes) {

        // Valida parâmetros
        if (!Array.isArray(requestedRows) && requestedRows != true) throw Error(`Invalid requestedRows.`)

        let requestedData = new Map()                                               // Informações de dados a serem requeridos para a API
        let mode = Array.isArray(requestedRows) ? "ROWS" : "COLUMNS"                // Define modo de execução
        if (columnIndexes == undefined) columnIndexes = this._getColumnIndexes()    // Obtém índices de colunas, caso não recebido

        // Monta os objetos de requisição
        switch (mode) {

            case "COLUMNS":
                let lastRow = this._table.getLastRow();
                for (let [colName, colIndex] of columnIndexes) {
                    requestedData.set(colName, {
                        gridRange: {
                            sheetId: this._tableID,
                            startRowIndex: 1,
                            startColumnIndex: colIndex,
                            endRowIndex: lastRow,
                            endColumnIndex: colIndex + 1
                        }
                    })
                }
                break;

            case "ROWS":
                let firstCol = Math.min(...columnIndexes.values())
                let lastCol = Math.max(...columnIndexes.values())
                for (let rowIndex of requestedRows) {
                    requestedData.set(rowIndex, {
                        gridRange: {
                            sheetId: this._tableID,
                            startRowIndex: rowIndex - 1,
                            endRowIndex: rowIndex,
                            startColumnIndex: firstCol,
                            endColumnIndex: lastCol + 1
                        }
                    })
                };
                break;

        }

        // Executa a requisição
        let APIresponse;
        try {

            // Requisita via API 
            APIresponse = Sheets.Spreadsheets.Values.batchGetByDataFilter(
                {
                    dataFilters: [...requestedData.values()],
                    majorDimension: mode,
                    valueRenderOption: "UNFORMATTED_VALUE",
                    dateTimeRenderOption: "FORMATTED_STRING"
                },
                this._sheetID,
            )
        } catch (e) {

            // Checa se é um erro de request grande demais
            if (e.message.includes("Response Code: 413. Message: response too large.")) mode = `${mode}-SAFETY`
            else throw e    // Se não for, lança erro
        }


        // Executa o pivoteamento dos dados
        switch (mode) {

            case "COLUMNS":

                // Prepara dados da coluna em um map cuja key é o índice da coluna 
                let columnsData = new Map(
                    APIresponse.valueRanges.map(range => {
                        let key = range.dataFilters[0].gridRange.startColumnIndex
                        let value = (range.valueRange.values && range.valueRange.values[0]) ? range.valueRange.values[0] : []
                        return [key, value]
                    })
                )

                let keyColumn = columnsData.get(columnIndexes.get(this._keyColumnName))  // Obtém a coluna de key

                // Para cada linha recebida
                keyColumn.forEach((key, rowInd) => {

                    if (key === undefined || key === null || String(key).trim() === "") return; // Ignora linhas sem key
                    let obj = {}                                                                // Cria um objeto de saída

                    // Para cada coluna solicitada, cria a propriedade e armazena o valor no objeto
                    columnIndexes.forEach((colIndex, colName) => obj[colName] = columnsData.get(colIndex)[rowInd] ?? null)

                    // Armazena resutados
                    this._data.set(String(key).trim(), obj)
                    this._loadedkeys.add(String(key).trim())
                    this._allkeys.add(String(key).trim())
                })
                break;

            case "COLUMNS-SAFETY":

                // Avisa o usuário sobre o uso do modo de segurança
                console.warn(`${this._log} - Too much data, activating safety mode. Consider requesting fewer columns or using minimal mode with .search() to increase speed.`)

                /**
                 * Obtém os valores de um intervalo da planilha utilizando a estrutura GridRange (0-indexed).
                 * Esta função atua como um intermediário da API Avançada do Sheets, com fallback no método nativo
                 * getRange do Apps Script. Retorna um array unidimensional (achatado).
                 * * @param {GoogleAppsScript.Spreadsheet.Sheet} table - A instância da aba da planilha (Sheet).
                 * * @param {Object} gridRange - Objeto contendo as coordenadas do intervalo.
                 * @param {number} gridRange.startRowIndex - Índice inicial da linha (0-indexed, inclusive).
                 * @param {number} gridRange.endRowIndex - Índice final da linha (0-indexed, exclusive).
                 * @param {number} gridRange.startColumnIndex - Índice inicial da coluna (0-indexed, inclusive).
                 * @param {number} gridRange.endColumnIndex - Índice final da coluna (0-indexed, exclusive).
                 * * @returns {any[]} Um array unidimensional contendo todos os valores do intervalo solicitado.
                 * @private
                 */
                function getValuesByGridRange(table, gridRange) {
                    try {

                        // Tenta requerer a API avançada
                        Utilities.sleep(500)    // Força aguardar para evitar erro 429
                        return Sheets.Spreadsheets.Values.batchGetByDataFilter(
                            {
                                dataFilters: [{ gridRange: gridRange }],
                                majorDimension: "ROWS",
                                valueRenderOption: "UNFORMATTED_VALUE",
                                dateTimeRenderOption: "FORMATTED_STRING"
                            },
                            table.getSheetId(),
                        ).valueRanges[0].valueRange.values
                    }
                    catch (e) {

                        // Fallback via SpreadsheetApp
                        return table.getRange(                                  // Obtém range
                            gridRange.startRowIndex + 1,                            // Converte para 1-indexed
                            gridRange.startColumnIndex + 1,                         // Converte para 1-indexed
                            gridRange.endRowIndex - gridRange.startRowIndex,        // Calcula total de linhas
                            gridRange.endColumnIndex - gridRange.startColumnIndex,  // Calcula total de colunas
                        ).getValues()                                           // Obtém e achata array
                    }
                }

                /**
                 * Agrupa colunas adjacentes (chunks) para reduzir o número de requisições à API.
                 *
                 * @param {Map<string, {gridRange: Object}>} gridRangesMap Mapa contendo os intervalos individuais de cada coluna.
                 * @returns {Map<string[], {gridRange: Object}>} Mapa com os intervalos fundidos, onde a chave é a lista de nomes das colunas agrupadas.
                 * @private
                 */
                function mergeGridRanges(gridRangesMap) {

                    const CHUNK_SIZE = 5    // Variável de tamanho máximo de chunk

                    // Junta todas as gridRanges na array, com nome das colunas
                    let allGridRanges = []
                    gridRangesMap.forEach((data, colName) => {

                        // Copia o objeto
                        let safeData = { ...data };
                        safeData.gridRange = { ...data.gridRange };

                        // Insere o dado de colName e armazena
                        safeData.gridRange.colName = colName;
                        allGridRanges.push(safeData);

                    })

                    // Ordena os gridRanges
                    let sortedGridRanges = allGridRanges.sort((a, b) => {
                        return a.gridRange.startColumnIndex - b.gridRange.startColumnIndex
                    })

                    // Mesclando gridRanges
                    let mergedGridRanges = new Map()
                    let mergingGridRange = undefined
                    let mergingColNames = undefined
                    while (sortedGridRanges.length != 0) {

                        // Obtém linha
                        let workingGridRange = sortedGridRanges.shift().gridRange

                        // Caso não haja um GridRange atualmente, criar e seguir para próx. loop
                        if (mergingGridRange == undefined) {
                            mergingGridRange = {
                                sheetId: workingGridRange.sheetId,
                                startRowIndex: workingGridRange.startRowIndex,
                                endRowIndex: workingGridRange.endRowIndex,
                                startColumnIndex: workingGridRange.startColumnIndex,
                                endColumnIndex: workingGridRange.endColumnIndex,
                            }
                            mergingColNames = [workingGridRange.colName]
                            continue;
                        }

                        // Mescla vizinhos
                        let validMerge = (mergingGridRange.endColumnIndex == workingGridRange.startColumnIndex)
                        if (validMerge) {
                            mergingGridRange.endColumnIndex = workingGridRange.endColumnIndex
                            mergingColNames.push(workingGridRange.colName);
                        }

                        // Caso tenha mesclado e atingido o tamanho ou caso não tenha sido mesclado
                        if ((validMerge && mergingColNames.length == CHUNK_SIZE) || (!validMerge)) {
                            mergedGridRanges.set(mergingColNames, { gridRange: mergingGridRange })
                            mergingGridRange = undefined
                            mergingColNames = undefined
                        }

                        // Caso não mesclado, guarda para nova iteração
                        if (!validMerge) sortedGridRanges.unshift({ gridRange: workingGridRange })
                    }

                    // Caso tenha ficado algum objeto pra trás, armazena ele
                    if (mergingGridRange != undefined) mergedGridRanges.set(mergingColNames, { gridRange: mergingGridRange })

                    // Retorna objeto mesclado
                    return mergedGridRanges
                }




                // Obtém dados de key
                let keyData = getValuesByGridRange(this._table, requestedData.get(this._keyColumnName).gridRange).map(([v]) => String(v).trim())

                // Registra IDs nos metadados de keys
                for (let key of keyData) {
                    if (key === undefined || key === null || key === "") continue;
                    this._loadedkeys.add(key)
                    this._allkeys.add(key)
                    this._data.set(key, {})
                    this._data.get(key)[this._keyColumnName] = key
                }


                requestedData.delete(this._keyColumnName)                   // Remove a requisição de coluna de key
                let mergedRequestedData = mergeGridRanges(requestedData)    // Mescla as requisições

                // Para cada conjunto de requisições
                mergedRequestedData.forEach(({ gridRange }, colNames) => {

                    let workingArray = getValuesByGridRange(this._table, gridRange)    // Obtém dados

                    // Para cada linha
                    workingArray.forEach((row, rowInd) => row.forEach((value, colInd) => {

                        let currentKey = keyData[rowInd]                            // Obtém key atual
                        if (!this._allkeys.has(currentKey)) return;                 // Se ID inválido, ignorar
                        this._data.get(currentKey)[colNames[colInd]] = value;       // Armazena valor na memória
                    }))
                })
                break;


            /*          
                        NAO TESTADO AINDA.
            
                        case "ROWS":
            
                            let rowsData = APIresponse.valueRanges.map(range => (range.valueRange.values && range.valueRange.values[0]) ? range.valueRange.values[0] : [])   // Prepara dados para leitura
                            let colOffset = Math.min(...columnIndexes.values())                                                             // Obtém o offset de colunas
                            let idIndex = columnIndexes.get(this._keyColumnName) - colOffset                                                // Obtém o índice do ID
            
                            // Para cada linha recebida
                            rowsData.forEach(row => {
            
                                if (row[idIndex] === undefined || row[idIndex] === null || String(row[idIndex]).trim() === "") return;  // Ignora linhas sem ID
                                let obj = {}                                                                                            // Cria um objeto de saída
            
                                // Para cada coluna solicitada, cria a propriedade e armazena o valor no objeto
                                columnIndexes.forEach((colInd, colName) => obj[String(colName).trim()] = row[colInd - colOffset] ?? null)
            
                                // Armazena resutados
                                this._data.set(String(row[idIndex]).trim(), obj)
                                this._loadedkeys.add(String(row[idIndex]).trim())
                                this._allkeys.add(String(row[idIndex]).trim())
                            })
                            break; */
        }
    }
}



