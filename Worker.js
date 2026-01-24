/* 
 * Este projeto busca intermediar as comunicações entre o JS e o Google Sheets
 * para otimizar a leitura e escrita do Sheets como uma database. O objetivo é
 * utilizar ua sintaxe análoga ao do processamento dos Maps do JS, mas com 
 * comandos de iniciar e encerrar uma transação.
 */

class Codex {


    /**
     * Creates a new Codex instance to manage a Google Sheets tab as a persistent key-value store.
     * This class adapts the Google Sheets API to function similarly to a JavaScript `Map`,
     *
     * @param {string} sheetId - The unique identifier of the Google Spreadsheet (extractable from the URL).
     * @param {string} tableName - The exact name of the tab (Sheet) to be managed.
     * @param {string} keyColumnName - The header name of the column that serves as the unique Primary Key (ID).
     * @param {Object} options - Configuration options for initialization.
     * @param {("full"|"minimal")} options.mode - Defines whether the spreadsheet should be pre-loaded into memory or requested on demand.
     * @param {string[]} [options.columns] - Array of specific column names to be eager-loaded during instantiation.
     * @param {boolean} [options.enableTypeInference] - Allow Codex to infer rich types, like Arrays and Objects, from data.
     *
     * @throws {Error} If the "Google Sheets API" Advanced Service is not enabled with the identifier "Sheets".
     * @throws {Error} If the spreadsheet or the specified tab cannot be accessed.
     *
     * @example
     * // Initialize connection to "Users" tab using "UserID" as key
     * 
     * // Not specifying "columns" will request all columns.
     * const dbAllColumns = new Codex("1BxiM...", "Users", "UserID", {
     *     mode: "minimal"
     * });
     * 
     * // mode: "minimal" retrieves line data on demand.
     * const dbOnDemand = new Codex("1BxiM...", "Users", "UserID", {
     *     mode: "minimal",
     *     columns: ["Price", "Stock"]
     * });
     * 
     * // mode: "full" retrieves all data from the rows at initialization.
     * const dbAllData = new Codex("1BxiM...", "Users", "UserID", {
     *     mode: "full",
     *     columns: ["Stock", "Description"]
     * });
     */
    constructor(sheetId, tableName, keyColumnName, options) {


        // FASE 0 - VALIDAÇÃO DE DEPENDÊNCIAS E DEFINIÇÃO DE VARS GLOBAIS


        /**
         * Prefixo identificador utilizado em mensagens de log e erros da instância.
         * @type {string}
         * @private
         */
        this._log = `[ CODEX | SheetID:"${sheetId}" | Table: "${tableName}" ]\n`;


        if (typeof Sheets === 'undefined') {
            throw new Error(
                `[CODEX] The "Google Sheets API" Advanced API is not enabled. ` +
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
        if (typeof sheetId !== 'string') throw Error(`${this._log} sheetID is not a string.`)


        /**
         * A instância da planilha (arquivo) do Google Sheets.
         * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
         * @private
         */
        this._sheet;                                                                            // Declara variável
        try { this._sheet = SpreadsheetApp.openById(this._sheetID) }                            // Carrega planilha
        catch (e) { throw Error(`${this._log} Failed to load spreadsheet. \n\n${e.stack}`) };   // Retorna algum erro


        /**
         * O nome da aba (página) dentro da planilha que será manipulada.
         * @type {string}
         * @private
         */
        this._tableName = tableName;
        if (typeof tableName !== 'string') throw Error(`${this._log} tableName is not a string.`)


        /**
         * A aba específica dentro da planilha que servirá como tabela.
         * @type {GoogleAppsScript.Spreadsheet.Sheet}
         * @private
         */
        this._table;                                                                        // Declara variável
        try { this._table = this._sheet.getSheetByName(this._tableName) }                   // Carrega página
        catch (e) { throw Error(`${this._log} Failed to load sheet/tab. \n\n${e.stack}`) }  // Retorna outros erros


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
        if (typeof keyColumnName !== 'string') throw Error(`${this._log} keyColumnName is not a string.`)



        /**
         * Objeto de configurações da instância do Worker.
         * @type {{
         * mode: ("minimal"|"full"),
         * columns: string[]
         * enableTypeInference: boolean
         * }}
         * @private
         */
        this._options = {
            mode: "minimal",
            columns: [],
            enableTypeInference: true,
            ...options
        };
        if (this._options.mode !== 'minimal' && this._options.mode !== 'full') throw Error(`${this._log} Invalid mode: ${this._options.mode}`)
        for (let c of this._options.columns) if (typeof c !== 'string') throw Error(`${this._log} Column "${c}" is not a string.`)



        /**
         * Mapa contendo os registros carregados da planilha.
         * Associa cada identificador único ao seu respectivo objeto de dados de linha.
         * @type {Map<string, Object>}
         * @private
         */
        this._data = new Map();

        /**
         * Mapa de cache contendo os proxies dos registros obtidos via funções de requisição.
         * Associa cada objeto original ao seu respectivo proxy.
         * @type {WeakMap<Object, Object>}
         * @private
         */
        this._proxies = new WeakMap();

        /**
         * Mapa que rastreia o status de sincronização das chaves alteradas na transação atual.
         * Associa a key do registro ao seu estado pendente para o próximo commit.
         * @type {Map<string, ("unmodified"|"new"|"modified"|"deleted")>}
         * @private
         */
        this._keys = new Map();


        /**
         * Marca se a planilha deve ser toda zerada.
         * @type {boolean}
         * @private
         */
        this._wipeOnCommit = false;


        // Carrega dados de índices de colunas
        let columnIndexes = this._getColumnIndexes()

        // Obtém dados baseados no modo de operação
        switch (this._options.mode) {

            case "minimal":

                try {
                    let keys = this._getRowIndexesByKey(true, columnIndexes)            // Obtém keys
                    this._keys = new Map([...keys.keys()].map(k => [k, "unmodified"]))  // Adiciona keys ao Map mestre
                }
                catch (e) { throw Error(`${this._log} Error to get values. \n\n${e.stack}`) }     // Retorna erros
                break;

            case "full":

                try { this._fetchNewData(true, columnIndexes) }                                 // Requisita dados
                catch (e) { throw Error(`${this._log} Error to get values. \n\n${e.stack}`) }   // Retorna erros
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

        // Retorna map de índices
        return columnIndexes
    }

    /**
     * Localiza os índices das linhas para chaves específicas ou para todas as chaves da planilha.
     * * @param {string|string[]|boolean} requestedKeys - As chaves a serem localizadas. 
     * Aceita uma string única, um array de strings ou `true` para mapear todas as chaves existentes.
     * @param {Map<string, number>} [columnIndexes] - Mapa opcional de cabeçalhos e índices. 
     * Se omitido, utiliza o mapeamento padrão da instância.
     * * @returns {Map<string, number>} Um Map onde a chave é o ID (string) e o valor é o índice da linha 0-based (number).
     * @private
     */
    _getRowIndexesByKey(requestedKeys, columnIndexes) {

        let rowIndexes = new Map();                     // Map para guardar índices
        let lastRow = this._table.getLastRow();         // Obtém última linha
        let keys_colIdx;                                // Var para guardar indice da coluna de keys
        let mode;

        // Valida parâmetros
        columnIndexes = columnIndexes === undefined ? this._getColumnIndexes() : columnIndexes
        mode = typeof requestedKeys === 'string' ? "SINGLE" : mode
        mode = Array.isArray(requestedKeys) ? "MULTI" : mode
        mode = requestedKeys == true ? "FULL" : mode
        requestedKeys = Array.isArray(requestedKeys) ? new Set([...requestedKeys]) : requestedKeys

        if (mode === undefined) throw Error(`Invalid requestedKeys.`)

        // Procura coluna de índices
        try { keys_colIdx = columnIndexes.get(this._keyColumnName) }                        // Procura pelo nome
        catch (e) { throw Error(`Error locating keyColumn: \n\n${e.stack}`) }    // Retorna outros erros
        if (keys_colIdx == undefined) throw Error(`keyColumn not found`)     // Se não achar coluna, lança erro

        try {

            // Modo rápido
            if (mode == "SINGLE") {
                let index = this._table.getRange(2, keys_colIdx + 1, lastRow - 1)   // Obtém range de keys
                    .createTextFinder(requestedKeys).matchEntireCell(true)          // Pesquisa na planilha
                    .findPrevious()                                                 // Obtém índice da última instância
                if (index === null) throw Error(`Error locating key.`)              // Se não tem key, retorna erro
                rowIndexes.set(requestedKeys, index.getRow() - 1)                   // Adiciona indice no Map
                return rowIndexes                                                   // Encerra execução
            }

            if (lastRow >= 2) { // Se planilha não está vazia

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

                keys_rawValues.forEach((k, i) => {                                                  // Para cada key
                    let trimKey = String(k).trim()                                                      // Limpa key
                    if (k == "" || k == null || k == undefined) return;                                 // Ignora keys vazias
                    if (mode === "FULL" || requestedKeys.has(trimKey)) rowIndexes.set(trimKey, i + 1)   // Armazena keys com índice
                })

                // Encerra execução
                if (rowIndexes.size === 0) throw Error(`Error locating keys.`)
                return rowIndexes
            }
        } catch (e) { throw Error(`Error to get keys: \n\n${e.stack}`) }  // Retorna outros erros
    }

    /**
     * Requisita novos dados da planilha via API avançada e realiza o pivoteamento para o cache interno.
     * Suporta busca completa ou por chaves específicas.
     * * @param {string[]|boolean} requestedKeys - Chaves para serem buscadas, ou true para realizar a requisição de colunas completas.
     * @param {Map<string, number>} [columnIndexes] - Mapa contendo os nomes das colunas e seus respectivos 
     * índices. Caso omitido, utiliza o mapeamento padrão da instância.
     * @throws {Error} Se o parâmetro requestedKeys não for um array nem o valor booleano true.
     * @private
     */
    _fetchNewData(requestedKeys, columnIndexes) {

        // HELPERS

        /**
         * Obtém os valores de um intervalo da planilha utilizando a estrutura GridRange (0-indexed).
         * Esta função atua como um intermediário da API Avançada do Sheets, com fallback no método nativo
         * getRange do Apps Script. Retorna um array similar ao Range.getValues().
         * * @param {GoogleAppsScript.Spreadsheet.Sheet} table - A instância da aba da planilha (Sheet).
         * * @param {Object} gridRange - Objeto contendo as coordenadas do intervalo.
         * @param {number} gridRange.startRowIndex - Índice inicial da linha (0-indexed, inclusive).
         * @param {number} gridRange.endRowIndex - Índice final da linha (0-indexed, exclusive).
         * @param {number} gridRange.startColumnIndex - Índice inicial da coluna (0-indexed, inclusive).
         * @param {number} gridRange.endColumnIndex - Índice final da coluna (0-indexed, exclusive).
         * * @returns {any[[]]} Um array contendo todos os valores do intervalo solicitado.
         * @private
         */
        function SAFEMODE_getValuesByGridRange(table, gridRange) {
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
         * Agrupa dimensões adjacentes (chunks) para reduzir o número de requisições à API.
         *
         * @param {Map<string, {gridRange: Object}>} gridRangesMap Mapa contendo os intervalos individuais de cada dimensão.
         * @param {"ROWS"|"COLUMNS"} dimension Define a orientação da mesclagem
         * @returns {Map<string[], {gridRange: Object}>} Mapa com os intervalos fundidos, onde a chave é a lista de nomes das colunas agrupadas.
         * @private
         */
        function SAFEMODE_mergeGridRanges(gridRangesMap, dimension) {

            // Variável de tamanho máximo de chunk
            const CHUNK_SIZE = dimension.includes("ROWS") ? 5 : 10

            // Define o nome das propriedades para cada dimensão
            let startDimensionIndex = dimension.includes("ROWS") ? "startRowIndex" : "startColumnIndex"
            let endDimensionIndex = dimension.includes("ROWS") ? "endRowIndex" : "endColumnIndex"


            // Junta todas as gridRanges na array, mantendo as keys
            let allGridRanges = []
            gridRangesMap.forEach((data, key) => {

                // Copia o objeto
                let safeData = { ...data };
                safeData.gridRange = { ...data.gridRange };

                // Insere o dado de key e armazena
                safeData.gridRange.key = key;
                allGridRanges.push(safeData);

            })

            // Ordena os gridRanges
            let sortedGridRanges = allGridRanges.sort((a, b) => {
                return a.gridRange[startDimensionIndex] - b.gridRange[startDimensionIndex]
            })

            // Mesclando gridRanges
            let mergedGridRanges = new Map()
            let mergingGridRange = undefined
            let mergingKeys = undefined
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
                    mergingKeys = [workingGridRange.key]
                    continue;
                }

                // Mescla vizinhos
                let validMerge = (mergingGridRange[endDimensionIndex] == workingGridRange[startDimensionIndex])
                if (validMerge) {
                    mergingGridRange[endDimensionIndex] = workingGridRange[endDimensionIndex]
                    mergingKeys.push(workingGridRange.key);
                }

                // Caso tenha mesclado e atingido o tamanho ou caso não tenha sido mesclado
                if ((validMerge && mergingKeys.length == CHUNK_SIZE) || (!validMerge)) {
                    mergedGridRanges.set(mergingKeys, { gridRange: mergingGridRange })
                    mergingGridRange = undefined
                    mergingKeys = undefined
                }

                // Caso não mesclado, guarda para nova iteração
                if (!validMerge) sortedGridRanges.unshift({ gridRange: workingGridRange })
            }

            // Caso tenha ficado algum objeto pra trás, armazena ele
            if (mergingGridRange != undefined) mergedGridRanges.set(mergingKeys, { gridRange: mergingGridRange })

            // Retorna objeto mesclado
            return mergedGridRanges
        }






        // Valida parâmetros
        if (!Array.isArray(requestedKeys) && requestedKeys != true) throw Error(`Invalid requestedKeys.`)

        let requestedData = new Map()                                               // Informações de dados a serem requeridos para a API
        let mode = Array.isArray(requestedKeys) ? "ROWS" : "COLUMNS"                // Define modo de execução
        if (columnIndexes == undefined) columnIndexes = this._getColumnIndexes()    // Obtém índices de colunas, caso não recebido
        let requestedRows = mode === "ROWS" ?                                       // Caso modo de linhas
            [...this._getRowIndexesByKey(requestedKeys).values()] :                     // Obtém índice das linhas
            undefined                                                                  // Caso não, mantém indefinido

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
                            startRowIndex: rowIndex,
                            endRowIndex: rowIndex + 1,
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

                    // Caso solicitado, converte valores dos objetos
                    if (this._options.enableTypeInference) obj = this._typeGStoJS(obj)

                    // Armazena resutados
                    this._data.set(String(key).trim(), obj)
                    this._keys.set(String(key).trim(), "unmodified")
                })
                break;

            case "COLUMNS-SAFETY":

                // Avisa o usuário sobre o uso do modo de segurança
                console.warn(`${this._log} Too much data, activating safety mode. Consider requesting fewer columns or using minimal mode with .search() to increase speed.`)

                // Obtém dados de key
                let keyData = SAFEMODE_getValuesByGridRange(this._table, requestedData.get(this._keyColumnName).gridRange).map(([v]) => String(v).trim())

                // Registra IDs nos metadados de keys
                for (let key of keyData) {
                    if (key === undefined || key === null || key === "") continue;
                    this._keys.set(key, "unmodified")
                    this._data.set(key, {})
                    this._data.get(key)[this._keyColumnName] = key
                }


                requestedData.delete(this._keyColumnName)                                       // Remove a requisição de coluna de key
                let mergedRequestedColumns = SAFEMODE_mergeGridRanges(requestedData, "COLUMNS") // Mescla as requisições

                // Para cada conjunto de requisições
                mergedRequestedColumns.forEach(({ gridRange }, colNames) => {

                    let workingArray = SAFEMODE_getValuesByGridRange(this._table, gridRange)    // Obtém dados

                    // Para cada linha
                    workingArray.forEach((row, rowInd) => row.forEach((value, colInd) => {

                        let currentKey = keyData[rowInd]                                        // Obtém key atual
                        if (!this._keys.has(currentKey)) return;                                // Se ID inválido, ignorar
                        if (this._options.enableTypeInference) value = this._typeGStoJS(value)  // Caso solicitado, converte valores dos objetos
                        this._data.get(currentKey)[colNames[colInd]] = value;                   // Armazena valor na memória
                    }))
                })
                break;

            case "ROWS":

                // Prepara dados para leitura
                let rowsData = APIresponse.valueRanges.map(range => (range.valueRange.values && range.valueRange.values[0]) ? range.valueRange.values[0] : [])

                let colOffset = Math.min(...columnIndexes.values())                     // Obtém o offset de colunas
                let keyIndex = columnIndexes.get(this._keyColumnName) - colOffset       // Obtém o índice das keys

                // Para cada linha recebida
                rowsData.forEach(row => {

                    if (row[keyIndex] === undefined || row[keyIndex] === null || String(row[keyIndex]).trim() === "") return;   // Ignora linhas sem keys
                    let obj = {}                                                                                                // Cria um objeto de saída

                    // Para cada coluna solicitada, cria a propriedade e armazena o valor no objeto
                    columnIndexes.forEach((colInd, colName) => obj[String(colName).trim()] = row[colInd - colOffset] ?? null)

                    // Caso solicitado, converte valores dos objetos
                    if (this._options.enableTypeInference) obj = this._typeGStoJS(obj)

                    // Armazena resutados
                    this._data.set(String(row[keyIndex]).trim(), obj)
                    this._keys.set(String(row[keyIndex]).trim(), "unmodified")
                })
                break;

            case "ROWS-SAFETY":

                // Avisa o usuário sobre o uso do modo de segurança
                console.warn(`${this._log} Too much data, activating safety mode. Consider requesting fewer rows or using minimal mode with .search() to increase speed.`)

                let mergedRequestedRows = SAFEMODE_mergeGridRanges(requestedData, "ROWS")       // Mescla as requisições
                let safe_colOffset = Math.min(...columnIndexes.values())                        // Obtém o offset de colunas
                let safe_keyIndex = columnIndexes.get(this._keyColumnName) - safe_colOffset     // Obtém o índice das keys

                // Para cada conjunto de requisições
                mergedRequestedRows.forEach(({ gridRange }) => {

                    let workingArray = SAFEMODE_getValuesByGridRange(this._table, gridRange)    // Obtém dados
                    workingArray.forEach((row) => {                                             // Para cada linha

                        let currentKey = String(row[safe_keyIndex]).trim()                                  // Obtém key atual
                        if (currentKey === undefined || currentKey === null || currentKey === "") return;   // Ignora linhas sem keys
                        let obj = {}                                                                        // Cria objeto de saída

                        // Para cada coluna solicitada, cria a propriedade e armazena o valor no objeto
                        columnIndexes.forEach((colInd, colName) => obj[String(colName).trim()] = row[colInd - safe_colOffset] ?? null)

                        // Caso solicitado, converte valores dos objetos
                        if (this._options.enableTypeInference) obj = this._typeGStoJS(obj)

                        // Armazena resutados
                        this._data.set(String(row[safe_keyIndex]).trim(), obj)
                        this._keys.set(String(row[safe_keyIndex]).trim(), "unmodified")
                    })
                })
                break;
        }
    }

    /**
    * Define a new status for the key.
    * @param {string} key The key of the entry to update status.
    * @param {"new"|"modified"|"deleted"} newState The status of the key.
    * @private
    */
    _setKeyAs(key, newState) {

        // Obtém estado atual da chave
        let actualState = this._keys.get(key)

        // Caso não tenha sido modificado, adiciona estado e encerra
        if (actualState === "unmodified") {
            this._keys.set(key, newState)
            return undefined
        }

        // Caso estado anterior seja igual ao novo, encerra
        if (actualState === newState) return;

        // Age conforme o estado atual
        switch (actualState) {

            case "new":
                if (newState == "deleted") this._keys.delete(key);
                // if (newState == "modified") deve manter o estado como "new" 
                break;

            case "modified":
                // if (newState == "new") não é uma operação válida
                if (newState == "deleted") this._keys.set(key, "deleted");
                break;
            case "deleted":
                // Adicionar uma chave deletada a reativa modificando o valor.
                if (newState == "new") this._keys.set(key, "modified");
            // if (newState == "modified") não reabilita a chave. 
        }
    }

    /**
     * Valida e converte dados para armazenamento no Google Sheets.
     * * Esta função atua como um motor de processamento recursivo que converte tipos complexos 
     * do JavaScript (como Map, Set, BigInt e RegExp) em formatos e tamanhos compatíveis com
     * as células da planilha.
     * @param {*} value - O dado a ser processado (Primitivos, Coleções ou Objetos).
     * @param {boolean} [runConversion=false] - Se verdadeiro, efetivamente converte o dado. 
     * Se falso, valida o dado para o commit .
     * @param {number} [depth=0] - Uso interno para recursão.
     * @returns {*} O valor enviado:
     * - `runConversion = true`: Retorna strings (JSON ou Primitivos) prontas para o Sheets.
     * - `runConversion = false`: Retorna o mesmo valor enviado.
     * * @throws {Error} Se a profundidade de aninhamento exceder 25 níveis.
     * @throws {Error} Se uma string resultante (JSON ou texto) ultrapassar 50.000 caracteres.
     * @throws {Error} Se chaves de um `Map` não forem do tipo `string` ou `number`.
     * @throws {Error} Se o tipo de dado não for suportado pela biblioteca.
     * @private
     */
    _typeJStoGS(value, runConversion = false, depth = 0) {

        let convertedValue;

        // Valida dado baseado no tipo
        switch (typeof value) {

            // Não necessita validar
            case 'number':
            case 'boolean':
                return value;

            case 'undefined':
                // Retorna undefined ou string vazia (commit)
                if (runConversion) { return "" } else { return undefined };


            case 'string':
                if (value.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)
                return value;

            case 'bigint':

                // Obtém maiores números possíveis em Integer
                const maxint = BigInt(Number.MAX_SAFE_INTEGER);
                const minint = BigInt(Number.MIN_SAFE_INTEGER);

                // Se não compatível com Integer, converter para String
                if (value > maxint || value < minint) {
                    convertedValue = String(value);
                    if (convertedValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)
                    if (runConversion) { return convertedValue } else { return value };
                }

                // Converte para Integer seguramente
                if (runConversion) { return Number(value) } else { return value };

            case 'object':

                // NULL
                if (value === null) {
                    if (runConversion) { return "" } else { return null };     // Retorna nulo ou string vazia (commit)
                }

                // DATAS
                if (value instanceof Date) {
                    if (runConversion) { return isNaN(value.getTime()) ? "Invalid date" : value }  // Caso commit, limpa data inválida
                    else { return value }                                                       // Se não, retorna valor     
                }

                // REGEX
                if (value instanceof RegExp) {
                    convertedValue = value.toString()
                    if (convertedValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)
                    if (runConversion) { return convertedValue } else { return value };
                }

                // ARRAY ou SET
                if (value instanceof Array || value instanceof Set) {

                    convertedValue = [...value]                                                                     // Cria cópia de segurança
                    if (depth < 25) convertedValue = convertedValue.map(v => this._typeJStoGS(v, true, depth + 1))  // Limpa até 25 camadas
                    if (depth == 25) throw Error(`The inputed array contains more than 25 levels of depth.`)        // Para de converter acima de 25 camadas
                    if (depth != 0) return convertedValue                                                           // Caso em recursão, retorna valor convertido

                    // Valida tamanho da string
                    convertedValue = JSON.stringify(convertedValue)
                    if (convertedValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)

                    // Encerra execução
                    if (runConversion) { return convertedValue } else { return value };
                }

                // OBJETO LITERAL OU MAP
                if (Object.prototype.toString.call(value) === '[object Object]' || value instanceof Map) {

                    // Obtém o encadeamento chave / valor
                    if (value instanceof Map) {
                        convertedValue = [...value.entries()]
                        // Executa uma validação simples nas keys
                        for (let [key] of convertedValue) {
                            if (typeof key !== 'number' && typeof key !== 'string') throw Error(`Maps with not-string or not-number keys are not supported.`)
                        }
                    }
                    else convertedValue = Object.entries(value)

                    if (depth < 25) convertedValue = convertedValue.map(([k, v]) => [k, this._typeJStoGS(v, true, depth + 1)])  // Limpa até 25 camadas
                    if (depth == 25) throw Error(`The inputed object contains more than 25 levels of depth.`)                   // Para de converter acima de 25 camadas
                    if (depth != 0) return Object.fromEntries(convertedValue)                                                                       // Caso em recursão, retorna valor convertido

                    // Valida tamanho da string
                    convertedValue = JSON.stringify(Object.fromEntries(convertedValue))
                    if (convertedValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)

                    // Encerra execução
                    if (runConversion) { return convertedValue } else { return value };

                }

            default:
                throw Error(`Value ${value} not supported.`)
        }
    }

    /**
         * Converte dados brutos vindos do Google Sheets para tipos nativos do JavaScript.
         * * Esta função é a contraparte simétrica de `_typeJStoGS`. Ela analisa o valor bruto 
         * recebido (geralmente de uma célula da planilha) e tenta identificar se ele representa 
         * uma estrutura complexa que foi serializada, como JSON (Arrays e Objetos), 
         * Expressões Regulares (RegExp) ou strings de Data em formato ISO.
         * * @param {*} value - O valor bruto a ser processado.
         * @returns {*} O valor reidratado para o tipo nativo mais rico identificado.
         * @private
         */
    _typeGStoJS(value) {
        switch (typeof value) {

            // Não processar casos nativos
            case "number":
            case "boolean":
            case "undefined":
                return value

            // Caso objeto, desambiguar:
            case "object":

                // NULL 
                if (value === null) return null

                // DATA
                if (value instanceof Date) return value

                // ARRAY
                if (value instanceof Array) {
                    return value.map(v => this._typeGStoJS(v))
                }

                // OBJETO LITERAL
                if (Object.prototype.toString.call(value) === '[object Object]') {
                    let entries = Object.entries(value)                                     // Desmonta
                    let parsedEntries = entries.map(([k, v]) => [k, this._typeGStoJS(v)])   // Roda recursivamente
                    return Object.fromEntries(parsedEntries)                                // Remonta
                }

            // Caso string, desambiguar
            case "string":

                // undefined
                if (value === "") return undefined

                // DATA ISO
                let regex_DateISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}.\d{3}Z$/
                if (regex_DateISO.test(value)) try {
                    let date = new Date(value)
                    if (!isNaN(date.getTime())) return date
                } catch (e) { }

                // ARRAYS E OBJECTS
                let regex_JSON = /^\s*[\{\[][\s\S]*[\}\]]\s*$/
                if (regex_JSON.test(value)) try {
                    let object = JSON.parse(value)
                    return this._typeGStoJS(object)
                } catch (e) { }

                // REGEX
                let regex_RegexString = /^\/((?:\\\/|[^\/])+)\/([gimuyvd]*)$/
                if (regex_RegexString.test(value)) try {
                    let [fullmatch, pattern, flags] = value.match(regex_RegexString)
                    return new RegExp(pattern, flags)
                } catch (e) { }

            // Padrão, retornar valor recebido
            default: return value

        }
    }

    /**
     * Cria um Proxy recursivo para monitoramento de mutações e rastreamento de estado.
     * * @param {object} value - O objeto, array ou estrutura mutável a ser monitorada.
     * @param {string} key - A Chave Primária (ID) da linha à qual este dado pertence.
     * @returns {object} Um Proxy que se comporta como o objeto original, mas rastreia mudanças.
     * @throws {Error} Se houver tentativa de modificar a coluna de Chave Primária.
     * @throws {Error} Se uma operação ilegal for detectada em tipos não suportados.
     * @private
     */
    _createProxy(value, key) {

        // Retorna proxy do cache se existir
        if (this._proxies.has(value)) return this._proxies.get(value)

        // Cria a trap de edições para commits
        let proxyHandlers = {
            set: (ogObj, colName, value) => {

                // Impede escrita de valores na coluna de keys.
                if (colName === this._keyColumnName) throw Error(`Key values are not writable.`)

                let cleanValue = this._typeJStoGS(value)        // Garante que a informação é compatível
                this._keys.set(key, "modified")                 // Marca o objeto como modificado
                return Reflect.set(ogObj, colName, cleanValue)  // Edita o objeto
            },

            get: (ogObj, colName) => {

                // Obtém objeto e alterna comportamento conforme tipo
                let value = Reflect.get(ogObj, colName)
                switch (typeof value) {

                    // Tipos primitivos não usam proxy
                    case "string":
                    case "number":
                    case "bigint":
                    case "boolean":
                    case "undefined":
                        return value;

                    // Funções não usam proxies e precisam da referecia original no this
                    case "function":
                        return (...args) => {
                            this._keys.set(key, "modified")
                            return value.apply(ogObj, args)
                        }

                    // Variar comportamento para proxies
                    case "object":

                        if (value === null) return null                 // NULL
                        if (value instanceof RegExp) return value       // REGEX
                        return this._createProxy(value, key)   // OUTROS

                    // Tipos que não deveriam existir retornam erro
                    default:
                        throw Error(`Illegal operation.`)
                }
            }
        }

        // Cria novo proxy
        let newProxy = new Proxy(value, proxyHandlers)
        this._proxies.set(value, newProxy)
        return newProxy
    }



    // MÉTODOS PÚBLICOS

    /**
     * Removes all elements from the Codex instance and schedules a full cleanup 
     * of the spreadsheet on the next commit.
     */
    clear() {
        this._wipeOnCommit = true;  // Marca planilha para exclusão
        this._keys.clear();         // Limpa histórico de mudanças
        this._data.clear();         // Limpa memória da instancia
    }

    /**
     * Removes the specified element from the Codex instance by key.
     * Schedules the deletion of the corresponding row in the Google Sheets on the next commit.
     * * @param {string} key The key of the element to remove.
     * @returns {boolean} `true` if an element in the Codex object existed and has been removed, or `false` if the element does not exist.
     */
    delete(key) {

        key = String(key).trim()                // Formata key
        let isDeleteable = this._keys.has(key)   // Verifica se há um dado a ser excluido

        // Caso deletável
        if (isDeleteable) {
            this._data.delete(key)          // Remove da memória
            this._setKeyAs(key, "deleted")  // Marca como deletado
        }

        // Retorna se dado está excluído
        return isDeleteable
    }

    /**
     * Checks if a specific key exists in the instance.
     * @param {string} key - The unique identifier (ID) to check.
     * @returns {boolean} `true` if the key exists and is active; `false` otherwise.
     */
    has(key) {
        let keyStatus = this._keys.get(String(key).trim())  // Obtém estado
        if (keyStatus === undefined) return false           // Se não existe, false
        if (keyStatus === "deleted") return false           // Se deletado, false
        return true                                         // Retorna que existe
    }


    /**
     * Retrieves a record by its unique Primary Key.
     * * @param {string|number} key - The unique identifier (ID) of the record.
     * @returns {Object|undefined} The data associated with the key, or `undefined` if the key does not exist or is marked as deleted.
     * * @example
     * const user = db.get("user_01");
     * if (user) {
     * user.lastLogin = new Date(); // Automatically marked as 'modified'
     * }
     */
    get(key) {
        key = String(key).trim()                        // Formata a key
        let keyStatus = this._keys.get(key)             // Obtém estado da key
        if (keyStatus === undefined) return undefined   // Se não existe, encerra
        if (keyStatus === "deleted") return undefined   // Se deletada, encerra
        let keyLoaded = this._data.has(key)             // Verifica se carregado
        if (!keyLoaded) this._fetchNewData([key])       // Requisita o load do dado
        let requestedData = this._data.get(key)         // Carrega o dado em uma var local
        return this._createProxy(requestedData, key)    // Cria proxy do objeto e retorna.
    }

    /**
     * Returns a generator that yields all active Primary Keys in the store.
     * * @yields {string} The next active Primary Key.
     * @returns {IterableIterator<string>} An iterable iterator of non-deleted keys.
     */
    *keys() {
        for (const [key, status] of this._keys) {   // Para cada key
            if (status !== "deleted") yield key     // Retorna sob demanda as keys
        }
    }
}



