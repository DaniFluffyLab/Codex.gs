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
        try {
            this._table = this._sheet.getSheetByName(this._tableName)                       // Carrega página
            if (this._table === null) throw Error(`Sheet not found`)                        // Lança erro se não houver página
        }
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
            mode: options.mode ?? "minimal",
            columns: new Set(options.columns ?? []),
            enableTypeInference: options.enableTypeInference ?? true,
        };
        if (this._options.mode !== 'minimal' && this._options.mode !== 'full') throw Error(`${this._log} Invalid mode: ${this._options.mode}`)
        for (let c of this._options.columns) { if (typeof c !== 'string') throw Error(`${this._log} Column "${c}" is not a string.`) }



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

        // Vars de metadados para detectar proxies 
        this._isCdxProxy = Symbol("isCodexProxy")
        this._cdxProxyTarget = Symbol("getTarget")

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

        // Lança erro se aba completamente vazia
        if (lastColumn === 0) throw Error(`The sheet "${this._tableName}" is empty (no headers found).`);

        let columnArray = this._table.getRange(1, 1, 1, lastColumn)     // Seleciona cabeçalho
            .getValues()[0]                                             // Obtém dados
        columnIndexes = new Map(columnArray.map((v, i) => [v, i]))      // Insere dados dos índices no Map

        // Valida a existência de uma coluna de keys
        if (!columnIndexes.has(this._keyColumnName)) { throw Error(`Primary Key column "${this._keyColumnName}" does not exist.`); }

        // Caso não hajam parâmetros sobre quais colunas obter, alimentar com nome de todas as colunas
        if (this._options.columns.size === 0) { this._options.columns = new Set(columnIndexes.keys()) }

        let filteredColumnIndexes = new Map()                                                   // Cria Map para colunas filtradas
        filteredColumnIndexes.set(this._keyColumnName, columnIndexes.get(this._keyColumnName))  // Garante coluna de key
        for (let columnName of this._options.columns) {                                         // Para cada coluna requisitada:
            let hasInvalid = !columnIndexes.has(columnName)                                         // Verifica a validade da colunas
            if (hasInvalid) { throw Error(`Requested column "${columnName}" do not exist.`); }      // Lança erro se inválido 
            filteredColumnIndexes.set(columnName, columnIndexes.get(columnName))                    // Armazena seu valor no Map de filtradas
        }

        // Retorna map de índices
        return filteredColumnIndexes
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
                let index = this._table.getRange(2, keys_colIdx + 1, lastRow - 1)       // Obtém range de keys
                    .createTextFinder(requestedKeys).matchEntireCell(true)              // Pesquisa na planilha
                    .findPrevious()                                                     // Obtém índice da última instância
                if (index !== null) rowIndexes.set(requestedKeys, index.getRow() - 1)  // Adiciona indice no Map
                return rowIndexes                                                       // Encerra execução
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
                ).valueRanges[0].valueRange.values

                if (keys_rawValues && keys_rawValues[0]) keys_rawValues[0].forEach((k, i) => {           // Para cada key
                    let trimKey = String(k).trim()                                                      // Limpa key
                    if (k == "" || k == null || k == undefined) return;                                 // Ignora keys vazias
                    if (mode === "FULL" || requestedKeys.has(trimKey)) rowIndexes.set(trimKey, i + 1)   // Armazena keys com índice
                })

                // Encerra execução
                return rowIndexes
            }
        } catch (e) { throw Error(`Error to get keys: \n\n${e.stack}`) }  // Retorna outros erros
    }

    /**
     * Requisita novos dados da planilha via API avançada e realiza o pivoteamento para o cache interno.
     * Suporta busca completa ou por chaves específicas.
     * * @param {string[]|number[]|boolean} requestedKeysOrRows - Chaves ou linhas para serem buscadas, ou true para realizar a requisição de colunas completas.
     * @param {Map<string, number>} [columnIndexes] - Mapa contendo os nomes das colunas e seus respectivos 
     * índices. Caso omitido, utiliza o mapeamento padrão da instância.
     * @throws {Error} Se o parâmetro requestedKeys não for um array nem o valor booleano true.
     * @private
     */
    _fetchNewData(requestedKeysOrRows, columnIndexes) {

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
                let response = Sheets.Spreadsheets.Values.batchGetByDataFilter(
                    {
                        dataFilters: [{ gridRange: gridRange }],
                        majorDimension: "ROWS",
                        valueRenderOption: "UNFORMATTED_VALUE",
                        dateTimeRenderOption: "FORMATTED_STRING"
                    },
                    table.getSheetId(),
                ).valueRanges[0].valueRange.values

                // Garante alguma resposta
                return response ?? []
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
        if (!Array.isArray(requestedKeysOrRows) && requestedKeysOrRows != true) throw Error(`Invalid requestedKeysOrRows.`)

        let requestedData = new Map()                                               // Informações de dados a serem requeridos para a API
        let mode = Array.isArray(requestedKeysOrRows) ? "ROWS" : "COLUMNS"          // Define modo de execução
        if (columnIndexes == undefined) columnIndexes = this._getColumnIndexes()    // Obtém índices de colunas, caso não recebido
        let lastRow = this._table.getLastRow();                                     // Obtém última linha
        let requestedRows;                                                          // Var para valores de linhas

        // Obtém índices das linhas
        if (mode === 'ROWS') {

            // Testa se array está vazia
            if (requestedKeysOrRows.length === 0) return;

            // Testa se todos são do mesmo tipo
            let type = typeof requestedKeysOrRows[0]
            if (!requestedKeysOrRows.every(v => (typeof v === type))) {
                throw Error(`requestedKeysOrRows must be a uniform array of strings (PKs) or numbers (Indexes).`)
            }

            // Alterna entre tipos
            switch (type) {

                case 'number':

                    // Testa se índices são válidos
                    if (requestedKeysOrRows.some(v => (v >= lastRow || v < 1))) {
                        throw Error(`requestedKeysOrRows must be more than 0 and less than last row index.`)
                    }

                    // Popula eles na array
                    requestedRows = [...requestedKeysOrRows]
                    break;

                case 'string':

                    // Obtém os índices com função auxiliar
                    requestedRows = [...this._getRowIndexesByKey(requestedKeysOrRows).values()]
                    break;

                default:
                    throw Error(`requestedKeysOrRows must be a uniform array of strings (PKs) or numbers (Indexes).`)
            }

        }

        // Monta os objetos de requisição
        switch (mode) {

            case "COLUMNS":
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
                console.warn(`${this._log} Too much data, activating safety mode. Consider requesting fewer columns or using minimal mode with Codex.search() to increase speed.`)

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
                console.warn(`${this._log} Too much data, activating safety mode. Consider requesting fewer rows or using minimal mode with Codex.search() to increase speed.`)

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

        // Caso estado anterior seja igual ao novo, encerra
        if (actualState === newState) return;

        // Age conforme o estado atual
        switch (actualState) {

            case undefined:
                if (newState == "new") this._keys.set(key, "new");
                if (newState == "modified") this._keys.set(key, "new");
                // Recebeu "deleted" => nada muda
                break;

            case "new":
                // Recebeu "new" => nada muda
                // Recebeu "modified" => nada muda
                if (newState == "deleted") this._keys.delete(key);
                break;

            case "unmodified":
                if (newState == "new") this._keys.set(key, "modified");
                if (newState == "modified") this._keys.set(key, "modified");
                if (newState == "deleted") this._keys.set(key, "deleted");
                break;

            case "modified":
                // Recebeu "new" => nada muda
                // Recebeu "modified" => nada muda
                if (newState == "deleted") this._keys.set(key, "deleted");
                break;

            case "deleted":
                if (newState == "new") this._keys.set(key, "modified");
                // Recebeu "modified" => nada muda
                // Recebeu "deleted" => nada muda
                break;

        }
    }

    /**
     * Valida e converte dados para armazenamento no Google Sheets.
     * * Esta função atua como um motor de processamento recursivo que converte tipos complexos 
     * do JavaScript (como Map, Set, BigInt e RegExp) em formatos e tamanhos compatíveis com
     * as células da planilha.
     * @param {*} value - O dado a ser processado (Primitivos, Coleções ou Objetos).
     * @param {'test'|'clone'|'commit'} [mode='test'] - Altera o modo de operação:
     * - `test`: Apenas valida a compatibilidade e retorna o objeto original.
     * - `clone`: Útil para objetos, também cria uma cópia do objeto original.
     * - `commit`: Efetivamente converte os objetos para serem submetidos ao GSheets.
     * @param {number} [depth=0] - Uso interno para recursão.
     * @returns {*} O valor enviado:
     * - `test`: Retorna o objeto original. No caso de CdxProxies, retorna o objeto origem.
     * - `clone`: Retorna uma cópia do objeto original.
     * - `commit`: Retorna os valores nos tipos suportados pelo GSheet.
     * * @throws {Error} Se a profundidade de aninhamento exceder 25 níveis.
     * @throws {Error} Se uma string resultante (JSON ou texto) ultrapassar 50.000 caracteres.
     * @throws {Error} Se chaves de um `Map` não forem do tipo `string` ou `number`.
     * @throws {Error} Se o tipo de dado não for suportado pela biblioteca.
     * @private
     */
    _typeJStoGS(value, mode = 'test', depth = 0) {

        // Vars usadas no switch
        let convertedValue, jsonValue;

        // Valida dado baseado no tipo
        switch (typeof value) {

            // Não necessita validar
            case 'number':
            case 'boolean':
                return value;

            case 'undefined':
                // Retorna undefined ou string vazia (commit)
                if (mode === 'commit') { return "" } else { return undefined };


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
                    if (mode === 'commit') { return convertedValue } else { return value };
                }

                // Converte para Integer seguramente
                if (mode === 'commit') { return Number(value) } else { return value };

            case 'object':



                // Caso seja um proxy, busca trabalhar com os dados originais
                if (value && value[this._isCdxProxy]) value = value[this._cdxProxyTarget]



                // NULL
                if (value === null) {
                    if (mode === 'commit') { return "" } else { return null };     // Retorna nulo ou string vazia (commit)
                }




                // DATAS
                if (value instanceof Date) {
                    if (mode === 'commit') { return isNaN(value.getTime()) ? "Invalid date" : value }   // Caso commit, limpa data inválida
                    if (mode === 'clone') { return new Date(value) }                                    // Caso clone, retorna data clonada
                    if (mode === 'test') { return value }                                               // Caso teste, retorna valor     
                }




                // REGEX
                if (value instanceof RegExp) {
                    convertedValue = value.toString()
                    if (convertedValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)
                    if (mode === 'commit') { return convertedValue }    // Caso commit, envia o regex convertido
                    if (mode === 'clone') { return new RegExp(value) }  // Caso clone, retorna novo regex
                    if (mode === 'test') { return value }               // Caso teste, retorna valor 
                }





                // ARRAY ou SET no modo commit
                if ((value instanceof Set || value instanceof Array) && mode === 'commit') {

                    convertedValue = [...value]                                                                         // Cria cópia de segurança
                    if (depth < 25) convertedValue = convertedValue.map(v => this._typeJStoGS(v, 'commit', depth + 1))  // Limpa até 25 camadas
                    if (depth == 25) throw Error(`The inputed object contains more than 25 levels of depth.`)           // Para de converter acima de 25 camadas
                    if (depth != 0) return convertedValue                                                               // Caso em recursão, retorna valor convertido

                    // Valida tamanho da string
                    jsonValue = JSON.stringify(convertedValue)
                    if (jsonValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)

                    // Encerra execução retornando JSON
                    if (mode === 'commit') { return jsonValue }
                }


                // ARRAY ou SET sem ser commit
                if ((value instanceof Set || value instanceof Array) && mode !== 'commit') {

                    this._typeJStoGS(value, 'commit')       // Valida o objeto
                    if (mode === 'test') { return value }   // Devolve o valor caso não precise de um clone

                    // Clona os objetos
                    convertedValue = [...value]                                                         // Clona a primeira camada 
                    convertedValue = convertedValue.map((v => this._typeJStoGS(v, mode, depth + 1)))   // Clona as posteriores

                    // Encerra execução
                    if (mode === 'clone' && value instanceof Set) { return new Set(convertedValue) }
                    if (mode === 'clone' && value instanceof Array) { return convertedValue }
                }





                // MAP no modo commit
                if (value instanceof Map && mode === 'commit') {

                    // Obtém o encadeamento chave / valor
                    convertedValue = [...value.entries()]
                    for (let [key] of convertedValue) {
                        if (typeof key !== 'number' && typeof key !== 'string') throw Error(`Maps with not-string or not-number keys are not supported.`)
                    }

                    if (depth < 25) convertedValue = convertedValue.map(([k, v]) => [k, this._typeJStoGS(v, 'commit', depth + 1)])  // Limpa até 25 camadas
                    if (depth == 25) throw Error(`The inputed object contains more than 25 levels of depth.`)                       // Para de converter acima de 25 camadas
                    if (depth != 0) return Object.fromEntries(convertedValue)                                                       // Caso em recursão, retorna valor convertido

                    // Valida tamanho da string
                    jsonValue = JSON.stringify(Object.fromEntries(convertedValue))
                    if (jsonValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)

                    // Encerra execução
                    return jsonValue;
                }


                // MAP sem ser commit
                if (value instanceof Map && mode !== 'commit') {

                    this._typeJStoGS(value, 'commit')       // Valida o objeto
                    if (mode === 'test') { return value }   // Devolve o valor caso não precise de um clone

                    // Clona os objetos
                    convertedValue = [...value.entries()]                                                       // Clona a primeira camada 
                    convertedValue = convertedValue.map(([k, v]) => [k, this._typeJStoGS(v, mode, depth + 1)])  // Clona as posteriores

                    // Encerra execução
                    if (mode === 'clone') { return new Map(convertedValue) }
                }




                // OBJETO LITERAL no modo commit
                if (Object.prototype.toString.call(value) === '[object Object]' && mode === 'commit') {

                    // Obtém o encadeamento chave / valor
                    convertedValue = Object.entries(value)

                    if (depth < 25) convertedValue = convertedValue.map(([k, v]) => [k, this._typeJStoGS(v, 'commit', depth + 1)])  // Limpa até 25 camadas
                    if (depth == 25) throw Error(`The inputed object contains more than 25 levels of depth.`)                       // Para de converter acima de 25 camadas
                    if (depth != 0) return Object.fromEntries(convertedValue)                                                       // Caso em recursão, retorna valor convertido
                    convertedValue = Object.fromEntries(convertedValue)                                                             // Fora da recursão, reconverte em objeto    

                    // Valida tamanho da string
                    jsonValue = JSON.stringify(convertedValue)
                    if (jsonValue.length > 50000) throw Error(`The input contains more than the maximum limit of 50,000 characters in a single cell.`)

                    // Encerra execução
                    if (mode === 'commit') { return jsonValue }     // Caso commit, retorna json
                }

                // OBJETO LITERAL sem ser commit
                if (Object.prototype.toString.call(value) === '[object Object]' && mode !== 'commit') {

                    this._typeJStoGS(value, 'commit')       // Valida o objeto
                    if (mode === 'test') { return value }   // Devolve o valor caso não precise de um clone

                    // Clona os objetos
                    convertedValue = Object.entries(value)                                                          // Clona a primeira camada 
                    convertedValue = convertedValue.map((([k, v]) => [k, this._typeJStoGS(v, mode, depth + 1)]))    // Clona as posteriores

                    // Encerra execução
                    if (mode === 'clone') { return Object.fromEntries(convertedValue) }
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
                break;

            // Caso string, desambiguar
            case "string":

                // undefined
                if (value === "") return undefined

                // DATA ISO
                let regex_DateISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{3})?Z$/
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
        }

        // Fallback, retornar valor recebido
        return String(value)
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
                this._setKeyAs(key, "modified")                 // Marca o objeto como modificado
                return Reflect.set(ogObj, colName, cleanValue)  // Edita o objeto
            },

            deleteProperty: (ogObj, colName) => {

                // Impede escrita de valores na coluna de keys.
                if (colName === this._keyColumnName) throw Error(`Key values are not deleteable.`)

                this._setKeyAs(key, "modified")                 // Marca o objeto como modificado
                return Reflect.deleteProperty(ogObj, colName)   // Deleta o valor no objeto
            },

            get: (ogObj, colName) => {

                // Comportamento de requisição de metadados
                if (colName === this._isCdxProxy) return true       // Valida que isso é uma Proxy 
                if (colName === this._cdxProxyTarget) return ogObj  // Devolve o objeto original

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
                            this._setKeyAs(key, "modified")
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
        try {

            this._wipeOnCommit = true;  // Marca planilha para exclusão
            this._keys.clear();         // Limpa histórico de mudanças
            this._data.clear();         // Limpa memória da instancia

        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }

    /**
     * Removes the specified element from the Codex instance by key.
     * Schedules the deletion of the corresponding row in the Google Sheets on the next commit.
     * * @param {string} key The key of the element to remove.
     * @returns {boolean} `true` if an element in the Codex object existed and has been removed, or `false` if the element does not exist.
     */
    delete(key) {
        try {

            key = String(key).trim()          // Formata key
            let status = this._keys.get(key)  // Verifica se há um dado a ser excluido

            // Varia comportamento
            switch (status) {

                // Nada é feito se já está excluído
                case undefined:
                case "deleted":
                    return false;

                default:
                    this._data.delete(key)          // Remove da memória
                    this._setKeyAs(key, "deleted")  // Marca como deletado
                    return true
            }
        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }

    /**
     * Checks if a specific key exists in the instance.
     * @param {string} key - The unique identifier (ID) to check.
     * @returns {boolean} `true` if the key exists and is active; `false` otherwise.
     */
    has(key) {
        try {

            let keyStatus = this._keys.get(String(key).trim())  // Obtém estado
            if (keyStatus === undefined) return false           // Se não existe, false
            if (keyStatus === "deleted") return false           // Se deletado, false
            return true                                         // Retorna que existe

        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }

    /**
     * Retrieves a record by its unique Primary Key.
     * * @param {string|number} key - The unique identifier (ID) of the record.
     * @returns {Object|undefined} The object associated with the key, with properties being the columns name,
     * or `undefined` if the key does not exist or is marked as deleted.
     * * @example
     * const user = db.get("user_01");
     * if (user) {
     * user.lastLogin = new Date(); // Automatically marked as 'modified'
     * }
     */
    get(key) {
        try {

            key = String(key).trim()                        // Formata a key
            let keyStatus = this._keys.get(key)             // Obtém estado da key
            if (keyStatus === undefined) return undefined   // Se não existe, encerra
            if (keyStatus === "deleted") return undefined   // Se deletada, encerra
            let keyLoaded = this._data.has(key)             // Verifica se carregado
            if (!keyLoaded) this._fetchNewData([key])       // Requisita o load do dado
            let requestedData = this._data.get(key)         // Carrega o dado em uma var local
            if (!requestedData) return undefined            // Se não achar, retorna undefined
            return this._createProxy(requestedData, key)    // Cria proxy do objeto e retorna.

        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }

    /**
     * Returns a iterator that contains all active Primary Keys in the store.
     * * @yields {string} The next active Primary Key.
     * @returns {IterableIterator<string>} An iterable iterator of non-deleted keys.
     */
    *keys() {
        try {

            for (const [key, status] of this._keys) {   // Para cada key
                if (status !== "deleted") yield key     // Retorna sob demanda as keys
            }

        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }

    /**
     * Returns a iterator that contains all active values in the store. Only avaliable on mode = full
     * * @yields {string} The next active value.
     * @returns {IterableIterator<object>} An iterable iterator of non-deleted values.
     * @throws {Error} If Codex is not in mode = full.
     * 
     */
    *values() {
        try {

            // Rejeita uso do método sem estar no modo full.
            if (this._options.mode != "full") throw Error(`Method Codex.values() is only avaliable on mode = full. Use Codex.search() instead.`)

            for (const [key, status] of this._keys) {           // Para cada key
                if (status !== "deleted") yield this.get(key)   // Retorna sob demanda os valores
            }

        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }

    /**
     * Returns a iterator that contains all active entries in the store. Only avaliable on mode = full
     * * @yields {string} The next active entries.
     * @returns {IterableIterator<[string, object]>} An iterable iterator of non-deleted entries.
     * @throws {Error} If Codex is not in mode = full.
     * 
     */
    *entries() {
        try {

            // Rejeita uso do método sem estar no modo full.
            if (this._options.mode != "full") throw Error(`Method Codex.entries() is only avaliable on mode = full. Use Codex.search() instead.`)

            for (const [key, status] of this._keys) {                   // Para cada key
                if (status !== "deleted") yield [key, this.get(key)]    // Retorna sob demanda as chave/valores
            }

        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }


    /**
     * Performs a structured search across the dataset using a specific matching strategy.
     * * This method acts as a Generator, lazily yielding records that match **all** the provided criteria (logical AND).
     * In "minimal" mode, it uses Google Sheets' native search (TextFinder) to locate rows 
     * before fetching data into memory.
     * * @param {("fullstring"|"partialstring"|"regex")} mode - The matching strategy to be applied:
     * - `"fullstring"`: Checks for exact equality (case-sensitive).
     * - `"partialstring"`: Checks if the value contains the substring (case-insensitive).
     * - `"regex"`: Matches using a Regular Expression (must be compatible with Google Sheets TextFinder Class).
     * * @param {Object.<string, string|RegExp>} search_for - A key-value object defining the filters.
     * - **Keys:** Must be valid column names defined in the schema.
     * - **Values:** The criteria to match against. Must be a `string` for string modes or a `RegExp` object for regex mode.
     * * @yields {Object} The next matching record, allowing for direct modification.
     * @returns {Generator<Object>} A generator that yields matching records one by one.
     * * @throws {Error} If `mode` is invalid or `search_for` is empty/not an object.
     * @throws {Error} If a column specified in `search_for` does not exist in the table.
     * @throws {Error} If the value type provided does not match the expected `mode` (e.g., passing a string when `regex` is expected).
     * @throws {Error} If a provided RegExp is incompatible with Google Sheets' TextFinder Class.
     * * @example
     * // 1. Exact match (Find active users in the "IT" department)
     * for (const user of db.search("fullstring", { Department: "IT", Status: "Active" })) {
     * console.log(user.Name);
     * }
     * * @example
     * // 2. Partial match (Find products containing "Apple" in the name)
     * // Matches "Apple", "Pineapple", "Apple Pie" (Case Insensitive)
     * for (const product of db.search("partialstring", { ProductName: "Apple" })) {
     * product.Stock -= 1; // You can directly modify the result
     * }
     * * @example
     * // 3. Regex match (Find emails ending in @gmail.com or @yahoo.com)
     * const emailPattern = /@(gmail|yahoo)\.com$/;
     * for (const lead of db.search("regex", { Email: emailPattern })) {
     * // ...
     * }
     */    
    *search(mode, search_for) {

        // HELPER
        let match = (value, condition, mode) => {

            // Caso seja um proxy, busca trabalhar com os dados originais
            if (value && value[this._isCdxProxy]) value = value[this._cdxProxyTarget]

            // DATAS (sem suporte)
            if (value instanceof Date) return false

            // REGEX (sem suporte)
            if (value instanceof RegExp) return false

            // ARRAY, SET, MAP (testar)
            if (value instanceof Set || value instanceof Array || value instanceof Map) {

                for (let v of value) {                              // Para cada valor
                    if (match(v, condition, mode)) return true;     // Se valor true, encerrar com true
                }
                return false;                                       // Se nada for true, encerrar com false
            }

            // OBJETO LITERAL (testar)
            if (Object.prototype.toString.call(value) === '[object Object]') {

                for (let v of Object.entries(value)) {              // Para cada valor
                    if (match(v, condition, mode)) return true;     // Se valor true, encerrar com true
                }
                return false;                                       // Se nada for true, encerrar com false
            }

            // Converte para texto
            let convertedValue = String(value)

            // Executa comparação
            switch (mode) {
                case 'fullstring': return condition === convertedValue
                case 'partialstring': return convertedValue.toLowerCase().includes(condition.toLowerCase())
                case 'regex': return condition.test(convertedValue)
            }
        }



        // ETAPA DE VALIDAÇÃO

        // Valida o modo de operação
        if (mode !== "fullstring" && mode !== "partialstring" && mode !== "regex") {
            throw Error(`Invalid mode: ${mode}`)
        }

        // Valida se search_for é um objeto
        if (Object.prototype.toString.call(search_for) !== '[object Object]') {
            throw Error(`search_for is not an literal object.`)
        }

        // Converte para Map, se válido
        search_for = new Map(Object.entries(search_for))

        // Verifica se search_for é vazio
        if (search_for.size === 0) {
            throw Error(`search_for can't be empty.`)
        }

        // Valida valores com base no tipo
        for (let value of search_for.values()) switch (mode) {

            // Caso string
            case 'fullstring':
            case 'partialstring':

                // Valida se é msm uma string
                if (typeof value !== 'string') throw Error(`${value} is not an string.`)
                break;


            // Caso Regex
            case 'regex':

                // Valida se é msm um Regex
                if (!(value instanceof RegExp)) throw Error(`${value} is not an regex.`)

                // Testa a compatibilidade com GSheets
                try { this._table.getRange(1, 1).createTextFinder(value.source).useRegularExpression(true).findNext() }
                catch (e) { throw Error(`${value.source} is not an regex compatible with Google Sheets.`) }

                break;
        }

        // Verifica se tem alguma propriedade inválida
        if (![...search_for.keys()].every(k => this._options.columns.has(k))) {
            throw Error(`Some properties does not exist in Sheet or constructor.`)
        }



        // ETAPA DE CACHING (apenas minimal)

        if (this._options.mode === 'minimal') {

            let lastRow = this._table.getLastRow();         // Obtém última linha
            if (lastRow < 2) return;                        // Para execução se não tem linhas

            let columnIndexes = this._getColumnIndexes()    // Obtém índices das colunas
            let columnRanges = new Map()                    // Prepara para receber ranges das colunas
            let queries = []                                // Prepara para buscar na planilha
            let wip = undefined                             // Prepara var para trabalhos em loop

            // Converte índice de colunas em ranges
            for (let [c, i] of columnIndexes) columnRanges.set(c, `R2C${i + 1}:R${lastRow}C${i + 1}`)

            // Para cada filtro solicitado
            for (let [columnName, filter] of search_for) switch (mode) {

                case 'partialstring':

                    // Efetua a busca
                    wip = this._table
                        .getRange(columnRanges.get(columnName))
                        .createTextFinder(filter)
                        .useRegularExpression(false)
                        .ignoreDiacritics(true)
                        .matchCase(false)
                        .matchEntireCell(false)
                        .findAll()

                    // Caso finder vazio, encerrar execução
                    if (wip.length == 0) return

                    // Armazenar finder no array de queries
                    queries.push(wip)
                    break;


                case 'fullstring':

                    // Efetua a busca
                    wip = this._table
                        .getRange(columnRanges.get(columnName))
                        .createTextFinder(filter)
                        .useRegularExpression(false)
                        .ignoreDiacritics(false)
                        .matchCase(true)
                        .matchEntireCell(true)
                        .findAll()

                    // Caso finder vazio, encerrar execução
                    if (wip.length == 0) return

                    // Armazenar finder no array de queries
                    queries.push(wip)
                    break;

                case 'regex':

                    // Efetua a busca
                    wip = this._table
                        .getRange(columnRanges.get(columnName))
                        .createTextFinder(filter.source)
                        .useRegularExpression(true)
                        .matchCase(!filter.ignoreCase)
                        .findAll()

                    // Caso finder vazio, encerrar execução
                    if (wip.length == 0) return

                    // Armazenar finder no array de queries
                    queries.push(wip)
                    break;
            }

            // Obtém o menor query
            let smallQuery = queries.reduce((small, actual) => {
                return (actual.length < small.length) ? actual : small
            })

            // Limpa var de queries para receber índices
            queries = []

            // Armazena todos os índices
            for (let row of smallQuery) queries.push(row.getRow() - 1)

            // Requisita esses dados na memória
            this._fetchNewData(queries, columnIndexes)
        }




        // ETAPA DE ITERAÇÃO

        // Para cada valor
        for (let [key, value] of this._data) {

            // Se key deletada, pular
            if (this._keys.get(key) === 'deleted') continue

            // Cria var de teste
            let filterPasses = true

            // Para cada filtro solicitado
            for (let [colName, condition] of search_for) {

                // Executa um AND com o dado
                filterPasses = match(value[colName], condition, mode)

                // Se o filtro nõo passar, parar imediatamente
                if (!filterPasses) break; 
            }

            // Se filtro não passou, pular
            if (!filterPasses) continue

            // Devolve resultado ao iterador
            yield this.get(key)
        }
    }


    /**
     * Adds or updates a record in the local memory, staging it for the next transaction commit.
     * @param {string|number} key - The Primary Key for the record.
     * @param {Object} value - The literal object containing the data to be stored.
     * @returns {Codex} The Codex instance (for method chaining).
     * @throws {Error} If `key` is not a string or number.
     * @throws {Error} If `value` is not a literal object.
     * @throws {Error} If `value` contains properties not defined in the Codex schema/columns.
     * @throws {Error} If the `key` property inside `value` differs from the `key` argument.
     */
    set(key, value) {
        try {

            // Fase 0 de validação: keys
            if (typeof key !== 'string' && typeof key !== 'number') { throw Error(`key must be a string or number`) }
            key = String(key).trim()

            // Fase 1 de validação: é um objeto válido?
            let validValue = this._typeJStoGS(value, 'clone')
            if (Object.prototype.toString.call(validValue) !== '[object Object]') {
                throw Error(`Not an literal object.`)
            }

            // Fase 2 de validação: tem alguma propriedade inválida?
            if (!Object.keys(validValue).every(k => this._options.columns.has(k))) {
                throw Error(`Some properties does not exist in Sheet or constructor.`)
            }

            // Fase 3 de validação: keys no objeto
            let objKey = validValue[this._keyColumnName]
            if (objKey === undefined) { validValue[this._keyColumnName] = key }
            if (validValue[this._keyColumnName] !== key) { throw Error(`Key property can't be different to key argument`) }

            // Insere na array
            this._data.set(key, validValue)
            this._setKeyAs(key, "new")
            return this

        } catch (e) {
            // Retorna erro.
            throw Error(`${this._log} ${e.stack}`)
        }
    }

}



