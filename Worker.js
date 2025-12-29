/** 
 * Este projeto busca intermediar as comunicações entre o JS e o Google Sheets
 * para otimizar a leitura e escrita do Sheets como uma database. O objetivo é
 * utilizar ua sintaxe análoga ao do processamento dos Maps do JS, mas com 
 * comandos de iniciar e encerrar uma transação.
 */

class CodexWorker {

    /**
    * Define a new status for the key.
    * @param {string} key The key of the entry to update status.
    * @param {string} newState The status of the key: "new", "modified", "deleted".
    * @private
    */
    _markAs(key, newState) {

        // Verifica se o estado solicitado é válido
        const validStates = ["new", "modified", "deleted"];
        if (!validStates.includes(newState)) throw Error("Not a valid state.")

        // Obtém estado atual da chave
        let actualState = this._keyStatus.get(key)

        // Caso não possua estado definido, adicionar
        if (actualState == undefined) {
            this._keyStatus.set(key, newState)
            return undefined
        }

        // Age conforme o estado atual
        switch (actualState) {

            case "new":
                if (newState == "deleted") this._keyStatus.delete(key);
                // if (newState == "modified") deve manter o estado como "new" 
                break;

            case "modified":
                // if (newState == "add") não é uma operação válida
                if (newState == "deleted") this._keyStatus.set(key, "deleted");
                break;
            case "deleted":
                // Adicionar uma chave deletada a reativa modificando o valor.
                if (newState == "new") this._keyStatus.set(key, "modified");
            // if (newState == "modified") não reabilita a chave. 
        }
    }

    /**
     * Commit all the alterations and delete this worker.
     */
    commit() { }

    /**
     * Generate a new random unique hex value for use as key in this data.
     * @returns {string} A new random unique key 
     */
    genKey() { }

    /**
     * Returns a boolean value indicating whether an entry with the specified key
     * exists in this data or not.
     * 
     * @param {string} key The key of the entry to test for presence.
     * @returns {bool} Returns "true" if an entry with the specified key exists
     * in the data; otherwise "false".
     */
    has(key) { }

    /**
     * Returns the value corresponding to the key in this data, or undefined
     * if there is none.
     * 
     * @param {string} key The key of the value to return from the data.
     * 
     * @returns The value associated with the specified key in the data.
     * If the key can't be founded, "undefined" is returned.
     */
    get(key) { }

    /**
     * Adds a new entry with a specified key and value to this data, or updates an
     * existing entry if the key already exists.
     * 
     * @param {string} key The key of the entry to add to or modify within the data.
     * @param {*} value The value of the entry to add or modify within the data.
     */
    set(key, value) { }

    /** 
     * Removes the entry specified by the key from this data.
     * 
     * @param {string} key The key of the entry to remove from the data. 
     * 
     * @returns {bool} "true" if an entry in the data has been removed succesfully.
     * "false" if the key is not found in the data.
     */
    delete(key) { }

    /** 
     * Removes all values from data.
     */
    clear() { }

    /**
     * Returns a new Iterator object that contains the keys for each element in the 
     * data in insertion order.
     * @returns A new iterable iterator object.
     */
    keys() { }



    constructor(sheetId, tableName, keyColumnName, options) {

        // FASE 0 - VALIDAÇÃO DE DEPENDÊNCIAS E DEFINIÇÃO DE VARS GLOBAIS
        if (typeof Sheets === 'undefined') {
            throw new Error(
                `[Codex] O Serviço Avançado "Google Sheets API" não está ativo. ` +
                `Para utilizar a biblioteca Codex, ative-o com o identificador "Sheets" em: ` +
                `Editor do GAS > Serviços (+) > Google Sheets API. ` +
                `Documentação: https://developers.google.com/apps-script/guides/services/advanced`
            );
        }

        /** @private Planilha origem */
        this._sheet;

        /** @private Página na planilha */
        this._table;

        /** @private Nome no cabeçalho para coluna de keys */
        this._keyColumnName = keyColumnName

        /** @private Objeto de configurações */
        this._options = {
            mode: "minimal", // minimal - filtered - full
            filters: [], // Apenas caso mode = filtered
            columns: [],
            ...options
        }

        /** @private Set com todas as keys */
        this._allkeys;

        /** @private Set com keys carregadas */
        this._loadedkeys;
        
        /** @private Map com todos os dados carregados */
        this._data = new Map();

        /** @private Set com keys alteradas */
        this._keyStatus = new Map();

        let log = `[Codex] SheetID:"${sheetId}"; Table: "${tableName}"`




        // FASE 1 - CARREGA A API DO GOOGLE

        try { this._sheet = SpreadsheetApp.openById(sheetId) }                          // Carrega planilha
        catch (e) { throw Error(`${log} - Erro ao carregar a planilha: ${e.stack}`) }   // Retorna erro

        try { this._table = this._sheet.getSheetByName(tableName) }                 // Carrega página
        catch (e) { throw Error(`${log} - Erro ao carregar a página: ${e.stack}`) } // Retorna outros erros

        // Armazena tamanho da planilha
        let shDims = { rows: this._table.getLastRow(), cols: this._table.getLastColumn() }


        // FASE 2 - CARREGA COLUNAS

        let colsNames;   // Cria Var para nomes das colunas
        try {
            let lastCol = shDims.cols                           // Obtém última coluna
            if (lastCol == 0) throw Error("Não há colunas")     // Lança erro se sem colunas
            colsNames = this._table.getRange(1, 1, 1, lastCol)  // Seleciona cabeçalho
                .getValues()[0]                                 // Obtém dados
        }
        catch (e) { throw Error(`${log} - Erro ao obter dados das colunas: ${e.stack}`) }  // Retorna erros




        // FASE 3 - CARREGA KEYS

        let keys_colIdx;                                                                    // Cria var para guardar índice
        try { keys_colIdx = colsNames.indexOf(this._keyColumnName) }                        // Procura pelo nome
        catch (e) { throw Error(`${log} - Erro ao procurar pela keyColumn: ${e.stack}`) }   // Retorna outros erros
        if (keys_colIdx == -1) throw Error(`${log} - keyColumn não encontrada`)             // Se não achar coluna, lança erro

        this._allkeys = new Set()      // Cria Set para guardar keys
        try {
            let lastRow = shDims.rows                                       // Obtém última linha
            if (lastRow >= 2) {                                             // Se planilha não está vazia
                let keys_rawValues = this._table                            // Acessa tabela
                    .getRange(2, keys_colIdx + 1, lastRow - 1, 1)           // Seleciona coluna de keys
                    .getValues()                                            // Obtém matriz
                for (let [k] of keys_rawValues) {                           // Para cada key
                    if (k == "" || k == null || k == undefined) continue    // Ignora keys vazias
                    this._allkeys.add(String(k).trim())                     // Adiciona Key ao Set mestre
                }
            }
        }
        catch (e) { throw Error(`${log} - Erro ao obter keys: ${e.stack}`) }  // Retorna outros erros




        // FASE 4 - OBTÉM ÍNDICES DE COLUNAS  

        let colsToAnalyze = new Set([keyColumnName, ...colsNames])          // Cria variável para array de colunas a analizar
        let colsRanges = new Map()                                          // Cria Map para índices das colunas
        let nameToIndex = new Map(colsNames.map((name, i) => [name, i]));   // Map temporário para armazenar índices originais

        // Caso hajam parâmetros sobre quais colunas obter
        if (this._options.columns.length != 0) {
            let hasInvalid = this._options.columns.some(item => !colsToAnalyze.has(item));      // Verifica a validade das colunas
            if (hasInvalid) throw Error(`${log} - Foram solicitadas colunas inexitentes.`);     // Lança erro se inválido 
            colsToAnalyze = new Set([keyColumnName, ...this._options.columns])                  // Prepara para analisar as colunas requisitadas
        }

        // Para cada coluna a analisar
        colsToAnalyze.forEach(col => {
            let index = nameToIndex.get(col);                                       // Obtém índice
            if (colsRanges.has(col)) return;                                        // Caso coluna ja exista, ignorar
            colsRanges.set(col, `'${tableName}'!R2C${index + 1}:C${index + 1}`);    // Armazena range
        });



        // FASE 5 - OBTEM DADOS

        switch (this._options.mode) {

            // Modo mínimo
            case "minimal":
                break;

            // Modo filtrado
            case "filtered":
                // ainda não sei
                break;

            case "full":

                // Caso não tenha dados, ignorar
                if (this._allkeys.size == 0) break;

                // Armazena os ranges em array
                let allRanges = Array.from(colsRanges.values())
                let allCols = Array.from(colsRanges.keys())

                // Solicita a API os dados
                let apiResponse = Sheets.Spreadsheets.Values.batchGet(sheetId, {
                    ranges: allRanges,
                    valueRenderOption: "UNFORMATTED_VALUE",
                    dateTimeRenderOption: "FORMATTED_STRING",
                    majorDimension: "COLUMNS",
                    fields: "valueRanges/values"
                });

                // Valida os dados recebidos
                if (!apiResponse.valueRanges) {
                    throw Error(`${log} - A API não retornou dados para os intervalos solicitados.`);
                }

                // Armazena os dados recebidos
                let requestedData = apiResponse.valueRanges.map(range => range.values || []);

                // Para cada linha
                for (let [rowInd, id] of requestedData[0][0].entries()) {

                    let obj = {}    // Cria objeto

                    // Para cada coluna
                    for (let [colInd, colName] of allCols.entries()) {
                        // Adiciona valor no objeto
                        let colValues = requestedData[colInd][0] || []; 
                        obj[colName] = colValues[rowInd] ?? undefined
                    }

                    // Adiciona linha no Map
                    this._data.set(String(id).trim(), obj)

                }

                break;

        }


    }


}



