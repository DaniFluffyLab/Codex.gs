// CODEX.GS v0.1- https://codex.danifluffy.dev
// Library to manage spreadsheets with an ORM correlated to JS Maps.
// Created by danifluffy.dev

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
    constructor(sheetId, tableName, keyColumnName, options = {}) {



        // FASE 0 - VALIDAÇÃO DE DEPENDÊNCIAS E DEFINIÇÃO DE VARS GLOBAIS

        /**
         * Prefixo identificador utilizado em mensagens de log e erros da instância.
         * @type {string}
         * @private
         */
        this._logSign = `[ CODEX | SheetID:"${sheetId}" | Table: "${tableName}" ]\n`;




        // Verifica se a API Sheets está disponível
        if (typeof Sheets === 'undefined') {
            throw this._log(200);
        };



        /**
         * O identificador único (ID) da planilha de origem.
         * @type {string}
         * @private
         */
        this._sheetID = sheetId;
        if (typeof sheetId !== 'string') throw this._log(244, this._sheetID)



        /**
         * A instância da planilha (arquivo) do Google Sheets.
         * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
         * @private
         */
        this._sheet;                                                    // Declara variável
        try { this._sheet = SpreadsheetApp.openById(this._sheetID) }    // Carrega planilha
        catch (e) { throw this._log(300) };                             // Retorna algum erro



        /**
         * O nome da aba (página) dentro da planilha que será manipulada.
         * @type {string}
         * @private
         */
        this._tableName = tableName;
        if (typeof tableName !== 'string') throw this._log(244, this._tableName)



        /**
         * A aba específica dentro da planilha que servirá como tabela.
         * @type {GoogleAppsScript.Spreadsheet.Sheet}
         * @private
         */
        this._table;                                                        // Declara variável
        try { this._table = this._sheet.getSheetByName(this._tableName) }   // Carrega página
        catch (e) { throw this._log(301) }                                  // Retorna outros erros
        if (this._table === null) throw this._log(201, this._tableName)     // Lança erro se não houver página



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
        if (typeof keyColumnName !== 'string') throw this._log(244, this._tableName)



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
        // Testa validade do modo de operação
        if (this._options.mode !== 'minimal' && this._options.mode !== 'full') throw this._log(241, this._options.mode)
        // Testa se colunas são strings
        for (let c of this._options.columns) { if (typeof c !== 'string') throw this._log(244, c) }



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

        /**
         * Marca se instância já foi commitada.
         * @type {boolean}
         * @private
         */
        this._commited = false



        // FASE 1: INICIA O CONSTRUTOR

        try {

            let lock = this._locker('reader', 'lock')   // Busca obter trava da planilha
            if (!lock) throw this._log(302)             // Impede a execução caso não consiga

            // Carrega dados de índices de colunas
            let columnIndexes = this._getColumnIndexes()

            // Obtém dados baseados no modo de operação
            switch (this._options.mode) {

                case "minimal":
                    try {
                        let keys = this._getRowIndexesByKey(true, columnIndexes)            // Obtém keys
                        this._keys = new Map([...keys.keys()].map(k => [k, "unmodified"]))  // Adiciona keys ao Map mestre
                    }
                    catch (e) { throw this._log(500, { message: "Error to get values.", stack: e.stack }) }     // Retorna erros
                    break;

                case "full":

                    try { this._fetchNewData(true, columnIndexes) }     // Requisita dados
                    catch (e) { throw this._log(500, { message: "Error to get values.", stack: e.stack }) }   // Retorna erros
                    break;
            }
        }

        // Libera o cadeado
        finally { this._locker('reader', 'release') }

    }


    /**
     * Gerencia o controle de concorrência da planilha utilizando o padrão Reader-Writer Lock.
     * * Esta função coordena o acesso simultâneo de múltiplas instâncias do Codex, garantindo que:
     * 1. Múltiplos leitores possam acessar os dados simultaneamente (Shared Lock).
     * 2. Um escritor tenha acesso exclusivo, impedindo novas leituras e outras escritas (Exclusive Lock).
     * 3. Locks órfãos (causados por crashes de instâncias anteriores) sejam limpos automaticamente 
     * após 6 minutos (Garbage Collection).
     * * A atomicidade das operações de metadados é garantida pelo uso do `LockService` nativo, 
     * enquanto o estado do lock é persistido de forma invisível via `DeveloperMetadata`.
     * * @param {("reader"|"writer")} role - O papel da instância:
     * - 'reader': Requer acesso para leitura. Permite outros leitores, mas espera por escritores.
     * - 'writer': Requer acesso para escrita. Exige que não haja nenhum leitor ou escritor ativo.
     * @param {("lock"|"release")} action - A ação a ser executada:
     * - 'lock': Tenta adquirir a autorização de acesso (com timeout de aprox. 90 segundos).
     * - 'release': Libera o acesso e atualiza os contadores de estado.
     * * @returns {boolean} Retorna `true` se a operação foi concluída com sucesso ou `false` 
     * em caso de timeout (excesso de tentativas fracassadas).
     * @private
     */
    _locker(role, action) {

        // HELPERS

        /**
         * Armazena metadados na tabela consultada
         * @param {string} key Chave para identificação
         * @param {number|boolean|string|Date} value Valor a armazenar
         * @private
         */
        let setMetadata = (key, value) => {
            for (let metadata of this._table.createDeveloperMetadataFinder().withKey(key).find()) metadata.remove() // Limpa keys antigas
            let visibility = SpreadsheetApp.DeveloperMetadataVisibility.PROJECT                                     // Define visibilidade
            if (value instanceof Date) this._table.addDeveloperMetadata(key, value.toISOString(), visibility)       // Adiciona nova key de data
            else this._table.addDeveloperMetadata(key, String(value), visibility)                                   // Adiciona nova key regular
        }

        /**
         * Consulta metadados na tabela consultada
         * @param {string} key Chave a consulta
         * @private
         */
        let getMetadata = (key) => {

            // Obtém a primeira key
            let metadata = this._table.createDeveloperMetadataFinder().withKey(key).find()[0]

            if (!metadata) return undefined     // Retorna indefinido se a key não existe
            let value = metadata.getValue()     // Obtém valor real

            // Tenta converter como número
            if (!isNaN(value)) return Number(value)

            // Tenta converter como data
            let valueAsDate = new Date(value)
            if (!isNaN(valueAsDate.getTime())) return valueAsDate

            // Tenta converter como bool
            switch (value) {
                case "true": return true;    // Retorna bool verdadeiro
                case "false": return false;   // Retorna bool falso
            }

            // Retorna valor em string
            return value
        }


        let sentLog = false                       // Var para não repetir log
        let failedTries = 0
        let locker = LockService.getScriptLock()    // Obtém o LockService
        const k = {
            age: "Codex_lockerAge",
            readers: "Codex_readersCount",
            writing: "Codex_isWriting"
        }



        // ETAPA PARA LIMPEZA DE LOCKERS VELHOS

        while (!locker.tryLock(3000)) {         // Tenta obter cadeado do LockService
            failedTries++;                      // Soma uma tentativa fracassada
            if (failedTries > 30) return false  // Se mais que 30 tentativas, desistir
        }
        let lockerAge = getMetadata(k.age) || new Date(0)   // Obtém idade do locker
        let now = (new Date()).getTime()                    // Obtém momento atual

        // Se um locker é velho demais, resetar dados
        if (lockerAge.getTime() < (now - 360000)) {
            setMetadata(k.readers, 0)
            setMetadata(k.writing, false)
        }

        // Renova idade do locker
        setMetadata(k.age, new Date())
        locker.releaseLock()



        // ETAPA PARA LIBERAÇÃO DE USO

        if (action === "release") switch (role) {

            case "reader":
                while (!locker.tryLock(3000)) {         // Tenta obter cadeado do LockService
                    failedTries++;                      // Soma uma tentativa fracassada
                    if (failedTries > 30) return false  // Se mais que 30 tentativas, desistir
                }
                let readersCount = Math.max(0, (getMetadata(k.readers) || 1) - 1)   // Obtém leitores - 1
                setMetadata(k.readers, readersCount)                                // Grava novos leitores
                locker.releaseLock()                                                // Destranca cadeado
                return true;                                                        // Encerra execução

            case "writer":
                setMetadata(k.writing, false)   // Define estado de escrita
                locker.releaseLock()            // Destranca cadeado
                return true;                    // Encerra execução
        }



        // ETAPA PARA REQUERER AUTORIZAÇÃO DE USO

        if (action === "lock") while (true) {

            // Caso tabela esteja ocupada escrevendo [Checagem 1]
            if (getMetadata(k.writing)) {
                if (!sentLog) { this._log(100); sentLog = true }    // Avisar que está ocupada
                Utilities.sleep(3000)                                   // Espera 3 segundos
                continue;                                               // Tenta de novo
            }

            while (!locker.tryLock(3000)) {         // Tenta obter cadeado do LockService
                failedTries++;                      // Soma uma tentativa fracassada
                if (failedTries > 30) return false  // Se mais que 30 tentativas, desistir
            }

            let isWriting = getMetadata(k.writing)      // Checa se tabela está ocupada com escrita
            let readersCount = getMetadata(k.readers)   // Checa se existem leitores

            // Caso tabela esteja ocupada [Checagem 2]
            if (isWriting || (role === "writer" && readersCount !== 0)) {
                if (!sentLog) { this._log(100); sentLog = true }    // Avisar que está ocupada
                Utilities.sleep(3000)                                   // Espera 3 segundos
                locker.releaseLock()                                    // Destranca o cadeado
                continue;                                               // Tenta de novo
            }

            // Caso não esteja, para cada modo
            switch (role) {

                case ("reader"):
                    readersCount++                          // Adiciona mais um leitor
                    setMetadata(k.readers, readersCount)    // Salva essa informação
                    locker.releaseLock()                    // Destranca o cadeado
                    return true;                            // Libera execução

                case ("writer"):

                    setMetadata(k.writing, true)    // Salva a informação de tranca
                    return true;                    // Libera execução
            }
        }
    }

    /**
     * Emite log, avisos e erros da biblioteca.
     * @param {number} code - O código identificador da ocorrência.
     * @param {*} [info] - Contexto dinâmico para a mensagem. 
     * * @returns {Error|null} Retorna um objeto `Error` formatado para ser lançado via `throw`,
     * ou `null` caso a mensagem seja apenas um aviso operacional (`console.warn`).
     * @private
     */
    _log(code, info) {

        // Define prefixo
        let prefix = `${this._logSign} ${code} -`

        switch (code) {

            // Avisos operacionais
            case 100: console.warn(`${prefix} Another Codex instance is running a critical task. Awaiting...`); return null;
            case 101: console.warn(`${prefix} Too much data, activating safety mode. Consider requesting fewer columns or using minimal mode with Codex.search() to increase speed.`); return null;
            case 102: console.warn(`${prefix} Duplicate header found: ${info}`); return null;
            case 110: console.warn(`${prefix} "${info.key}": { "${info.colName}": "${info.value}" } ignored: Column "${info.colName}" doesn't exist anymore on table.`); return null;
            case 111: console.warn(`${prefix} "${info.key}": { "${info.colName}": "${info.value}" } ignored: Row "${info.key}" doesn't exist anymore on table.`); return null;
            case 120: console.log(`${prefix} Committing...`); return null;
            case 121: console.log(`${prefix} No changes detected, nothing to commit.`); return null;

            // Erros do usuário do construtor
            case 200: return Error(`${prefix} ` +
                `The "Google Sheets API" Advanced API is not enabled. To use Codex library, you need ` +
                `to activate it with identifier "Sheets".\n` +
                `Documentation: https://developers.google.com/apps-script/guides/services/advanced`)
            case 201: return Error(`${prefix} Sheet "${info}" not found.`)


            // Erros do usuário ao obter colunas
            case 210: return Error(`${prefix} Sheet "${info}" is empty (no headers found).`)
            case 211: return Error(`${prefix} Primary Key column "${info}" does not exist.`)
            case 212: return Error(`${prefix} Requested column "${info}" does not exist.`)

            // Erros do usuário ao obter linhas
            case 220: return Error(`${prefix} Unable to get indexes from keys. \n\n${info.stack}`)

            // Erros do usuário ao armazenar dados
            case 230: return Error(`${prefix} The input contains more than the maximum limit of 50,000 characters in a single cell.`)
            case 231: return Error(`${prefix} The input object contains more than 25 levels of depth.`)
            case 232: return Error(`${prefix} Maps with not-string or not-number keys are not supported.`)
            case 234: return Error(`${prefix} Key values are not editable.`)
            case 239: return Error(`${prefix} Value ${info} not supported.`)

            // Erros do usuário de uso incorreto
            case 240: return Error(`${prefix} Method ${info} is only available on mode = full. Use Codex.search() instead.`)
            case 241: return Error(`${prefix} Invalid mode: ${info}`)
            case 242: return Error(`${prefix} "${info}" is not a literal object.`)
            case 243: return Error(`${prefix} "${info}" can't be empty.`)
            case 244: return Error(`${prefix} "${info}" is not a string.`)
            case 245: return Error(`${prefix} "${info}" is not a regex.`)
            case 246: return Error(`${prefix} "${info}" is not a regex compatible with Google Sheets.`)
            case 247: return Error(`${prefix} "${info}" is not a string or a number.`)

            // Erro de usuário ao interagir com Codex commitado
            case 250: return Error(`${prefix} Unable to execute action: This instance has been already committed.`)



            // Erros de API no construtor
            case 300: return Error(`${prefix} Failed to load spreadsheet. \n\n${info.stack}`)
            case 301: return Error(`${prefix} Failed to load sheet/tab. \n\n${info.stack}`)
            case 302: return Error(`${prefix} Failed to lock sheet.`)

            // Erros da API avançada
            case 310: return Error(`${prefix} SheetsAPI Error. \n\n${info.stack}`)
            case 311: return Error(`${prefix} DriveApp Error. \n\n${info.stack}`)
            case 312: return Error(`${prefix} TextFinder Error. \n\n${info.stack}`)



            // Bug na Codex ao obter linhas
            case 400: return Error(`${prefix} Invalid "requestedKeys".`)
            case 401: return Error(`${prefix} Error locating "keyColumn": \n\n${info.stack}`)

            // Bug na Codex ao carregar dados
            case 410: return Error(`${prefix} Invalid "requestedKeysOrRows".`)
            case 411: return Error(`${prefix} "requestedKeysOrRows" must be a uniform array of strings (PKs) or numbers (Indexes).`)
            case 412: return Error(`${prefix} "requestedKeysOrRows" must be more than 0 and less than last row index.`)

            // Bug no Codex por sintaxe incorreta
            case 420: return Error(`${prefix} Invalid mode: ${info}`)
            case 421: return Error(`${prefix} Inconsistent values.`)


            // Erro desconhecido
            case 500: return Error(`${prefix} ${info.message || `Unhandled error.`} \n\n${info.stack}`)

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

        let columnIndexes = new Map();                                  // Cria var para índices das colunas
        let lastColumn = this._table.getLastColumn()                    // Obtém última coluna

        // Lança erro se aba completamente vazia
        if (lastColumn === 0) { throw this._log(210, this._tableName); }

        let columnArray = this._table.getRange(1, 1, 1, lastColumn)     // Seleciona cabeçalho
            .getValues()[0]                                             // Obtém dados

        columnArray.forEach((header, index) => {
            if (header === undefined || header === null || header === "") return    // Ignora headers vazias
            let normalizedHeader = String(header).trim()                            // Normaliza header
            if (columnIndexes.has(normalizedHeader)) this._log(102)                 // Avisa se header duplicada
            columnIndexes.set(normalizedHeader, index)                              // Insere dados dos índices no Map
        })

        // Valida a existência de uma coluna de keys
        if (!columnIndexes.has(this._keyColumnName)) { throw this._log(211, this._keyColumnName); }

        // Caso não hajam parâmetros sobre quais colunas obter, alimentar com nome de todas as colunas
        if (this._options.columns.size === 0) { this._options.columns = new Set(columnIndexes.keys()) }

        let filteredColumnIndexes = new Map()                                                   // Cria Map para colunas filtradas
        filteredColumnIndexes.set(this._keyColumnName, columnIndexes.get(this._keyColumnName))  // Garante coluna de key
        for (let columnName of this._options.columns) {                                         // Para cada coluna requisitada:
            let hasInvalid = !columnIndexes.has(columnName)                                         // Verifica a validade da colunas
            if (hasInvalid) { throw this._log(212, columnName) }                                    // Lança erro se inválido 
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

        // Caso haja bug interno, enviar erro
        if (mode === undefined) throw this._log(400)

        // Procura coluna de índices
        try { keys_colIdx = columnIndexes.get(this._keyColumnName) }            // Procura pelo nome
        catch (e) { throw this._log(401, e) }                                   // Retorna outros erros
        if (keys_colIdx == undefined) throw this._log(211, this._keyColumnName) // Se não achar coluna, lança erro

        try {

            // Modo rápido
            if (mode == "SINGLE") {
                let index = this._table.getRange(2, keys_colIdx + 1, lastRow - 1)       // Obtém range de keys
                    .createTextFinder(requestedKeys).matchEntireCell(true)              // Pesquisa na planilha
                    .findPrevious()                                                     // Obtém índice da última instância
                if (index !== null) rowIndexes.set(requestedKeys, index.getRow() - 1)   // Adiciona indice no Map
                return rowIndexes                                                       // Encerra execução
            }

            // Se planilha não está vazia
            if (lastRow >= 2) {

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

            }

            // Encerra execução
            return rowIndexes

        } catch (e) { throw this._log(220, e) }  // Retorna outros erros
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
        let SAFEMODE_getValuesByGridRange = (table, gridRange) => {
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
         * Agrupa dimensões adjacentes respeitando um limite de segurança, garantindo
         * que nenhum bloco mesclado exceda o número máximo de chaves originais.
         * * @param {Map<string, {gridRange: Object}>} gridRangesMap 
         * @param {"ROWS"|"COLUMNS"} dimension 
         * @returns {Map<string[], {gridRange: Object}>}
         * @private
         */
        let SAFEMODE_mergeGridRanges = (gridRangesMap, dimension) => {

            // Define o tamanho do lote de segurança
            const CHUNK_SIZE = dimension.includes("ROWS") ? 5 : 10
            const orientation = dimension.includes("ROWS") ? 'rows' : 'columns'

            // Converte o Map em Array para mesclagem
            let rawItems = [];
            gridRangesMap.forEach((data, key) => {
                rawItems.push({
                    startRow: data.gridRange.startRowIndex,
                    endRow: data.gridRange.endRowIndex,
                    startCol: data.gridRange.startColumnIndex,
                    endCol: data.gridRange.endColumnIndex,
                    values: [key]
                });
            });

            // Pré-ordena os itens para fatiar
            rawItems.sort((a, b) => {
                switch (orientation) {
                    case 'columns':
                        if (a.startRow !== b.startRow) return a.startRow - b.startRow
                        return a.startCol - b.startCol
                    case 'rows':
                        if (a.startCol !== b.startCol) return a.startCol - b.startCol
                        return a.startRow - b.startRow
                }
            });

            let resultMap = new Map();

            // Fatia e mescla em lotes
            for (let i = 0; i < rawItems.length; i += CHUNK_SIZE) {

                // Pega um subconjunto seguro
                let chunk = rawItems.slice(i, i + CHUNK_SIZE);

                // Mescla subconjunto
                let mergedChunk = this._mergeRequests(orientation, chunk);

                // Converte o Array mesclado em Map
                mergedChunk.forEach(item => {
                    let gridRangeObj = {
                        sheetId: this._tableID,
                        startRowIndex: item.startRow,
                        endRowIndex: item.endRow,
                        startColumnIndex: item.startCol,
                        endColumnIndex: item.endCol
                    }
                    resultMap.set(item.values, { gridRange: gridRangeObj })
                })
            }

            // Retorna resultado
            return resultMap
        }






        // Valida parâmetros
        if (!Array.isArray(requestedKeysOrRows) && requestedKeysOrRows != true) throw this._log(410)

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
                throw this._log(411)
            }

            // Alterna entre tipos
            switch (type) {

                case 'number':

                    // Testa se índices são válidos
                    if (requestedKeysOrRows.some(v => (v >= lastRow || v < 1))) {
                        throw this._log(412)
                    }

                    // Popula eles na array
                    requestedRows = [...requestedKeysOrRows]
                    break;

                case 'string':

                    // Obtém os índices com função auxiliar
                    requestedRows = [...this._getRowIndexesByKey(requestedKeysOrRows).values()]
                    break;

                default:
                    throw this._log(411)
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
            else throw this._log(310, e)    // Se não for, lança erro
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
                this._log(101)

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
                this._log(101)

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
                if (value.length > 50000) throw this._log(230)
                return value;

            case 'bigint':

                // Obtém maiores números possíveis em Integer
                const maxint = BigInt(Number.MAX_SAFE_INTEGER);
                const minint = BigInt(Number.MIN_SAFE_INTEGER);

                // Se não compatível com Integer, converter para String
                if (value > maxint || value < minint) {
                    convertedValue = String(value);
                    if (convertedValue.length > 50000) throw this._log(230)
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
                    if (convertedValue.length > 50000) throw this._log(230)
                    if (mode === 'commit') { return convertedValue }    // Caso commit, envia o regex convertido
                    if (mode === 'clone') { return new RegExp(value) }  // Caso clone, retorna novo regex
                    if (mode === 'test') { return value }               // Caso teste, retorna valor 
                }





                // ARRAY ou SET no modo commit
                if ((value instanceof Set || value instanceof Array) && mode === 'commit') {

                    convertedValue = [...value]                                                                         // Cria cópia de segurança
                    if (depth < 25) convertedValue = convertedValue.map(v => this._typeJStoGS(v, 'commit', depth + 1))  // Limpa até 25 camadas
                    if (depth == 25) throw this._log(231)                                                               // Para de converter acima de 25 camadas
                    if (depth != 0) return convertedValue                                                               // Caso em recursão, retorna valor convertido

                    // Valida tamanho da string
                    jsonValue = JSON.stringify(convertedValue)
                    if (jsonValue.length > 50000) throw this._log(230)

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
                        if (typeof key !== 'number' && typeof key !== 'string') throw this._log(232)
                    }

                    if (depth < 25) convertedValue = convertedValue.map(([k, v]) => [k, this._typeJStoGS(v, 'commit', depth + 1)])  // Limpa até 25 camadas
                    if (depth == 25) throw this._log(231)                                                                           // Para de converter acima de 25 camadas
                    if (depth != 0) return Object.fromEntries(convertedValue)                                                       // Caso em recursão, retorna valor convertido

                    // Valida tamanho da string
                    jsonValue = JSON.stringify(Object.fromEntries(convertedValue))
                    if (jsonValue.length > 50000) throw this._log(230)

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
                    if (depth == 25) throw this._log(231)                                                                           // Para de converter acima de 25 camadas
                    if (depth != 0) return Object.fromEntries(convertedValue)                                                       // Caso em recursão, retorna valor convertido
                    convertedValue = Object.fromEntries(convertedValue)                                                             // Fora da recursão, reconverte em objeto    

                    // Valida tamanho da string
                    jsonValue = JSON.stringify(convertedValue)
                    if (jsonValue.length > 50000) throw this._log(230)

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
                throw this._log(239, value)
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
                if (colName === this._keyColumnName) throw this._log(234)

                let cleanValue = this._typeJStoGS(value)        // Garante que a informação é compatível
                this._setKeyAs(key, "modified")                 // Marca o objeto como modificado
                return Reflect.set(ogObj, colName, cleanValue)  // Edita o objeto
            },

            deleteProperty: (ogObj, colName) => {

                // Impede escrita de valores na coluna de keys.
                if (colName === this._keyColumnName) throw this._log(234)

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
                        throw this._log(500, { stack: "Illegal operation." })
                }
            }
        }

        // Cria novo proxy
        let newProxy = new Proxy(value, proxyHandlers)
        this._proxies.set(value, newProxy)
        return newProxy
    }

    /**
     * Cria um backup preventivo da planilha no Google Drive na pasta Codex Backups, 
     * inserindo metadados técnicos na descrição do arquivo de backup.
     * @private
     * @throws {Error} Se o script não tiver permissões de acesso ao Drive ou se a cota de armazenamento for excedida.
     */
    _createBackup() {
        try {
            // Obtém esse arquivo via DriveApp
            let thisFile = DriveApp.getFileById(this._sheetID)
            let thisFolder = thisFile.getParents().next()
            let timeZone = Session.getScriptTimeZone()
            let now = Utilities.formatDate(new Date(), timeZone, "yyyy/MM/dd HH:mm:ss")

            // Obtém pasta de backup
            let bkpFolderIt = thisFolder.getFoldersByName("Codex Backups")
            let bkpFolder = bkpFolderIt.hasNext() ? bkpFolderIt.next() : thisFolder.createFolder("Codex Backups")
            let bkpName = `[Codex Bkp ${now} ${timeZone}] ${thisFile.getName()}`

            // Cria backup
            let bkpFile = thisFile.makeCopy(bkpName, bkpFolder)
            bkpFile.setDescription(
                "[CODEX BACKUP]\n" +
                `Date/Time: ${now} ${timeZone}\n` +
                `Original file: https://drive.google.com/open?id=${this._sheetID}\n` +
                `Table edited: ${this._tableName}`
            )
        }

        // Loga erros de API
        catch (e) { throw this._log(311, e) }
    }

    /**
     * Funde intervalos adjacentes (Horizontais ou Verticais) para requests.
     * * @param {'columns'|'rows'} orientation - Direção da mesclagem.
     * @param {Array<{startRow: number, endRow: number, startCol: number, endCol: number, values: Array}>} items 
     * @returns {Array} Array com os itens mesclados.
     */
    _mergeRequests(orientation, items) {

        // Caso não tenha itens, ignorar
        if (!items || items.length === 0) return [];

        // Verifica se é para usar values
        let hasValues = (items[0].values !== undefined);

        // Testa consistencia
        let isConsistent = items.every(item => (item.values !== undefined) === hasValues);
        if (!isConsistent) throw this._log(421)

        // Ordena os itens com base na regra de orientação
        switch (orientation) {

            // Ordena por colunas
            case 'columns':
                items.sort((a, b) => {
                    if (a.startRow !== b.startRow) return a.startRow - b.startRow;  // Ordena linhas
                    return a.startCol - b.startCol;                 // Se mesma linha, ordena colunas
                }); break;

            // Ordena por linhas
            case 'rows':
                items.sort((a, b) => {
                    if (a.startCol !== b.startCol) return a.startCol - b.startCol;  // Ordena colunas
                    return a.startRow - b.startRow;                 // Se mesma coluna, ordena linhas
                }); break;

            // Emite erro de modo inválido
            default:
                throw this._log(420, orientation)
        }

        // Prepara para processamento
        let merged = [];
        let prev = null;

        // Para cada item
        for (let curr of items) {

            // Caso não tenha um item prévio, criar e seguir para próximo loop
            if (!prev) {
                prev = { ...curr, values: curr.values ? [...curr.values] : undefined };
                continue;
            }

            // Prepara para testar vizinhança
            let isNeighbor = false;

            // Alterna comportamento entre modos
            switch (orientation) {

                // Operando por colunas
                case 'columns':
                    if (prev.startRow === curr.startRow &&  // Testa se começam na mesma linha
                        prev.endRow === curr.endRow &&      // Testa se terminam na mesma linha
                        prev.endCol === curr.startCol       // Testa vizinhança de colunas
                    ) { isNeighbor = true }
                    break;

                // Operando por linhas
                case 'rows':
                    if (prev.startCol === curr.startCol &&  // Testa se começam na mesma coluna
                        prev.endCol === curr.endCol &&      // Testa se terminam na mesma coluna
                        prev.endRow === curr.startRow       // Testa vizinhança de linhas
                    ) { isNeighbor = true }
                    break;
            }

            // Mescla se vizinhos
            if (isNeighbor) {

                // Estende a coordenada
                switch (orientation) {
                    case "columns": prev.endCol = curr.endCol; break;
                    case "rows": prev.endRow = curr.endRow; break;
                }

                // Concatena os valores (se existirem)
                if (hasValues) prev.values.push(...curr.values)
            }

            // Caso não sejam vizinhos
            else {
                merged.push(prev);                                                      // Salva o anterior
                prev = { ...curr, values: curr.values ? [...curr.values] : undefined }; // Inicia novo
            }
        }

        // Salva o último item que sobrou no buffer
        if (prev) merged.push(prev);

        // Retorna dados mesclados
        return merged;
    }



    // MÉTODOS PÚBLICOS

    /**
     * Removes all elements from the Codex instance and schedules a full cleanup 
     * of the spreadsheet on the next commit.
     */
    clear() {

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        try {
            this._wipeOnCommit = true;  // Marca planilha para exclusão
            this._keys.clear();         // Limpa histórico de mudanças
            this._data.clear();         // Limpa memória da instancia
        }

        // Retorna erros não conhecidos
        catch (e) { throw this._log(500, e) }
    }

    /**
     * Removes the specified element from the Codex instance by key.
     * Schedules the deletion of the corresponding row in the Google Sheets on the next commit.
     * * @param {string} key The key of the element to remove.
     * @returns {boolean} `true` if an element in the Codex object existed and has been removed, or `false` if the element does not exist.
     */
    delete(key) {

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

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
        }

        // Retorna erros não conhecidos
        catch (e) { throw this._log(500, e) }
    }

    /**
     * Checks if a specific key exists in the instance.
     * @param {string} key - The unique identifier (ID) to check.
     * @returns {boolean} `true` if the key exists and is active; `false` otherwise.
     */
    has(key) {

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        try {
            let keyStatus = this._keys.get(String(key).trim())  // Obtém estado
            if (keyStatus === undefined) return false           // Se não existe, false
            if (keyStatus === "deleted") return false           // Se deletado, false
            return true                                         // Retorna que existe
        }

        // Retorna erros não conhecidos
        catch (e) { throw this._log(500, e) }
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

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        // Define vars
        let keyStatus, keyLoaded, requestedData

        // Obtém dados da key solicitada
        key = String(key).trim()                        // Formata a key
        keyStatus = this._keys.get(key)                 // Obtém estado da key
        if (keyStatus === undefined) return undefined   // Se não existe, encerra
        if (keyStatus === "deleted") return undefined   // Se deletada, encerra
        keyLoaded = this._data.has(key)                 // Verifica se carregado

        // Requisita o fetch do dado se não existe
        if (!keyLoaded) this._fetchNewData([key])

        // Carrega o dado
        requestedData = this._data.get(key)     // Carrega o dado em uma var local
        if (!requestedData) return undefined    // Se não achar, retorna undefined

        // Cria proxy do objeto e retorna.
        return this._createProxy(requestedData, key)
    }

    /**
     * Returns a iterator that contains all active Primary Keys in the store.
     * * @yields {string} The next active Primary Key.
     * @returns {IterableIterator<string>} An iterable iterator of non-deleted keys.
     */
    *keys() {

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        try {

            for (const [key, status] of this._keys) {   // Para cada key
                if (status !== "deleted") yield key     // Retorna sob demanda as keys
            }

        } catch (e) { throw this._log(500, e) }     // Retorna erros não conhecidos
    }

    /**
     * Returns a iterator that contains all active values in the store. Only available on mode = full
     * * @yields {object} The next active value.
     * @returns {IterableIterator<object>} An iterable iterator of non-deleted values.
     * @throws {Error} If Codex is not in mode = full.
     * 
     */
    *values() {

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        // Rejeita uso do método sem estar no modo full.
        if (this._options.mode != "full") throw this._log(240, "Codex.values()")

        for (const [key, status] of this._keys) {           // Para cada key
            if (status !== "deleted") yield this.get(key)   // Retorna sob demanda os valores
        }
    }

    /**
     * Returns a iterator that contains all active entries in the store. Only available on mode = full
     * * @yields {string} The next active entries.
     * @returns {IterableIterator<[string, Object]>} An iterable iterator of non-deleted entries.
     * @throws {Error} If Codex is not in mode = full.
     * 
     */
    *entries() {

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        // Rejeita uso do método sem estar no modo full.
        if (this._options.mode != "full") throw this._log(240, "Codex.entries()")

        for (const [key, status] of this._keys) {                   // Para cada key
            if (status !== "deleted") yield [key, this.get(key)]    // Retorna sob demanda as chave/valores
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

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        // HELPER
        let match = (value, condition, mode) => {
            try {

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
            } catch (e) { throw this._log(500, e) }     // Retorna erros desconhecidos
        }

        // ETAPA DE VALIDAÇÃO

        // Valida o modo de operação
        if (mode !== "fullstring" && mode !== "partialstring" && mode !== "regex") {
            throw this._log(241, mode)
        }

        // Valida se search_for é um objeto
        if (Object.prototype.toString.call(search_for) !== '[object Object]') {
            throw this._log(242, "search_for")
        }

        // Converte para Map, se válido
        search_for = new Map(Object.entries(search_for))

        // Verifica se search_for é vazio
        if (search_for.size === 0) {
            throw this._log(243, "search_for")
        }

        // Valida valores com base no tipo
        for (let value of search_for.values()) switch (mode) {

            // Caso string
            case 'fullstring':
            case 'partialstring':

                // Valida se é msm uma string
                if (typeof value !== 'string') throw this._log(244, value)
                break;


            // Caso Regex
            case 'regex':

                // Valida se é msm um Regex
                if (!(value instanceof RegExp)) throw this._log(245, value)

                // Testa a compatibilidade com GSheets
                try { this._table.getRange(1, 1).createTextFinder(value.source).useRegularExpression(true).findNext() }
                catch (e) { throw this._log(246, value.source) }

                break;
        }

        // Verifica se tem alguma propriedade inválida
        for (let key of search_for.keys()) {
            if (!this._options.columns.has(key)) throw this._log(212, key)
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

            // Cria todos os queries
            try {

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

            } catch (e) { throw this._log(312, e) }  // Retorno erro

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

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        // Fase 0 de validação: keys
        if (typeof key !== 'string' && typeof key !== 'number') { throw Error(`key must be a string or number`) }
        key = String(key).trim()

        // Fase 1 de validação: é um objeto válido?
        let validValue = this._typeJStoGS(value, 'clone')
        if (Object.prototype.toString.call(validValue) !== '[object Object]') {
            throw this._log(242, key)
        }

        // Fase 2 de validação: tem alguma propriedade inválida?
        for (let key of Object.keys(validValue)) {
            if (!this._options.columns.has(key)) throw this._log(212, key)
        }

        // Fase 3 de validação: keys no objeto
        let objKey = validValue[this._keyColumnName]
        if (objKey === undefined) { validValue[this._keyColumnName] = key }
        if (validValue[this._keyColumnName] !== key) { throw this._log(234) }

        // Insere na array
        this._data.set(key, validValue)
        this._setKeyAs(key, "new")
        return this

    }

    /**
     * Synchronizes all local changes to the remote Google Sheets file and clear the memory.
     * * @param {boolean} [enableBackup=false] - If set to `true`, creates a timestamped duplicate of the sheet before applying changes.
     * @throws {Error} [Code 250] If the instance has already been committed (prevents double-submission).
     * @throws {Error} [Code 302] If the write lock cannot be acquired.
     * @throws {Error} [Code 310] If the Google Sheets API `batchUpdate` fails for any batch.
     * @returns {void}
     */
    commit(enableBackup = false) {

        // Impede execução caso commitado
        if (this._commited) throw this._log(250)

        // HELPERS

        /**
         * Converte um valor do Codex para a estrutura CellData da API avançada do Google Sheets.
         * Esta função atua como um complemento da typeJStoGS, traduzindo os tipos primitivos do
         * JavaScript e o esquema exigido pelo método `batchUpdate`. Ela normaliza os dados,
         * trata a conversão cronológica para o sistema de números de série do Sheets (base Lotus 1-2-3)
         * e define máscaras de formatação explícitas para garantir a integridade visual na planilha.
         * * @param {*} value - O valor original proveniente do Codex (String, Number, Boolean, Date ou JSON).
         * @returns {Object} Um objeto compatível com `GoogleAppsScript.Sheets.Schema.CellData`.
         * Retorna um objeto vazio `{}` para strings vazias para representar células nulas.
         * @private
         */
        let api_createCellData = (value) => {

            // Converte valor
            let convertedValue = this._typeJStoGS(value, 'commit')

            // Retorna com base no tipo
            switch (typeof convertedValue) {


                case "string":
                    if (convertedValue === "") return {}

                    return {
                        "userEnteredValue": { "stringValue": convertedValue },
                        "userEnteredFormat": { "numberFormat": { "type": "TEXT" } }
                    }


                case "number":
                    return {
                        "userEnteredValue": { "numberValue": convertedValue },
                        "userEnteredFormat": { "numberFormat": { "type": "NUMBER" } }
                    }


                case "boolean":
                    return { "userEnteredValue": { "boolValue": convertedValue } }


                case "object":

                    // Obtém o zero do Google Sheets
                    let dateZeroLotus123 = Date.UTC(1899, 11, 30)

                    // Obtém a data atual no mesmo formato
                    let convertedDateInUTC = Date.UTC(
                        convertedValue.getFullYear(),
                        convertedValue.getMonth(),
                        convertedValue.getDate(),
                        convertedValue.getHours(),
                        convertedValue.getMinutes(),
                        convertedValue.getSeconds(),
                        convertedValue.getMilliseconds()
                    );

                    // Obtém a data compatível com o Sheets
                    let convertedDate = (convertedDateInUTC - dateZeroLotus123) / 86400000

                    // Retorna valor
                    return {
                        "userEnteredValue": { "numberValue": convertedDate },
                        "userEnteredFormat": { "numberFormat": { "type": "DATE_TIME" } }
                    }
            }
        }

        /**
         * Gera e otimiza o lote de requisições para exclusão física de linhas na planilha.
         * Esta função processa as chaves marcadas como 'deleted', aplica ordenação decrescente
         * e funde intervalos vizinhos em um único comando de exclusão de intervalo.
         * * @param {Map<string, number>} rowIndexes - Mapa contendo a relação atual entre 
         * as Chaves Primárias (IDs) e seus respectivos índices físicos de linha (0-based).
         * * @returns {Object[]} Um array de objetos de requisição `deleteDimension` formatados, 
         * ordenados e otimizados, prontos para serem inseridos no `batchUpdate`.
         */
        let api_prepareDeleteRequests = (rowIndexes) => {

            // Armazena valores para mesclagem
            let rawValues = []

            // Cria os valores de exclusão
            for (let [key, status] of this._keys) {

                // Ignorar entradas não-validas
                if (status !== 'deleted') continue

                let rowIndex = rowIndexes.get(key)      // Obtém índice da linha
                if (rowIndex === undefined) continue;   // Caso essa linha já não exista, ignorar

                // Adiciona o request de exclusão
                rawValues.push({
                    startRow: rowIndex,
                    endRow: rowIndex + 1,
                    startCol: 0,    // Dummy, apenas para mesclagem
                    endCol: 0,      // Dummy, apenas para mesclagem
                })
            }

            // Mescla linhas
            let mergedItems = this._mergeRequests('rows', rawValues)

            // Ordena as requisições de forma decrescente
            mergedItems.sort((a, b) => {
                return b.startRow - a.startRow
            })

            // Remapeia valores para encerrar
            return mergedItems.map(item => ({
                deleteDimension: {
                    range: {
                        sheetId: this._tableID,
                        dimension: "ROWS",
                        startIndex: item.startRow,
                        endIndex: item.endRow
                    }
                }
            }));
        }

        /**
         * Gera o objeto de requisição para adicionar novas linhas.
         * * @param {Map<string, number>} columnIndexes - Mapa contendo a relação entre 
         * os nomes das colunas (cabeçalhos) e seus índices numéricos (0-based).
         * * @returns {Object[]} Um array contendo o objeto de comando `appendCells` formatado, 
         * ou um array vazio caso não haja novos registros.
         */
        let api_prepareAddRequests = (columnIndexes) => {

            // Prepara vars para adicionar linhas
            let requestAdd = []

            // Itera sobre os valores
            for (let [key, status] of this._keys) {

                // Ignorar entradas não-validas
                if (status !== 'new') continue

                let data = this._data.get(key)                      // Obtém dados
                let values = new Array(columnIndexes.size).fill({}) // Prepara para receber os valores

                // Para cada coluna
                for (let [colName, value] of Object.entries(data)) {
                    value = api_createCellData(value)       // Converte dado para formato da API
                    let index = columnIndexes.get(colName)  // Obtém índice da coluna

                    // Caso falhe em obter índices, alertar e ignorar
                    if (index === undefined) { this._log(110, { key: key, colName: colName, value: value }); continue; }

                    // Adiciona item na array
                    values[index] = value
                }
                requestAdd.push({ "values": values })   // Armazena valor no array
            }

            // Retorna undefined se está vazio
            if (requestAdd.length === 0) return []
            else return [{
                appendCells: {
                    sheetId: this._tableID,
                    rows: requestAdd,
                    fields: "userEnteredValue, userEnteredFormat"
                }
            }]
        }

        /**
         * Gera o lote de requisições para atualização de células.
         * @param {Map<string, number>} columnIndexes - Índices das colunas.
         * @param {Map<string, number>} rowIndexes - Índices das linhas.
         * @returns {Object[]} Array de requisições `updateCells`.
         */
        let api_prepareUpdateRequests = (columnIndexes, rowIndexes) => {

            // Prepara para receber valores
            let rawData = []

            // Para cada entrada
            for (let [key, status] of this._keys) {

                if (status !== 'modified') continue     // Ignorar entradas não-validas
                let data = this._data.get(key)          // Obtém dados

                // Para cada coluna
                for (let [colName, value] of Object.entries(data)) {
                    value = api_createCellData(value)           // Converte dado para formato da API
                    let indexCol = columnIndexes.get(colName)   // Obtém índice da coluna
                    let indexRow = rowIndexes.get(key)          // Obtém índice da linha

                    // Caso falhe em obter índices, alertar e ignorar
                    if (indexCol === undefined) { this._log(110, { key: key, colName: colName, value: value }); continue; }
                    if (indexRow === undefined) { this._log(111, { key: key, colName: colName, value: value }); continue; }

                    // Adiciona estrutura de dados
                    rawData.push({
                        startRow: indexRow,
                        endRow: indexRow + 1,
                        startCol: indexCol,
                        endCol: indexCol + 1,
                        values: [value]
                    })
                }
            }

            // Mescla colunas
            let mergedColumns = this._mergeRequests('columns', rawData)

            // Prepara linha para API
            let rowReadyForAPI = mergedColumns.map(data => ({
                ...data,
                values: [{ values: data.values }]
            }))

            // Mescla linhas
            let mergedRows = this._mergeRequests('rows', rowReadyForAPI)

            // Retorna valor construído para API
            return mergedRows.map(item => ({
                updateCells: {
                    range: {
                        sheetId: this._tableID,
                        startRowIndex: item.startRow,
                        endRowIndex: item.endRow,
                        startColumnIndex: item.startCol,
                        endColumnIndex: item.endCol
                    },
                    rows: item.values,
                    fields: "userEnteredValue,userEnteredFormat"
                }
            }));
        }

        /**
         * Divide um array de requisições em lotes menores para respeitar
         * o limite de memória e payload da API do Google Sheets.
         * * @param {Object[]} payLoad - Array contendo todas as requisições geradas.
         * @returns {Object[][]} Um array de arrays (lotes de requisições).
         */
        let api_splitPayload = (payLoad) => {

            let receivedPayloads = [...payLoad]
            let allPayloads = []                // Buffer de todos os payloads
            let prevPayload = []                // Buffer de apenas o payload atual
            let lengthPrevPayload = 0           // Avalia o tamanho do payload
            let limitLengthPayload = 5000000    // Limita o payload a 5MB

            for (let curr of receivedPayloads) {

                // Obtém tamanho
                let json = JSON.stringify(curr)
                let potentialLength = lengthPrevPayload + json.length

                // Se tamanho vai exceder
                if (potentialLength > limitLengthPayload) {
                    allPayloads.push(prevPayload)   // Armazena anterior ao buffer de retorno
                    prevPayload = []                // Reseta buffer
                    lengthPrevPayload = 0           // Zera counter de tamanho
                }

                prevPayload.push(curr)              // Adiciona payLoad ao buffer
                lengthPrevPayload += json.length    // Adiciona tamanho ao counter
            }

            // Garante buffer limpo
            if (prevPayload.length !== 0) allPayloads.push(prevPayload)

            // Retorna payloads divididos
            return allPayloads
        }






        let lock = this._locker('writer', 'lock')   // Busca obter trava da planilha
        if (!lock) throw this._log(302)             // Impede a execução caso não consiga


        try {
            // Cria backup caso requisitado
            if (enableBackup) this._createBackup()



            // ETAPA DE MONTAGEM DO OBJETO DE REQUEST

            // Limpa valores que não irão ser commitados
            for (let [k, v] of this._keys) if (v === 'unmodified') {
                this._keys.delete(k)
                this._data.delete(k)
            }

            // Monta variáveis
            let rowIndexes = this._getRowIndexesByKey([...this._keys.keys()])
            let columnIndexes = this._getColumnIndexes()
            let lastRow = this._table.getLastRow()

            // Prepara para obter requests
            let requestDelete = []
            let requestAdd = []
            let requestUpdate = []

            // Caso se queira limpar toda a planilha,
            // adiciona um request de exclusão total
            if (this._wipeOnCommit) requestDelete.unshift({
                deleteDimension: {
                    range: {
                        sheetId: this._tableID,
                        dimension: "ROWS",
                        startIndex: 1,
                        endIndex: lastRow
                    }
                }
            })

            // Caso não tenha Wipe, adiciona requests normais de update e exclusão
            if (!this._wipeOnCommit) {
                requestDelete = api_prepareDeleteRequests(rowIndexes)
                requestUpdate = api_prepareUpdateRequests(columnIndexes, rowIndexes)

            }

            // Adiciona requests de novas linhas em ambos os casos
            requestAdd = api_prepareAddRequests(columnIndexes)

            // Junta todas as requests na ordem que precisam ser chamadas
            let fullRequests = [...requestUpdate, ...requestDelete, ...requestAdd]

            // Se há requisições
            if (fullRequests.length !== 0) {

                let splittedPayloads = api_splitPayload(fullRequests)   // Divide o payload em chunks seguras
                this._log(120)                                          // Notifica o commit

                // Executa os requests à API
                for (let [index, payload] of splittedPayloads.entries()) {

                    try { Sheets.Spreadsheets.batchUpdate({ requests: payload }, this._sheetID) }
                    catch (e) { throw this._log(310, e) }

                }
            }
            // Loga aviso de commit vazio
            else { this._log(121) }


            // Limpa memória
            this._data.clear()
            this._keys.clear()
            this._commited = true

        }

        // Libera o cadeado
        finally { this._locker('writer', 'release') }


    }

}



