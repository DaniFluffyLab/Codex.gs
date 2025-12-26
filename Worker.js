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



    constructor(sheetId, tableName, keyColumnName, fullLoad = false) {

        /** @private ID da planilha origem */
        this._sheetId = sheetId

        /** @private Nome da página na planilha */
        this._tableName = tableName

        /** @private Nome no cabeçalho para coluna de keys */
        this._keyColumnName = keyColumnName

        /** @private Bool para informar se worker foi carregado completamente */
        this._fullLoad = fullLoad

        /** @private Map com todos os dados carregados */
        this._data = new Map()

        /** @private Set com todas as keys */
        this._keys = new Set()

        /** @private Set com keys alteradas */
        this._keyStatus = new Map()
    }


}


let teste = new CodexWorker()


