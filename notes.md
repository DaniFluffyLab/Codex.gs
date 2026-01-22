### **Resumo Estratégico: Codex Proxy Driver**

**Objetivo:** Eliminar a necessidade de a usuária chamar `.set()` manualmente sempre que alterar uma propriedade de um objeto obtido por `.get()`, além de garantir que apenas dados compatíveis com planilhas sejam armazenados.

#### **1\. O Interceptador Principal (Trap `set`)**

Ao executar o método `get(key)`, a `Codex` não retorna o objeto original do `this._data`, mas sim um `new Proxy(target, handler)`.

* **Ação:** Sempre que `objeto.propriedade = valor` é executado, o Proxy intercepta a operação.  
* **Rastreamento:** Dispara automaticamente o método interno `_markAs(key, "modified")`.  
* **Sanitização (Gatekeeper):** Antes de salvar no Map, o Proxy valida o tipo:  
  * Rejeita `functions` e `symbols`.  
  * Converte `Objects/Arrays` complexos em `JSON strings`.  
  * Preserva `Numbers`, `Strings`, `Booleans` e `Dates`.

#### **2\. O Problema da Mutação de Objetos (Ex: `Date`)**

Objetos como `Date` são alterados por métodos (ex: `.setFullYear()`) e não por atribuição direta. Isso "bypassaria" o interceptador `set` tradicional.

**Solução: Interceptação de Métodos (Nested Proxy)**

* No interceptador `get` do Proxy principal, se a propriedade acessada for uma `instanceof Date`, a `Codex` retorna um **segundo Proxy** (Sub-Proxy) para aquele objeto de data.  
* Esse Sub-Proxy monitora o acesso a funções. Se o nome da função começar com `"set"` (como `.setHours`), o Proxy:  
  1. Marca a chave pai como `"modified"`.  
  2. Executa a função original da data.

#### **3\. Benefícios para a Arquitetura**

* **Invisibilidade:** A usuária interage com os dados como se fossem objetos JS puros, enquanto a `Codex` gerencia o estado "dirty" nos bastidores.  
* **Proteção do Driver:** O `commit()` nunca falha por tipos de dados inválidos, pois o Proxy barrou ou converteu a entrada na origem.  
* **Performance:** Evita clonagens profundas (`deep clone`) de grandes volumes de dados, processando a lógica apenas no momento da escrita.

---

**Nota para implementação futura:** Ao implementar a interceptação de métodos em `Date`, lembre-se de usar `.bind(target)` ou `.apply()` para manter o contexto do objeto original, evitando o erro de *Illegal Invocation*.









# Adicionar futuramente:

```javascript
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
     * Returns a new Iterator object that contains the keys for each element in the 
     * data in insertion order.
     * @returns A new iterable iterator object.
     */
    keys() { }









```