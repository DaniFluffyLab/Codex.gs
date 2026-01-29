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




---

# 🔍 Codex Method: `.search()` (Hybrid Implementation)

## 📋 Descrição Geral

O método `.search()` transforma o Codex de um simples armazenamento Chave-Valor em um motor de busca consultável. Ele utiliza uma estratégia **Híbrida de Duas Camadas**:

1. **Broad Filter (Camada Planilha):** Triagem de alta performance usando a busca nativa do Google para reduzir o conjunto de dados.
2. **Fine Filter (Camada Memória):** Refinamento lógico rigoroso para garantir a integridade dos resultados (AND lógico) e suporte a tipos complexos.

---

## 🛠 Especificação da API

### Assinatura

`*search(mode, search_for)`

### Parâmetros

* **`mode`**: `fullstring` | `partialstring` | `regex`
* **`search_for`**: `{ [colName: string]: any }` (Ex: `{ nome: "Dani", codServ: 123 }`)

### Retorno

* **`IterableIterator<[string, object]>`**: Um Generator que emite pares `[key, Proxy]`, idêntico ao `.entries()`.

---

## ⚙️ Fluxo de Execução (Modo Minimal / Lazy)

### 1. Triagem de Cardinalidade (Otimização)

Em vez de iterar sobre todos os critérios, o sistema identifica o **caminho de menor resistência**:

1. Para cada propriedade em `search_for`:
* Configura um `Sheet.createTextFinder()` restrito à coluna específica (`withRange`).
* Configura o finder conforme o `mode` (Regex, Case Sensitive, etc).
* Executa `.findAll()` e armazena o array de `Ranges`.


2. **Short-Circuit:** Se qualquer critério retornar 0 resultados, o método encerra imediatamente (busca AND com zero ocorrências é sempre vazia).
3. **Vencedor:** Identifica qual critério retornou o **menor** número de `Ranges`.

### 2. Sincronização de Cache (`Hydration`)

1. Itera apenas sobre os resultados do "Vencedor".
2. Extrai as Chaves Primárias (`PKs`) dessas linhas.
3. Dispara o método interno `_fetchNewData(keysToLoad)` para garantir que esses registros (e apenas eles) estejam na memória RAM.

---

## ⚙️ Fluxo de Execução (Pós-Hydration / Modo Full)

### 3. Filtro Fino (AND Lógico)

Com os candidatos carregados na memória, o sistema executa a validação final:

1. Para cada registro candidato na memória:
* Valida se **todos** os critérios de `search_for` são atendidos (Lógica AND).
* Aplica a comparação baseada no `mode`:
* `fullstring`: Igualdade estrita (após conversão de tipo).
* `partialstring`: `.includes()` ou similar.
* `regex`: Teste de expressão regular contra o valor (ou stringify de objetos/arrays).




2. **Yield:** Se aprovado, emite o par `[key, _createProxy(data, key)]`.

---

## 💡 Observações Técnicas Importantes

> * **Performance:** A busca por cardinalidade evita o estouro de cota da API do Google ao não carregar linhas desnecessárias para a RAM.
> * **Falsos Positivos:** O `TextFinder` pode encontrar termos dentro de strings JSON (ex: encontrar `"SP"` dentro de uma URL). O "Filtro Fino" em memória é o que garante que o resultado seja 100% preciso.
> * **Reatividade:** Como os resultados são emitidos via `_createProxy`, o usuário pode editar os resultados da busca diretamente, e as mudanças serão capturadas para o próximo `commit()`.
> 
> 






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