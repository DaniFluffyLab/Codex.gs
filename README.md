# Codex.gs

> [!WARNING]
> A biblioteca Codex está tecnicamente funcional, mas este projeto ainda está sendo ativamente testado e corrigido bugs, portanto **cuidado ao utilizar em sistemas em produção**. Caso você encontre algum bug, fique à vontade para abrir uma issue!

**Codex.gs** é uma biblioteca do Google Apps Script (GAS) para manipular uma planilha do Google Sheets de forma análoga a um [Map do Javascript](https://developer.mozilla.org/pt-BR/docs/Web/JavaScript/Reference/Global_Objects/Map). Ao criar um objeto Codex, você pode manipular sua planilha por pares de chave-objeto, que representam as linhas da planilha, e onde cada propriedade do objeto representa uma coluna.

Por exemplo, imagine uma planilha de ID _"abcABC1234"_ com uma aba chamada "_Users_":

| username       | fullname      | lastLogin        | id  | userconfigs                                                   |
|----------------|---------------|------------------|-----|---------------------------------------------------------------|
| danifluffycat  | Dani ^^       | 06/02/2026 17:55 | 111 | {"level": 5,"followers": [586,469],"lastView": "homepage"}    |
| ladradebatatas | Laurinha      | 22/11/2025 07:20 | 586 | {"level": 10,"followers": [111,469],"lastView": "followList"} |
| patocarambola  | PatoCarambola | 12/01/2026 12:14 | 469 | {"level": 7,"followers": [111,586],"lastView": "community"}   |

Digitando o código abaixo no GAS:

```javascript
let users = new Codex("abcABC1234", "Users", "id")
let activeUser = users.get("111")
console.log(activeUser)
```

Você leria no seu console:

```javascript
{
    username: "danifluffycat",
    fullname: "Dani ^^",
    lastLogin: [object Date],  // Convertida para um objeto nativo 
    id: 111,
    userconfigs: {  // JSON convertido para objeto literal
        level: 5,
        followers: [586, 469],
        lastView: "homepage" 
    }
}
```

E poderia editar os dados como um objeto literal comum:

```javascript
activeuser.lastLogin = new Date()
activeuser.userconfigs.level++

// Solicita salvamento 
users.commit()
```

## Sobre este projeto
A Classe _SpreadsheetApp_ do GAS, é super funcional para um script rápido. Entretanto, conforme o projeto vai se tornando mais complexo, a classe se torna muito verborrágica: você precisa mapear onde armazenar seus dados navegando em vetores, validar dados que serão gravados nas células para o sheets não quebrar, adicionar diversas linhas com novos dados, etc.  
Além disso, qualquer edição realizada por ela é aplicada em tempo real na planilha. Isso a torna muito lenta, já que cada edição cria uma requisição aos servidores do Google; e se o script quebra por qualquer motivo durante um loop de search-replace, você está com uma planilha parcialmente alterada e _potencialmente quebrada_. Atomicidade de alterações só é possível via API Avançada do Sheets.

A SpreadsheetApp é uma boa API para uma planilha, mas não me é um CRUD prático.

Sendo assim, os objetivos da Codex são:

- **Diminuir a lentidão nos processos de leitura e escrita:**  
Toda a Codex é pensada para diminuir chamadas à API do Google, porque essas tendem a ser o ponto de gargalo de qualquer projeto do GAS.

- **Simplificar transações de Criação, Leitura, Escrita e Exclusão:**  
Imitando o comportamento de um Map de objetos, adicionar, ler e escrever dados se torna tão simples quanto ler e escrever as propriedades de um objeto literal que você busca em um Map.

- **Facilitar busca de dados com funções de Query:**  
Para iterar entre todos os usuários inativos, por exemplo, basta iterar sobre os valores de um `Codex.search({active: true})`

- **Otimizar operações na memória do GAS:**  
A VM do Google Apps Script não consegue lidar com muitos dados ao mesmo tempo na memória.  
Com o modo _'minimal'_, obtém-se os dados apenas quando iterados, e pode-se obter apenas um tipo de dado específico com uma query em `Codex.search()`.  
Além disso, o modo _'full'_ permite obter apenas as colunas relevantes para o contexto em que a Codex está sendo instanciada.  
Ao salvar os dados, toda informação é desindexada para o Garbage Collector do JS limpar a memória da execução e evitar travamentos.

- **Compatibilizar dados não-nativos com Google Sheets**:  
_BigInts_ se tornam _Integer_ ou _Strings_, _RegEx_ se tornam _Strings_, _Arrays_ e _Objetos literais_ se tornam _Strings JSON_, e vice-versa.  

> [!NOTE]
> Teoricamente Maps e Sets sõo compatíveis, mas como JSONs não suportam essas estruturas de dados nativamente, a Codex as converte no commit para Objetos literais e Arrays, respectivamente

- **Atomizar transações de dados:**  
Nada é alterado na planilha até que se execute `Codex.commit()`. Caso o script quebre durante a transação, ou caso aconteça algo incomum nos servidores do Google, seus dados estão seguros.  
Além disso, duas instâncias de _Codex_ podem ler uma planilha ao mesmo tempo, mas _não commitam_ ao mesmo tempo, para evitar que uma instância conflite com a outra.



## To-do
- [x] _~~Implantar todos os requests de API~~_
- [x] _~~Implantar todos os métodos análogos ao Map~~_
- [ ] (atual) Testar e validar todas as funções e todos os modos de operação
- [ ] Testar atomicidade do commit
- [ ] Modificar modo minimal para não haver necessidade de request de keys na construção da instância
- [ ] Lançar v1
- [ ] Escrever documentação e um readme mais interessante pro projeto

