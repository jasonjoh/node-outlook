# Wrapper do Node.js para a biblioteca de clientes de APIs do Office 365

Esta biblioteca fornece uma implementação leve das APIs eo [Mail](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations), [Calendário](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations), e [Contatos](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) do Outlook.

Para obter um aplicativo de exemplo que usa essa biblioteca, confira [introdução à API de Email do Outlook e Node.js](https://github.com/jasonjoh/node-tutorial).

> A versão original dessa biblioteca era um wrapper simples que se destina a habilitar o uso das [Bibliotecas de Clientes de APIs do Microsoft Office 365 para aplicativos Cordova](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) de um aplicativo node. js. Essa versão ainda inclui interfaces antigas, mas nenhum desenvolvimento será realizado nessa parte da biblioteca. É recomendável que os aplicativos usem as novas interfaces avançando.

## Software necessário

- [Bibliotecas de Clientes de APIs do Microsoft 365 para aplicativos Cordova ](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) (Inclusas)
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## Colaboração

Ver [CONTRIBUIÇÃO.md](CONTRIBUTING.md)

## Instalação

A instalação deve ser feita via NPM:

```Shell
npm install node-outlook
```

## Uso

Após a instalação, adicione o seguinte ao arquivo de origem:

```js
var outlook = require("node-outlook");
```

### Nova interface

A nova interface é documentada em uma referência [simples](reference/node-outlook.md) (cortesia de [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown)).

#### Configuração

A configuração da biblioteca é feita por meio da `base`do namespace:

- `outlook.base.setApiEndpoint`- Use-o para substituir o ponto de extremidade da API. O valor padrão usa o ponto de extremidade do Outlook v1.0: `https://outlook.office.com/api/v1.0`.
- `outlook.base.setAnchorMailbox`- Defina-o para o endereço SMTP do usuário afim de habilitar o ponto de extremidade da API para rotear de forma eficiente as solicitações da API.
- `outlook.base.setPreferredTimeZone` - Use isso para especificar um fuso horário para o servidor usar afim de retornar valores de data/hora na API de calendário.

#### Fazer chamadas da API

A biblioteca tem um namespace para cada API.

- `outlook.mail` - API do email
- `outlook.calendar` - API do calendário
- `Outlook. Contacts` – API de contatos

Cada namespace tem funções mínimas (mais por vir). O uso é semelhante entre os namespaces. Por exemplo, esta é a maneira de chamar a função `GetMessages` no namespace `outlook.mail`:

```js
// Specify an OData query parameters to include in the request
var queryParams = {
  '$select': 'Subject,ReceivedDateTime,From',
  '$orderby': 'ReceivedDateTime desc',
  '$top': 10
};

// Set the API endpoint to use the v2.0 endpoint
outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
// Set the anchor mailbox to the user's SMTP address
outlook.base.setAnchorMailbox(email);

outlook.mail.getMessages({token: token, odataParams: queryParams},
  function(error, result){
    if (error) {
      console.log('getMessages returned an error: ' + error);
    }
    else if (result) {
      console.log('getMessages returned ' + result.value.length + ' messages.');
      result.value.forEach(function(message) {
        console.log('  Subject: ' + message.Subject);
        var from = message.From ? message.From.EmailAddress.Name : "NONE";
        console.log('  From: ' + from);
        console.log('  Received: ' + message.ReceivedDateTime.toString());
      });
    }
  });
```

#### Fazer chamadas de API brutas

Se a biblioteca não implementar uma função que faça o que você precisa, você pode usar o método `outlook.base.makeApiCall` para requisitar qualquer chamada de API implementada no servidor. Veja as implementações de qualquer método nos namespaces `outlook.mail`, `outlook.calendar`ou `outlook.contacts` para obter um exemplo de como usar esse método.

### Interface antiga

> Como lembrete, a interface antiga não está mais sendo desenvolvida. É recomendável usar a nova interface.

Você pode criar um objeto `OutlookServices.Client`, assim:

```js
var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office.com/api/v2.0',
  authHelper.getAccessTokenFn('https://outlook.office.com/', session));
```

Onde `authHelper.getAccessTokenFn` é um método de retorno de chamada que você implementa para fornecer o token de acesso OAuth2 necessário.

## Direitos autorais

Copyright (c) Microsoft. Todos os direitos reservados.

Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.

----------
Conecte-se comigo no Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Siga o [Blog Outlook Dev](https://blogs.msdn.microsoft.com/exchangedev/)