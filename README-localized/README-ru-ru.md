# Оболочка Node.js для клиентской библиотеки API Office 365

Эта библиотека предоставляет облегченную реализацию Outlook [mail](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations), [календарей](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations) и [контактов](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) API.

Образец приложения, в котором используется эта библиотека, см. в статье [начало работы с почтовыми API Outlook и узла Node.js](https://github.com/jasonjoh/node-tutorial).

> Первоначальная версия этой библиотеки представляла собой простую оболочку, предназначенную для использования клиентских библиотек API-интерфейсов [Microsoft Office 365 для приложений Cordova из приложения](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) Node.js. Эта версия все еще включает в себя старые интерфейсы, но дальнейшая разработка этой части этой библиотеки не ведется. Рекомендуется, чтобы приложения использовали новые интерфейсы для продвижения вперед.

## Необходимое программное обеспечение

- [Клиентские библиотеки Microsoft Office 365 API для приложений Cordova](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) (в комплекте)
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## Участие

Дополнительные сведения см. в статье [CONTRIBUTING.md](CONTRIBUTING.md)

## Установка

Установку следует выполнять с помощью NPM:

```Shell
npm install node-outlook
```

## Использование

После установки добавьте следующее в ваш исходный файл:

```js
var outlook = require("node-outlook");
```

### Новый интерфейс

Новый интерфейс задокументирован в простой [справке](reference/node-outlook.md) (любезно предоставлено [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown)).

#### Настройка

Настройка библиотеки осуществляется через `базовое` пространство имен:

- `outlook.base.setApiEndpoint` - используйте это для переопределения конечной точки API. Значение по умолчанию использует конечную точку Outlook v1.0: `https://outlook.office.com/api/v1.0`.
- `outlook.base.setAnchorMailbox` - установите для этого SMTP-адрес пользователя, чтобы конечная точка API могла эффективно направлять запросы API.
- `outlook.base.setPreferredTimeZone` - используйте этот параметр, чтобы указать часовой пояс, который сервер будет использовать для возврата значений даты / времени в API Календаря.

#### Выполнение вызовов API

В библиотеке есть пространство имен для каждого API.

- `Outlook. mail` — API почты
- `Outlook. календарь` — API календаря.
- `Outlook. Контакты` — API контактов.

Каждое пространство имен имеет минимальные функции (еще не все). Использование аналогично между пространствами имен. Например, вот как вы вызываете функцию `getMessages` в пространстве имен `outlook.mail`:

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

#### Создание необработанных вызовов API

Если библиотека не реализует функцию, которая делает то, что вам нужно, вы можете использовать метод `outlook.base.makeApiCall` для вызова любого вызова API, реализованного на сервере. Посмотрите реализации любых методов в пространствах имен `outlook.mail`, `outlook.calendar` или `outlook.contacts` для примера использования этого метода.

### Старый интерфейс

> Напомним, что старый интерфейс больше не разрабатывается. Рекомендуется использовать новый интерфейс.

Вы можете создать объект `OutlookServices.Client` следующим образом:

```js
var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office.com/api/v2.0',
  authHelper.getAccessTokenFn('https://outlook.office.com/', session));
```

Где authHelper.`getAccessTokenFn` - это метод обратного вызова, который вы реализуете для предоставления необходимого токена доступа OAuth2.

## Авторские права

(c) Корпорация Майкрософт (Microsoft Corporation). Все права защищены.

Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).

----------
Работайте со мной на сайте Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Следуйте [блогу разработчиков Outlook](https://blogs.msdn.microsoft.com/exchangedev/)