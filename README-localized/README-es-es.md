# Contenedor de Node.js para la biblioteca cliente de API de Office 365

Esta biblioteca proporciona una implementación ligera de las API de Outlook [Correo](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations), [Calendario](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations) y [Contactos](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations).

Para obtener una aplicación de ejemplo que usa esta biblioteca, consulte [Introducción a la API de Correo de Outlook y Node.js](https://github.com/jasonjoh/node-tutorial).

> La versión original de esta biblioteca era un contenedor simple diseñado para permitir el uso de las [bibliotecas cliente de API de Microsoft Office 365 para aplicaciones de Cordova](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) desde una aplicación de Node.js. Esta versión aun incluye las interfaces antiguas, pero no se realizará ningún desarrollo adicional en esa parte de la biblioteca. Se recomienda que las aplicaciones usen las nuevas interfaces en el futuro.

## Software necesario

- [Bibliotecas cliente de API de Microsoft Office 365 para aplicaciones de Cordova](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) (incluidas)
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## Colaboradores

Consulte [CONTRIBUTING.md](CONTRIBUTING.md)

## Instalación

La instalación debe realizarse mediante NPM:

```Shell
npm install node-outlook
```

## Uso

Una vez instalado, agregue lo siguiente al archivo de origen:

```js
var outlook = require("node-outlook");
```

### Interfaz nueva

La interfaz nueva se documenta en una simple [referencia](reference/node-outlook.md) (cortesía de [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown)).

#### Configuración

La configuración de la biblioteca se realiza mediante el espacio de nombres `base`:

- `outlook.base.setApiEndpoint` - Úselo para invalidar el punto de conexión de API. El valor predeterminado usa el punto de conexión de la versión 1.0 de Outlook: `https://outlook.office.com/api/v1.0`.
- `outlook.base.setAnchorMailbox` - Configure este valor como la dirección SMTP del usuario para permitir que el punto de conexión de la API pueda enrutar de forma eficaz las solicitudes de la API.
- `outlook.base.setPreferredTimeZone` - Úselo para especificar una zona horaria para que el servidor pueda devolver valores de fecha y hora en la API de Calendario.

#### Realizar llamadas de API

La biblioteca tiene un espacio de nombres para cada API.

- `outlook.mail` - La API de Correo
- `outlook.calendar` - La API de Calendario
- `outlook.contacts` - La API de Contactos

Cada espacio de nombres tiene funciones mínimas (próximamente tendrá más). Los espacios de nombres se usan de forma similar. Por ejemplo, así es cómo llama a la función de `getMessages` en el espacio de nombres `Outlook.mail`:

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

#### Realizar llamadas de API sin procesar

Si la biblioteca no implementa una función que realiza lo que necesita, puede usar el método `outlook.base.makeApiCall` para llamar a cualquier llamada de API implementada en el servidor. Consulte las implementaciones de cualquiera de los métodos que aparecen en los espacios de nombres `outlook.mail`, `outlook.calendar`, o `outlook.contacts` para ver un ejemplo de cómo usar este método.

### Interfaz anterior

> Para recordar, la interfaz anterior no se seguirá desarrollando. Se recomienda usar la interfaz nueva.

Puede crear un objeto `OutlookServices.Client` como el siguiente:

```js
var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office.com/api/v2.0',
  authHelper.getAccessTokenFn('https://outlook.office.com/', session));
```

Donde `authHelper.getAccessTokenFn` es el método de devolución de llamada que implementa para proporcionar el token de acceso de OAuth2 necesario.

## Copyright

Copyright (c) Microsoft. Todos los derechos reservados.

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, consulte [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.

----------
Conecte conmigo en Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Siga el [blog de desarrollo de Outlook](https://blogs.msdn.microsoft.com/exchangedev/)