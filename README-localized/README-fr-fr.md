# Wrapper Node.js pour la bibliothèque client des API Office 365

Cette bibliothèque offre une implémentation légère des API[Courrier](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations), [Calendrier](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations) et [Contacts](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) Outlook.

Pour consulter un exemple d’application qui utilise cette bibliothèque, voir [Prise en main de l’API de messagerie Outlook et du node.js](https://github.com/jasonjoh/node-tutorial).

> La version d’origine de cette bibliothèque était un simple Wrapper conçu pour permettre l’utilisation des [bibliothèques clientes des API Microsoft Office 365 pour les applications Cordova](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) à partir d’une application node.js. Cette version inclut tout de même les anciennes interfaces, mais aucun développement n’est effectué sur cette partie de la bibliothèque. Nous vous recommandons d’utiliser la nouvelle interface pour les applications.

## Logiciels requis

- [Bibliothèques clientes Microsoft Office 365 pour les applications Cordova](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) (incluses)
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## Contribution

Voir [CONTRIBUTING.md](CONTRIBUTING.md)

## Installation

L’installation doit être effectuée via NPM :

```Shell
npm install node-outlook
```

## Utilisation

Une fois l’installation terminée, ajoutez ce qui suit à votre fichier source :

```js
var outlook = require("node-outlook");
```

### Nouvelle interface

La nouvelle interface est documentée dans une simple  référence (courtoisie de [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown)).

#### Configuration

La configuration de la bibliothèque est effectuée via l’espace de noms de`base` :

- `outlook.base.setApiEndpoint` \- utilisez ceci pour remplacer le point de terminaison de l’API. La valeur par défaut utilise le point de terminaison Outlook v 1.0 : `https://outlook.office.com/api/v1.0`.
- `outlook.base.setAnchorMailbox`\- définissez ceci comme adresse SMTP de l’utilisateur pour activer le point de terminaison de l’API afin d’acheminer les demandes API de façon efficace.
- `outlook.base.setPreferredTimeZone`\- utilisez cette valeur pour spécifier un fuseau horaire que le serveur doit utiliser pour renvoyer des valeurs de date/heure dans l’API de calendrier.

#### Exécution d’appels API

La bibliothèque dispose d’un espace de noms pour chaque API.

- `outlook.mail` \- l’API E-Mail 
- `outlook.calendar` \- l’API Calendrier
- `outlook.contacts` \- l’API Contacts

Chaque espace de noms possède des fonctions minimales (plus à venir). L’utilisation est similaire entre les espaces de noms. Par exemple, voici comment appeler la fonction getMessages dans l’espace de noms `outlook.mail` :

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

#### Appels API bruts

Si la bibliothèque n’implémente pas une fonction pour laquelle vous avez besoin, vous pouvez utiliser la méthode `outlook.base.makeApiCall` pour appeler un appel API implémenté sur le serveur. Consultez les implémentations de toute méthode dans les espaces de noms `outlook.mail`, `outlook.calendar`ou `outlook.contacts` pour voir un exemple d’utilisation de cette méthode.

### Ancienne interface

> En guise de rappel, l’ancienne interface n’est plus développée. Nous vous recommandons d’utiliser la nouvelle interface.

Vous pouvez créer un objet`OutlookServices.Client` comme suit :

```js
var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office.com/api/v2.0',
  authHelper.getAccessTokenFn('https://outlook.office.com/', session));
```

Où `authHelper.getAccessTokenFn` est une méthode de rappel que vous implémentez pour fournir le jeton d’accès Oauth2 nécessaire.

## Copyright

Copyright (c) Microsoft. Tous droits réservés.

Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.

----------
Suivez-moi sur Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Suivez le [blog développeurs Outlook](https://blogs.msdn.microsoft.com/exchangedev/)