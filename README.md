# Node.js Wrapper for Office 365 APIs Client Library #

This library is a simple wrapper intended to enable use of the [Microsoft Office 365 APIs Client Libraries for Cordova Applications](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) from a Node.js app.

## Tricks used ##

- To load in the exchange.js file, which is not a Node module, I used the method described [here](http://stackoverflow.com/questions/5171213/load-vanilla-javascript-libraries-into-node-js).
- The exchange.js file uses the AJAX XMLHttpRequest object for sending requests. To get it working quickly without having to modify exchange.js, I used the node-XMLHttpRequest module.

## Required software ##

- [Microsoft Office 365 APIs Client Libraries for Cordova Applications](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) (Included)
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## Installation ##

Installing should be done via NPM:

    npm install node-outlook

## Usage ##

Once installed, add the following to your source file:

    var outlook = require("node-outlook");

You can then create an `OutlookServices.Client` object like so:

    var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office365.com/api/v1.0',  
      authHelper.getAccessTokenFn('https://outlook.office365.com/', session)); 

Where `authHelper.getAccessTokenFn` is a callback method you implement to provide the needed OAuth2 access token.

For more details on the syntax of using this library, see:

- [Common mail tasks](https://msdn.microsoft.com/office/office365/HowTo/common-mail-tasks-client-library)
- [Common calendar tasks](https://msdn.microsoft.com/en-us/office/office365/howto/common-calendar-tasks-client-library)
- [Common contacts tasks](https://msdn.microsoft.com/en-us/office/office365/howto/common-contacts-tasks-client-library)

## Sample App ##

For a sample app that uses this library, see the [node-mail sample](https://github.com/jasonjoh/node-mail).

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Exchange Dev Blog](http://blogs.msdn.com/b/exchangedev/)