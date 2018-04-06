# Node.js Wrapper for Office 365 APIs Client Library

This library provides a light-weight implementation of the Outlook [Mail](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations), [Calendar](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations), and [Contacts](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) APIs.

For a sample app that uses this library, see [Getting Started with the Outlook Mail API and Node.js](https://github.com/jasonjoh/node-tutorial).

> The original version of this library was a simple wrapper intended to enable use of the [Microsoft Office 365 APIs Client Libraries for Cordova Applications](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) from a Node.js app. This version still includes the old interfaces, but no further development is being made on that part of this library. It is recommended that applications use the new interfaces moving forward.

## Required software

- [Microsoft Office 365 APIs Client Libraries for Cordova Applications](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) (Included)
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md)

## Installation

Installing should be done via NPM:

```Shell
npm install node-outlook
```

## Usage

Once installed, add the following to your source file:

```js
var outlook = require("node-outlook");
```

### New interface

The new interface is documented in a simple [reference](reference/node-outlook.md) (courtesy of [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown)).

#### Configuration

Configuration of the library is done via the `base` namespace:

- `outlook.base.setApiEndpoint` - Use this to override the API endpoint. The default value uses the Outlook v1.0 endpoint: `https://outlook.office.com/api/v1.0`.
- `outlook.base.setAnchorMailbox` - Set this to the user's SMTP address to enable the API endpoint to efficiently route API requests.
- `outlook.base.setPreferredTimeZone` - Use this to specify a time zone for the server to use to return date/time values in the Calendar API.

#### Making API calls

The library has a namespace for each API.

- `outlook.mail` - The Mail API
- `outlook.calendar` - The Calendar API
- `outlook.contacts` - The Contacts API

Each namespace has minimal functions (more to come). Usage is similar between the namespaces. For example, this is how you call the `getMessages` function in the `outlook.mail` namespace:

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

#### Making raw API calls

If the library does not implement a function that does what you need, you can use the `outlook.base.makeApiCall` method to call any API call implemented on the server. See the implementations of any methods in the `outlook.mail`, `outlook.calendar`, or `outlook.contacts` namespaces for an example of how to use this method.

### Old interface

> As a reminder, the old interface is no longer being developed. It's recommended that you use the new interface.

You can create an `OutlookServices.Client` object like so:

```js
var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office.com/api/v2.0',
  authHelper.getAccessTokenFn('https://outlook.office.com/', session));
```

Where `authHelper.getAccessTokenFn` is a callback method you implement to provide the needed OAuth2 access token.

## Copyright

Copyright (c) Microsoft. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Outlook Dev Blog](https://blogs.msdn.microsoft.com/exchangedev/)