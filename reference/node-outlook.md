## Modules

<dl>
<dt><a href="#module_base">base</a></dt>
<dd></dd>
<dt><a href="#module_mail">mail</a></dt>
<dd></dd>
<dt><a href="#module_calendar">calendar</a></dt>
<dd></dd>
<dt><a href="#module_contacts">contacts</a></dt>
<dd></dd>
</dl>

<a name="module_base"></a>

## base

* [base](#module_base)
    * [.makeApiCall(parameters, [callback])](#module_base.makeApiCall)
    * [.getUser(parameters, [callback])](#module_base.getUser)
    * [.setTraceFunc(traceFunc)](#module_base.setTraceFunc)
    * [.setFiddlerEnabled(enabled)](#module_base.setFiddlerEnabled)
    * [.apiEndpoint()](#module_base.apiEndpoint) ⇒ <code>string</code>
    * [.setApiEndpoint(newEndPoint)](#module_base.setApiEndpoint)
    * [.anchorMailbox()](#module_base.anchorMailbox) ⇒ <code>string</code>
    * [.setAnchorMailbox(newAnchor)](#module_base.setAnchorMailbox)
    * [.preferredTimeZone()](#module_base.preferredTimeZone) ⇒ <code>string</code>
    * [.setPreferredTimeZone(preferredTimeZone)](#module_base.setPreferredTimeZone)

<a name="module_base.makeApiCall"></a>

### base.makeApiCall(parameters, [callback])
Used to do the actual send of a REST request to the REST endpoint.

**Kind**: static method of [<code>base</code>](#module_base)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.url | <code>string</code> | The full URL of the API endpoint |
| parameters.token | <code>string</code> | The access token for authentication |
| [parameters.user.email] | <code>string</code> | The user's SMTP email address, used to set the `X-AnchorMailbox` header. |
| [parameters.user.timezone] | <code>string</code> | The user's time zone, used to set the `outlook.timezone` `Prefer` header. |
| [parameters.method] | <code>string</code> | Used to specify the HTTP method. Default is 'GET'. |
| [parameters.query] | <code>object</code> | An object containing key/value pairs. The pairs will be serialized into a query string. |
| [parameters.payload] | <code>object</code> | A JSON-serializable object representing the request body. |
| [parameters.headers] | <code>object</code> | A JSON-serializable object representing custom headers to send with the request. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_base.getUser"></a>

### base.getUser(parameters, [callback])
Used to get information about a user.

**Kind**: static method of [<code>base</code>](#module_base)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// Set up oData parametersvar queryParams = {  '$select': 'DisplayName, EmailAddress',};outlook.base.getUser({token: token, odataParams: queryParams},  function(error, result) {    if (error) {      console.log('getUser returned an error: ' + error);    }    else if (result) {      console.log('User name:', result.DisplayName);      console.log('User email:', result.EmailAddress);    }  });
```
<a name="module_base.setTraceFunc"></a>

### base.setTraceFunc(traceFunc)
Used to provide a tracing function.

**Kind**: static method of [<code>base</code>](#module_base)  

| Param | Type | Description |
| --- | --- | --- |
| traceFunc | <code>function</code> | A function that takes a string parameter. The string parameter contains the text to add to the trace. |

<a name="module_base.setFiddlerEnabled"></a>

### base.setFiddlerEnabled(enabled)
Used to enable network sniffing with Fiddler.

**Kind**: static method of [<code>base</code>](#module_base)  

| Param | Type | Description |
| --- | --- | --- |
| enabled | <code>boolean</code> | `true` to enable default Fiddler proxy and disable SSL verification. `false` to disable proxy and enable SSL verification. |

<a name="module_base.apiEndpoint"></a>

### base.apiEndpoint() ⇒ <code>string</code>
Gets the API endpoint URL.

**Kind**: static method of [<code>base</code>](#module_base)  
<a name="module_base.setApiEndpoint"></a>

### base.setApiEndpoint(newEndPoint)
Sets the API endpoint URL. If not called, the default of `https://outlook.office.com/api/v1.0` is used.

**Kind**: static method of [<code>base</code>](#module_base)  

| Param | Type | Description |
| --- | --- | --- |
| newEndPoint | <code>string</code> | The API endpoint URL to use. |

<a name="module_base.anchorMailbox"></a>

### base.anchorMailbox() ⇒ <code>string</code>
Gets the default anchor mailbox address.

**Kind**: static method of [<code>base</code>](#module_base)  
<a name="module_base.setAnchorMailbox"></a>

### base.setAnchorMailbox(newAnchor)
Sets the default anchor mailbox address.

**Kind**: static method of [<code>base</code>](#module_base)  

| Param | Type | Description |
| --- | --- | --- |
| newAnchor | <code>string</code> | The SMTP address to send in the `X-Anchor-Mailbox` header. |

<a name="module_base.preferredTimeZone"></a>

### base.preferredTimeZone() ⇒ <code>string</code>
Gets the default preferred time zone.

**Kind**: static method of [<code>base</code>](#module_base)  
<a name="module_base.setPreferredTimeZone"></a>

### base.setPreferredTimeZone(preferredTimeZone)
Sets the default preferred time zone.

**Kind**: static method of [<code>base</code>](#module_base)  

| Param | Type | Description |
| --- | --- | --- |
| preferredTimeZone | <code>string</code> | The time zone in which the server should return date time values. |

<a name="module_mail"></a>

## mail

* [mail](#module_mail)
    * [.getMessages(parameters, [callback])](#module_mail.getMessages)
    * [.getMessage(parameters, [callback])](#module_mail.getMessage)
    * [.getMessageAttachments(parameters, [callback])](#module_mail.getMessageAttachments)
    * [.createMessage(parameters, [callback])](#module_mail.createMessage)
    * [.updateMessage(parameters, [callback])](#module_mail.updateMessage)
    * [.deleteMessage(parameters, [callback])](#module_mail.deleteMessage)
    * [.sendNewMessage(parameters, [callback])](#module_mail.sendNewMessage)
    * [.sendDraftMessage(parameters, [callback])](#module_mail.sendDraftMessage)
    * [.syncMessages(parameters, [callback])](#module_mail.syncMessages)
    * [.replyToMessage(parameters, [callback])](#module_mail.replyToMessage)
    * [.replyToAllMessage(parameters, [callback])](#module_mail.replyToAllMessage)

<a name="module_mail.getMessages"></a>

### mail.getMessages(parameters, [callback])
Used to get messages from a folder.

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.folderId] | <code>string</code> | The folder id. If absent, the API calls the `/User/Messages` endpoint. Valid values of this parameter are: - The `Id` property of a `MailFolder` entity - `Inbox` - `Drafts` - `SentItems` - `DeletedItems` |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// Set up oData parametersvar queryParams = {  '$select': 'Subject,ReceivedDateTime,From',  '$orderby': 'ReceivedDateTime desc',  '$top': 20};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.getMessages({token: token, folderId: 'Inbox', odataParams: queryParams, user: userInfo},  function(error, result){    if (error) {      console.log('getMessages returned an error: ' + error);    }    else if (result) {      console.log('getMessages returned ' + result.value.length + ' messages.');      result.value.forEach(function(message) {        console.log('  Subject:', message.Subject);        console.log('  Received:', message.ReceivedDateTime.toString());        console.log('  From:', message.From ? message.From.EmailAddress.Name : 'EMPTY');      });    }  });
```
<a name="module_mail.getMessage"></a>

### mail.getMessage(parameters, [callback])
Used to get a specific message.

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the message to retrieve. This could be// from a previous call to getMessagesvar msgId = 'AAMkADVhYTYwNzk...';// Set up oData parametersvar queryParams = {  '$select': 'Subject,ReceivedDateTime,From'};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.getMessage({token: token, messageId: msgId, odataParams: queryParams, user: userInfo},  function(error, result){    if (error) {      console.log('getMessage returned an error: ' + error);    }    else if (result) {      console.log('  Subject:', result.Subject);      console.log('  Received:', result.ReceivedDateTime.toString());      console.log('  From:', result.From ? result.From.EmailAddress.Name : 'EMPTY');    }  });
```
<a name="module_mail.getMessageAttachments"></a>

### mail.getMessageAttachments(parameters, [callback])
Get all attachments from a message

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the message to retrieve. This could be// from a previous call to getMessagesvar msgId = 'AAMkADVhYTYwNzk...';// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.getMessageAttachments({token: token, messageId: msgId, user: userInfo},  function(error, result){    if (error) {      console.log('getMessageAttachments returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_mail.createMessage"></a>

### mail.createMessage(parameters, [callback])
Create a new message

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.message | <code>object</code> | The JSON-serializable message |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.folderId] | <code>string</code> | The folder id. If absent, the API calls the `/User/Messages` endpoint. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';var newMsg = {  Subject: 'Did you see last night\'s game?',  Importance: 'Low',  Body: {    ContentType: 'HTML',    Content: 'They were <b>awesome</b>!'  },  ToRecipients: [    {      EmailAddress: {        Address: 'azizh@contoso.com'      }    }  ]};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.createMessage({token: token, message: newMsg, user: userInfo},  function(error, result){    if (error) {      console.log('createMessage returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_mail.updateMessage"></a>

### mail.updateMessage(parameters, [callback])
Update a specific message.

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| parameters.update | <code>object</code> | The JSON-serializable update payload |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the message to update. This could be// from a previous call to getMessagesvar msgId = 'AAMkADVhYTYwNzk...';// Mark the message unreadvar update = {  IsRead: false,};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.updateMessage({token: token, messageId: msgId, update: update, user: userInfo},  function(error, result){    if (error) {      console.log('updateMessage returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_mail.deleteMessage"></a>

### mail.deleteMessage(parameters, [callback])
Delete a specific message.

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the message to delete. This could be// from a previous call to getMessagesvar msgId = 'AAMkADVhYTYwNzk...';// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.deleteMessage({token: token, messageId: msgId, user: userInfo},  function(error, result){    if (error) {      console.log('deleteMessage returned an error: ' + error);    }    else if (result) {      console.log('SUCCESS');    }  });
```
<a name="module_mail.sendNewMessage"></a>

### mail.sendNewMessage(parameters, [callback])
Sends a new message

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.message | <code>object</code> | The JSON-serializable message |
| [parameters.saveToSentItems] | <code>boolean</code> | Set to false to bypass saving a copy to the Sent Items folder. Default is true. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';var newMsg = {  Subject: 'Did you see last night\'s game?',  Importance: 'Low',  Body: {    ContentType: 'HTML',    Content: 'They were <b>awesome</b>!'  },  ToRecipients: [    {      EmailAddress: {        Address: 'azizh@contoso.com'      }    }  ]};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.sendNewMessage({token: token, message: newMsg, user: userInfo},  function(error, result){    if (error) {      console.log('sendNewMessage returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_mail.sendDraftMessage"></a>

### mail.sendDraftMessage(parameters, [callback])
Sends a draft message.

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the message to send. This could be// from a previous call to getMessagesvar msgId = 'AAMkADVhYTYwNzk...';// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.mail.sendDraftMessage({token: token, messageId: msgId, user: userInfo},  function(error, result){    if (error) {      console.log('sendDraftMessage returned an error: ' + error);    }    else if (result) {      console.log('SUCCESS');    }  });
```
<a name="module_mail.syncMessages"></a>

### mail.syncMessages(parameters, [callback])
Syncs messages in a folder.

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.pageSize] | <code>Number</code> | The maximum number of results to return in each call. Defaults to 50. |
| [parameters.skipToken] | <code>string</code> | The value to pass in the `skipToken` query parameter in the API call. |
| [parameters.deltaToken] | <code>string</code> | The value to pass in the `deltaToken` query parameter in the API call. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.folderId] | <code>string</code> | The folder id. If absent, the API calls the `/User/Messages` endpoint. Valid values of this parameter are: - The `Id` property of a `MailFolder` entity - `Inbox` - `Drafts` - `SentItems` - `DeletedItems` |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the beta endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/beta');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};var syncMsgParams = {  '$select': 'Subject,ReceivedDateTime,From,BodyPreview,IsRead',  '$orderby': 'ReceivedDateTime desc'};var apiOptions = {  token: token,  folderId: 'Inbox',  odataParams: syncMsgParams,  user: userinfo,  pageSize: 20};outlook.mail.syncMessages(apiOptions, function(error, messages) {  if (error) {    console.log('syncMessages returned an error:', error);  } else {    // Do something with the messages.value array    // Then get the @odata.deltaLink    var delta = messages['@odata.deltaLink'];    // Handle deltaLink value appropriately:    // In general, if the deltaLink has a $skiptoken, that means there are more    // "pages" in the sync results, you should call syncMessages again, passing    // the $skiptoken value in the apiOptions.skipToken. If on the other hand,    // the deltaLink has a $deltatoken, that means the sync is complete, and you should    // store the $deltatoken value for future syncs.    //    // The one exception to this rule is on the initial sync (when you call with no skip or delta tokens).    // In this case you always get a $deltatoken back, even if there are more results. In this case, you should    // immediately call syncMessages again, passing the $deltatoken value in apiOptions.deltaToken.  }}
```
<a name="module_mail.replyToMessage"></a>

### mail.replyToMessage(parameters, [callback])
Reply to sender

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The ID of the message to reply to. |
| parameters.comment | <code>string</code> | The comment to include.  Can be an empty string. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';var comment = "Sounds great! See you tomorrow.";var messageId = "AAMkAGE0Mz8DmAAA=";outlook.mail.replyToMessage({token: token, comment: comment, messageId: messageId},  function(error, result){    if (error) {      console.log('replyToMessage returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_mail.replyToAllMessage"></a>

### mail.replyToAllMessage(parameters, [callback])
Reply to all

**Kind**: static method of [<code>mail</code>](#module_mail)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The ID of the message to reply to. |
| parameters.comment | <code>string</code> | The comment to include.  Can be an empty string. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';var comment = "Sounds great! See you tomorrow.";var messageId = "AAMkAGE0Mz8DmAAA=";outlook.mail.replyToAllMessage({token: token, comment: comment, messageId: messageId},  function(error, result){    if (error) {      console.log('replyToMessage returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_calendar"></a>

## calendar

* [calendar](#module_calendar)
    * [.getEvents(parameters, [callback])](#module_calendar.getEvents)
    * [.syncEvents(parameters, [callback])](#module_calendar.syncEvents)
    * [.getEvent(parameters, [callback])](#module_calendar.getEvent)
    * [.createEvent(parameters, [callback])](#module_calendar.createEvent)
    * [.updateEvent(parameters, [callback])](#module_calendar.updateEvent)
    * [.deleteEvent(parameters, [callback])](#module_calendar.deleteEvent)

<a name="module_calendar.getEvents"></a>

### calendar.getEvents(parameters, [callback])
Used to get events from a calendar.

**Kind**: static method of [<code>calendar</code>](#module_calendar)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.calendarId] | <code>string</code> | The calendar id. If absent, the API calls the `/User/Events` endpoint. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// Set up oData parametersvar queryParams = {  '$select': 'Subject,Start,End',  '$orderby': 'Start/DateTime desc',  '$top': 20};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.calendar.getEvents({token: token, folderId: 'Inbox', odataParams: queryParams, user: userInfo},  function(error, result){    if (error) {      console.log('getEvents returned an error: ' + error);    }    else if (result) {      console.log('getEvents returned ' + result.value.length + ' events.');      result.value.forEach(function(event) {        console.log('  Subject:', event.Subject);        console.log('  Start:', event.Start.DateTime.toString());        console.log('  End:', event.End.DateTime.toString());      });    }  });
```
<a name="module_calendar.syncEvents"></a>

### calendar.syncEvents(parameters, [callback])
Syncs events of a calendar.

**Kind**: static method of [<code>calendar</code>](#module_calendar)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.startDateTime | <code>string</code> | The start time and date for the calendar view in ISO 8601 format without a timezone designator. Time zone is assumed to be UTC unless the `Prefer: outlook.timezone` header is sent in the request. |
| parameters.endDateTime | <code>string</code> | The end time and date for the calendar view in ISO 8601 format without a timezone designator. Time zone is assumed to be UTC unless the `Prefer: outlook.timezone` header is sent in the request. |
| [parameters.skipToken] | <code>string</code> | The value to pass in the `skipToken` query parameter in the API call. |
| [parameters.deltaToken] | <code>string</code> | The value to pass in the `deltaToken` query parameter in the API call. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.calendarId] | <code>string</code> | The calendar id. If absent, the API calls the `/User/calendarview` endpoint. Valid values of this parameter are: - The `Id` property of a `Calendar` entity - `Primary`, the primary calendar is used. It is used by default if no Id is specified |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the 2.0 version of the apioutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};// You have to specify a time windowvar startDateTime = "2017-01-01";var endDateTime = "2017-12-31";var apiOptions = {  token: token,  calendarId: 'calendar_id', // If none specified, the Primary calendar will be used  user: userinfo,  startDateTime: startDateTime,  endDateTime: endDateTime};outlook.calendar.syncEvents(apiOptions, function(error, events) {  if (error) {    console.log('syncEvents returned an error:', error);  } else {    // Do something with the events.value array    // Then get the @odata.deltaLink    var delta = messages['@odata.deltaLink'];    // Handle deltaLink value appropriately:    // In general, if the deltaLink has a $skiptoken, that means there are more    // "pages" in the sync results, you should call syncEvents again, passing    // the $skiptoken value in the apiOptions.skipToken. If on the other hand,    // the deltaLink has a $deltatoken, that means the sync is complete, and you should    // store the $deltatoken value for future syncs.    //    // The one exception to this rule is on the initial sync (when you call with no skip or delta tokens).    // In this case you always get a $deltatoken back, even if there are more results. In this case, you should    // immediately call syncMessages again, passing the $deltatoken value in apiOptions.deltaToken.  }}
```
<a name="module_calendar.getEvent"></a>

### calendar.getEvent(parameters, [callback])
Used to get a specific event.

**Kind**: static method of [<code>calendar</code>](#module_calendar)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.eventId | <code>string</code> | The Id of the event. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the event to retrieve. This could be// from a previous call to getEventsvar eventId = 'AAMkADVhYTYwNzk...';// Set up oData parametersvar queryParams = {  '$select': 'Subject,Start,End'};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.calendar.getEvent({token: token, eventId: eventId, odataParams: queryParams, user: userInfo},  function(error, result){    if (error) {      console.log('getEvent returned an error: ' + error);    }    else if (result) {      console.log('  Subject:', result.Subject);      console.log('  Start:', result.Start.DateTime.toString());      console.log('  End:', result.End.DateTime.toString());    }  });
```
<a name="module_calendar.createEvent"></a>

### calendar.createEvent(parameters, [callback])
Create a new event

**Kind**: static method of [<code>calendar</code>](#module_calendar)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.event | <code>object</code> | The JSON-serializable event |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.calendarId] | <code>string</code> | The calendar id. If absent, the API calls the `/User/Events` endpoint. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';var newEvent = {  "Subject": "Discuss the Calendar REST API",  "Body": {    "ContentType": "HTML",    "Content": "I think it will meet our requirements!"  },  "Start": {    "DateTime": "2016-02-03T18:00:00",    "TimeZone": "Eastern Standard Time"  },  "End": {    "DateTime": "2016-02-03T19:00:00",    "TimeZone": "Eastern Standard Time"  },  "Attendees": [    {      "EmailAddress": {        "Address": "allieb@contoso.com",        "Name": "Allie Bellew"      },      "Type": "Required"    }  ]};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.calendar.createEvent({token: token, event: newEvent, user: userInfo},  function(error, result){    if (error) {      console.log('createEvent returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_calendar.updateEvent"></a>

### calendar.updateEvent(parameters, [callback])
Update a specific event.

**Kind**: static method of [<code>calendar</code>](#module_calendar)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.eventId | <code>string</code> | The Id of the event. |
| parameters.update | <code>object</code> | The JSON-serializable update payload |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the event to update. This could be// from a previous call to getEventsvar eventId = 'AAMkADVhYTYwNzk...';// Update the locationvar update = {  Location: {    DisplayName: 'Conference Room 2'  }};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.calendar.updateEvent({token: token, eventId: eventId, update: update, user: userInfo},  function(error, result){    if (error) {      console.log('updateEvent returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_calendar.deleteEvent"></a>

### calendar.deleteEvent(parameters, [callback])
Delete a specific event.

**Kind**: static method of [<code>calendar</code>](#module_calendar)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.eventId | <code>string</code> | The Id of the event. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the event to delete. This could be// from a previous call to getEventsvar eventId = 'AAMkADVhYTYwNzk...';// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.calendar.deleteEvent({token: token, eventId: eventId, user: userInfo},  function(error, result){    if (error) {      console.log('deleteEvent returned an error: ' + error);    }    else if (result) {      console.log('SUCCESS');    }  });
```
<a name="module_contacts"></a>

## contacts

* [contacts](#module_contacts)
    * [.getContacts(parameters, [callback])](#module_contacts.getContacts)
    * [.getContact(parameters, [callback])](#module_contacts.getContact)
    * [.createContact(parameters, [callback])](#module_contacts.createContact)
    * [.updateContact(parameters, [callback])](#module_contacts.updateContact)
    * [.deleteContact(parameters, [callback])](#module_contacts.deleteContact)

<a name="module_contacts.getContacts"></a>

### contacts.getContacts(parameters, [callback])
Used to get contacts from a contact folder.

**Kind**: static method of [<code>contacts</code>](#module_contacts)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.contactFolderId] | <code>string</code> | The contact folder id. If absent, the API calls the `/User/Contacts` endpoint. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// Set up oData parametersvar queryParams = {  '$select': 'GivenName,Surname,EmailAddresses',  '$orderby': 'CreatedDateTime desc',  '$top': 20};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.contacts.getContacts({token: token, odataParams: queryParams, user: userInfo},  function(error, result){    if (error) {      console.log('getContacts returned an error: ' + error);    }    else if (result) {      console.log('getContacts returned ' + result.value.length + ' contacts.');      result.value.forEach(function(contact) {        console.log('  GivenName:', contact.GivenName);        console.log('  Surname:', contact.Surname);        console.log('  Email Address:', contact.EmailAddresses[0] ? contact.EmailAddresses[0].Address : "NONE");      });    }  });
```
<a name="module_contacts.getContact"></a>

### contacts.getContact(parameters, [callback])
Used to get a specific contact.

**Kind**: static method of [<code>contacts</code>](#module_contacts)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contactId | <code>string</code> | The Id of the contact. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the contact to retrieve. This could be// from a previous call to getContactsvar contactId = 'AAMkADVhYTYwNzk...';// Set up oData parametersvar queryParams = {  '$select': 'GivenName,Surname,EmailAddresses'};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.contacts.getContact({token: token, contactId: contactId, odataParams: queryParams, user: userInfo},  function(error, result){    if (error) {      console.log('getContact returned an error: ' + error);    }    else if (result) {      console.log('  GivenName:', result.GivenName);      console.log('  Surname:', result.Surname);      console.log('  Email Address:', result.EmailAddresses[0] ? result.EmailAddresses[0].Address : "NONE");    }  });
```
<a name="module_contacts.createContact"></a>

### contacts.createContact(parameters, [callback])
Create a new contact

**Kind**: static method of [<code>contacts</code>](#module_contacts)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contact | <code>object</code> | The JSON-serializable contact |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.contactFolderId] | <code>string</code> | The contact folder id. If absent, the API calls the `/User/Contacts` endpoint. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';var newContact = {  "GivenName": "Pavel",  "Surname": "Bansky",  "EmailAddresses": [    {      "Address": "pavelb@contoso.com",      "Name": "Pavel Bansky"    }  ],  "BusinessPhones": [    "+1 732 555 0102"  ]};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.contacts.createContact({token: token, contact: newContact, user: userInfo},  function(error, result){    if (error) {      console.log('createContact returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_contacts.updateContact"></a>

### contacts.updateContact(parameters, [callback])
Update a specific contact.

**Kind**: static method of [<code>contacts</code>](#module_contacts)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contactId | <code>string</code> | The Id of the contact. |
| parameters.update | <code>object</code> | The JSON-serializable update payload |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the contact to update. This could be// from a previous call to getContactsvar contactId = 'AAMkADVhYTYwNzk...';// Change the mobile numbervar update = {  MobilePhone1: '425-555-1212',};// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.contacts.updateContact({token: token, contactId: contactId, update: update, user: userInfo},  function(error, result){    if (error) {      console.log('updateContact returned an error: ' + error);    }    else if (result) {      console.log(JSON.stringify(result, null, 2));    }  });
```
<a name="module_contacts.deleteContact"></a>

### contacts.deleteContact(parameters, [callback])
Delete a specific contact.

**Kind**: static method of [<code>contacts</code>](#module_contacts)  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contactId | <code>string</code> | The Id of the contact. |
| [parameters.useMe] | <code>boolean</code> | If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the `/Me` segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

**Example**  
```js
var outlook = require('node-outlook');// Set the API endpoint to use the v2.0 endpointoutlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');// This is the oAuth tokenvar token = 'eyJ0eXAiOiJKV1Q...';// The Id property of the contact to delete. This could be// from a previous call to getContactsvar contactId = 'AAMkADVhYTYwNzk...';// Pass the user's email addressvar userInfo = {  email: 'sarad@contoso.com'};outlook.contacts.deleteContact({token: token, contactId: contactId, user: userInfo},  function(error, result){    if (error) {      console.log('deleteContact returned an error: ' + error);    }    else if (result) {      console.log('SUCCESS');    }  });
```
