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

**Kind**: static method of <code>[base](#module_base)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.url | <code>string</code> | The full URL of the API endpoint |
| parameters.token | <code>string</code> | The access token for authentication |
| [parameters.user.email] | <code>string</code> | The user's SMTP email address, used to set the X-AnchorMailbox header. |
| [parameters.user.timezone] | <code>string</code> | The user's time zone, used to set the outlook.timezone Prefer header. |
| [parameters.method] | <code>string</code> | Used to specify the HTTP method. Default is 'GET'. |
| [parameters.query] | <code>object</code> | An object containing key/value pairs. The pairs will be serialized into a query string. |
| [parameters.payload] | <code>object</code> | : A JSON-serializable object representing the request body. |
| [parameters.headers] | <code>object</code> | : A JSON-serializable object representing custom headers to send with the request. |
| [callback] | <code>function</code> | : A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_base.setTraceFunc"></a>
### base.setTraceFunc(traceFunc)
Used to provide a tracing function.

**Kind**: static method of <code>[base](#module_base)</code>  

| Param | Type | Description |
| --- | --- | --- |
| traceFunc | <code>function</code> | A function that takes a string parameter. The string parameter contains the text to add to the trace. |

<a name="module_base.setFiddlerEnabled"></a>
### base.setFiddlerEnabled(enabled)
Used to enable network sniffing with Fiddler.

**Kind**: static method of <code>[base](#module_base)</code>  

| Param | Type | Description |
| --- | --- | --- |
| enabled | <code>boolean</code> | : `true` to enable default Fiddler proxy and disable SSL verification. `false` to disable proxy and enable SSL verification. |

<a name="module_base.apiEndpoint"></a>
### base.apiEndpoint() ⇒ <code>string</code>
Gets the API endpoint URL.

**Kind**: static method of <code>[base](#module_base)</code>  
<a name="module_base.setApiEndpoint"></a>
### base.setApiEndpoint(newEndPoint)
Sets the API endpoint URL. If not called, the default of `https://outlook.office.com/api/v1.0` is used.

**Kind**: static method of <code>[base](#module_base)</code>  

| Param | Type | Description |
| --- | --- | --- |
| newEndPoint | <code>string</code> | The API endpoint URL to use. |

<a name="module_base.anchorMailbox"></a>
### base.anchorMailbox() ⇒ <code>string</code>
Gets the default anchor mailbox address.

**Kind**: static method of <code>[base](#module_base)</code>  
<a name="module_base.setAnchorMailbox"></a>
### base.setAnchorMailbox(newAnchor)
Sets the default anchor mailbox address.

**Kind**: static method of <code>[base](#module_base)</code>  

| Param | Type | Description |
| --- | --- | --- |
| newAnchor | <code>string</code> | The SMTP address to send in the `X-Anchor-Mailbox` header. |

<a name="module_base.preferredTimeZone"></a>
### base.preferredTimeZone() ⇒ <code>string</code>
Gets the default preferred time zone.

**Kind**: static method of <code>[base](#module_base)</code>  
<a name="module_base.setPreferredTimeZone"></a>
### base.setPreferredTimeZone(preferredTimeZone)
Sets the default preferred time zone.

**Kind**: static method of <code>[base](#module_base)</code>  

| Param | Type | Description |
| --- | --- | --- |
| preferredTimeZone | <code>string</code> | The time zone in which the server should return date time values. |

<a name="module_mail"></a>
## mail

* [mail](#module_mail)
    * [.getMessages(parameters, [callback])](#module_mail.getMessages)
    * [.getMessage(parameters, [callback])](#module_mail.getMessage)
    * [.createMessage(parameters, [callback])](#module_mail.createMessage)
    * [.updateMessage(parameters, [callback])](#module_mail.updateMessage)
    * [.deleteMessage(parameters, [callback])](#module_mail.deleteMessage)
    * [.sendNewMessage(parameters, [callback])](#module_mail.sendNewMessage)
    * [.sendDraftMessage(parameters, [callback])](#module_mail.sendDraftMessage)

<a name="module_mail.getMessages"></a>
### mail.getMessages(parameters, [callback])
Used to get messages from a folder.

**Kind**: static method of <code>[mail](#module_mail)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.folderId] | <code>string</code> | The folder id. If absent, the API calls the `/User/Messages` endpoint. Valid values of this parameter are: - The `Id` property of a `MailFolder` entity - `Inbox` - `Drafts` - `SentItems` - `DeletedItems` |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_mail.getMessage"></a>
### mail.getMessage(parameters, [callback])
Used to get a specific message.

**Kind**: static method of <code>[mail](#module_mail)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_mail.createMessage"></a>
### mail.createMessage(parameters, [callback])
Create a new message

**Kind**: static method of <code>[mail](#module_mail)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.message | <code>object</code> | : The JSON-serializable message |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.folderId] | <code>string</code> | The folder id. If absent, the API calls the `/User/Messages` endpoint. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_mail.updateMessage"></a>
### mail.updateMessage(parameters, [callback])
Update a specific message.

**Kind**: static method of <code>[mail](#module_mail)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| parameters.update | <code>object</code> | : The JSON-serializable update payload |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_mail.deleteMessage"></a>
### mail.deleteMessage(parameters, [callback])
Delete a specific message.

**Kind**: static method of <code>[mail](#module_mail)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_mail.sendNewMessage"></a>
### mail.sendNewMessage(parameters, [callback])
Sends a new message

**Kind**: static method of <code>[mail](#module_mail)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.message | <code>object</code> | : The JSON-serializable message |
| [parameters.saveToSentItems] | <code>boolean</code> | : Set to false to bypass saving a copy to the Sent Items folder. Default is true. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_mail.sendDraftMessage"></a>
### mail.sendDraftMessage(parameters, [callback])
Sends a draft message.

**Kind**: static method of <code>[mail](#module_mail)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.messageId | <code>string</code> | The Id of the message. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_calendar"></a>
## calendar

* [calendar](#module_calendar)
    * [.getEvents(parameters, [callback])](#module_calendar.getEvents)
    * [.getEvent(parameters, [callback])](#module_calendar.getEvent)
    * [.createEvent(parameters, [callback])](#module_calendar.createEvent)
    * [.updateEvent(parameters, [callback])](#module_calendar.updateEvent)
    * [.deleteEvent(parameters, [callback])](#module_calendar.deleteEvent)

<a name="module_calendar.getEvents"></a>
### calendar.getEvents(parameters, [callback])
Used to get events from a calendar.

**Kind**: static method of <code>[calendar](#module_calendar)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.calendarId] | <code>string</code> | The calendar id. If absent, the API calls the `/User/Events` endpoint. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_calendar.getEvent"></a>
### calendar.getEvent(parameters, [callback])
Used to get a specific event.

**Kind**: static method of <code>[calendar](#module_calendar)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.eventId | <code>string</code> | The Id of the event. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_calendar.createEvent"></a>
### calendar.createEvent(parameters, [callback])
Create a new event

**Kind**: static method of <code>[calendar](#module_calendar)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.event | <code>object</code> | : The JSON-serializable event |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.calendarId] | <code>string</code> | The calendar id. If absent, the API calls the `/User/Events` endpoint. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_calendar.updateEvent"></a>
### calendar.updateEvent(parameters, [callback])
Update a specific event.

**Kind**: static method of <code>[calendar](#module_calendar)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.eventId | <code>string</code> | The Id of the event. |
| parameters.update | <code>object</code> | : The JSON-serializable update payload |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_calendar.deleteEvent"></a>
### calendar.deleteEvent(parameters, [callback])
Delete a specific event.

**Kind**: static method of <code>[calendar](#module_calendar)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.eventId | <code>string</code> | The Id of the event. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

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

**Kind**: static method of <code>[contacts](#module_contacts)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segement instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segement is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.contactFolderId] | <code>string</code> | The contact folder id. If absent, the API calls the `/User/Contacts` endpoint. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_contacts.getContact"></a>
### contacts.getContact(parameters, [callback])
Used to get a specific contact.

**Kind**: static method of <code>[contacts](#module_contacts)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contactId | <code>string</code> | The Id of the contact. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_contacts.createContact"></a>
### contacts.createContact(parameters, [callback])
Create a new contact

**Kind**: static method of <code>[contacts](#module_contacts)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contact | <code>object</code> | : The JSON-serializable contact |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.contactFolderId] | <code>string</code> | The contact folder id. If absent, the API calls the `/User/Contacts` endpoint. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_contacts.updateContact"></a>
### contacts.updateContact(parameters, [callback])
Update a specific contact.

**Kind**: static method of <code>[contacts](#module_contacts)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contactId | <code>string</code> | The Id of the contact. |
| parameters.update | <code>object</code> | : The JSON-serializable update payload |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [parameters.odataParams] | <code>object</code> | An object containing key/value pairs representing OData query parameters. See [Use OData query parameters](https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters) for details. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

<a name="module_contacts.deleteContact"></a>
### contacts.deleteContact(parameters, [callback])
Delete a specific contact.

**Kind**: static method of <code>[contacts](#module_contacts)</code>  

| Param | Type | Description |
| --- | --- | --- |
| parameters | <code>object</code> | An object containing all of the relevant parameters. Possible values: |
| parameters.token | <code>string</code> | The access token. |
| parameters.contactId | <code>string</code> | The Id of the contact. |
| [parameters.useMe] | <code>boolean</code> | If true, use the /Me segment instead of the /Users/<email> segment. This parameter defaults to false and is ignored if the parameters.user.email parameter isn't provided (the /Me segment is always used in this case). |
| [parameters.user.email] | <code>string</code> | The SMTP address of the user. If absent, the '/Me' segment is used in the API URL. |
| [parameters.user.timezone] | <code>string</code> | The timezone of the user. |
| [callback] | <code>function</code> | A callback function that is called when the function completes. It should have the signature `function (error, result)`. |

