// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var base = require('./version-2.js');
var utilities = require('./utilities.js');

/**
 * @module mail
 */

module.exports = {
  /**
   * Used to get messages from a folder.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.folderId] {string} The folder id. If absent, the API calls the `/User/Messages` endpoint. Valid values of this parameter are:
   *
   * - The `Id` property of a `MailFolder` entity
   * - `Inbox`
   * - `Drafts`
   * - `SentItems`
   * - `DeletedItems`
   *
   * @param [parameters.odataParams] {object} An object containing key/value pairs representing OData query parameters. See [Use OData query parameters]{@link https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters} for details.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * // Set up oData parameters
   * var queryParams = {
   *   '$select': 'Subject,ReceivedDateTime,From',
   *   '$orderby': 'ReceivedDateTime desc',
   *   '$top': 20
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.getMessages({token: token, folderId: 'Inbox', odataParams: queryParams, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('getMessages returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('getMessages returned ' + result.value.length + ' messages.');
   *       result.value.forEach(function(message) {
   *         console.log('  Subject:', message.Subject);
   *         console.log('  Received:', message.ReceivedDateTime.toString());
   *         console.log('  From:', message.From ? message.From.EmailAddress.Name : 'EMPTY');
   *       });
   *     }
   *   });
   */
  getMessages: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);
    var folderSpec = parameters.folderId === undefined ? '' : getFolderSegment() + parameters.folderId;

    var requestUrl = base.apiEndpoint() + userSpec + folderSpec + '/Messages';

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user
    };

    if (parameters.odataParams !== undefined) {
      apiOptions['query'] = parameters.odataParams;
    }

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 200) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Used to get a specific message.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.messageId {string} The Id of the message.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.odataParams] {object} An object containing key/value pairs representing OData query parameters. See [Use OData query parameters]{@link https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters} for details.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * // The Id property of the message to retrieve. This could be
   * // from a previous call to getMessages
   * var msgId = 'AAMkADVhYTYwNzk...';
   *
   * // Set up oData parameters
   * var queryParams = {
   *   '$select': 'Subject,ReceivedDateTime,From'
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.getMessage({token: token, messageId: msgId, odataParams: queryParams, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('getMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('  Subject:', result.Subject);
   *       console.log('  Received:', result.ReceivedDateTime.toString());
   *       console.log('  From:', result.From ? result.From.EmailAddress.Name : 'EMPTY');
   *     }
   *   });
   */
  getMessage: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/Messages/' + parameters.messageId;

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user
    };

    if (parameters.odataParams !== undefined) {
      apiOptions['query'] = parameters.odataParams;
    }

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 200) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Get all attachments from a message
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.messageId {string} The Id of the message.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.odataParams] {object} An object containing key/value pairs representing OData query parameters. See [Use OData query parameters]{@link https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters} for details.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * // The Id property of the message to retrieve. This could be
   * // from a previous call to getMessages
   * var msgId = 'AAMkADVhYTYwNzk...';
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.getMessageAttachments({token: token, messageId: msgId, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('getMessageAttachments returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  getMessageAttachments: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);
    var requestUrl = base.apiEndpoint() + userSpec + '/messages/' + parameters.messageId + '/attachments';
    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user
    };

    if (parameters.odataParams !== undefined) {
      apiOptions['query'] = parameters.odataParams;
    }

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      } else if (response.statusCode !== 200) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      } else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Create a new message
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.message {object} The JSON-serializable message
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.folderId] {string} The folder id. If absent, the API calls the `/User/Messages` endpoint.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * var newMsg = {
   *   Subject: 'Did you see last night\'s game?',
   *   Importance: 'Low',
   *   Body: {
   *     ContentType: 'HTML',
   *     Content: 'They were <b>awesome</b>!'
   *   },
   *   ToRecipients: [
   *     {
   *       EmailAddress: {
   *         Address: 'azizh@contoso.com'
   *       }
   *     }
   *   ]
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.createMessage({token: token, message: newMsg, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('createMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  createMessage: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);
    var folderSpec = parameters.folderId === undefined ? '' : getFolderSegment() + parameters.folderId;

    var requestUrl = base.apiEndpoint() + userSpec + folderSpec + '/Messages';

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      payload: parameters.message,
      method: 'POST'
    };

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 201) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Update a specific message.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.messageId {string} The Id of the message.
   * @param parameters.update {object} The JSON-serializable update payload
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.odataParams] {object} An object containing key/value pairs representing OData query parameters. See [Use OData query parameters]{@link https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters} for details.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * // The Id property of the message to update. This could be
   * // from a previous call to getMessages
   * var msgId = 'AAMkADVhYTYwNzk...';
   *
   * // Mark the message unread
   * var update = {
   *   IsRead: false,
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.updateMessage({token: token, messageId: msgId, update: update, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('updateMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  updateMessage: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/Messages/' + parameters.messageId;

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      payload: parameters.update,
      method: 'PATCH'
    };

    if (parameters.odataParams !== undefined) {
      apiOptions['query'] = parameters.odataParams;
    }

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 200) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Delete a specific message.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.messageId {string} The Id of the message.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * // The Id property of the message to delete. This could be
   * // from a previous call to getMessages
   * var msgId = 'AAMkADVhYTYwNzk...';
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.deleteMessage({token: token, messageId: msgId, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('deleteMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('SUCCESS');
   *     }
   *   });
   */
  deleteMessage: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/Messages/' + parameters.messageId;

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      method: 'DELETE'
    };

    if (parameters.odataParams !== undefined) {
      apiOptions['query'] = parameters.odataParams;
    }

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 204) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Sends a new message
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.message {object} The JSON-serializable message
   * @param [parameters.saveToSentItems] {boolean} Set to false to bypass saving a copy to the Sent Items folder. Default is true.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * var newMsg = {
   *   Subject: 'Did you see last night\'s game?',
   *   Importance: 'Low',
   *   Body: {
   *     ContentType: 'HTML',
   *     Content: 'They were <b>awesome</b>!'
   *   },
   *   ToRecipients: [
   *     {
   *       EmailAddress: {
   *         Address: 'azizh@contoso.com'
   *       }
   *     }
   *   ]
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.sendNewMessage({token: token, message: newMsg, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('sendNewMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  sendNewMessage: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/sendmail';

    var payload = {
      Message: parameters.message,
      SaveToSentItems: parameters.saveToSentItems !== undefined ? parameters.saveToSentItems : 'true'
    };

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      payload: payload,
      method: 'POST'
    };

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 202) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Sends a draft message.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.messageId {string} The Id of the message.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * // The Id property of the message to send. This could be
   * // from a previous call to getMessages
   * var msgId = 'AAMkADVhYTYwNzk...';
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.mail.sendDraftMessage({token: token, messageId: msgId, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('sendDraftMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('SUCCESS');
   *     }
   *   });
   */
  sendDraftMessage: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/Messages/' + parameters.messageId + '/send';

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      method: 'POST'
    };

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 202) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Syncs messages in a folder.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param [parameters.pageSize] {Number} The maximum number of results to return in each call. Defaults to 50.
   * @param [parameters.skipToken] {string} The value to pass in the `skipToken` query parameter in the API call.
   * @param [parameters.deltaToken] {string} The value to pass in the `deltaToken` query parameter in the API call.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.folderId] {string} The folder id. If absent, the API calls the `/User/Messages` endpoint. Valid values of this parameter are:
   *
   * - The `Id` property of a `MailFolder` entity
   * - `Inbox`
   * - `Drafts`
   * - `SentItems`
   * - `DeletedItems`
   *
   * @param [parameters.odataParams] {object} An object containing key/value pairs representing OData query parameters. See [Use OData query parameters]{@link https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters} for details.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the beta endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/beta');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * var syncMsgParams = {
   *   '$select': 'Subject,ReceivedDateTime,From,BodyPreview,IsRead',
   *   '$orderby': 'ReceivedDateTime desc'
   * };
   *
   * var apiOptions = {
   *   token: token,
   *   folderId: 'Inbox',
   *   odataParams: syncMsgParams,
   *   user: userinfo,
   *   pageSize: 20
   * };
   *
   * outlook.mail.syncMessages(apiOptions, function(error, messages) {
   *   if (error) {
   *     console.log('syncMessages returned an error:', error);
   *   } else {
   *     // Do something with the messages.value array
   *     // Then get the @odata.deltaLink
   *     var delta = messages['@odata.deltaLink'];
   *
   *     // Handle deltaLink value appropriately:
   *     // In general, if the deltaLink has a $skiptoken, that means there are more
   *     // "pages" in the sync results, you should call syncMessages again, passing
   *     // the $skiptoken value in the apiOptions.skipToken. If on the other hand,
   *     // the deltaLink has a $deltatoken, that means the sync is complete, and you should
   *     // store the $deltatoken value for future syncs.
   *     //
   *     // The one exception to this rule is on the initial sync (when you call with no skip or delta tokens).
   *     // In this case you always get a $deltatoken back, even if there are more results. In this case, you should
   *     // immediately call syncMessages again, passing the $deltatoken value in apiOptions.deltaToken.
   *   }
   * }
   */

  syncMessages: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);
    var folderSpec = parameters.folderId === undefined ? '' : getFolderSegment() + parameters.folderId;

    var requestUrl = base.apiEndpoint() + userSpec + folderSpec + '/Messages';

    var query = parameters.odataParams || {};
    if (parameters.skipToken) {
      query['$skiptoken'] = parameters.skipToken;
    }

    if (parameters.deltaToken) {
      query['$deltatoken'] = parameters.deltaToken;
    }

    var headers = {
      Prefer: [
        'odata.track-changes',
        'odata.maxpagesize=' + (parameters.pageSize === undefined ? '50' : parameters.pageSize.toString())
      ]
    };

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      query: query,
      headers: headers
    };

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 200) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Reply to sender
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.messageId {string} The ID of the message to reply to.
   * @param parameters.comment {string} The comment to include.  Can be an empty string.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * var comment = "Sounds great! See you tomorrow.";
   * var messageId = "AAMkAGE0Mz8DmAAA=";
   *
   * outlook.mail.replyToMessage({token: token, comment: comment, messageId: messageId},
   *   function(error, result){
   *     if (error) {
   *       console.log('replyToMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  replyToMessage: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/messages/'+ parameters.messageId + '/reply' ;

    var payload = {
      Comment: parameters.comment,
    };

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      payload: payload,
      method: 'POST'
    };

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 202) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  },

  /**
   * Reply to all
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.messageId {string} The ID of the message to reply to.
   * @param parameters.comment {string} The comment to include.  Can be an empty string.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   *
   * @example var outlook = require('node-outlook');
   *
   * // Set the API endpoint to use the v2.0 endpoint
   * outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
   *
   * // This is the oAuth token
   * var token = 'eyJ0eXAiOiJKV1Q...';
   *
   * var comment = "Sounds great! See you tomorrow.";
   * var messageId = "AAMkAGE0Mz8DmAAA=";
   *
   * outlook.mail.replyToAllMessage({token: token, comment: comment, messageId: messageId},
   *   function(error, result){
   *     if (error) {
   *       console.log('replyToMessage returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  replyToAllMessage: function(params, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/messages/'+ parameters.messageId + '/replyall' ;

    var payload = {
      Comment: parameters.comment,
    };

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      payload: payload,
      method: 'POST'
    };

    base.makeApiCall(apiOptions, function(error, response) {
      if (error) {
        if (typeof callback === 'function') {
          callback(error, response);
        }
      }
      else if (response.statusCode !== 202) {
        if (typeof callback === 'function') {
          callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
        }
      }
      else {
        if (typeof callback === 'function') {
          callback(null, response.body);
        }
      }
    });
  }
};

/**
 * Helper function to return the correct name for the folders segment
 * of the request URL. /Me/Folders became /Me/MailFolders in the beta and
 * 2.0 endpoints.
 * @private
 */
var getFolderSegment = function() {
  if (base.apiEndpoint().toLowerCase().indexOf('/api/v1.0') > 0){
    return '/Folders/';
  }

  return '/MailFolders/'
}