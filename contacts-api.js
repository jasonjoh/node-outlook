// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

var base = require('./version-2.js');
var utilities = require('./utilities.js');

/**
 * @module contacts
 */

module.exports = {
  /**
   * Used to get contacts from a contact folder.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.contactFolderId] {string} The contact folder id. If absent, the API calls the `/User/Contacts` endpoint.
   *
   * @param [parameters.odataParams] {object} An object containing key/value pairs representing OData query parameters. See [Use OData query parameters]{@link https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters} for details.
   *
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
   *   '$select': 'GivenName,Surname,EmailAddresses',
   *   '$orderby': 'CreatedDateTime desc',
   *   '$top': 20
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.contacts.getContacts({token: token, odataParams: queryParams, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('getContacts returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('getContacts returned ' + result.value.length + ' contacts.');
   *       result.value.forEach(function(contact) {
   *         console.log('  GivenName:', contact.GivenName);
   *         console.log('  Surname:', contact.Surname);
   *         console.log('  Email Address:', contact.EmailAddresses[0] ? contact.EmailAddresses[0].Address : "NONE");
   *       });
   *     }
   *   });
   */
  getContacts: function(parameters, callback){
    var userSpec = utilities.getUserSegment(parameters);
    var contactFolderSpec = parameters.folderId === undefined ? '' : '/ContactFolders/' + parameters.folderId;

    var requestUrl = base.apiEndpoint() + userSpec + contactFolderSpec + '/Contacts';

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
   * Used to get a specific contact.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.contactId {string} The Id of the contact.
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
   * // The Id property of the contact to retrieve. This could be
   * // from a previous call to getContacts
   * var contactId = 'AAMkADVhYTYwNzk...';
   *
   * // Set up oData parameters
   * var queryParams = {
   *   '$select': 'GivenName,Surname,EmailAddresses'
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.contacts.getContact({token: token, contactId: contactId, odataParams: queryParams, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('getContact returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('  GivenName:', result.GivenName);
   *       console.log('  Surname:', result.Surname);
   *       console.log('  Email Address:', result.EmailAddresses[0] ? result.EmailAddresses[0].Address : "NONE");
   *     }
   *   });
   */
  getContact: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/Contacts/' + parameters.contactId;

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
   * Create a new contact
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.contact {object} The JSON-serializable contact
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
   * @param [parameters.user.timezone] {string} The timezone of the user.
   * @param [parameters.contactFolderId] {string} The contact folder id. If absent, the API calls the `/User/Contacts` endpoint.
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
   * var newContact = {
   *   "GivenName": "Pavel",
   *   "Surname": "Bansky",
   *   "EmailAddresses": [
   *     {
   *       "Address": "pavelb@contoso.com",
   *       "Name": "Pavel Bansky"
   *     }
   *   ],
   *   "BusinessPhones": [
   *     "+1 732 555 0102"
   *   ]
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.contacts.createContact({token: token, contact: newContact, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('createContact returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  createContact: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);
    var folderSpec = parameters.folderId === undefined ? '' : '/ContactFolders/' + parameters.folderId;

    var requestUrl = base.apiEndpoint() + userSpec + folderSpec + '/Contacts';

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user,
      payload: parameters.contact,
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
   * Update a specific contact.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.contactId {string} The Id of the contact.
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
   * // The Id property of the contact to update. This could be
   * // from a previous call to getContacts
   * var contactId = 'AAMkADVhYTYwNzk...';
   *
   * // Change the mobile number
   * var update = {
   *   MobilePhone1: '425-555-1212',
   * };
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.contacts.updateContact({token: token, contactId: contactId, update: update, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('updateContact returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log(JSON.stringify(result, null, 2));
   *     }
   *   });
   */
  updateContact: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/Contacts/' + parameters.contactId;

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
   * Delete a specific contact.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param parameters.contactId {string} The Id of the contact.
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
   * // The Id property of the contact to delete. This could be
   * // from a previous call to getContacts
   * var contactId = 'AAMkADVhYTYwNzk...';
   *
   * // Pass the user's email address
   * var userInfo = {
   *   email: 'sarad@contoso.com'
   * };
   *
   * outlook.contacts.deleteContact({token: token, contactId: contactId, user: userInfo},
   *   function(error, result){
   *     if (error) {
   *       console.log('deleteContact returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('SUCCESS');
   *     }
   *   });
   */
  deleteContact: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = base.apiEndpoint() + userSpec + '/Contacts/' + parameters.contactId;

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
  }
};