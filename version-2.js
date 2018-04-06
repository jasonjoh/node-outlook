// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var request = require('request');
var uuid = require('node-uuid');
var utilities = require('./utilities.js');

var fiddlerEnabled = false;
var traceFunction = undefined;
var endpoint = 'https://outlook.office.com/api/v1.0';
var defaultAnchor = '';
var defaultTimeZone = '';

/**
 * @module base
 */

module.exports = {
  /**
   * Used to do the actual send of a REST request to the REST endpoint.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.url {string} The full URL of the API endpoint
   * @param parameters.token {string} The access token for authentication
   * @param [parameters.user.email] {string} The user's SMTP email address, used to set the `X-AnchorMailbox` header.
   * @param [parameters.user.timezone] {string} The user's time zone, used to set the `outlook.timezone` `Prefer` header.
   * @param [parameters.method] {string} Used to specify the HTTP method. Default is 'GET'.
   * @param [parameters.query] {object} An object containing key/value pairs. The pairs will be serialized into a query string.
   * @param [parameters.payload] {object} A JSON-serializable object representing the request body.
   * @param [parameters.headers] {object} A JSON-serializable object representing custom headers to send with the request.
   * @param [callback] {function} A callback function that is called when the function completes. It should have the signature `function (error, result)`.
   */
  makeApiCall: function (parameters, callback) {
    // Check required parameters
    if (parameters.url === undefined || parameters.token === undefined) {
      trace('makeApiCall - ERROR: Missing required parameter');
      if (typeof callback === 'function') {
        callback('ERROR: You must include the \'url\' and \'token\' parameters.');
      }
      return;
    }

    var method = parameters.method === undefined ? 'GET' : parameters.method;

    trace('url: ' + parameters.url);
    trace('token: ' + parameters.token);
    trace('method: ' + method);

    var auth = {
      'bearer': parameters.token
    };

    var headers = parameters.headers || {};
    headers['Accept'] = headers['Accept'] || 'application/json';
    headers['User-Agent'] = headers['User-Agent'] || 'node-outlook/2.0';
    headers['client-request-id'] = headers['client-request-id'] || uuid.v4();
    headers['return-client-request-id'] = headers['return-client-request-id'] || 'true';

    // Determine if we have an anchor mailbox to use
    // Passed parameter has greater priority than module-level default
    var anchorMbx = '';
    if (parameters.user && parameters.user.email && parameters.user.email.length > 0) {
      anchorMbx = parameters.user.email;
    }
    else {
      anchorMbx = defaultAnchor;
    }

    if (anchorMbx.length > 0) {
      headers['X-Anchor-Mailbox'] = anchorMbx;
    }

    // Determine if we have a time zone to use
    // Passed parameter has greater priority than module-level default
    var timezone = '';
    if (parameters.user && parameters.user.timezone && parameters.user.timezone.length > 0) {
      timezone = parameters.user.timezone;
    }
    else {
      timezone = defaultTimeZone;
    }

    if (timezone.length > 0) {
      headers['Prefer'] = headers['Prefer'] || [];
      headers['Prefer'].push('outlook.timezone = "' + timezone + '"');
    }

    var options = {
      method: method,
      url: parameters.url,
      headers: headers,
      auth: auth,
      json: true
    };

    if (parameters.query !== undefined) {
      trace('query:' + JSON.stringify(parameters.query));
      options['qs'] = parameters.query;
    }

    if (fiddlerEnabled) {
      options['proxy'] = 'http://127.0.0.1:8888';
      options['strictSSL'] = false;
    }

    if (method.toUpperCase() === 'POST' || method.toUpperCase() === 'PATCH') {
      if (parameters.payload !== undefined) {
        trace('payload:' + JSON.stringify(parameters.payload));
      }
      options['body'] = parameters.payload;
    }

    request(options, function(error, response, body) {
      if (typeof callback === 'function') {
        callback(error, response);
      }
    });
  },

  /**
   * Used to get information about a user.
   *
   * @param parameters {object} An object containing all of the relevant parameters. Possible values:
   * @param parameters.token {string} The access token.
   * @param [parameters.useMe] {boolean} If true, use the `/Me` segment instead of the `/Users/<email>` segment. This parameter defaults to false and is ignored if the `parameters.user.email` parameter isn't provided (the `/Me` segment is always used in this case).
   * @param [parameters.user.email] {string} The SMTP address of the user. If absent, the `/Me` segment is used in the API URL.
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
   *   '$select': 'DisplayName, EmailAddress',
   * };
   *
   * outlook.base.getUser({token: token, odataParams: queryParams},
   *   function(error, result) {
   *     if (error) {
   *       console.log('getUser returned an error: ' + error);
   *     }
   *     else if (result) {
   *       console.log('User name:', result.DisplayName);
   *       console.log('User email:', result.EmailAddress);
   *     }
   *   });
   */
  getUser: function(parameters, callback) {
    var userSpec = utilities.getUserSegment(parameters);

    var requestUrl = this.apiEndpoint() + userSpec;

    var apiOptions = {
      url: requestUrl,
      token: parameters.token,
      user: parameters.user
    };

    if (parameters.odataParams !== undefined) {
      apiOptions['query'] = parameters.odataParams;
    }

    this.makeApiCall(apiOptions, function(error, response) {
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
   * Used to provide a tracing function.
   *
   * @param traceFunc {function} A function that takes a string parameter. The string parameter contains the text to add to the trace.
   */
  setTraceFunc: function(traceFunc) {
    traceFunction = traceFunc;
  },

  /**
   * Used to enable network sniffing with Fiddler.
   *
   * @param enabled {boolean} `true` to enable default Fiddler proxy and disable SSL verification. `false` to disable proxy and enable SSL verification.
   */
  setFiddlerEnabled: function(enabled) {
    fiddlerEnabled = enabled;
  },

  /**
   * Gets the API endpoint URL.
   * @return {string}
   */
  apiEndpoint: function() {
    return endpoint;
  },

  /**
   * Sets the API endpoint URL. If not called, the default of `https://outlook.office.com/api/v1.0` is used.
   *
   * @param newEndPoint {string} The API endpoint URL to use.
   */
  setApiEndpoint: function(newEndPoint) {
    endpoint = newEndPoint;
  },

  /**
   * Gets the default anchor mailbox address.
   * @return {string}
   */
  anchorMailbox: function() {
    return defaultAnchor;
  },

  /**
   * Sets the default anchor mailbox address.
   *
   * @param newAnchor {string} The SMTP address to send in the `X-Anchor-Mailbox` header.
   */
  setAnchorMailbox: function(newAnchor) {
    defaultAnchor = newAnchor;
  },

  /**
   * Gets the default preferred time zone.
   * @return {string}
   */
  preferredTimeZone: function() {
    return defaultTimeZone;
  },

  /**
   * Sets the default preferred time zone.
   *
   * @param preferredTimeZone {string} The time zone in which the server should return date time values.
   */
  setPreferredTimeZone: function(preferredTimeZone) {
    defaultTimeZone = preferredTimeZone;
  }
};

/**
 * @private
 */
function trace(message) {
  if (typeof traceFunction === 'function') {
    traceFunction(message);
  }
}