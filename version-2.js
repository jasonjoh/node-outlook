// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var request = require('request');
var uuid = require('node-uuid');

var fiddlerEnabled = false;
var traceFunction = undefined;
var endpoint = 'https://outlook.office365.com/api/v1.0';

module.exports = {
  /*
    makeApiCall
    
    Used to do the actual send of a REST request to the REST endpoint.
    
    Parameters:
    - parameters(object)(required): An object containing all of the relevant parameters.
      Possible values:
      - url (string)(required):     The full URL of the API endpoint
      - token (string)(required):   The access token for authentication
      - method (string)(optional):  Used to specify the HTTP method. Default is 'GET'.
      - query (object)(optional):   An object containing key/value pairs. The pairs will be
                                    serialized into a query string.
      - payload (object)(optional): A JSON-serializable object representing the request body.
      
    - callback(function)(optional): A callback function that is called when the function completes.
                                    It should have the signature function (error, result).
  */
  makeApiCall: function (parameters, callback) {
    // Check required parameters
    if (parameters.url === undefined || parameters.token == undefined) {
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
    
    var headers = {
      'Accept': 'application/json',
      'User-Agent': 'node-outlook/2.0',
      'client-request-id': uuid.v4(),
      'return-client-request-id': 'true'
    };
    
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
  
  /*
    setTraceFunc
    
    Used to provide a tracing function.
    
    Parameters:
    - traceFunc(function)(required): A function that takes a string parameter. The string
                                     parameter contains the text to add to the trace.
  */
  setTraceFunc: function(traceFunc) {
    traceFunction = traceFunc;
  },
  
  /*
    setFiddlerEnabled
    
    Used to enable network sniffing with Fiddler.
    
    Parameters:
    - enabled(bool)(required): 'true' to enable default Fiddler proxy and disable SSL verification.
                               'false' to disable proxy and enable SSL verification.
  */
  setFiddlerEnabled: function(enabled) {
    fiddlerEnabled = enabled;
  },
  
  apiEndpoint: function() {
    return endpoint;
  },
  
  setApiEndpoint: function(newEndPoint) {
    endpoint = newEndPoint;
  }
};

function trace(message) {
  if (typeof traceFunction === 'function') {
    traceFunction(message);
  }
}

/*
  MIT License: 

  Permission is hereby granted, free of charge, to any person obtaining 
  a copy of this software and associated documentation files (the 
  ""Software""), to deal in the Software without restriction, including 
  without limitation the rights to use, copy, modify, merge, publish, 
  distribute, sublicense, and/or sell copies of the Software, and to 
  permit persons to whom the Software is furnished to do so, subject to 
  the following conditions: 

  The above copyright notice and this permission notice shall be 
  included in all copies or substantial portions of the Software. 

  THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
  LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
  OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
  WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/