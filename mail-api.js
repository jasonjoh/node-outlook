// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var base = require('./version-2.js');

module.exports = {
  /*
    getMessages
    
    Used to get messages from a folder.
    
    Parameters:
    - parameters(object)(required): An object containing all of the relevant parameters.
      Possible values:
      - token (string)(required):       The access token.
      - user (string)(optional):        The SMTP address of the user. If absent, the 'Me'
                                        is used.
      - folderId (string)(optional):    The folder id. If absent, the API calls the /User/Messages endpoint.
        Valid value types:
        - The 'Id' property of a Folder entity
        - A well-known folder name (Inbox, Drafts, SentItems, DeletedItems)
        
      - odataParams (object)(optional): An object containing key/value pairs representing OData
                                        query parameters. See:
                                        https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters
      
    - callback(function)(optional): A callback function that is called when the function completes.
                                    It should have the signature function (error, result).
  */
  getMessages: function(parameters, callback){
    var userSpec = parameters.user === undefined ? '/Me' : '/Users/' + parameters.user;
    var folderSpec = parameters.folderId === undefined ? '' : '/folders/' + parameters.folderId;
    
    var requestUrl = base.apiEndpoint() + userSpec + folderSpec + '/Messages';
    
    var apiOptions = {
      url: requestUrl,
      token: parameters.token
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
          callback('REST request returned ' + response.statusCode, response);
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