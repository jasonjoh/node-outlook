// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var fs = require("fs");
var path = require("path");
var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

var currentDirectory = path.dirname(fs.realpathSync(__filename));

exchangefile = fs.readFileSync(path.join(currentDirectory, './exchange-lib/exchange.js'), 'utf8');
eval(exchangefile);
utilityfile = fs.readFileSync(path.join(currentDirectory, './exchange-lib/utility.js'), 'utf8');
eval(utilityfile);

exports.Microsoft = Microsoft;
exports.base = require('./version-2.js');
exports.mail = require('./mail-api.js');
exports.calendar = require('./calendar-api.js');
exports.contacts = require('./contacts-api.js');

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