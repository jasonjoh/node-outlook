// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var fs = require("fs");
var path = require("path");
var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

var currentDirectory = path.dirname(fs.realpathSync(__filename));

var exchangefile = fs.readFileSync(path.join(currentDirectory, './exchange-lib/exchange.js'), 'utf8');
eval(exchangefile);
var utilityfile = fs.readFileSync(path.join(currentDirectory, './exchange-lib/utility.js'), 'utf8');
eval(utilityfile);

exports.Microsoft = Microsoft;
exports.base = require('./version-2.js');
exports.mail = require('./mail-api.js');
exports.calendar = require('./calendar-api.js');
exports.contacts = require('./contacts-api.js');