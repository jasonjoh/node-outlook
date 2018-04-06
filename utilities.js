// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

module.exports = {
  useMeSegment: function(parameters) {
    return parameters.useMe || parameters.user === undefined || parameters.user.email === undefined || parameters.user.email.length <= 0;
  },

  getUserSegment: function(parameters) {
    return this.useMeSegment(parameters) ? '/Me' : '/Users/' + parameters.user.email;
  }
}