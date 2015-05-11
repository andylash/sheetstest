"use strict";

if (Meteor.isClient) {
  Template.hello.helpers({
    'results': function() {
      return Template.instance().results.get();
    },
    value: function() {
      return this.numericValue || this.value;
    }
  });

  Template.hello.created = function() {
    this.results = new ReactiveVar();
  };


  Template.hello.events({
    'click .js-get-data': function(event, template) {
      var spreadsheet = template.$('.js-spreadsheet-name').val();
      var worksheetId = template.$('.js-worksheet-id').val();
      var range = template.$('.js-range').val();

      console.log("Calling getData with spreadsheet: " + spreadsheet + " worksheetId " + worksheetId +
        " range " + range);

      Meteor.call('getData', spreadsheet, worksheetId, range, function(error, result) {
        if (error) {
          console.log("Error: " + error.reason);
        } else {
          console.log("Success! " + EJSON.stringify(result, {
            indent: true
          }));
          template.results.set(result);
        }
      });
    }
  });
}

var getNumberForChar = function(char) {
  return char.toUpperCase().charCodeAt() - 'A'.charCodeAt() + 1;
};

var rangeRegexp = new RegExp(/^([A-Za-z])(\d+)(:?)(?:([A-Za-z])(\d+))?$/);

var parseRange = function(input) {
  var retVal = {};

  // Legal values are: A1, A1:, A1:B2
  var match = input.match(rangeRegexp);

  if (match) {
    retVal['min-col'] = getNumberForChar(match[1]);
    retVal['min-row'] = parseInt(match[2], 10);
    if (match[4] && match[5]) {
      retVal['max-col'] = getNumberForChar(match[4]);
      retVal['max-row'] = parseInt(match[5], 10);
    } else if (! match[3]) {
      retVal['max-col'] = retVal['min-col'];
      retVal['max-row'] = retVal['min-row'];
    }
  }
  return retVal;
};



if (Meteor.isServer) {
  Meteor.startup(function() {
    // code to run on server at startup
  });
  Meteor.methods({
    'getData': function(sheetkey, worksheetId, range) {
      check(sheetkey, String);
      check(worksheetId, String);
      check(range, String);


      var GoogleSpreadsheet = Meteor.npmRequire('google-spreadsheet');

      var sheet = new GoogleSpreadsheet(sheetkey);
      // var setAuth = Meteor.wrapAsync(sheet.setAuth, sheet);
      // setAuth('andy@opstarts.com','kj@zbdCvogQB#uFar68G');

      var cellOptions = parseRange(range);


      var getCells = Meteor.wrapAsync(sheet.getCells, sheet);
      var data = getCells(worksheetId, cellOptions);
      console.log(EJSON.stringify(data, {
        indent: true
      }));

      // numericValue || value

      return data;
    }

  });
}
