var fs = require('fs');
var fsExtra = require('fs-extra');
var json2csv = require('json2csv');
var Papa = require('papaparse');
var path = require('path');
var XLSX = require('xlsx');

var sourceExportPath = path.join(__dirname, 'export.csv');

var stream = fs.createReadStream(sourceExportPath, 'UTF-8');

var townCityFolderPath = path.join(__dirname, 'exports', 'towns-cities');
var countyFolderPath = path.join(__dirname, 'exports', 'counties');
var typeRatingFolderPath = path.join(__dirname, 'exports', 'types-and-ratings');
var routeFolderPath = path.join(__dirname, 'exports', 'routes');

fsExtra.emptyDirSync(townCityFolderPath);
fsExtra.emptyDirSync(countyFolderPath);
fsExtra.emptyDirSync(typeRatingFolderPath);
fsExtra.emptyDirSync(routeFolderPath);

var headers = [
  'Organisation Name',
  'Town / City',
  'County',
  'Type & Rating',
  'Route',
  'Google',
  'Facebook',
  'LinkedIn',
  'Bing',
  'Gov UK',
];

var townCityHash = {};
var countyHash = {};
var typeRatingHash = {};
var routeHash = {};

function processRow(data) {
  var orgName = data['Organisation Name'];
  var townCity = data['Town/City'];
  var townCityKey = generateKey(townCity);
  var county = data['County'];
  var countyKey = generateKey(county);
  var typeRating = data['Type & Rating'];
  var typeRatingKey = generateKey(typeRating);
  var route = data['Route'];
  var routeKey = generateKey(route);

  var row = [
    orgName,
    townCity,
    county,
    typeRating,
    route,
    'https://www.google.co.uk/search?q=' + encodeURIComponent(orgName),
    'https://www.facebook.com/search/top?q=' + encodeURIComponent(orgName),
    'https://www.linkedin.com/search/results/companies/?keywords=' + encodeURIComponent(orgName),
    'https://www.bing.com/search?q=' + encodeURIComponent(orgName),
    'https://find-and-update.company-information.service.gov.uk/search?q=' + encodeURIComponent(orgName),
  ].reduce(function (result, value, index) {
    result[headers[index]] = value;

    return result;
  }, {});

  if (!townCityHash.hasOwnProperty(townCityKey)) {
    townCityHash[townCityKey] = [];
  }
  townCityHash[townCityKey].push(row);

  if (!countyHash.hasOwnProperty(countyKey)) {
    countyHash[countyKey] = [];
  }
  countyHash[countyKey].push(row);

  if (!typeRatingHash.hasOwnProperty(typeRatingKey)) {
    typeRatingHash[typeRatingKey] = [];
  }
  typeRatingHash[typeRatingKey].push(row);

  if (!routeHash.hasOwnProperty(routeKey)) {
    routeHash[routeKey] = [];
  }
  routeHash[routeKey].push(row);
}

function generateKey(text) {
  if (!text) {
    return 'unknown';
  }

  return (
    text
      // To lower case
      .toLowerCase()

      // Only keep letters
      .replace(/[^a-zA-Z0-9 ]/g, '')

      // Trim
      .replace(/ {5}/g, ' ')
      .replace(/ {4}/g, ' ')
      .replace(/ {3}/g, ' ')
      .replace(/ {2}/g, ' ')
      .trim()

      // Replace spaces with underscores
      .replace(/ /g, '_')
  );
}

function attachLinksToXlsx(xlsxSheet) {
  Object.keys(xlsxSheet.Sheets.Sheet1).forEach(function (key) {
    var column = key[0];
    var row = parseInt(key.substr(1));

    if (row === 1) {
      return;
    }

    if (['f', 'g', 'h', 'i', 'j'].indexOf(column.toLowerCase()) !== -1) {
      xlsxSheet.Sheets.Sheet1[key].l = { Target: xlsxSheet.Sheets.Sheet1[key].v };
    }
  });
}

// Instructions for reading data
Papa.parse(stream, {
  header: true,

  step: function (results, parser) {
    processRow(results.data);
  },

  error: function () {
    console.log('err: ', arguments);
  },

  complete: function () {
    Object.keys(townCityHash).forEach(function (key) {
      var townCityXlsxFile = path.join(townCityFolderPath, key + '.xlsx');

      fs.closeSync(fs.openSync(townCityXlsxFile, 'w'));

      try {
        var townCityCsv = json2csv({ data: townCityHash[key], fields: headers });
        var townCityXlsx = XLSX.read(townCityCsv, { type: 'binary' });

        attachLinksToXlsx(townCityXlsx);

        fs.appendFileSync(townCityXlsxFile, XLSX.write(townCityXlsx, { type: 'buffer', bookType: 'xlsx' }));
      } catch (err) {
        console.error(err);
      }
    });

    Object.keys(countyHash).forEach(function (key) {
      var countyXlsxFile = path.join(countyFolderPath, key + '.xlsx');

      fs.closeSync(fs.openSync(countyXlsxFile, 'w'));

      try {
        var countyCsv = json2csv({ data: countyHash[key], fields: headers });
        var countyXlsx = XLSX.read(countyCsv, { type: 'binary' });

        attachLinksToXlsx(countyXlsx);

        fs.appendFileSync(countyXlsxFile, XLSX.write(countyXlsx, { type: 'buffer', bookType: 'xlsx' }));
      } catch (err) {
        console.error(err);
      }
    });

    Object.keys(typeRatingHash).forEach(function (key) {
      var typeRatingXlsxFile = path.join(typeRatingFolderPath, key + '.xlsx');

      fs.closeSync(fs.openSync(typeRatingXlsxFile, 'w'));

      try {
        var typeRatingCsv = json2csv({ data: typeRatingHash[key], fields: headers });
        var typeRatingXlsx = XLSX.read(typeRatingCsv, { type: 'binary' });

        attachLinksToXlsx(typeRatingXlsx);

        fs.appendFileSync(typeRatingXlsxFile, XLSX.write(typeRatingXlsx, { type: 'buffer', bookType: 'xlsx' }));
      } catch (err) {
        console.error(err);
      }
    });

    Object.keys(routeHash).forEach(function (key) {
      var routeXlsxFile = path.join(routeFolderPath, key + '.xlsx');

      fs.closeSync(fs.openSync(routeXlsxFile, 'w'));

      try {
        var routeCsv = json2csv({ data: routeHash[key], fields: headers });
        var routeXlsx = XLSX.read(routeCsv, { type: 'binary' });

        attachLinksToXlsx(routeXlsx);

        fs.appendFileSync(routeXlsxFile, XLSX.write(routeXlsx, { type: 'buffer', bookType: 'xlsx' }));
      } catch (err) {
        console.error(err);
      }
    });
  },
});
