var fs = require('fs');
var fsExtra = require('fs-extra');
var json2csv = require('json2csv');
var Papa = require('papaparse');
var path = require('path');
var XLSX = require('xlsx');

var config = {
  generateTownCity: true,
  generateCounty: true,
  generateTypeRating: true,
  generateRoute: true,
  bookType: 'xlsx',
  extension: '.xlsx',
  maxRowsPerFile: 500,
};

var sourceExportPath = path.join(__dirname, 'export.csv');

var stream = fs.createReadStream(sourceExportPath, 'UTF-8');

var exportsFolderPath = path.join(__dirname, 'exports');
fsExtra.ensureDirSync(exportsFolderPath);

var townCityFolderPath = path.join(exportsFolderPath, 'towns-cities');
if (config.generateTownCity) {
  fsExtra.ensureDirSync(townCityFolderPath);
  fsExtra.emptyDirSync(townCityFolderPath);
}

var countyFolderPath = path.join(exportsFolderPath, 'counties');
if (config.generateCounty) {
  fsExtra.ensureDirSync(countyFolderPath);
  fsExtra.emptyDirSync(countyFolderPath);
}

var typeRatingFolderPath = path.join(exportsFolderPath, 'types-and-ratings');
if (config.generateTypeRating) {
  fsExtra.ensureDirSync(typeRatingFolderPath);
  fsExtra.emptyDirSync(typeRatingFolderPath);
}

var routeFolderPath = path.join(exportsFolderPath, 'routes');
if (config.generateRoute) {
  fsExtra.ensureDirSync(routeFolderPath);
  fsExtra.emptyDirSync(routeFolderPath);
}

var sentenceCaseRegex = /(^\w{1}|\.\s*\w{1})/gi;

var headers = ['Organisation Name', 'Town / City', 'County', 'Type & Rating', 'Route', 'Google', 'LinkedIn', 'Gov UK'];

var townCityHash = {};
var countyHash = {};
var typeRatingHash = {};
var routeHash = {};

function addRow(key, data, hash) {
  var orgName = formatValue(data['Organisation Name']);
  var encodedName = orgName ? encodeURIComponent(orgName) : '';

  var row = [
    orgName,
    formatValue(data['Town/City']),
    formatValue(data['County']),
    formatValue(data['Type & Rating']),
    formatValue(data['Route']),
    encodedName ? 'https://www.google.co.uk/search?q=' + encodedName : encodedName,
    encodedName ? 'https://www.linkedin.com/search/results/companies/?keywords=' + encodedName : encodedName,
    encodedName ? 'https://find-and-update.company-information.service.gov.uk/search?q=' + encodedName : encodedName,
  ].reduce(function (result, value, index) {
    result[headers[index]] = value;

    return result;
  }, {});

  var allKey = key + 'All';
  if (!hash.hasOwnProperty(allKey)) {
    hash[allKey] = [];
  }
  hash[allKey].push(hash[allKey].length);

  var groupNumber = Math.ceil(hash[allKey].length / config.maxRowsPerFile);
  var groupKey = groupNumber > 1 ? key + groupNumber.toString() : key;
  if (!hash.hasOwnProperty(groupKey)) {
    hash[groupKey] = [];
  }
  hash[groupKey].push(row);
}

function sanitiseValue(text) {
  if (!text) {
    return '';
  }

  return text
    .replace(/ {7}/g, ' ')
    .replace(/ {6}/g, ' ')
    .replace(/ {5}/g, ' ')
    .replace(/ {4}/g, ' ')
    .replace(/ {3}/g, ' ')
    .replace(/ {2}/g, ' ')
    .trim();
}

function generateKey(text) {
  text = sanitiseValue(text);

  if (!text) {
    return 'unknown';
  }

  return text
    .toLowerCase()
    .replace(/\\&/g, ' and ')
    .replace(/\//g, ' or ')
    .replace(/[^A-Za-z0-9 ]/g, '')
    .replace(/ /g, '_');
}

function formatValue(text) {
  text = sanitiseValue(text);

  if (!text) {
    return '';
  }

  return text
    .replace(/[^A-Za-z0-9_\- ]/g, '')
    .toLowerCase()
    .replace(sentenceCaseRegex, function (firstLetter) {
      return firstLetter.toUpperCase();
    });
}

function attachLinks(xlsxSheet) {
  Object.keys(xlsxSheet.Sheets.Sheet1).forEach(function (key) {
    var column = key[0];
    var row = parseInt(key.substr(1));

    if (row === 1) {
      return;
    }

    if (
      ['f', 'g', 'h'].indexOf(column.toLowerCase()) !== -1 &&
      xlsxSheet.Sheets.Sheet1[key].v &&
      xlsxSheet.Sheets.Sheet1[key].v.indexOf('://') !== -1
    ) {
      xlsxSheet.Sheets.Sheet1[key].l = { Target: xlsxSheet.Sheets.Sheet1[key].v };
    }
  });
}

// Instructions for reading data
Papa.parse(stream, {
  header: true,

  step: function (results, parser) {
    if (config.generateTownCity) {
      addRow(generateKey(results.data['Town/City']), results.data, townCityHash);
    }
    if (config.generateCounty) {
      addRow(generateKey(results.data['County']), results.data, countyHash);
    }
    if (config.generateTypeRating) {
      addRow(generateKey(results.data['Type & Rating']), results.data, typeRatingHash);
    }
    if (config.generateRoute) {
      addRow(generateKey(results.data['Route']), results.data, routeHash);
    }
  },

  error: function () {
    console.log('err: ', arguments);
  },

  complete: function () {
    if (config.generateTownCity) {
      Object.keys(townCityHash).forEach(function (key) {
        if (key.indexOf('All') !== -1) {
          return;
        }

        var townCityFile = path.join(townCityFolderPath, key + config.extension);

        fs.closeSync(fs.openSync(townCityFile, 'w'));

        try {
          var townCityCsv = json2csv({ data: townCityHash[key], fields: headers });
          var townCityXlsx = XLSX.read(townCityCsv, { type: 'string' });

          attachLinks(townCityXlsx);

          XLSX.writeFile(townCityXlsx, townCityFile, { compression: false });
        } catch (err) {
          console.error(err);
        }
      });
    }

    if (config.generateCounty) {
      Object.keys(countyHash).forEach(function (key) {
        if (key.indexOf('All') !== -1) {
          return;
        }

        var countyFile = path.join(countyFolderPath, key + config.extension);

        fs.closeSync(fs.openSync(countyFile, 'w'));

        try {
          var countyCsv = json2csv({ data: countyHash[key], fields: headers });
          var countyXlsx = XLSX.read(countyCsv, { type: 'string' });

          attachLinks(countyXlsx);

          XLSX.writeFile(countyXlsx, countyFile, { compression: false });
        } catch (err) {
          console.error(err);
        }
      });
    }

    if (config.generateTypeRating) {
      Object.keys(typeRatingHash).forEach(function (key) {
        if (key.indexOf('All') !== -1) {
          return;
        }

        var typeRatingFile = path.join(typeRatingFolderPath, key + config.extension);

        fs.closeSync(fs.openSync(typeRatingFile, 'w'));

        try {
          var typeRatingCsv = json2csv({ data: typeRatingHash[key], fields: headers });
          var typeRatingXlsx = XLSX.read(typeRatingCsv, { type: 'string' });

          attachLinks(typeRatingXlsx);

          XLSX.writeFile(typeRatingXlsx, typeRatingFile, { compression: false });
        } catch (err) {
          console.error(err);
        }
      });
    }

    if (config.generateRoute) {
      Object.keys(routeHash).forEach(function (key) {
        if (key.indexOf('All') !== -1) {
          return;
        }

        var routeFile = path.join(routeFolderPath, key + config.extension);

        fs.closeSync(fs.openSync(routeFile, 'w'));

        try {
          var routeCsv = json2csv({ data: routeHash[key], fields: headers });
          var routeXlsx = XLSX.read(routeCsv, { type: 'string' });

          attachLinks(routeXlsx);

          XLSX.writeFile(routeXlsx, routeFile, { compression: false });
        } catch (err) {
          console.error(err);
        }
      });
    }
  },
});
