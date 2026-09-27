/**
 * Copyright (c) 2014-2019 Guyon Roche
 * LICENCE: MIT - please refer to LICENSE file included with this module
 */

if (parseInt(process.versions.node.split('.')[0], 10) < 10) {
  throw new Error(
    'For node versions older than 10, please use the ES5 import.'
  );
}

module.exports = require('./lib/exceljs.nodejs.js');
