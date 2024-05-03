'use strict';

const fs = require('node:fs');
const officeprops = require('officeprops');
console.log(officeprops);

void async function () {
  const { editable, readOnly } = await officeprops.getData(fs.readFileSync('./test.docx'));
  console.log({ editable, readOnly });
}();
