const fs = require('fs');
const {PassThrough} = require('readable-stream');

const testXformHelper = require('../test-xform-helper');

const TableXform = verquire('xlsx/xform/table/table-xform');
const parseSax = verquire('utils/parse-sax');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: require('./data/table.1.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
    parsedModel: require('./data/table.1.3.json'),
    tests: ['render', 'renderIn', 'parse'],
  },
];

// read and write a table the way XLSX does (parse; prepare + toXml)
const roundTrip = async xml => {
  const stream = new PassThrough();
  stream.write(xml);
  stream.end();
  const xform = new TableXform();
  const model = await xform.parse(parseSax(stream));
  xform.prepare(model, {});
  return xform.toXml(model);
};

const table = (attributes, filterColumns) =>
  `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B3"${attributes}>` +
  `<autoFilter ref="A1:B3">${filterColumns}</autoFilter>` +
  '<tableColumns count="2"><tableColumn id="1" name="key"/><tableColumn id="2" name="value"/></tableColumns>' +
  '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
  '</table>';

describe('TableXform', () => {
  testXformHelper(expectations);

  it('keeps the defaults of a table saved by Excel on a round trip', async () => {
    // Excel leaves out headerRowCount (default 1), totalsRowCount (default 0) and filterColumn
    const xml = await roundTrip(table('', ''));

    expect(xml).to.not.include('headerRowCount');
    expect(xml).to.not.include('totalsRowCount');
    expect(xml).to.not.include('totalsRowShown');
    expect(xml).to.not.include('hiddenButton="1"');
  });

  it('keeps explicit table settings on a round trip', async () => {
    const xml = await roundTrip(
      table(
        ' headerRowCount="0" totalsRowCount="1"',
        '<filterColumn colId="0" hiddenButton="1"/>'
      )
    );

    expect(xml).to.include('headerRowCount="0"');
    expect(xml).to.include('totalsRowCount="1"');
    expect(xml).to.include('<filterColumn colId="0" hiddenButton="1"/>');
  });
});
