const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 2125 - spliceRows remove last row', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet();
    ws.addRows([['1st'], ['2nd'], ['3rd']]);

    ws.spliceRows(ws.rowCount, 1);

    expect(ws.getRow(ws.rowCount).getCell(1).value).to.equal('2nd');
  });

  it('issue 2125 - spliceRows remove several rows at the end', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet();
    ws.addRows([['1st'], ['2nd'], ['3rd'], ['4th']]);

    ws.spliceRows(3, 2);

    expect(ws.rowCount).to.equal(2);
    expect(ws.getRow(2).getCell(1).value).to.equal('2nd');
  });

  it('issue 2125 - spliceRows remove more rows than there are at the end', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet();
    ws.addRows([['1st'], ['2nd'], ['3rd'], ['4th']]);

    ws.spliceRows(3, 10);

    expect(ws.rowCount).to.equal(2);
    expect(ws.getRow(2).getCell(1).value).to.equal('2nd');
  });

  it('issue 2125 - spliceRows replace rows at the end with fewer rows', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet();
    ws.addRows([['1st'], ['2nd'], ['3rd'], ['4th']]);

    ws.spliceRows(3, 2, ['new']);

    expect(ws.rowCount).to.equal(3);
    expect(ws.getRow(3).getCell(1).value).to.equal('new');
  });
});
