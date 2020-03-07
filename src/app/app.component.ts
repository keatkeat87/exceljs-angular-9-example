import { Component } from '@angular/core';

import 'core-js/modules/es.promise';
import 'core-js/modules/es.object.assign';
import 'core-js/modules/es.object.keys';
import 'regenerator-runtime/runtime';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'exceljs';

  async download() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');
    worksheet.columns = [
      { key: 'string', width: 20, header: 'String' },
      { key: 'textarea', width: 20, header: 'Textarea', style: { alignment: { wrapText: true } } },
      { key: 'number', width: 20, header: 'Number', style: { numFmt: '#,##0.00#######' } },
      { key: 'boolean', width: 20, header: 'Boolean' },
      { key: 'date', width: 20, header: 'Date', style: { numFmt: 'dd-MM-yyyy' } },
      { key: 'datetime', width: 20, header: 'Datetime', style: { numFmt: 'dd-MM-yyyy hh:mm AM/PM' } }
    ];

    worksheet.addRow({
      string: 'string',
      textarea: 'textarea\r\ntextarea',
      number: 55.1,
      boolean: true,
      date: new Date(2020, 0, 1),
      datetime: new Date(Date.UTC(2020, 0, 1, 14, 55, 12)),
    });

    worksheet.addRow({
      string: 'string',
      textarea: 'textarea\r\ntextarea',
      number: 55.123,
      boolean: true,
      date: new Date(2020, 0, 1),
      datetime: new Date(2020, 0, 1, 14, 50, 32, 123)
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    fs.saveAs(blob, 'CarData.xlsx');
  }
}
