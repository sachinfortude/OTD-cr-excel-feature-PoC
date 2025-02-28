import { Injectable, NotFoundException } from '@nestjs/common';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { Response } from 'express';

@Injectable()
export class ExcelService {
  private readonly templatePath = path.join(
    __dirname,
    '../../templates/template.xlsx',
  );

  async exportToExcel(data: any[], res: Response): Promise<void> {
    if (!fs.existsSync(this.templatePath)) {
      throw new NotFoundException('Template file not found');
    }

    // Choose the implementation based requirements
    await this.useImplementation1(data, res);
    // await this.useImplementation2(data, res);
  }

  // IMPLEMENTATION 01 - Load the template and populate with data
  // Issue of this approach -> When modifying the existing template file, ExcelJS does not fully preserve excel features like filtering, data validation, leading to corruption in the exported file.
  private async useImplementation1(data: any[], res: Response): Promise<void> {
    const workbook = await this.loadExcelTemplate();
    const worksheet = this.getWorksheet(
      workbook,
      'Forecast Failure-Template(New)',
    );
    const headers = this.extractHeaders(worksheet);

    this.populateWorksheetWithData(worksheet, data, headers);
    await this.sendExcelResponse(workbook, res);
  }

  // IMPLEMENTATION 02 - Load the template, create a new workbook based on the template and populate with data
  private async useImplementation2(data: any[], res: Response): Promise<void> {
    const templateWorkbook = await this.loadExcelTemplate();
    const templateForecastSheet = templateWorkbook.getWorksheet(
      'Forecast Failure-Template(New)',
    );
    const templateMasterSheet = templateWorkbook.getWorksheet('MASTER Data ');

    if (!templateForecastSheet || !templateMasterSheet) {
      throw new Error('Required worksheets not found in template.');
    }

    const newWorkbook = new ExcelJS.Workbook(); // Create a new workbook

    // Copy Forecast Failure-Template(New) sheet
    const newForecastSheet = newWorkbook.addWorksheet(
      'Forecast Failure-Template(New)',
    );
    this.copySheet(templateForecastSheet, newForecastSheet);

    // Copy MASTER Data sheet
    const newMasterSheet = newWorkbook.addWorksheet('MASTER Data');
    this.copySheet(templateMasterSheet, newMasterSheet);

    const headers = this.extractHeaders(newForecastSheet); // Extract headers
    this.makeHeadersBold(newForecastSheet);
    this.makeHeadersBold(newMasterSheet);
    this.applyFillColorToHeaders(newForecastSheet);
    this.applyFillColorToHeaders(newMasterSheet);

    this.populateWorksheetWithData(newForecastSheet, data, headers); // Populate new Forecast sheet with data
    this.applyFillColorsToDataCells(newForecastSheet);
    this.addDataValidation(newForecastSheet); // Add dropdown values

    await this.sendExcelResponse(newWorkbook, res);
  }

  private async loadExcelTemplate(): Promise<ExcelJS.Workbook> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(this.templatePath);
    return workbook;
  }

  private copySheet(
    sourceSheet: ExcelJS.Worksheet,
    targetSheet: ExcelJS.Worksheet,
  ): void {
    sourceSheet.columns.forEach((col, index) => {
      const targetCol = targetSheet.getColumn(index + 1);
      if (col.width) {
        targetCol.width = col.width;
      }
    });

    sourceSheet.eachRow((row, rowIndex) => {
      const newRow = targetSheet.getRow(rowIndex);

      row.eachCell((cell, colIndex) => {
        const newCell = newRow.getCell(colIndex);
        newCell.value = cell.value;
      });

      newRow.commit();
    });
  }

  private extractHeaders(worksheet: ExcelJS.Worksheet): string[] {
    let headers: string[] = worksheet.getRow(1).values as string[];
    return headers
      ? headers
          .map((header) =>
            typeof header === 'string' ? header.trim() : header,
          )
          .filter((header) => header.length > 0)
      : [];
  }

  private makeHeadersBold(sheet: ExcelJS.Worksheet): void {
    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
    });
    headerRow.commit();
  }

  private applyFillColorToHeaders(sheet: ExcelJS.Worksheet): void {
    const columnColorMapping: { [key: string]: { [key: number]: string } } = {
      'Forecast Failure-Template(New)': {
        20: '#FFF2CC',
        21: '#FFF2CC',
        22: '#FFF2CC',
        23: '#FFF2CC',
        24: '#FFF2CC',
        25: '#FFF2CC',
        27: '#FFF2CC',
        29: '#FFF2CC',
      },
      'MASTER Data': {
        1: '#A5A5A5',
        2: '#A5A5A5',
        3: '#A5A5A5',
        4: '#A5A5A5',
        5: '#A5A5A5',
        7: '#A5A5A5',
        9: '#A5A5A5',
      },
    };

    const sheetName = sheet.name;
    const colors = columnColorMapping[sheetName];

    if (!colors) return;

    const headerRow = sheet.getRow(1);

    headerRow.eachCell((cell, colNumber) => {
      if (colors[colNumber]) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: colors[colNumber].replace('#', '') },
        };
      }
    });

    headerRow.commit();
  }

  private populateWorksheetWithData(
    worksheet: ExcelJS.Worksheet,
    data: any[],
    headers: string[],
  ): void {
    data.forEach((item, index) => {
      const row = worksheet.getRow(index + 2);
      headers.forEach((header, colIndex) => {
        row.getCell(colIndex + 1).value = item[header] || '';
      });
      row.commit();
    });
  }

  private applyFillColorsToDataCells(worksheet: ExcelJS.Worksheet) {
    worksheet.eachRow((row, rowIndex) => {
      [20, 21, 22, 23, 24, 25, 27, 29].forEach((colIndex) => {
        const cell = row.getCell(colIndex);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF2CC' },
        };
      });

      row.commit();
    });
  }

  private addDataValidation(sheet: ExcelJS.Worksheet): void {
    const dropdownMapping: { [key: number]: string } = {
      21: "'MASTER Data'!$B$2:$B$93", // cat1
      22: "'MASTER Data'!$C$2:$C$93", // cat2
      23: "'MASTER Data'!$D$2:$D$93", // cat3
      24: "'MASTER Data'!$E$2:$E$93", // cat4
      25: "'MASTER Data'!$I$2:$I$4", // segregation
      27: "'MASTER Data'!$G$2:$G$6", // failure impact
    };

    Object.entries(dropdownMapping).forEach(([colIndex, range]) => {
      const colNumber = parseInt(colIndex, 10);

      sheet.getColumn(colNumber).eachCell((cell, rowIndex) => {
        if (rowIndex > 1) {
          cell.dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [range],
          };
        }
      });
    });
  }

  private getWorksheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
  ): ExcelJS.Worksheet {
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(
        `Worksheet "${sheetName}" not found in the template file.`,
      );
    }
    return worksheet;
  }

  private async sendExcelResponse(
    workbook: ExcelJS.Workbook,
    res: Response,
  ): Promise<void> {
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=OTD_Failure_Categorization.xlsx',
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    await workbook.xlsx.write(res);
    res.end();
  }
}
