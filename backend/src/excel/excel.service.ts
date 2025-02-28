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

    // IMPLEMENTATION 01 - Load the template and populate with data
    // Issue of this approach -> ExcelJS is unintentionally removing some elements from the template—specifically AutoFilters and Data Validations (Dropdowns).
    const workbook = await this.loadExcelTemplate();
    const worksheet = this.getWorksheet(
      workbook,
      'Forecast Failure-Template(New)',
    );
    const headers = this.extractHeaders(worksheet);

    this.populateWorksheetWithData(worksheet, data, headers);
    await this.sendExcelResponse(workbook, res);
    // ----------------------------end of implementation 1-----------------------------------------

    // IMPLEMENTATION 02 - Load the template, create a new workbook based on the template and populate with data
    // Issue of this approach -> some styles are getting missed in the template
    // const templateWorkbook = await this.loadExcelTemplate();
    // const templateForecastSheet = templateWorkbook.getWorksheet(
    //   'Forecast Failure-Template(New)',
    // );
    // const templateMasterSheet = templateWorkbook.getWorksheet('MASTER Data ');

    // if (!templateForecastSheet || !templateMasterSheet) {
    //   throw new Error('Required worksheets not found in template.');
    // }

    // const newWorkbook = new ExcelJS.Workbook(); // Create a new workbook

    // // Copy Forecast Failure-Template(New) sheet
    // const newForecastSheet = newWorkbook.addWorksheet(
    //   'Forecast Failure-Template(New)',
    // );
    // this.copySheet(templateForecastSheet, newForecastSheet);

    // // Copy MASTER Data sheet
    // const newMasterSheet = newWorkbook.addWorksheet('MASTER Data');
    // this.copySheet(templateMasterSheet, newMasterSheet);

    // const headers = this.extractHeaders(newForecastSheet); // Extract headers
    // this.makeHeadersBold(newForecastSheet);
    // this.makeHeadersBold(newMasterSheet);

    // this.populateWorksheetWithData(newForecastSheet, data, headers); // Populate new Forecast sheet with data
    // this.addDataValidation(newForecastSheet); // Add dropdown values

    // await this.sendExcelResponse(newWorkbook, res);
    // ------------------------------------end of implementation 2-------------------------------------------------------------
  }

  // Copies one sheet to another, preserving styles
  private copySheet(
    sourceSheet: ExcelJS.Worksheet,
    targetSheet: ExcelJS.Worksheet,
  ): void {
    // Copy column widths
    sourceSheet.columns.forEach((col, index) => {
      const targetCol = targetSheet.getColumn(index + 1);
      if (col.width) {
        targetCol.width = col.width;
      }
    });

    // Copy each row
    sourceSheet.eachRow((row, rowIndex) => {
      const newRow = targetSheet.getRow(rowIndex);

      row.eachCell((cell, colIndex) => {
        const newCell = newRow.getCell(colIndex);

        // Copy cell values
        newCell.value = cell.value;
      });

      newRow.commit();
    });

    // // Force set fill color for specific columns (20, 21, 22, 23, 24, 25, 27, 29)
    // targetSheet.eachRow((row, rowIndex) => {
    //   [20, 21, 22, 23, 24, 25, 27, 29].forEach((colIndex) => {
    //     const cell = row.getCell(colIndex);
    //     cell.fill = {
    //       type: 'pattern',
    //       pattern: 'solid',
    //       fgColor: { argb: 'FFF2CC' },
    //     };
    //   });

    //   row.commit();
    // });
  }

  private addDataValidation(sheet: ExcelJS.Worksheet): void {
    // Define the mapping of column index → Dropdown range from "MASTER Data"
    const dropdownMapping: { [key: number]: string } = {
      21: "'MASTER Data'!$B$2:$B$93", // cat1
      22: "'MASTER Data'!$C$2:$C$93", // cat2
      23: "'MASTER Data'!$D$2:$D$93", // cat3
      24: "'MASTER Data'!$E$2:$E$93", // cat4
      25: "'MASTER Data'!$I$2:$I$4", // segregation
      27: "'MASTER Data'!$G$2:$G$6", // failure impact
    };

    // Apply the correct dropdown validation for each column
    Object.entries(dropdownMapping).forEach(([colIndex, range]) => {
      const colNumber = parseInt(colIndex, 10);

      sheet.getColumn(colNumber).eachCell((cell, rowIndex) => {
        if (rowIndex > 1) {
          // Avoid setting dropdown in the header row
          cell.dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [range], // Assign the correct range for this column
          };
        }
      });
    });
  }

  private async loadExcelTemplate(): Promise<ExcelJS.Workbook> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(this.templatePath);
    return workbook;
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
    const headerRow = sheet.getRow(1); // Get the first row (header)

    headerRow.eachCell((cell) => {
      cell.font = { bold: true }; // Apply bold styling
    });

    headerRow.commit(); // Ensure the changes are saved
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
