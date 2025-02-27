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

    const workbook = await this.loadExcelTemplate();
    const worksheet = this.getWorksheet(
      workbook,
      'Forecast Failure-Template(New)',
    );
    const headers = this.extractHeaders(worksheet);

    this.populateWorksheetWithData(worksheet, data, headers);
    await this.sendExcelResponse(workbook, res);
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
