import { BadRequestException, Injectable } from '@nestjs/common';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { Response } from 'express';

@Injectable()
export class ExcelService {
  private readonly templatePath = path.join(
    __dirname,
    '../../templates/template.xlsx',
  ); // Template file path

  // Export data to Excel using a template
  async exportToExcel(data: any[], res: Response): Promise<void> {
    if (!fs.existsSync(this.templatePath)) {
      throw new BadRequestException('Template file not found');
    }

    // Load the template file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(this.templatePath);
    const worksheet = workbook.getWorksheet('Forecast Failure-Template(New)');

    if (!worksheet) {
      throw new Error(
        'Worksheet "Forecast Failure-Template(New)" not found in the template file.',
      );
    }

    // Extract headers from the uploaded file, trim whitespace, and remove empty values
    let headers: string[] = worksheet.getRow(1).values as string[];
    if (headers) {
      headers = headers
        .map((header) => (typeof header === 'string' ? header.trim() : header))
        .filter((header) => header.length > 0);
    }

    // Append JSON data to worksheet
    data.forEach((item, index) => {
      const row = worksheet.getRow(index + 2); // Start from second row since Row 2 onwards is where the data should be written.
      headers.forEach((header, colIndex) => {
        row.getCell(colIndex + 1).value = item[header] || ''; // Fetches the correct cell in the row and Write Data to Each Cell
      });
      row.commit();
    });

    // Generate buffer and send response
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="exported_data.xlsx"',
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.send(buffer);
  }
}
