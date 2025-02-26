import {
  BadRequestException,
  Body,
  Controller,
  Post,
  Res,
} from '@nestjs/common';
import { ExcelService } from './excel.service';
import { Response } from 'express';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  // Export Excel file
  @Post('export')
  async exportToExcel(
    @Body() data: any[],
    @Res() res: Response,
  ): Promise<void> {
    if (!Array.isArray(data) || data.length === 0) {
      throw new BadRequestException('Invalid data format');
    }
    return this.excelService.exportToExcel(data, res);
  }
}
