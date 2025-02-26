import { Module } from '@nestjs/common';
import { ExcelModule } from './excel/excel.module';

@Module({
  imports: [ExcelModule],
  controllers: [],
  providers: [],
})
export class AppModule {}
