import ExcelJS from 'exceljs';
import { fetchFinancialData, fetchCampaignSkus, CampaignInfo } from '../utils/finance-rk';

export async function createFinanceRKSheet(workbook: ExcelJS.Workbook, apiKey: string, startDate: string, endDate: string) {
  const sheet = workbook.addWorksheet('Финансы РК');

  const headers = [
    'ID кампании',
    'Название кампании',
    'Дата',
    'Сумма',
    'Источник списания',
    'Тип операции',
    'Номер документа',
    'SKU ID',
    'Период отчета'
  ];

  sheet.addRow(headers);
  const hdr = sheet.getRow(1);
  hdr.font = { bold: true };
  hdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } } as any;
  [15,30,15,15,20,15,20,40,25].forEach((w,i)=>{ (sheet.getColumn(i+1).width as number) = w; });

  const data = await fetchFinancialData(apiKey, startDate, endDate);
  if (data.length === 0) return sheet;

  const uniq = new Map<number, CampaignInfo>();
  data.forEach(r => { if (!uniq.has(r.advertId)) uniq.set(r.advertId, { advertId: r.advertId, name: r.campName || 'Неизвестная кампания' }); });
  const skus = await fetchCampaignSkus(apiKey, Array.from(uniq.values()));
  const period = `${new Date(startDate).toLocaleDateString('ru-RU')} - ${new Date(endDate).toLocaleDateString('ru-RU')}`;

  for (const r of data) {
    sheet.addRow([
      r.advertId,
      r.campName || 'Неизвестная кампания',
      r.date,
      Number(r.sum) || 0,
      r.bill === 1 ? 'Счет' : 'Баланс',
      r.type,
      r.docNumber || '',
      skus.get(r.advertId) || 'Нет данных',
      period
    ]);
  }

  const sumCol = sheet.getColumn(4);
  sumCol.eachCell((cell, row) => {
    if (row === 1) return;
    (cell as any).numFmt = '[$-419]# ##0,00;[$-419]-# ##0,00';
    cell.value = Number(cell.value) || 0;
  });

  return sheet;
}

export async function attachFinanceRKToWorkbook(workbook: ExcelJS.Workbook, apiKey: string, startDate: string, endDate: string) {
  await createFinanceRKSheet(workbook, apiKey, startDate, endDate);
  return workbook;
}

