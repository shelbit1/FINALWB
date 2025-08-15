import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { fetchCampaignSkus, fetchFinancialData } from '../../../../finance-utils';

export const runtime = "nodejs";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { startDate, endDate, token } = body;

    if (!token || !startDate || !endDate) {
      return NextResponse.json({ 
        error: 'token, startDate и endDate обязательны' 
      }, { status: 400 });
    }

    console.log('🚀 Начало создания отчета "Финансы РК"...');
    const startTime = Date.now();

    // Получаем финансовые данные с буферными днями
    const financialData = await fetchFinancialData(token, startDate, endDate);

    if (financialData.length === 0) {
      return NextResponse.json({ 
        error: 'Нет данных за указанный период' 
      }, { status: 404 });
    }

    // Собираем уникальные кампании из финансовых данных
    const uniqueCampaigns = new Map();
    financialData.forEach(record => {
      if (!uniqueCampaigns.has(record.advertId)) {
        uniqueCampaigns.set(record.advertId, {
          advertId: record.advertId,
          name: record.campName || 'Неизвестная кампания'
        });
      }
    });
    
    // Получаем детальные данные артикулов для столбца SKU ID
    const campaignsArray = Array.from(uniqueCampaigns.values());
    console.log(`📊 Найдено ${campaignsArray.length} уникальных кампаний`);
    const skusMap = await fetchCampaignSkus(token, campaignsArray);
    
    // Создаем Excel файл (точно как у конкурента)
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([]);

    // Заголовки (точно как у конкурента)
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

    // Добавляем заголовки
    XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: 'A1' });

    // Настройка ширины колонок (как у конкурента)
    const columnWidths = [
      { wch: 15 }, // ID кампании
      { wch: 30 }, // Название кампании
      { wch: 15 }, // Дата
      { wch: 15 }, // Сумма
      { wch: 20 }, // Источник списания
      { wch: 15 }, // Тип операции
      { wch: 20 }, // Номер документа
      { wch: 40 }, // SKU ID
      { wch: 25 }  // Период отчета
    ];
    
    worksheet['!cols'] = columnWidths;

    // Подготавливаем данные для Excel (как у конкурента)
    const reportPeriod = `${new Date(startDate).toLocaleDateString('ru-RU')} - ${new Date(endDate).toLocaleDateString('ru-RU')}`;
    
    const excelData = financialData.map((record, index) => [
      record.advertId,
      record.campName || 'Неизвестная кампания',
      record.date,
      record.sum,
      record.bill === 1 ? 'Счет' : 'Баланс',
      record.type,
      record.docNumber,
      skusMap.get(record.advertId) || 'Нет данных',
      reportPeriod
    ]);

    // Добавляем данные
    XLSX.utils.sheet_add_aoa(worksheet, excelData, { origin: 'A2' });

    // Форматирование заголовков (как у конкурента)
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!worksheet[cellAddress]) continue;
      
      worksheet[cellAddress].s = {
        font: { bold: true },
        fill: { fgColor: { rgb: "E0E0E0" } }
      };
    }

    // Форматируем числовые колонки в российском формате (целые без десятичных, дробные с запятой)
    for (let row = 1; row <= financialData.length; row++) {
      const sumCell = XLSX.utils.encode_cell({ r: row, c: 3 }); // Колонка "Сумма"
      if (worksheet[sumCell]) {
        // Российский формат: 4092 для целых, 4092,10 для дробных (запятая как разделитель)
        worksheet[sumCell].z = '# ##0,##';
        // Убеждаемся, что значение является числом
        const cellValue = worksheet[sumCell].v;
        if (typeof cellValue === 'string') {
          const numValue = parseFloat(String(cellValue).replace(/[^\d.,]/g, '').replace(',', '.'));
          if (!isNaN(numValue)) {
            worksheet[sumCell].v = numValue;
            worksheet[sumCell].t = 'n';
          }
        }
      }
    }

    // Добавляем лист в книгу
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Финансы РК');

    // Генерируем файл
    const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });
    const fileName = `Финансы РК - ${startDate}–${endDate}.xlsx`;

    const endTime = Date.now();
    console.log(`✅ Отчет "Финансы РК" создан за ${endTime - startTime}ms. Записей: ${financialData.length}`);

    return new NextResponse(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(fileName)}"`,
      },
    });

  } catch (error) {
    console.error('❌ Ошибка генерации отчета "Финансы РК":', error);
    return NextResponse.json(
      { error: 'Ошибка при генерации отчета' },
      { status: 500 }
    );
  }
}
