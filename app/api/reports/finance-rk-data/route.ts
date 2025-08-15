import { NextRequest, NextResponse } from 'next/server';
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

    console.log('🚀 Получение сырых данных "Финансы РК"...');

    // Получаем финансовые данные с буферными днями
    const financialData = await fetchFinancialData(token, startDate, endDate);

    if (financialData.length === 0) {
      return NextResponse.json({ 
        fields: [
          'ID кампании',
          'Название кампании', 
          'Дата',
          'Сумма',
          'Источник списания',
          'Тип операции',
          'Номер документа',
          'SKU ID',
          'Период отчета'
        ],
        rows: []
      });
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
    const skusMap = await fetchCampaignSkus(token, campaignsArray);
    
    // Форматируем период отчета
    const reportPeriod = `${new Date(startDate).toLocaleDateString('ru-RU')} - ${new Date(endDate).toLocaleDateString('ru-RU')}`;
    
    // Подготавливаем сырые данные для Excel (без предварительного форматирования)
    const financeExcelData = financialData.map(record => ({
      "ID кампании": record.advertId,
      "Название кампании": record.campName || 'Неизвестная кампания',
      "Дата": record.date,
      "Сумма": record.sum, // Возвращаем как число, без форматирования
      "Источник списания": record.bill === 1 ? 'Счет' : 'Баланс',
      "Тип операции": record.type,
      "Номер документа": record.docNumber,
      "SKU ID": skusMap.get(record.advertId) || 'Нет данных',
      "Период отчета": reportPeriod
    }));

    const fields = [
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

    console.log(`✅ Сырые данные "Финансы РК" получены. Записей: ${financeExcelData.length}`);

    return NextResponse.json({ fields, rows: financeExcelData });

  } catch (error) {
    console.error('❌ Ошибка получения сырых данных "Финансы РК":', error);
    return NextResponse.json(
      { error: 'Ошибка при получении данных' },
      { status: 500 }
    );
  }
}
