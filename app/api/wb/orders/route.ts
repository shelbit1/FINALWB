import { NextRequest, NextResponse } from 'next/server';

export const runtime = 'nodejs';

interface OrderData {
  date: string;
  lastChangeDate: string;
  warehouseName: string;
  warehouseType: string;
  countryName: string;
  oblastOkrugName: string;
  regionName: string;
  supplierArticle: string;
  nmId: number;
  barcode: string;
  category: string;
  subject: string;
  brand: string;
  techSize: string;
  incomeID: number;
  isSupply: boolean;
  isRealization: boolean;
  totalPrice: number;
  discountPercent: number;
  spp: number;
  finishedPrice: number;
  priceWithDisc: number;
  isCancel: boolean;
  cancelDate: string;
  orderType: string;
  sticker: string;
  gNumber: string;
  srid: string;
}

/**
 * Получить данные о заказах из Wildberries API
 * Метод: /api/v1/supplier/orders
 * Документация: https://statistics-api.wildberries.ru/api/v1/supplier/orders
 */
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { token: requestToken, dateFrom, dateTo, flag = 0 } = body;

    // Используем токен из запроса или из переменных окружения
    const token = requestToken || process.env.WB_API_TOKEN;

    if (!token || !dateFrom) {
      return NextResponse.json(
        { error: 'token и dateFrom обязательны' },
        { status: 400 }
      );
    }

    console.log(`📦 Получение данных о заказах с ${dateFrom}${dateTo ? ` по ${dateTo}` : ''}...`);
    const startTime = Date.now();

    const allOrdersData: OrderData[] = [];
    let currentDateFrom = dateFrom;
    let hasMoreData = true;
    let requestCount = 0;
    const maxRequests = 100; // Защита от бесконечного цикла

    // Получаем данные пакетами (максимум 80000 строк за раз)
    while (hasMoreData && requestCount < maxRequests) {
      requestCount++;
      
      const url = `https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=${encodeURIComponent(currentDateFrom)}&flag=${flag}`;
      
      console.log(`📥 Запрос ${requestCount}: ${url}`);
      
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': token,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`❌ Ошибка получения заказов: ${response.status} ${response.statusText}`);
        console.error(`Ответ: ${errorText}`);
        
        return NextResponse.json(
          { error: `Ошибка получения данных: ${response.status} ${response.statusText}` },
          { status: response.status }
        );
      }

      const data: OrderData[] = await response.json();
      
      console.log(`✅ Получено ${data.length} записей о заказах`);
      
      // Если массив пустой, значит все данные получены
      if (data.length === 0) {
        hasMoreData = false;
        break;
      }

      allOrdersData.push(...data);

      // Если получено меньше строк, чем лимит (80000), значит это последняя пачка
      if (data.length < 80000) {
        hasMoreData = false;
        break;
      }

      // Для следующего запроса используем lastChangeDate последней строки
      const lastRecord = data[data.length - 1];
      if (lastRecord && lastRecord.lastChangeDate) {
        currentDateFrom = lastRecord.lastChangeDate;
        console.log(`📅 Следующий запрос с dateFrom: ${currentDateFrom}`);
      } else {
        hasMoreData = false;
      }

      // Пауза между запросами (лимит: 1 запрос в минуту)
      // Для безопасности ждем 61 секунду
      if (hasMoreData) {
        console.log('⏳ Ожидание 61 секунду перед следующим запросом...');
        await new Promise(resolve => setTimeout(resolve, 61000));
      }
    }

    const endTime = Date.now();
    console.log(`✅ Всего получено ${allOrdersData.length} записей о заказах за ${Math.round((endTime - startTime) / 1000)}с`);

    // Фильтруем данные по периоду, если указан dateTo
    let filteredOrdersData = allOrdersData;
    if (dateTo) {
      filteredOrdersData = allOrdersData.filter(item => {
        const itemDate = item.date ? item.date.split('T')[0] : '';
        return itemDate >= dateFrom && itemDate <= dateTo;
      });
      console.log(`📅 После фильтрации по периоду ${dateFrom} - ${dateTo}: ${filteredOrdersData.length} записей`);
    }

    // Формируем данные для Excel в русском формате
    const fields = [
      'Дата',
      'Дата изменения',
      'Склад',
      'Тип склада',
      'Страна',
      'Округ',
      'Регион',
      'Артикул продавца',
      'Артикул WB',
      'Штрихкод',
      'Категория',
      'Предмет',
      'Бренд',
      'Размер',
      'ID поставки',
      'Поставка',
      'Реализация',
      'Цена без скидки',
      'Скидка %',
      'СПП',
      'Цена после всех скидок',
      'Цена со скидкой',
      'Отменен',
      'Дата отмены',
      'Тип заказа',
      'Стикер',
      'Номер заказа',
      'SRID',
      'Количество'
    ];

    const rows = filteredOrdersData.map((item) => ({
      'Дата': item.date || '',
      'Дата изменения': item.lastChangeDate || '',
      'Склад': item.warehouseName || '',
      'Тип склада': item.warehouseType || '',
      'Страна': item.countryName || '',
      'Округ': item.oblastOkrugName || '',
      'Регион': item.regionName || '',
      'Артикул продавца': item.supplierArticle || '',
      'Артикул WB': item.nmId || 0,
      'Штрихкод': item.barcode || '',
      'Категория': item.category || '',
      'Предмет': item.subject || '',
      'Бренд': item.brand || '',
      'Размер': item.techSize || '',
      'ID поставки': item.incomeID || 0,
      'Поставка': item.isSupply ? 'Да' : 'Нет',
      'Реализация': item.isRealization ? 'Да' : 'Нет',
      'Цена без скидки': item.totalPrice || 0,
      'Скидка %': item.discountPercent || 0,
      'СПП': item.spp || 0,
      'Цена после всех скидок': item.finishedPrice || 0,
      'Цена со скидкой': item.priceWithDisc || 0,
      'Отменен': item.isCancel ? 'Да' : 'Нет',
      'Дата отмены': item.cancelDate || '',
      'Тип заказа': item.orderType || '',
      'Стикер': item.sticker || '',
      'Номер заказа': item.gNumber || '',
      'SRID': item.srid || '',
      'Количество': 1
    }));

    return NextResponse.json({ fields, rows });

  } catch (error) {
    console.error('❌ Ошибка получения данных о заказах:', error);
    const message = error instanceof Error ? error.message : 'Неизвестная ошибка';
    return NextResponse.json(
      { error: `Ошибка при получении данных: ${message}` },
      { status: 500 }
    );
  }
}

