import { NextRequest, NextResponse } from 'next/server';

export const runtime = 'nodejs';

interface SalesFunnelCard {
  nmID: number;
  vendorCode: string;
  brandName: string;
  tags: Array<{ id: number; name: string }>;
  object: {
    id: number;
    name: string;
  };
  photo: {
    big: string;
    tm: string;
  };
  statistics: {
    selectedPeriod: {
      begin: string;
      end: string;
      openCard: number;
      addToCart: number;
      orders: number;
      avgRubPrice: number;
      ordersSumRub: number;
      stockMpQty: number;
      stockWbQty: number;
      cancelSumRub: number;
      cancelCount: number;
      buyoutCount: number;
      buyoutSumRub: number;
      openCardPercent: number;
      addToCartPercent: number;
      cartToOrderPercent: number;
      buyoutsPercent: number;
    };
    previousPeriod: {
      begin: string;
      end: string;
      openCard: number;
      addToCart: number;
      orders: number;
      avgRubPrice: number;
      ordersSumRub: number;
      stockMpQty: number;
      stockWbQty: number;
      cancelSumRub: number;
      cancelCount: number;
      buyoutCount: number;
      buyoutSumRub: number;
      openCardPercent: number;
      addToCartPercent: number;
      cartToOrderPercent: number;
      buyoutsPercent: number;
    };
  };
}

interface SalesFunnelResponse {
  data: {
    page: number;
    isNextPage: boolean;
    cards: SalesFunnelCard[];
  };
  error: boolean;
  errorText: string;
  additionalErrors: unknown[];
}

/**
 * Получить данные воронки продаж из Wildberries API
 * Метод: /api/v2/nm-report/detail
 * Документация: https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail
 */
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { token: requestToken, startDate, endDate } = body;

    // Используем токен из запроса или из переменных окружения
    const token = requestToken || process.env.WB_API_TOKEN;

    if (!token || !startDate || !endDate) {
      return NextResponse.json(
        { error: 'token, startDate и endDate обязательны' },
        { status: 400 }
      );
    }

    console.log(`📊 Получение данных воронки продаж с ${startDate} по ${endDate}...`);
    const startTime = Date.now();

    const allCards: SalesFunnelCard[] = [];
    let currentPage = 1;
    let hasMorePages = true;

    // Получаем данные постранично
    while (hasMorePages) {
      const requestBody = {
        brandNames: [],
        objectIDs: [],
        tagIDs: [],
        nmIDs: [],
        timezone: "Europe/Moscow",
        period: {
          begin: `${startDate} 00:00:00`,
          end: `${endDate} 23:59:59`
        },
        orderBy: {
          field: "ordersSumRub",
          mode: "desc"
        },
        page: currentPage
      };

      console.log(`📥 Запрос страницы ${currentPage}...`);

      const response = await fetch('https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail', {
        method: 'POST',
        headers: {
          'Authorization': token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`❌ Ошибка получения воронки продаж: ${response.status} ${response.statusText}`);
        console.error(`Ответ: ${errorText}`);
        
        return NextResponse.json(
          { error: `Ошибка получения данных: ${response.status} ${response.statusText}` },
          { status: response.status }
        );
      }

      const data: SalesFunnelResponse = await response.json();

      if (data.error) {
        console.error(`❌ Ошибка API: ${data.errorText}`);
        return NextResponse.json(
          { error: data.errorText || 'Ошибка получения данных' },
          { status: 400 }
        );
      }

      console.log(`✅ Получено ${data.data.cards.length} карточек на странице ${currentPage}`);
      
      allCards.push(...data.data.cards);

      hasMorePages = data.data.isNextPage;
      currentPage++;

      // Пауза между запросами (лимит: 3 запроса в минуту, интервал 20 секунд)
      if (hasMorePages) {
        console.log('⏳ Ожидание 21 секунду перед следующим запросом...');
        await new Promise(resolve => setTimeout(resolve, 21000));
      }
    }

    const endTime = Date.now();
    console.log(`✅ Всего получено ${allCards.length} карточек за ${Math.round((endTime - startTime) / 1000)}с`);

    // Формируем данные для Excel в русском формате
    const fields = [
      'Артикул WB',
      'Артикул продавца',
      'Бренд',
      'Предмет',
      'Переходы в карточку',
      'Добавления в корзину',
      'Заказы',
      'Выкупы',
      'Отмены',
      'Средняя цена (₽)',
      'Сумма заказов (₽)',
      'Сумма выкупов (₽)',
      'Сумма отмен (₽)',
      'Остаток WB (шт)',
      'Остаток МП (шт)',
      '% Переходов в карточку',
      '% Добавлений в корзину',
      '% Корзина → Заказ',
      '% Выкупов',
      'Предыдущий период - Переходы',
      'Предыдущий период - В корзину',
      'Предыдущий период - Заказы',
      'Предыдущий период - Выкупы',
      'Предыдущий период - Отмены',
      'Предыдущий период - Сумма заказов (₽)',
      'Предыдущий период - Сумма выкупов (₽)',
      'Изменение переходов (%)',
      'Изменение добавлений (%)',
      'Изменение заказов (%)',
      'Изменение выкупов (%)'
    ];

    const rows = allCards.map((card) => {
      const selected = card.statistics.selectedPeriod;
      const previous = card.statistics.previousPeriod;
      
      // Расчет изменений в процентах
      const openCardChange = previous.openCard > 0 
        ? ((selected.openCard - previous.openCard) / previous.openCard * 100).toFixed(2)
        : 0;
      const addToCartChange = previous.addToCart > 0
        ? ((selected.addToCart - previous.addToCart) / previous.addToCart * 100).toFixed(2)
        : 0;
      const ordersChange = previous.orders > 0
        ? ((selected.orders - previous.orders) / previous.orders * 100).toFixed(2)
        : 0;
      const buyoutChange = previous.buyoutCount > 0
        ? ((selected.buyoutCount - previous.buyoutCount) / previous.buyoutCount * 100).toFixed(2)
        : 0;

      return {
        'Артикул WB': card.nmID || 0,
        'Артикул продавца': card.vendorCode || '',
        'Бренд': card.brandName || '',
        'Предмет': card.object?.name || '',
        'Переходы в карточку': selected.openCard || 0,
        'Добавления в корзину': selected.addToCart || 0,
        'Заказы': selected.orders || 0,
        'Выкупы': selected.buyoutCount || 0,
        'Отмены': selected.cancelCount || 0,
        'Средняя цена (₽)': selected.avgRubPrice || 0,
        'Сумма заказов (₽)': selected.ordersSumRub || 0,
        'Сумма выкупов (₽)': selected.buyoutSumRub || 0,
        'Сумма отмен (₽)': selected.cancelSumRub || 0,
        'Остаток WB (шт)': selected.stockWbQty || 0,
        'Остаток МП (шт)': selected.stockMpQty || 0,
        '% Переходов в карточку': selected.openCardPercent || 0,
        '% Добавлений в корзину': selected.addToCartPercent || 0,
        '% Корзина → Заказ': selected.cartToOrderPercent || 0,
        '% Выкупов': selected.buyoutsPercent || 0,
        'Предыдущий период - Переходы': previous.openCard || 0,
        'Предыдущий период - В корзину': previous.addToCart || 0,
        'Предыдущий период - Заказы': previous.orders || 0,
        'Предыдущий период - Выкупы': previous.buyoutCount || 0,
        'Предыдущий период - Отмены': previous.cancelCount || 0,
        'Предыдущий период - Сумма заказов (₽)': previous.ordersSumRub || 0,
        'Предыдущий период - Сумма выкупов (₽)': previous.buyoutSumRub || 0,
        'Изменение переходов (%)': openCardChange,
        'Изменение добавлений (%)': addToCartChange,
        'Изменение заказов (%)': ordersChange,
        'Изменение выкупов (%)': buyoutChange
      };
    });

    return NextResponse.json({ fields, rows });

  } catch (error) {
    console.error('❌ Ошибка получения данных воронки продаж:', error);
    const message = error instanceof Error ? error.message : 'Неизвестная ошибка';
    return NextResponse.json(
      { error: `Ошибка при получении данных: ${message}` },
      { status: 500 }
    );
  }
}

