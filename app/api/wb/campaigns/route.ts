import { NextRequest, NextResponse } from 'next/server';

export const runtime = "nodejs";

interface Campaign {
  advertId: number;
  type: number;
  status: number;
  name: string;
  createTime?: string;
  changeTime?: string;
  startTime?: string;
  endTime?: string;
  dailyBudget?: number;
  [key: string]: unknown;
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { token: requestToken } = body;

    // Используем токен из запроса или из переменных окружения
    const token = requestToken || process.env.WB_API_TOKEN;

    if (!token) {
      return NextResponse.json({ 
        error: 'Токен обязателен' 
      }, { status: 400 });
    }

    console.log('🚀 Получение списка рекламных кампаний...');

    // Шаг 1: Получаем список всех кампаний
    const countResponse = await fetch('https://advert-api.wildberries.ru/adv/v1/promotion/count', {
      method: 'GET',
      headers: {
        'Authorization': token,
        'Content-Type': 'application/json'
      }
    });

    if (!countResponse.ok) {
      console.error(`❌ Ошибка получения списка кампаний: ${countResponse.status}`);
      return NextResponse.json({ 
        error: `Ошибка API WB: ${countResponse.status}` 
      }, { status: countResponse.status });
    }

    const countData = await countResponse.json();
    console.log(`📊 Получено кампаний: ${countData.all || 0}`);

    // Собираем все ID кампаний из всех статусов и типов
    const allCampaignIds: number[] = [];
    
    if (countData.adverts && Array.isArray(countData.adverts)) {
      countData.adverts.forEach((advert: { advert_list?: Array<{ advertId: number }> }) => {
        if (advert.advert_list && Array.isArray(advert.advert_list)) {
          advert.advert_list.forEach((item: { advertId: number }) => {
            if (item.advertId) {
              allCampaignIds.push(item.advertId);
            }
          });
        }
      });
    }

    console.log(`📊 Найдено ID кампаний: ${allCampaignIds.length}`);

    if (allCampaignIds.length === 0) {
      return NextResponse.json({ 
        fields: [
          'ID кампании',
          'Название кампании',
          'Тип',
          'Статус',
          'Дата создания',
          'Дата изменения',
          'Дата начала',
          'Дата окончания',
          'Дневной бюджет'
        ],
        rows: []
      });
    }

    // Шаг 2: Получаем детальную информацию по кампаниям (порциями по 50)
    const allCampaigns: Campaign[] = [];
    const batchSize = 50;

    for (let i = 0; i < allCampaignIds.length; i += batchSize) {
      const batchIds = allCampaignIds.slice(i, i + batchSize);
      
      try {
        console.log(`📊 Запрос детальной информации для ${batchIds.length} кампаний (партия ${Math.floor(i / batchSize) + 1})...`);
        
        const advertsResponse = await fetch('https://advert-api.wildberries.ru/adv/v1/promotion/adverts', {
          method: 'POST',
          headers: {
            'Authorization': token,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(batchIds)
        });

        if (advertsResponse.ok) {
          const campaignsData = await advertsResponse.json();
          if (Array.isArray(campaignsData)) {
            allCampaigns.push(...campaignsData);
          }
        } else if (advertsResponse.status === 204) {
          console.log(`⚠️ Кампании не найдены для партии ${Math.floor(i / batchSize) + 1}`);
        } else {
          console.error(`❌ Ошибка получения деталей кампаний: ${advertsResponse.status}`);
        }
      } catch (error) {
        console.error(`❌ Ошибка при запросе деталей кампаний:`, error);
      }
      
      // Пауза между запросами для соблюдения лимита (5 запросов в секунду)
      if (i + batchSize < allCampaignIds.length) {
        await new Promise(resolve => setTimeout(resolve, 250));
      }
    }

    console.log(`✅ Получена детальная информация о ${allCampaigns.length} кампаниях`);

    // Функция для определения типа кампании
    const getCampaignType = (type: number): string => {
      const types: { [key: number]: string } = {
        4: 'В каталоге (устар.)',
        5: 'В карточке товара (устар.)',
        6: 'В поиске (устар.)',
        7: 'В рекомендациях (устар.)',
        8: 'Автоматическая',
        9: 'Аукцион'
      };
      return types[type] || `Тип ${type}`;
    };

    // Функция для определения статуса кампании
    const getCampaignStatus = (status: number): string => {
      const statuses: { [key: number]: string } = {
        '-1': 'Удалена',
        4: 'Готова к запуску',
        7: 'Завершена',
        8: 'Отменена',
        9: 'Активна',
        11: 'На паузе'
      };
      return statuses[status] || `Статус ${status}`;
    };

    // Форматируем данные для основного листа "РК"
    const campaignRows = allCampaigns.map((campaign: Campaign) => ({
      "ID кампании": campaign.advertId || '',
      "Название кампании": campaign.name || 'Без названия',
      "Тип": getCampaignType(campaign.type),
      "Статус": getCampaignStatus(campaign.status),
      "Дата создания": campaign.createTime ? new Date(campaign.createTime).toLocaleString('ru-RU') : '',
      "Дата изменения": campaign.changeTime ? new Date(campaign.changeTime).toLocaleString('ru-RU') : '',
      "Дата начала": campaign.startTime ? new Date(campaign.startTime).toLocaleString('ru-RU') : '',
      "Дата окончания": campaign.endTime ? new Date(campaign.endTime).toLocaleString('ru-RU') : '',
      "Дневной бюджет": campaign.dailyBudget || 0
    }));

    const fields = [
      'ID кампании',
      'Название кампании',
      'Тип',
      'Статус',
      'Дата создания',
      'Дата изменения',
      'Дата начала',
      'Дата окончания',
      'Дневной бюджет'
    ];

    console.log(`✅ Данные рекламных кампаний подготовлены: ${campaignRows.length} записей`);

    // Возвращаем и основные данные, и полную информацию о кампаниях
    return NextResponse.json({ 
      fields, 
      rows: campaignRows,
      detailedCampaigns: allCampaigns // Полная информация для листа "Информация о компаниях"
    });

  } catch (error) {
    console.error('❌ Ошибка получения данных рекламных кампаний:', error);
    return NextResponse.json(
      { error: 'Ошибка при получении данных рекламных кампаний' },
      { status: 500 }
    );
  }
}

