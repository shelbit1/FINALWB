import { NextRequest, NextResponse } from 'next/server';

export const runtime = "nodejs";

type StatsApiResponse = unknown;

function sleep(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function chunkArray<T>(arr: T[], size: number): T[][] {
  const chunks: T[][] = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

// Универсальная утилита безопасного приведения значения к строке
function safeStringify(value: unknown): string | number {
  const MAX_EXCEL_CELL_LENGTH = 32000; // Лимит Excel - 32767, берём с запасом
  
  if (value === null || value === undefined) return '';
  
  if (typeof value === 'object') {
    try {
      const jsonString = JSON.stringify(value);
      if (jsonString.length > MAX_EXCEL_CELL_LENGTH) {
        return jsonString.substring(0, MAX_EXCEL_CELL_LENGTH - 20) + '... (обрезано)';
      }
      return jsonString;
    } catch {
      const strValue = String(value);
      if (strValue.length > MAX_EXCEL_CELL_LENGTH) {
        return strValue.substring(0, MAX_EXCEL_CELL_LENGTH - 20) + '... (обрезано)';
      }
      return strValue;
    }
  }
  
  if (typeof value === 'string' && value.length > MAX_EXCEL_CELL_LENGTH) {
    return value.substring(0, MAX_EXCEL_CELL_LENGTH - 20) + '... (обрезано)';
  }
  
  return value as string | number;
}

// Функция для определения типа кампании
function getCampaignType(type: number | unknown): string {
  if (typeof type !== 'number') {
    return String(type || '');
  }
  
  const types: { [key: number]: string } = {
    4: 'В каталоге (устар.)',
    5: 'В карточке товара (устар.)',
    6: 'В поиске (устар.)',
    7: 'В рекомендациях (устар.)',
    8: 'Автоматическая',
    9: 'Аукцион'
  };
  
  return types[type] || `Тип ${type}`;
}

// Функция для получения SKU ID и типов кампаний из API promotion/adverts
async function fetchCampaignData(token: string, campaignIds: number[]): Promise<{
  skusMap: Map<number, string>;
  typesMap: Map<number, number>;
}> {
  const skusMap = new Map<number, string>();
  const typesMap = new Map<number, number>();
  
  try {
    console.log(`📊 Запрос данных для ${campaignIds.length} кампаний...`);
    
    const batchSize = 50; // Максимум 50 ID в запросе
    
    for (let i = 0; i < campaignIds.length; i += batchSize) {
      const batchIds = campaignIds.slice(i, i + batchSize);
      
      try {
        const response = await fetch('https://advert-api.wildberries.ru/adv/v1/promotion/adverts', {
          method: 'POST',
          headers: {
            'Authorization': token,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(batchIds)
        });
        
        if (response.ok) {
          const campaignsData = await response.json();
          if (Array.isArray(campaignsData)) {
            campaignsData.forEach(campaignData => {
              if (campaignData && campaignData.advertId) {
                const skus: (number | string)[] = [];
                
                // Сохраняем тип кампании
                if (typeof campaignData.type === 'number') {
                  typesMap.set(campaignData.advertId, campaignData.type);
                }
                
                // Для автоматических кампаний (type 8)
                if (campaignData.type === 8 && campaignData.autoParams && Array.isArray(campaignData.autoParams.nms)) {
                  skus.push(...campaignData.autoParams.nms);
                }
                
                // Для аукционных кампаний (type 9)
                if (campaignData.type === 9 && Array.isArray(campaignData.auction_multibids)) {
                  const auctionSkus = campaignData.auction_multibids
                    .map((bid: { nm: number }) => bid.nm)
                    .filter(Boolean);
                  skus.push(...auctionSkus);
                }
                
                // Общий параметр для разных типов кампаний
                if (Array.isArray(campaignData.unitedParams)) {
                  const unitedSkus = campaignData.unitedParams
                    .flatMap((p: { nms?: {nm: number}[] }) => p.nms || [])
                    .map((nm: {nm: number}) => nm.nm)
                    .filter(Boolean);
                  skus.push(...unitedSkus);
                }
                
                const uniqueSkus = [...new Set(skus)];
                const skusString = uniqueSkus.join(', ');
                skusMap.set(campaignData.advertId, skusString || '');
              }
            });
          }
        }
      } catch (error) {
        console.error(`❌ Ошибка при запросе данных для пакета кампаний:`, error);
      }
      
      // Пауза между пакетами
      if (campaignIds.length > i + batchSize) {
        await sleep(250);
      }
    }
    
    console.log(`✅ Получены данные для ${skusMap.size} кампаний (SKU) и ${typesMap.size} кампаний (типы)`);
    return { skusMap, typesMap };
  } catch (error) {
    console.error('❌ Ошибка при получении данных кампаний:', error);
    return { skusMap, typesMap };
  }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { token, dateFrom, dateTo, ids } = body as {
      token?: string;
      dateFrom?: string;
      dateTo?: string;
      ids?: number[];
    };

    if (!token || !dateFrom || !dateTo || !Array.isArray(ids) || ids.length === 0) {
      return NextResponse.json({
        error: 'token, dateFrom, dateTo и ids обязательны',
      }, { status: 400 });
    }

    // Проверка максимального периода 31 день
    const from = new Date(dateFrom);
    const to = new Date(dateTo);
    const diffDays = Math.ceil((to.getTime() - from.getTime()) / (1000 * 60 * 60 * 24)) + 1; // включительно
    if (diffDays > 31) {
      return NextResponse.json({ error: 'Максимальный период запроса — 31 день' }, { status: 400 });
    }

    // Бьем ids на чанки по 100 (ограничение метода)
    const batches = chunkArray(Array.from(new Set(ids)), 100);

    const allStats: Record<string, unknown>[] = [];

    for (let i = 0; i < batches.length; i++) {
      const batchIds = batches[i];
      const params = new URLSearchParams();
      params.set('ids', batchIds.join(','));
      params.set('beginDate', dateFrom);
      params.set('endDate', dateTo);

      const url = `https://advert-api.wildberries.ru/adv/v3/fullstats?${params.toString()}`;

      const res = await fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': token,
          'Content-Type': 'application/json',
        },
      });

      if (!res.ok) {
        const errText = await res.text().catch(() => '');
        return NextResponse.json({
          error: `Ошибка WB fullstats ${res.status}: ${errText || res.statusText}`,
        }, { status: res.status });
      }

      const json = (await res.json()) as StatsApiResponse;

      // Приводим ответ к единому виду массива кампаний
      let items: unknown[] = [];
      if (Array.isArray(json)) items = json;
      else if (json && typeof json === 'object') {
        const maybe = (json as Record<string, unknown>);
        if (Array.isArray(maybe.data)) items = maybe.data;
        else if (Array.isArray(maybe.adverts)) items = maybe.adverts as unknown[];
        else items = Object.values(maybe);
      }

      // Фильтруем и добавляем только объекты
      items.forEach(item => {
        if (typeof item === 'object' && item !== null) {
          allStats.push(item as Record<string, unknown>);
        }
      });

      // Соблюдаем лимиты API: до 3 запросов в минуту, всплеск 1 запрос
      if (i < batches.length - 1) {
        await sleep(20000); // 20 секунд между батчами
      }
    }

    // Получаем SKU ID и типы для всех кампаний
    const uniqueCampaignIds = Array.from(
      new Set(
        allStats
          .map(item => item?.advertId || item?.id)
          .filter((id): id is number => typeof id === 'number')
      )
    );
    const { skusMap: campaignSkusMap, typesMap: campaignTypesMap } = await fetchCampaignData(token, uniqueCampaignIds);
    
    // Преобразуем в плоские строки для Excel: строка на каждый день по каждой кампании
    const metricKeys = new Set<string>();
    const rows: Record<string, string | number>[] = [];

    // Функция для извлечения SKU ID из поля apps
    const extractSkuIds = (apps: unknown): string => {
      if (!apps) return '';
      
      // Если это массив
      if (Array.isArray(apps)) {
        const skuList: (string | number)[] = [];
        
        for (const item of apps) {
          if (!item) continue;
          
          // Если это число или строка - добавляем напрямую
          if (typeof item === 'number' || typeof item === 'string') {
            skuList.push(item);
            continue;
          }
          
          // Если это объект - пробуем извлечь различные поля
          if (typeof item === 'object') {
            const obj = item as Record<string, unknown>;
            
            // Пробуем разные варианты полей с артикулом
            const possibleFields = ['nm', 'nmId', 'nmID', 'sku', 'skuId', 'SKU', 'id', 'nmid'];
            
            for (const field of possibleFields) {
              if (field in obj && obj[field]) {
                const value = obj[field];
                if (typeof value === 'number' || typeof value === 'string') {
                  skuList.push(value);
                  break; // Нашли значение, переходим к следующему элементу
                }
              }
            }
          }
        }
        
        // Убираем дубликаты и пустые значения
        const uniqueSkus = Array.from(new Set(skuList.filter(Boolean)));
        return uniqueSkus.join(', ');
      }
      
      // Если это объект (не массив)
      if (typeof apps === 'object') {
        const obj = apps as Record<string, unknown>;
        const possibleFields = ['nm', 'nmId', 'nmID', 'sku', 'skuId', 'SKU', 'id', 'nmid'];
        
        for (const field of possibleFields) {
          if (field in obj && obj[field]) {
            const value = obj[field];
            if (typeof value === 'number' || typeof value === 'string') {
              return String(value);
            }
            // Если значение - массив, обрабатываем рекурсивно
            if (Array.isArray(value)) {
              return extractSkuIds(value);
            }
          }
        }
      }
      
      // Если ничего не подошло, возвращаем как есть
      return String(apps);
    };

    // Логирование структуры данных для отладки
    if (allStats.length > 0) {
      console.log('📊 Структура первого элемента allStats:', JSON.stringify(allStats[0]).substring(0, 1000));
      if (allStats[0]?.days && allStats[0].days.length > 0) {
        console.log('📊 Структура первого дня:', JSON.stringify(allStats[0].days[0]).substring(0, 1000));
      }
    }

    for (const item of allStats) {
      const advertId = (item && (item.advertId ?? item.id)) ?? '';
      const campaignId = typeof advertId === 'number' ? advertId : Number(advertId);
      
      // Получаем тип кампании из карты (из листа "РК")
      const type = campaignTypesMap.get(campaignId) ?? item?.type ?? '';
      
      const days = Array.isArray(item?.days) ? item.days : [];

      if (days.length > 0) {
        for (const day of days) {
          // Логирование для отладки структуры apps
          if (rows.length === 0) {
            console.log('📊 Структура day:', Object.keys(day || {}));
            console.log('📊 Значение apps:', day?.apps);
            console.log('📊 Значение appType:', day?.appType);
            console.log('📊 Все поля дня:', JSON.stringify(day).substring(0, 500));
          }
          
          // Пробуем извлечь SKU ID из поля apps
          let skuIdValue = extractSkuIds(day?.apps);
          
          // Если SKU ID не найден в apps, берем из карты кампаний
          if (!skuIdValue || skuIdValue === '') {
            skuIdValue = campaignSkusMap.get(campaignId) || '';
          }
          
          // Дополнительная проверка - если результат не строка, принудительно конвертируем
          if (typeof skuIdValue !== 'string') {
            console.warn('⚠️ SKU ID не строка, тип:', typeof skuIdValue, 'значение:', skuIdValue);
            skuIdValue = String(skuIdValue || '');
          }
          
          const row: Record<string, string | number> = {
            'ID кампании': advertId,
            'Тип': getCampaignType(type),
            'Дата': day?.date || '',
            'SKU ID': skuIdValue,
          };

          Object.keys(day || {}).forEach((k) => {
            if (k === 'date' || k === 'apps') return; // Пропускаем date и apps (apps уже обработан)
            const key = k;
            metricKeys.add(key);
            const value = (day as Record<string, unknown>)[k];
            row[key] = typeof value === 'number' ? value : (safeStringify(value) as string);
          });

          rows.push(row);
        }
      } else {
        // Если массив days отсутствует, кладём агрегированную запись
        
        // Пробуем извлечь SKU ID из поля apps
        let skuIdValue = extractSkuIds(item?.apps);
        
        // Если SKU ID не найден в apps, берем из карты кампаний
        if (!skuIdValue || skuIdValue === '') {
          skuIdValue = campaignSkusMap.get(campaignId) || '';
        }
        
        // Дополнительная проверка - если результат не строка, принудительно конвертируем
        if (typeof skuIdValue !== 'string') {
          console.warn('⚠️ SKU ID (no days) не строка, тип:', typeof skuIdValue, 'значение:', skuIdValue);
          skuIdValue = String(skuIdValue || '');
        }
        
        const row: Record<string, string | number> = {
          'ID кампании': advertId,
          'Тип': getCampaignType(type),
          'Дата': '',
          'SKU ID': skuIdValue,
        };
        Object.keys(item || {}).forEach((k) => {
          if (k === 'days' || k === 'apps') return; // Пропускаем days и apps
          const value = (item as Record<string, unknown>)[k];
          if (typeof value === 'object') return;
          metricKeys.add(k);
          row[k] = typeof value === 'number' ? value : (safeStringify(value) as string);
        });
        rows.push(row);
      }
    }

    const fields: string[] = ['ID кампании', 'Тип', 'Дата', 'SKU ID', ...Array.from(metricKeys).filter(k => !['ID кампании','Тип','Дата','SKU ID'].includes(k))];

    // Подсчитываем статистику заполнения данных
    const rowsWithSku = rows.filter(row => row['SKU ID'] && String(row['SKU ID']).trim() !== '').length;
    const rowsWithType = rows.filter(row => row['Тип'] && String(row['Тип']).trim() !== '').length;
    
    console.log(`✅ Подготовлено ${rows.length} строк со статистикой`);
    console.log(`📊 SKU ID заполнены в ${rowsWithSku} из ${rows.length} строк`);
    console.log(`📊 Тип кампании заполнен в ${rowsWithType} из ${rows.length} строк`);
    
    if (rows.length > 0) {
      console.log(`📊 Пример первой строки:`);
      console.log(`   - ID кампании: ${rows[0]['ID кампании']}`);
      console.log(`   - Тип: "${rows[0]['Тип']}"`);
      console.log(`   - SKU ID: "${rows[0]['SKU ID']}"`);
      console.log(`   - Дата: "${rows[0]['Дата']}"`);
    }

    return NextResponse.json({ fields, rows });

  } catch (error) {
    console.error('❌ Ошибка получения статистики кампаний:', error);
    return NextResponse.json({ error: 'Ошибка при получении статистики кампаний' }, { status: 500 });
  }
}


