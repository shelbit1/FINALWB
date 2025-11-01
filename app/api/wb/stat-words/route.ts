import { NextRequest, NextResponse } from 'next/server';

export const runtime = "nodejs";

function sleep(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { token, ids } = body as {
      token?: string;
      ids?: number[];
    };

    if (!token || !Array.isArray(ids) || ids.length === 0) {
      return NextResponse.json({
        error: 'token и ids обязательны',
      }, { status: 400 });
    }

    const allStats: Record<string, unknown>[] = [];

    // Лимит: 4 запроса в секунду (250 миллисекунд между запросами)
    for (let i = 0; i < ids.length; i++) {
      const campaignId = ids[i];
      
      try {
        const url = `https://advert-api.wildberries.ru/adv/v2/auto/stat-words?id=${campaignId}`;

        const res = await fetch(url, {
          method: 'GET',
          headers: {
            'Authorization': token,
            'Content-Type': 'application/json',
          },
        });

        if (res.ok) {
          const json = await res.json();
          
          // Добавляем campaignId к каждой записи
          if (Array.isArray(json)) {
            json.forEach((item: unknown) => {
              if (typeof item === 'object' && item !== null) {
                allStats.push({
                  campaignId,
                  ...(item as Record<string, unknown>)
                });
              }
            });
          } else if (json && typeof json === 'object') {
            // Если это объект, пробуем найти массив внутри
            const jsonObj = json as Record<string, unknown>;
            const data = jsonObj.data || jsonObj.keywords || json;
            if (Array.isArray(data)) {
              data.forEach((item: unknown) => {
                if (typeof item === 'object' && item !== null) {
                  allStats.push({
                    campaignId,
                    ...(item as Record<string, unknown>)
                  });
                }
              });
            } else {
              allStats.push({
                campaignId,
                ...json
              });
            }
          }
        } else if (res.status === 204) {
          // Кампания не найдена - пропускаем
          console.log(`⚠️ Кампания ${campaignId} не найдена (204)`);
        } else {
          console.error(`❌ Ошибка для кампании ${campaignId}: ${res.status}`);
        }
      } catch (error) {
        console.error(`❌ Ошибка запроса для кампании ${campaignId}:`, error);
      }
      
      // Пауза между запросами (250 мс для соблюдения лимита 4 запроса/сек)
      if (i < ids.length - 1) {
        await sleep(250);
      }
    }

    // Формируем структуру для Excel
    const rows: Record<string, string | number>[] = [];

    for (const item of allStats) {
      const row: Record<string, string | number> = {
        'ID кампании': item.campaignId || '',
      };

      // Добавляем все поля из ответа
      Object.keys(item).forEach((key) => {
        if (key === 'campaignId') return; // Уже добавлено
        
        const value = item[key];
        if (value === null || value === undefined) {
          row[key] = '';
        } else if (typeof value === 'object') {
          try {
            let jsonString = JSON.stringify(value);
            // Ограничиваем длину до 32000 символов (лимит Excel - 32767)
            if (jsonString.length > 32000) {
              jsonString = jsonString.substring(0, 31980) + '... (обрезано)';
            }
            row[key] = jsonString;
          } catch {
            let strValue = String(value);
            if (strValue.length > 32000) {
              strValue = strValue.substring(0, 31980) + '... (обрезано)';
            }
            row[key] = strValue;
          }
        } else {
          let strValue = String(value);
          if (strValue.length > 32000) {
            strValue = strValue.substring(0, 31980) + '... (обрезано)';
          }
          row[key] = strValue;
        }
      });

      rows.push(row);
    }

    // Собираем все уникальные ключи для заголовков
    const allKeys = new Set<string>(['ID кампании']);
    rows.forEach(row => {
      Object.keys(row).forEach(key => {
        if (key !== 'ID кампании') allKeys.add(key);
      });
    });

    const fields = Array.from(allKeys);

    console.log(`✅ Получена статистика по кластерам для ${ids.length} кампаний: ${rows.length} записей`);

    return NextResponse.json({ fields, rows });

  } catch (error) {
    console.error('❌ Ошибка получения статистики по кластерам фраз:', error);
    return NextResponse.json({ error: 'Ошибка при получении статистики по кластерам фраз' }, { status: 500 });
  }
}

