import { NextRequest, NextResponse } from 'next/server';

export const runtime = "nodejs";

function sleep(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { token: requestToken, dateFrom, dateTo, ids } = body as {
      token?: string;
      dateFrom?: string;
      dateTo?: string;
      ids?: number[];
    };

    // Используем токен из запроса или из переменных окружения
    const token = requestToken || process.env.WB_API_TOKEN;

    if (!token || !dateFrom || !dateTo || !Array.isArray(ids) || ids.length === 0) {
      return NextResponse.json({
        error: 'token, dateFrom, dateTo и ids обязательны',
      }, { status: 400 });
    }

    const allStats: Record<string, unknown>[] = [];

    // Лимит: 10 запросов в минуту (6 секунд между запросами)
    // API поддерживает до 100 items в одном запросе
    const batchSize = 100;
    
    for (let i = 0; i < ids.length; i += batchSize) {
      const batchIds = ids.slice(i, i + batchSize);
      
      try {
        const url = 'https://advert-api.wildberries.ru/adv/v0/normquery/stats';

        // Формируем items для всех кампаний в пакете
        const items = batchIds.map(id => ({ id }));

        const requestBody = {
          from: dateFrom,
          to: dateTo,
          items
        };

        console.log(`📊 Запрос статистики для ${batchIds.length} кампаний (пакет ${Math.floor(i / batchSize) + 1})...`);

        const res = await fetch(url, {
          method: 'POST',
          headers: {
            'Authorization': token,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody)
        });

        if (res.ok) {
          const json = await res.json();
          
          // Обрабатываем ответ
          if (Array.isArray(json)) {
            allStats.push(...json);
          } else if (json && typeof json === 'object') {
            // Если это объект, пробуем найти массив внутри
            const jsonObj = json as Record<string, unknown>;
            const data = jsonObj.data || jsonObj.stats || jsonObj.items || json;
            if (Array.isArray(data)) {
              allStats.push(...data);
            } else {
              allStats.push(json);
            }
          }
          console.log(`✅ Получено записей: ${Array.isArray(json) ? json.length : 1}`);
        } else if (res.status === 204) {
          // Нет данных для этого пакета - пропускаем
          console.log(`⚠️ Нет данных для пакета кампаний (204)`);
        } else if (res.status === 400) {
          console.error(`❌ Неправильный запрос: ${res.status}`);
          const errorText = await res.text().catch(() => '');
          console.error(`Детали ошибки: ${errorText}`);
        } else if (res.status === 401) {
          console.error(`❌ Не авторизован: ${res.status}`);
        } else if (res.status === 429) {
          console.error(`❌ Превышен лимит запросов: ${res.status}`);
          // Увеличиваем паузу при превышении лимита
          await sleep(10000);
        } else {
          console.error(`❌ Ошибка запроса: ${res.status}`);
        }
      } catch (error) {
        console.error(`❌ Ошибка запроса для пакета кампаний:`, error);
      }
      
      // Пауза между запросами (6 секунд для соблюдения лимита 10 запросов/минуту)
      if (i + batchSize < ids.length) {
        await sleep(6000);
      }
    }

    // Формируем структуру для Excel
    const rows: Record<string, string | number>[] = [];

    for (const item of allStats) {
      const row: Record<string, string | number> = {};

      // Добавляем все поля из ответа
      Object.keys(item).forEach((key) => {
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
    const allKeys = new Set<string>();
    rows.forEach(row => {
      Object.keys(row).forEach(key => {
        allKeys.add(key);
      });
    });

    const fields = Array.from(allKeys);

    console.log(`✅ Получена статистика поисковых кластеров для ${ids.length} кампаний: ${rows.length} записей`);

    return NextResponse.json({ fields, rows });

  } catch (error) {
    console.error('❌ Ошибка получения статистики поисковых кластеров:', error);
    return NextResponse.json({ error: 'Ошибка при получении статистики поисковых кластеров' }, { status: 500 });
  }
}

