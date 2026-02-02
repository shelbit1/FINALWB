export const runtime = "nodejs";

// Интерфейсы для данных номенклатуры
interface NomenclatureSize {
  chrtID: number;
  techSize: string;
  wbSize: string;
  skus: string[];
}

interface NomenclatureDimensions {
  length: number;
  width: number;
  height: number;
}

interface NomenclatureCard {
  nmID: number;
  imtID: number;
  vendorCode: string;
  brand: string;
  title: string;
  object: string;
  dimensions: NomenclatureDimensions;
  sizes: NomenclatureSize[];
  createdAt: string;
  updatedAt: string;
  isProhibited: boolean;
  photos: Record<string, unknown>[];
  video: string;
  status: number;
  tags: Record<string, unknown>[];
  characteristics: Record<string, unknown>[];
}

interface NomenclatureResponse {
  cards: NomenclatureCard[];
  cursor: {
    updatedAt: string;
    nmID: number;
    total: number;
  };
}

// Функция для чтения вложенных свойств объекта (аналог _readOBJ) - не используется
// function readObjectProperty(obj: Record<string, unknown>, key: string, separator: string = '.'): unknown {
//   if (!obj || !key) return '';
//   
//   let value: unknown = obj;
//   const keys = key.split(separator);
//   
//   try {
//     for (const k of keys) {
//       if (value === null || value === undefined) return '';
//       value = (value as Record<string, unknown>)[k];
//     }
//     return value !== undefined ? value : '';
//   } catch {
//     return '';
//   }
// }

export async function POST(request: Request) {
  try {
    const { token: requestToken } = (await request.json()) as {
      token?: string;
    };

    // Используем токен из запроса или из переменных окружения
    const token = requestToken || process.env.WB_API_TOKEN;

    if (!token) {
      return new Response(
        JSON.stringify({ error: "token обязателен" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    console.log("🚀 Начало получения номенклатуры товаров");

    const base = "https://content-api.wildberries.ru/content/v2/get/cards/list?locale=ru";
    
    // Определяем заголовки для Excel (убираем неиспользуемую переменную)
    // const headers = [
    //   'nmID',
    //   'imtID', 
    //   'vendorCode',
    //   'brand',
    //   'title',
    //   'object',
    //   'dimensions.length',
    //   'dimensions.width', 
    //   'dimensions.height',
    //   'calc_volume',
    //   'createdAt',
    //   'updatedAt',
    //   'isProhibited',
    //   'status',
    //   '[sizes.chrtID]',
    //   '[sizes.techSize]',
    //   '[sizes.wbSize]',
    //   '[sizes.skus]',
    //   'dt'
    // ];

    const allData: Record<string, unknown>[] = [];
    const cursor: Record<string, unknown> = { limit: 100 };
    let hasMore = true;
    let requestCount = 0;
    const maxRequests = 50; // Ограничение на количество запросов

    console.log("📊 Начинаем загрузку карточек товаров...");

    while (hasMore && requestCount < maxRequests) {
      requestCount++;
      console.log(`⏳ Запрос ${requestCount}, cursor:`, cursor);

      const requestBody = {
        settings: {
          sort: {
            ascending: false
          },
          filter: {
            withPhoto: -1 // -1 — все карточки товара
          },
          cursor
        }
      };

      const response = await fetch(base, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": token,
        },
        body: JSON.stringify(requestBody),
        cache: "no-store",
      });

      if (!response.ok) {
        const text = await response.text();
        return new Response(
          JSON.stringify({ error: text || `WB error ${response.status}` }),
          { status: response.status, headers: { "Content-Type": "application/json" } }
        );
      }

      const responseText = await response.text();
      console.log(`📊 Получен ответ номенклатуры (запрос ${requestCount}), длина: ${responseText.length} символов`);
      
      let data: NomenclatureResponse;
      try {
        data = JSON.parse(responseText) as NomenclatureResponse;
      } catch (parseError) {
        console.error(`❌ Ошибка парсинга JSON номенклатуры (запрос ${requestCount}):`, parseError);
        console.error(`Первые 500 символов ответа: ${responseText.substring(0, 500)}`);
        
        // Если это первый запрос, возвращаем ошибку
        if (requestCount === 1) {
          return new Response(
            JSON.stringify({ error: "Ошибка парсинга ответа от API Wildberries (номенклатура)" }),
            { status: 500, headers: { "Content-Type": "application/json" } }
          );
        }
        
        // Если уже есть данные, просто завершаем загрузку
        console.log(`⚠️ Не удалось распарсить ответ, но есть ${allData.length} записей - завершаем`);
        break;
      }
      
      if (!data.cards || !Array.isArray(data.cards)) {
        console.log("⚠️ Нет карточек в ответе");
        break;
      }

      console.log(`✅ Получено ${data.cards.length} карточек`);

      // Текущая дата в формате dd.MM.yyyy
      const currentDate = new Date().toLocaleDateString('ru-RU');

      // Обрабатываем каждую карточку
      data.cards.forEach((card) => {
        // Вычисляем объем (длина * ширина * высота / 1000)
        const calcVolume = (card.dimensions?.length || 0) * 
                          (card.dimensions?.width || 0) * 
                          (card.dimensions?.height || 0) / 1000;

        // Базовая строка с данными карточки
        const baseRow: Record<string, unknown> = {};
        
        // Заполняем основные поля
        baseRow['nmID'] = card.nmID;
        baseRow['imtID'] = card.imtID;
        baseRow['vendorCode'] = card.vendorCode;
        baseRow['brand'] = card.brand;
        baseRow['title'] = card.title;
        baseRow['object'] = card.object;
        baseRow['dimensions.length'] = card.dimensions?.length || 0;
        baseRow['dimensions.width'] = card.dimensions?.width || 0;
        baseRow['dimensions.height'] = card.dimensions?.height || 0;
        baseRow['calc_volume'] = calcVolume;
        baseRow['createdAt'] = card.createdAt;
        baseRow['updatedAt'] = card.updatedAt;
        baseRow['isProhibited'] = card.isProhibited;
        baseRow['status'] = card.status;
        baseRow['dt'] = currentDate;

        // Если размеров нет, добавляем одну строку
        if (!card.sizes || card.sizes.length === 0) {
          baseRow['[sizes.chrtID]'] = '';
          baseRow['[sizes.techSize]'] = '';
          baseRow['[sizes.wbSize]'] = '';
          baseRow['[sizes.skus]'] = '';
          allData.push(baseRow);
        } else {
          // Если есть размеры, создаем строку для каждого размера
          card.sizes.forEach((size) => {
            const sizeRow = { ...baseRow };
            sizeRow['[sizes.chrtID]'] = size.chrtID;
            sizeRow['[sizes.techSize]'] = size.techSize;
            sizeRow['[sizes.wbSize]'] = size.wbSize;
            sizeRow['[sizes.skus]'] = size.skus ? size.skus.join(';\n') : '';
            allData.push(sizeRow);
          });
        }
      });

      // Проверяем, есть ли еще данные для загрузки
      if (data.cursor.total < 100) {
        hasMore = false;
      } else {
        cursor.updatedAt = data.cursor.updatedAt;
        cursor.nmID = data.cursor.nmID;
        
        // Задержка между запросами для соблюдения лимитов API
        await new Promise(resolve => setTimeout(resolve, 1100));
      }
    }

    console.log(`✅ Загрузка завершена. Всего записей: ${allData.length}`);

    // Преобразуем данные в формат для Excel
    const rows = allData.map((item) => ({
      "ID товара": item.nmID,
      "ID предмета": item.imtID,
      "Артикул продавца": item.vendorCode,
      "Бренд": item.brand,
      "Наименование": item.title,
      "Предмет": item.object,
      "Длина (см)": item['dimensions.length'],
      "Ширина (см)": item['dimensions.width'],
      "Высота (см)": item['dimensions.height'],
      "Объем (л)": item.calc_volume,
      "Дата создания": item.createdAt,
      "Дата обновления": item.updatedAt,
      "Запрещен": item.isProhibited ? "Да" : "Нет",
      "Статус": item.status,
      "ID характеристики": item['[sizes.chrtID]'],
      "Технический размер": item['[sizes.techSize]'],
      "Размер WB": item['[sizes.wbSize]'],
      "SKU": item['[sizes.skus]'],
      "Дата выгрузки": item.dt,
      "Себестоимость": ""
    }));

    const fields = [
      "ID товара",
      "ID предмета", 
      "Артикул продавца",
      "Бренд",
      "Наименование",
      "Предмет",
      "Длина (см)",
      "Ширина (см)",
      "Высота (см)",
      "Объем (л)",
      "Дата создания",
      "Дата обновления",
      "Запрещен",
      "Статус",
      "ID характеристики",
      "Технический размер",
      "Размер WB",
      "SKU",
      "Дата выгрузки",
      "Себестоимость"
    ];

    return new Response(
      JSON.stringify({ fields, rows }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('❌ Ошибка в API номенклатуры:', message);
    return new Response(
      JSON.stringify({ error: message || "Internal Server Error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}
