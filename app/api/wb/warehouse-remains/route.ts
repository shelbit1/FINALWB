export const runtime = "nodejs";

// Интерфейсы для данных об остатках на складах
interface WarehouseItem {
  warehouseName: string;
  whId: number;
  quantity: number;
  inWayToClient: number;
  inWayFromClient: number;
}

interface WarehouseRemainsItem {
  brand: string;
  subjectName: string;
  vendorCode: string;
  nmId: number;
  barcode: string;
  techSize: string;
  volume?: number;
  warehouses: WarehouseItem[];
}

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

    console.log("🚀 Начало получения данных об остатках на складах");

    const base = "https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains";

    // 1) Создаем задание на генерацию отчета
    console.log("📊 Создание задания на генерацию отчета об остатках...");
    
    // Параметры для максимально детального отчета
    const params = new URLSearchParams({
      locale: "ru",
      groupByBrand: "true",
      groupBySubject: "true", 
      groupBySa: "true",
      groupByNm: "true",
      groupByBarcode: "true",
      groupBySize: "true",
      filterPics: "0", // не применять фильтр по фото
      filterVolume: "0" // не применять фильтр по объему
    });

    const startRes = await fetch(`${base}?${params.toString()}`, {
      method: "GET",
      headers: {
        Authorization: token,
        "Content-Type": "application/json",
      },
      cache: "no-store",
    });

    if (!startRes.ok) {
      const text = await startRes.text();
      return new Response(
        JSON.stringify({ error: text || `WB error ${startRes.status}` }),
        { status: startRes.status, headers: { "Content-Type": "application/json" } }
      );
    }

    const startJson = (await startRes.json()) as unknown;
    let taskId: string | undefined;
    
    if (
      typeof startJson === "object" &&
      startJson !== null &&
      "data" in (startJson as Record<string, unknown>)
    ) {
      const data = (startJson as Record<string, unknown>).data;
      if (typeof data === "object" && data !== null && "taskId" in data) {
        const t = (data as Record<string, unknown>).taskId;
        if (typeof t === "string") taskId = t;
      }
    }

    if (!taskId) {
      return new Response(
        JSON.stringify({ error: "Не удалось получить taskId" }),
        { status: 502, headers: { "Content-Type": "application/json" } }
      );
    }

    console.log(`📋 Получен taskId: ${taskId}`);

    // 2) Ожидание готовности отчета
    const statusUrl = `${base}/tasks/${encodeURIComponent(taskId)}/status`;
    let isDone = false;
    
    // максимум ~15 попыток по 5 секунд = ~75 секунд ожидания
    for (let i = 0; i < 15; i++) {
      console.log(`⏳ Проверка статуса, попытка ${i + 1}/15...`);
      
      const stRes = await fetch(statusUrl, {
        method: "GET",
        headers: {
          Authorization: token,
          "Content-Type": "application/json",
        },
        cache: "no-store",
      });

      if (!stRes.ok) {
        const text = await stRes.text();
        return new Response(
          JSON.stringify({ error: text || `WB error ${stRes.status}` }),
          { status: stRes.status, headers: { "Content-Type": "application/json" } }
        );
      }

      const stJson = (await stRes.json()) as unknown;
      let status: string | undefined;
      
      if (
        typeof stJson === "object" &&
        stJson !== null &&
        "data" in (stJson as Record<string, unknown>)
      ) {
        const data = (stJson as Record<string, unknown>).data;
        if (typeof data === "object" && data !== null && "status" in data) {
          const s = (data as Record<string, unknown>).status;
          if (typeof s === "string") status = s;
        }
      }

      console.log(`📊 Статус отчета: ${status}`);

      if (status === "done") {
        isDone = true;
        break;
      }
      
      // Ждем 5 секунд перед следующей проверкой
      if (i < 14) {
        await new Promise((r) => setTimeout(r, 5000));
      }
    }

    if (!isDone) {
      return new Response(
        JSON.stringify({ error: "Отчет не завершился в отведённое время (75 секунд)" }),
        { status: 504, headers: { "Content-Type": "application/json" } }
      );
    }

    // 3) Загрузка готового отчета
    console.log("📥 Загрузка готового отчета об остатках...");
    const downloadUrl = `${base}/tasks/${encodeURIComponent(taskId)}/download`;
    const dlRes = await fetch(downloadUrl, {
      method: "GET",
      headers: {
        Authorization: token,
        "Content-Type": "application/json",
      },
      cache: "no-store",
    });

    if (!dlRes.ok) {
      const text = await dlRes.text();
      return new Response(
        JSON.stringify({ error: text || `WB error ${dlRes.status}` }),
        { status: dlRes.status, headers: { "Content-Type": "application/json" } }
      );
    }

    const responseText = await dlRes.text();
    console.log(`📊 Получен ответ остатков, длина: ${responseText.length} символов`);
    
    let warehouseData: unknown;
    try {
      warehouseData = JSON.parse(responseText);
    } catch (parseError) {
      console.error(`❌ Ошибка парсинга JSON остатков:`, parseError);
      console.error(`Первые 500 символов ответа: ${responseText.substring(0, 500)}`);
      return new Response(
        JSON.stringify({ error: "Ошибка парсинга ответа от API Wildberries (остатки)" }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
    }
    
    if (!Array.isArray(warehouseData)) {
      return new Response(
        JSON.stringify({ error: "Некорректный формат ответа download" }),
        { status: 502, headers: { "Content-Type": "application/json" } }
      );
    }

    const items = warehouseData as WarehouseRemainsItem[];
    console.log(`✅ Получено ${items.length} записей об остатках на складах`);

    // Преобразуем данные в плоский формат для Excel
    const rows: Record<string, unknown>[] = [];
    const currentDate = new Date().toLocaleDateString('ru-RU');

    items.forEach((item) => {
      if (!item.warehouses || item.warehouses.length === 0) {
        // Если нет данных по складам, добавляем одну строку без склада
        rows.push({
          "Бренд": item.brand || "",
          "Предмет": item.subjectName || "",
          "Артикул продавца": item.vendorCode || "",
          "Артикул WB": item.nmId || "",
          "Штрихкод": item.barcode || "",
          "Размер": item.techSize || "",
          "Объем (л)": item.volume || "",
          "Название склада": "",
          "ID склада": "",
          "Количество": 0,
          "В пути к клиенту": 0,
          "В пути от клиента": 0,
          "Дата выгрузки": currentDate
        });
      } else {
        // Для каждого склада создаем отдельную строку
        item.warehouses.forEach((warehouse) => {
          rows.push({
            "Бренд": item.brand || "",
            "Предмет": item.subjectName || "",
            "Артикул продавца": item.vendorCode || "",
            "Артикул WB": item.nmId || "",
            "Штрихкод": item.barcode || "",
            "Размер": item.techSize || "",
            "Объем (л)": item.volume || "",
            "Название склада": warehouse.warehouseName || "",
            "ID склада": warehouse.whId || "",
            "Количество": warehouse.quantity || 0,
            "В пути к клиенту": warehouse.inWayToClient || 0,
            "В пути от клиента": warehouse.inWayFromClient || 0,
            "Дата выгрузки": currentDate
          });
        });
      }
    });

    const fields = [
      "Бренд",
      "Предмет",
      "Артикул продавца",
      "Артикул WB",
      "Штрихкод",
      "Размер",
      "Объем (л)",
      "Название склада",
      "ID склада",
      "Количество",
      "В пути к клиенту",
      "В пути от клиента",
      "Дата выгрузки"
    ];

    return new Response(
      JSON.stringify({ fields, rows }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('❌ Ошибка в API остатков на складах:', message);
    return new Response(
      JSON.stringify({ error: message || "Internal Server Error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}
