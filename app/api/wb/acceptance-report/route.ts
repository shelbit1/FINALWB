export const runtime = "nodejs";

// Интерфейс для данных платной приемки
interface AcceptanceItem {
  count: number;
  giCreateDate: string; // Дата создания GI
  incomeId: number; // Income ID
  nmID: number; // Артикул WB
  shkCreateDate: string; // Дата создания ШК
  subjectName: string; // Предмет
  total: number; // Сумма (руб)
}

export async function POST(request: Request) {
  try {
    const { token, dateFrom, dateTo } = (await request.json()) as {
      token?: string;
      dateFrom?: string;
      dateTo?: string;
    };

    if (!token || !dateFrom || !dateTo) {
      return new Response(
        JSON.stringify({ error: "token, dateFrom и dateTo обязательны" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    console.log(`🚀 Начало получения данных платной приемки: ${dateFrom} - ${dateTo}`);

    const base = "https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report";

    // 1) Создаем задание на генерацию отчета
    console.log("📊 Создание задания на генерацию отчета...");
    const startRes = await fetch(
      `${base}?dateFrom=${encodeURIComponent(dateFrom)}&dateTo=${encodeURIComponent(dateTo)}`,
      {
        method: "GET",
        headers: {
          Authorization: token,
          "Content-Type": "application/json",
        },
        cache: "no-store",
      }
    );

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
    
    // максимум ~12 попыток по 5 секунд = ~60 секунд ожидания
    for (let i = 0; i < 12; i++) {
      console.log(`⏳ Проверка статуса, попытка ${i + 1}/12...`);
      
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
      if (i < 11) {
        await new Promise((r) => setTimeout(r, 5000));
      }
    }

    if (!isDone) {
      return new Response(
        JSON.stringify({ error: "Отчет не завершился в отведённое время (60 секунд)" }),
        { status: 504, headers: { "Content-Type": "application/json" } }
      );
    }

    // 3) Загрузка готового отчета
    console.log("📥 Загрузка готового отчета...");
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

    const acceptanceData = (await dlRes.json()) as unknown;
    
    if (!Array.isArray(acceptanceData)) {
      return new Response(
        JSON.stringify({ error: "Некорректный формат ответа download" }),
        { status: 502, headers: { "Content-Type": "application/json" } }
      );
    }

    const items = acceptanceData as AcceptanceItem[];
    console.log(`✅ Получено ${items.length} записей платной приемки (до фильтрации)`);

    // Применяем фильтрацию по дате создания ШК (период -1 день)
    // Если выбран период с 4 по 10 августа, то для "Дата создания ШК" берем с 3 по 9 августа
    const startDate = new Date(dateFrom);
    const endDate = new Date(dateTo);
    
    // Смещаем период на -1 день для фильтрации по shkCreateDate
    const filterStartDate = new Date(startDate);
    filterStartDate.setDate(filterStartDate.getDate() - 1);
    const filterEndDate = new Date(endDate);
    filterEndDate.setDate(filterEndDate.getDate() - 1);
    
    console.log(`🔍 Фильтрация по "Дата создания ШК": с ${filterStartDate.toISOString().split('T')[0]} по ${filterEndDate.toISOString().split('T')[0]}`);

    const filteredItems = items.filter((item) => {
      if (!item.shkCreateDate) return false;
      
      // Преобразуем дату создания ШК в формат Date
      const shkDate = new Date(item.shkCreateDate);
      
      // Проверяем, что дата создания ШК входит в смещенный период
      return shkDate >= filterStartDate && shkDate <= filterEndDate;
    });

    console.log(`✅ После фильтрации осталось ${filteredItems.length} записей`);

    // Форматируем период отчета
    const reportPeriod = `${new Date(dateFrom).toLocaleDateString('ru-RU')} - ${new Date(dateTo).toLocaleDateString('ru-RU')}`;

    // Преобразуем данные в формат для Excel
    const rows = filteredItems.map((item) => ({
      "Кол-во": item.count,
      "Дата создания GI": item.giCreateDate,
      "Income ID": item.incomeId,
      "Артикул WB": item.nmID,
      "Дата создания ШК": item.shkCreateDate,
      "Предмет": item.subjectName,
      "Сумма (руб)": item.total,
      "Дата отчета": reportPeriod,
      "Номер отчета": "" // Пока пустое поле, будет заполняться при сопоставлении
    }));

    const fields = [
      "Кол-во",
      "Дата создания GI",
      "Income ID",
      "Артикул WB",
      "Дата создания ШК",
      "Предмет",
      "Сумма (руб)",
      "Дата отчета",
      "Номер отчета"
    ];

    return new Response(
      JSON.stringify({ fields, rows }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('❌ Ошибка в API платной приемки:', message);
    return new Response(
      JSON.stringify({ error: message || "Internal Server Error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}