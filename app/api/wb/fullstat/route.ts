export const runtime = "nodejs";

type JsonRecord = Record<string, unknown>;

const MAX_RK_PER_REQUEST = 30;

const FIELD_ORDER: string[] = [
  "date",
  "advertId",
  "typ",
  "nmId",
  "name",
  "boosterStats",
  "views",
  "clicks",
  "sum",
  "atbs",
  "orders",
  "shks",
  "sum_price",
  "ctr",
  "cpc",
  "cr",
  "Номер документа",
];

function formatDateDdMmYyyy(input: string): string {
  const d = new Date(input);
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = d.getUTCFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function chunkArray<T>(arr: T[], size: number): T[][] {
  const result: T[][] = [];
  for (let i = 0; i < arr.length; i += size) {
    result.push(arr.slice(i, i + size));
  }
  return result;
}

export async function POST(request: Request) {
  try {
    const { token: requestToken, dateFrom, dateTo } = (await request.json()) as {
      token?: string;
      dateFrom?: string; // YYYY-MM-DD
      dateTo?: string; // YYYY-MM-DD
    };

    // Используем токен из запроса или из переменных окружения
    const token = requestToken || process.env.WB_API_TOKEN;

    if (!token || !dateFrom || !dateTo) {
      return new Response(
        JSON.stringify({ error: "token, dateFrom и dateTo обязательны" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    // 1) Получить список РК с нужными статусами
    const countRes = await fetch(
      "https://advert-api.wildberries.ru/adv/v1/promotion/count",
      {
        method: "GET",
        headers: { Authorization: token },
        cache: "no-store",
      }
    );
    if (!countRes.ok) {
      const text = await countRes.text();
      return new Response(
        JSON.stringify({ error: text || `WB error ${countRes.status}` }),
        { status: countRes.status, headers: { "Content-Type": "application/json" } }
      );
    }
    const countJson = (await countRes.json()) as JsonRecord;
    const adverts = Array.isArray((countJson as JsonRecord).adverts)
      ? ((countJson as JsonRecord).adverts as unknown[])
      : [];

    const allowedStatuses = new Set<number>([7, 9, 11]);
    const ids: Array<{ id: number; typ: number; status: number }> = [];
    for (const adv of adverts as JsonRecord[]) {
      const statusVal = typeof adv.status === "number" ? adv.status : -1;
      if (!allowedStatuses.has(statusVal)) continue;
      const advertList = Array.isArray(adv.advert_list)
        ? (adv.advert_list as unknown[])
        : [];
      const typ = typeof adv.type === "number" ? adv.type : 0;
      for (const el of advertList as JsonRecord[]) {
        const advertId = typeof el.advertId === "number" ? el.advertId : undefined;
        if (typeof advertId === "number") {
          ids.push({ id: advertId, typ, status: statusVal });
        }
      }
    }

    if (ids.length === 0) {
      return new Response(
        JSON.stringify({ fields: FIELD_ORDER, rows: [] }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      );
    }

    const interval = { begin: dateFrom, end: dateTo };

    // Получаем данные updNum для всех РК с расширенным периодом
    const updMapping: Record<number, string> = {};
    try {
      // Добавляем буфер ±1 день для корректного получения номеров документов
      const bufferDays = 1;
      const startDate = new Date(dateFrom);
      startDate.setDate(startDate.getDate() - bufferDays);
      const endDate = new Date(dateTo);  
      endDate.setDate(endDate.getDate() + bufferDays);
      
      const fromParam = startDate.toISOString().split('T')[0];
      const toParam = endDate.toISOString().split('T')[0];
      
      const updRes = await fetch(
        `https://advert-api.wildberries.ru/adv/v1/upd?from=${fromParam}&to=${toParam}`,
        {
          method: "GET",
          headers: { Authorization: token },
          cache: "no-store",
        }
      );
      if (updRes.ok) {
        const updData = (await updRes.json()) as unknown;
        if (Array.isArray(updData)) {
          for (const item of updData as JsonRecord[]) {
            const advertId = typeof item.advertId === "number" ? item.advertId : undefined;
            const updNum = typeof item.updNum === "string" ? item.updNum : "";
            if (typeof advertId === "number") {
              updMapping[advertId] = updNum;
            }
          }
        }
      } else {
        console.error("Ошибка запроса updNum:", updRes.status, updRes.statusText);
      }
    } catch (error) {
      // Игнорируем ошибки получения updNum, просто оставляем поле пустым
      console.error("Ошибка получения updNum:", error);
    }
    
    // Логирование для отладки
    console.log(`updMapping содержит: ${Object.keys(updMapping).length} записей`);
    if (Object.keys(updMapping).length > 0) {
      console.log("Пример updMapping:", Object.entries(updMapping).slice(0, 5));
    } else {
      console.log("updMapping пуст - номера документов будут пустыми");
    }

    const rows: Array<Record<string, unknown>> = [];

    // 2) Бежим чанками по fullstats
    for (const chunk of chunkArray(ids, MAX_RK_PER_REQUEST)) {
      const payload = chunk.map((c) => ({ id: c.id, interval }));
      const fsRes = await fetch(
        "https://advert-api.wildberries.ru/adv/v2/fullstats",
        {
          method: "POST",
          headers: {
            Authorization: token,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(payload),
          cache: "no-store",
        }
      );
      if (!fsRes.ok) {
        const text = await fsRes.text();
        return new Response(
          JSON.stringify({ error: text || `WB error ${fsRes.status}` }),
          { status: fsRes.status, headers: { "Content-Type": "application/json" } }
        );
      }

      const statsUnknown = (await fsRes.json()) as unknown;
      const stats = Array.isArray(statsUnknown)
        ? (statsUnknown as JsonRecord[])
        : [];

      for (const advStat of stats) {
        const advertId = typeof advStat.advertId === "number" ? advStat.advertId : 0;
        const typ = chunk.find((c) => c.id === advertId)?.id
          ? ids.find((i) => i.id === advertId)?.typ ?? 0
          : 0;

        // boosterStats → карта по дате и nmId
        const boosterMap: Record<string, Record<number, number>> = {};
        if (Array.isArray((advStat as JsonRecord).boosterStats)) {
          for (const b of (advStat as JsonRecord).boosterStats as JsonRecord[]) {
            const dateStr = typeof b.date === "string" ? formatDateDdMmYyyy(b.date) : undefined;
            const nm = typeof b.nm === "number" ? b.nm : undefined;
            const avg = typeof b.avg_position === "number" ? b.avg_position : undefined;
            if (dateStr && typeof nm === "number" && typeof avg === "number") {
              if (!boosterMap[dateStr]) boosterMap[dateStr] = {};
              boosterMap[dateStr][nm] = avg;
            }
          }
        }

        // Агрегация по дням и приложениям
        const sums: Record<number, Record<string, Record<string, unknown>>> = {};
        const days = Array.isArray((advStat as JsonRecord).days)
          ? ((advStat as JsonRecord).days as JsonRecord[])
          : [];
        for (const day of days) {
          const date = typeof day.date === "string" ? formatDateDdMmYyyy(day.date) : undefined;
          if (!date) continue;
          const apps = Array.isArray(day.apps) ? (day.apps as JsonRecord[]) : [];
          for (const app of apps) {
            const appType = typeof app.appType === "number" ? app.appType : -1;
            const nmArr = Array.isArray(app.nm) ? (app.nm as JsonRecord[]) : [];

            if (appType === 0 && nmArr.length > 0) {
              const views = typeof app.views === "number" ? app.views / nmArr.length : 0;
              const clicks = typeof app.clicks === "number" ? app.clicks / nmArr.length : 0;
              const sum = typeof app.sum === "number" ? app.sum / nmArr.length : 0;
              const atbs = typeof app.atbs === "number" ? app.atbs / nmArr.length : 0;
              const orders = typeof app.orders === "number" ? app.orders / nmArr.length : 0;
              const shks = typeof app.shks === "number" ? app.shks / nmArr.length : 0;
              const sum_price = typeof app.sum_price === "number" ? app.sum_price / nmArr.length : 0;
              for (const n of nmArr) {
                n.views = views;
                n.clicks = clicks;
                n.sum = sum;
                n.atbs = atbs;
                n.orders = orders;
                n.shks = shks;
                n.sum_price = sum_price;
              }
            }

            for (const n of nmArr) {
              const nmId = typeof n.nmId === "number" ? n.nmId : undefined;
              if (!nmId) continue;
              const name = typeof n.name === "string" ? n.name : "";
              const add = (key: string): number => (typeof (n as JsonRecord)[key] === "number" ? ((n as JsonRecord)[key] as number) : 0);

              if (!sums[nmId]) sums[nmId] = {};
              if (!sums[nmId][date]) {
                sums[nmId][date] = {
                  name,
                  nmId,
                  date,
                  advertId,
                  typ,
                  boosterStats: boosterMap[date]?.[nmId],
                  views: add("views"),
                  clicks: add("clicks"),
                  sum: add("sum"),
                  atbs: add("atbs"),
                  orders: add("orders"),
                  shks: add("shks"),
                  sum_price: add("sum_price"),
                } as Record<string, unknown>;
              } else {
                const cur = sums[nmId][date];
                cur.views = (Number(cur.views) || 0) + add("views");
                cur.clicks = (Number(cur.clicks) || 0) + add("clicks");
                cur.sum = (Number(cur.sum) || 0) + add("sum");
                cur.atbs = (Number(cur.atbs) || 0) + add("atbs");
                cur.orders = (Number(cur.orders) || 0) + add("orders");
                cur.shks = (Number(cur.shks) || 0) + add("shks");
                cur.sum_price = (Number(cur.sum_price) || 0) + add("sum_price");
              }
            }
          }
        }

        // Подготовить строки по FIELD_ORDER
        for (const nmIdStr of Object.keys(sums)) {
          const nmId = Number(nmIdStr);
          const byDate = sums[nmId];
          for (const date of Object.keys(byDate)) {
            const stat = byDate[date];
            const clicks = Number(stat.clicks) || 0;
            const views = Number(stat.views) || 0;
            const sum = Number(stat.sum) || 0;
            const orders = Number(stat.orders) || 0;

            const ctr = views > 0 ? Number(((clicks / views) * 100).toFixed(2)) : "";
            const cpc = clicks > 0 ? sum / clicks : "";
            const cr = views > 0 ? orders / views : "";

            const advertId = Number(stat.advertId) || 0;
            const docNumber = updMapping[advertId] || "";
            
            // Логирование для отладки номеров документов
            if (!docNumber && advertId > 0) {
              console.log(`Номер документа не найден для advertId: ${advertId}`);
            }
            
            const row: Record<string, unknown> = {
              ...stat,
              ctr,
              cpc,
              cr,
              "Номер документа": docNumber,
            };
            const ordered: Record<string, unknown> = {};
            for (const key of FIELD_ORDER) ordered[key] = row[key] ?? "";
            rows.push(ordered);
          }
        }
      }
    }

    return new Response(
      JSON.stringify({ fields: FIELD_ORDER, rows }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    return new Response(
      JSON.stringify({ error: message || "Internal Server Error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}

