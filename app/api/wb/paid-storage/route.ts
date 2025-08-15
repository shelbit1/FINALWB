export const runtime = "nodejs";

type PaidStorageRow = Record<string, unknown>;

export async function POST(request: Request) {
  try {
    const { token, dateFrom, dateTo } = (await request.json()) as {
      token?: string;
      dateFrom?: string; // YYYY-MM-DD
      dateTo?: string; // YYYY-MM-DD
    };

    if (!token || !dateFrom || !dateTo) {
      return new Response(
        JSON.stringify({ error: "token, dateFrom и dateTo обязательны" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const base = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage";

    // 1) Старт задачи
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
    let taskId: string | number | undefined;
    if (
      typeof startJson === "object" &&
      startJson !== null &&
      "data" in (startJson as Record<string, unknown>)
    ) {
      const data = (startJson as Record<string, unknown>).data;
      if (typeof data === "object" && data !== null && "taskId" in data) {
        const t = (data as Record<string, unknown>).taskId;
        if (typeof t === "string" || typeof t === "number") taskId = t;
      }
    }

    if (!taskId) {
      return new Response(
        JSON.stringify({ error: "Не удалось получить taskId" }),
        { status: 502, headers: { "Content-Type": "application/json" } }
      );
    }

    // 2) Ожидание готовности
    const statusUrl = `${base}/tasks/${encodeURIComponent(String(taskId))}/status`;
    let isDone = false;
    // максимум ~20 попыток по 3 секунды = ~60 секунд ожидания
    for (let i = 0; i < 20; i++) {
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

      if (status === "done") {
        isDone = true;
        break;
      }
      await new Promise((r) => setTimeout(r, 3000));
    }

    if (!isDone) {
      return new Response(
        JSON.stringify({ error: "Задача не завершилась в отведённое время" }),
        { status: 504, headers: { "Content-Type": "application/json" } }
      );
    }

    // 3) Загрузка
    const downloadUrl = `${base}/tasks/${encodeURIComponent(String(taskId))}/download`;
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

    const rowsUnknown = (await dlRes.json()) as unknown;
    if (!Array.isArray(rowsUnknown)) {
      return new Response(
        JSON.stringify({ error: "Некорректный формат ответа download" }),
        { status: 502, headers: { "Content-Type": "application/json" } }
      );
    }
    const rows = rowsUnknown as PaidStorageRow[];

    const fields: string[] = rows.length > 0 ? Object.keys(rows[0]) : [];

    return new Response(
      JSON.stringify({ fields, rows }),
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

