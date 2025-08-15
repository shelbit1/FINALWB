export const runtime = "nodejs";

type ReportItem = Record<string, unknown>;

const FIELD_ORDER: string[] = [
  "realizationreport_id",
  "date_from",
  "date_to",
  "create_dt",
  "currency_name",
  "suppliercontract_code",
  "rrd_id",
  "gi_id",
  "subject_name",
  "nm_id",
  "brand_name",
  "sa_name",
  "ts_name",
  "barcode",
  "doc_type_name",
  "quantity",
  "retail_price",
  "retail_amount",
  "sale_percent",
  "commission_percent",
  "office_name",
  "supplier_oper_name",
  "order_dt",
  "sale_dt",
  "rr_dt",
  "shk_id",
  "retail_price_withdisc_rub",
  "delivery_amount",
  "return_amount",
  "delivery_rub",
  "gi_box_type_name",
  "product_discount_for_report",
  "supplier_promo",
  "rid",
  "ppvz_spp_prc",
  "ppvz_kvw_prc_base",
  "ppvz_kvw_prc",
  "sup_rating_prc_up",
  "is_kgvp_v2",
  "ppvz_sales_commission",
  "ppvz_for_pay",
  "ppvz_reward",
  "acquiring_fee",
  "acquiring_percent",
  "acquiring_bank",
  "ppvz_vw",
  "ppvz_vw_nds",
  "ppvz_office_id",
  "ppvz_office_name",
  "ppvz_supplier_id",
  "ppvz_supplier_name",
  "ppvz_inn",
  "declaration_number",
  "bonus_type_name",
  "sticker_id",
  "site_country",
  "penalty",
  "additional_payment",
  "rebill_logistic_cost",
  "rebill_logistic_org",
  "kiz",
  "storage_fee",
  "deduction",
  "acceptance",
  "srid",
  "report_type",
];

function mapItem(raw: Record<string, unknown>): ReportItem {
  const item: Record<string, unknown> = { ...raw };
  if (!item.bonus_type_name) item.bonus_type_name = "";
  if (!item.rebill_logistic_cost) item.rebill_logistic_cost = 0;
  if (!item.rebill_logistic_org) item.rebill_logistic_org = "";
  if (!item.kiz) item.kiz = "";

  const mapped: ReportItem = {};
  for (const key of FIELD_ORDER) {
    mapped[key] = item[key] ?? null;
  }
  return mapped;
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

    const endpoint =
      "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod";
    const limit = 30000;
    let rrdid = 0;
    const result: ReportItem[] = [];

    // Пагинация по rrdid
    // Останавливаемся, когда приходит пустой массив
    // Между запросами сортируем по rr_dt по аналогии с GAS
    // Без задержки: серверный код, риски лимитов минимальны при единичной выгрузке
    // При необходимости можно добавить задержку 5s
    while (true) {
      const params = new URLSearchParams({
        dateFrom,
        dateTo,
        limit: String(limit),
        rrdid: String(rrdid),
      });

      const res = await fetch(`${endpoint}?${params.toString()}`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        cache: "no-store",
      });

      if (!res.ok) {
        const text = await res.text();
        return new Response(
          JSON.stringify({ error: text || `WB error ${res.status}` }),
          { status: res.status, headers: { "Content-Type": "application/json" } }
        );
      }

      const dataUnknown = (await res.json()) as unknown;
      if (!Array.isArray(dataUnknown) || dataUnknown.length === 0) {
        break;
      }
      const data = dataUnknown as Array<Record<string, unknown>>;

      data.sort((a, b) => {
        const aDt = typeof a.rr_dt === "string" ? a.rr_dt : "";
        const bDt = typeof b.rr_dt === "string" ? b.rr_dt : "";
        return new Date(aDt).getTime() - new Date(bDt).getTime();
      });

      let lastRrdid = 0;
      for (const raw of data) {
        const mapped = mapItem(raw);
        result.push(mapped);
        const rrd = raw.rrd_id;
        lastRrdid = typeof rrd === "number" ? rrd : lastRrdid;
      }

      // Если rrd_id не изменился, чтобы избежать бесконечного цикла, выходим
      if (lastRrdid === rrdid) {
        break;
      }
      rrdid = lastRrdid;
    }

    return new Response(JSON.stringify({ fields: FIELD_ORDER, rows: result }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    return new Response(
      JSON.stringify({ error: message || "Internal Server Error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}

