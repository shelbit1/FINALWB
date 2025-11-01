export const runtime = "nodejs";

// Тип для данных РНП (не используется явно, но определяет структуру)
type RnpItemStructure = {
  realizationreport_id: number;
  date_from: string;
  date_to: string;
  create_dt: string;
  currency_name: string;
  suppliercontract_code: string;
  rrd_id: number;
  gi_id: number;
  subject_name: string;
  nm_id: number;
  brand_name: string;
  sa_name: string;
  ts_name: string;
  barcode: string;
  doc_type_name: string;
  quantity: number;
  retail_price: number;
  retail_amount: number;
  sale_percent: number;
  commission_percent: number;
  office_name: string;
  supplier_oper_name: string;
  order_dt: string;
  sale_dt: string;
  rr_dt: string;
  shk_id: number;
  retail_price_withdisc_rub: number;
  delivery_amount: number;
  return_amount: number;
  delivery_rub: number;
  gi_box_type_name: string;
  product_discount_for_report: number;
  supplier_promo: number;
  rid: number;
  ppvz_spp_prc: number;
  ppvz_kvw_prc_base: number;
  ppvz_kvw_prc: number;
  sup_rating_prc_up: number;
  is_kgvp_v2: number;
  ppvz_sales_commission: number;
  ppvz_for_pay: number;
  ppvz_reward: number;
  acquiring_fee: number;
  acquiring_percent: number;
  acquiring_bank: string;
  ppvz_vw: number;
  ppvz_vw_nds: number;
  ppvz_office_id: number;
  ppvz_office_name: string;
  ppvz_supplier_id: number;
  ppvz_supplier_name: string;
  ppvz_inn: string;
  declaration_number: string;
  bonus_type_name: string;
  sticker_id: string;
  site_country: string;
  penalty: number;
  additional_payment: number;
  rebill_logistic_cost: number;
  rebill_logistic_org: string;
  kiz: string;
  storage_fee: number;
  deduction: number;
  acceptance: number;
  srid: string;
  report_type: number;
}

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
  "report_type"
];

function mapRnpItem(raw: Record<string, unknown>): Record<string, unknown> {
  const item: Record<string, unknown> = { ...raw };
  
  // Устанавливаем значения по умолчанию для отсутствующих полей
  if (!item.bonus_type_name) item.bonus_type_name = "";
  if (!item.rebill_logistic_cost) item.rebill_logistic_cost = 0;
  if (!item.rebill_logistic_org) item.rebill_logistic_org = "";
  if (!item.kiz) item.kiz = "";
  if (!item.sticker_id) item.sticker_id = "";
  if (!item.declaration_number) item.declaration_number = "";
  if (!item.acquiring_bank) item.acquiring_bank = "";
  if (!item.srid) item.srid = "";

  const mapped: Record<string, unknown> = {};
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

    console.log(`🚀 Начало получения данных РНП (ежедневные отчеты): ${dateFrom} - ${dateTo}`);

    const endpoint = "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod";
    const limit = 30000;
    let rrdid = 0;
    const result: Record<string, unknown>[] = [];

    // Пагинация по rrdid для получения всех данных с параметром period=daily
    while (true) {
      const params = new URLSearchParams({
        dateFrom,
        dateTo,
        limit: String(limit),
        rrdid: String(rrdid),
        period: "daily", // Используем ежедневные отчеты для детализации
      });

      console.log(`📊 Запрос РНП данных (daily) с rrdid: ${rrdid}`);

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
        console.error(`❌ Ошибка API РНП: ${res.status}`, text);
        return new Response(
          JSON.stringify({ error: text || `WB error ${res.status}` }),
          { status: res.status, headers: { "Content-Type": "application/json" } }
        );
      }

      const dataUnknown = (await res.json()) as unknown;
      if (!Array.isArray(dataUnknown) || dataUnknown.length === 0) {
        console.log(`✅ Получено данных РНП (ежедневные отчеты): ${result.length} записей`);
        break;
      }
      
      const data = dataUnknown as Array<Record<string, unknown>>;

      // Сортируем по дате отчета для корректной пагинации
      data.sort((a, b) => {
        const aDt = typeof a.rr_dt === "string" ? a.rr_dt : "";
        const bDt = typeof b.rr_dt === "string" ? b.rr_dt : "";
        return new Date(aDt).getTime() - new Date(bDt).getTime();
      });

      let lastRrdid = 0;
      for (const raw of data) {
        const mapped = mapRnpItem(raw);
        result.push(mapped);
        const rrd = raw.rrd_id;
        lastRrdid = typeof rrd === "number" ? rrd : lastRrdid;
      }

      // Если rrd_id не изменился, выходим из цикла
      if (lastRrdid === rrdid) {
        console.log(`✅ Завершение пагинации РНП (ежедневные отчеты). Всего записей: ${result.length}`);
        break;
      }
      rrdid = lastRrdid;

      // Небольшая задержка между запросами для соблюдения лимитов API (1 запрос в минуту)
      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    // Преобразуем данные в формат для Excel с русскими заголовками
    const rows = result.map((item) => ({
      "ID отчета": item.realizationreport_id,
      "Дата начала": item.date_from,
      "Дата окончания": item.date_to,
      "Дата создания": item.create_dt,
      "Валюта": item.currency_name,
      "Договор": item.suppliercontract_code,
      "RRD ID": item.rrd_id,
      "GI ID": item.gi_id,
      "Предмет": item.subject_name,
      "Артикул WB": item.nm_id,
      "Бренд": item.brand_name,
      "Артикул продавца": item.sa_name,
      "Размер": item.ts_name,
      "Штрихкод": item.barcode,
      "Тип документа": item.doc_type_name,
      "Количество": item.quantity,
      "Цена розничная": item.retail_price,
      "Сумма продаж": item.retail_amount,
      "Скидка продавца": item.sale_percent,
      "Комиссия": item.commission_percent,
      "Склад": item.office_name,
      "Тип операции": item.supplier_oper_name,
      "Дата заказа": item.order_dt,
      "Дата продажи": item.sale_dt,
      "Дата отчета": item.rr_dt,
      "ШК ID": item.shk_id,
      "Цена со скидкой": item.retail_price_withdisc_rub,
      "Доставка": item.delivery_amount,
      "Возврат": item.return_amount,
      "Доставка руб": item.delivery_rub,
      "Тип коробки": item.gi_box_type_name,
      "Скидка товара": item.product_discount_for_report,
      "Промо продавца": item.supplier_promo,
      "RID": item.rid,
      "СПП": item.ppvz_spp_prc,
      "КВВ базовый": item.ppvz_kvw_prc_base,
      "КВВ": item.ppvz_kvw_prc,
      "Рейтинг": item.sup_rating_prc_up,
      "КГВП v2": item.is_kgvp_v2,
      "Комиссия продаж": item.ppvz_sales_commission,
      "К доплате": item.ppvz_for_pay,
      "Вознаграждение": item.ppvz_reward,
      "Эквайринг": item.acquiring_fee,
      "Процент эквайринга": item.acquiring_percent,
      "Банк эквайринга": item.acquiring_bank,
      "ВВ": item.ppvz_vw,
      "ВВ с НДС": item.ppvz_vw_nds,
      "ID офиса": item.ppvz_office_id,
      "Офис": item.ppvz_office_name,
      "ID поставщика": item.ppvz_supplier_id,
      "Поставщик": item.ppvz_supplier_name,
      "ИНН": item.ppvz_inn,
      "Номер декларации": item.declaration_number,
      "Тип бонуса": item.bonus_type_name,
      "ID стикера": item.sticker_id,
      "Страна": item.site_country,
      "Штраф": item.penalty,
      "Доплата": item.additional_payment,
      "Перевыставление логистики": item.rebill_logistic_cost,
      "Организация логистики": item.rebill_logistic_org,
      "КИЗ": item.kiz,
      "Хранение": item.storage_fee,
      "Удержание": item.deduction,
      "Приемка": item.acceptance,
      "SRID": item.srid,
      "Тип отчета": item.report_type
    }));

    const fields = [
      "ID отчета",
      "Дата начала",
      "Дата окончания", 
      "Дата создания",
      "Валюта",
      "Договор",
      "RRD ID",
      "GI ID",
      "Предмет",
      "Артикул WB",
      "Бренд",
      "Артикул продавца",
      "Размер",
      "Штрихкод",
      "Тип документа",
      "Количество",
      "Цена розничная",
      "Сумма продаж",
      "Скидка продавца",
      "Комиссия",
      "Склад",
      "Тип операции",
      "Дата заказа",
      "Дата продажи",
      "Дата отчета",
      "ШК ID",
      "Цена со скидкой",
      "Доставка",
      "Возврат",
      "Доставка руб",
      "Тип коробки",
      "Скидка товара",
      "Промо продавца",
      "RID",
      "СПП",
      "КВВ базовый",
      "КВВ",
      "Рейтинг",
      "КГВП v2",
      "Комиссия продаж",
      "К доплате",
      "Вознаграждение",
      "Эквайринг",
      "Процент эквайринга",
      "Банк эквайринга",
      "ВВ",
      "ВВ с НДС",
      "ID офиса",
      "Офис",
      "ID поставщика",
      "Поставщик",
      "ИНН",
      "Номер декларации",
      "Тип бонуса",
      "ID стикера",
      "Страна",
      "Штраф",
      "Доплата",
      "Перевыставление логистики",
      "Организация логистики",
      "КИЗ",
      "Хранение",
      "Удержание",
      "Приемка",
      "SRID",
      "Тип отчета"
    ];

    return new Response(
      JSON.stringify({ fields, rows }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('❌ Ошибка в API РНП:', message);
    return new Response(
      JSON.stringify({ error: message || "Internal Server Error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}
