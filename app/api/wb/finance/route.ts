export const runtime = "nodejs";

import { fetchFinancialData, fetchCampaignSkus } from "@/utils/finance-rk";

const FIELD_ORDER: string[] = [
  "advertId",
  "campName",
  "date", 
  "sum",
  "bill",
  "type",
  "docNumber",
  "skuIds",
  "reportPeriod"
];

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

    console.log('🚀 Начало получения финансовых данных РК...');
    const financeStartTime = Date.now();

    // Получаем финансовые данные с буферными днями
    const financialData = await fetchFinancialData(token, dateFrom, dateTo);

    console.log(`📊 Проверка данных для листа "Финансы РК". Финансовых записей: ${financialData.length}`);

    if (financialData.length > 0) {
      // Собираем уникальные кампании из финансовых данных
      const uniqueCampaigns = new Map<number, { advertId: number; name: string }>();
      financialData.forEach(record => {
        if (!uniqueCampaigns.has(record.advertId)) {
          uniqueCampaigns.set(record.advertId, {
            advertId: record.advertId,
            name: record.campName || 'Неизвестная кампания'
          });
        }
      });
      
      console.log(`📊 Найдено уникальных кампаний: ${uniqueCampaigns.size}`);
      
      // Получаем детальные данные артикулов для столбца SKU ID
      const campaignsArray = Array.from(uniqueCampaigns.values());
      console.log(`📊 Получение детальных данных артикулов для кампаний...`);
      const skusMap = await fetchCampaignSkus(token, campaignsArray);
      console.log(`📊 Получено детальных данных артикулов: ${skusMap.size} кампаний`);
      
      // Подготавливаем данные для листа "Финансы РК"
      const reportPeriod = `${new Date(dateFrom).toLocaleDateString('ru-RU')} - ${new Date(dateTo).toLocaleDateString('ru-RU')}`;
      
      const rows = financialData.map(record => {
        const ordered: Record<string, unknown> = {};
        
        const rowData = {
          advertId: record.advertId,
          campName: record.campName || 'Неизвестная кампания',
          date: record.date,
          sum: record.sum,
          bill: record.bill === 1 ? 'Счет' : 'Баланс',
          type: record.type,
          docNumber: record.docNumber,
          skuIds: skusMap.get(record.advertId) || 'Нет данных',
          reportPeriod: reportPeriod
        };
        
        for (const key of FIELD_ORDER) {
          ordered[key] = (rowData as any)[key] ?? "";
        }
        
        return ordered;
      });

      console.log(`✅ Лист "Финансы РК" подготовлен за ${Date.now() - financeStartTime}ms с ${rows.length} записями`);

      return new Response(
        JSON.stringify({ fields: FIELD_ORDER, rows }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      );
    } else {
      console.log("ℹ️ Нет данных для создания листа 'Финансы РК' - возвращаем пустой результат");
      
      return new Response(
        JSON.stringify({ fields: FIELD_ORDER, rows: [] }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      );
    }

  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('❌ Ошибка в API финансов РК:', message);
    return new Response(
      JSON.stringify({ error: message || "Internal Server Error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}