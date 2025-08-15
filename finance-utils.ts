// Интерфейсы для финансовых данных (соответствует логике конкурента)
export interface FinancialData {
  advertId: number;
  date: string;
  sum: number;
  bill: number;
  type: string;
  docNumber: string;
  campName?: string;
}

export interface CampaignInfo {
  advertId: number;
  name: string;
}

// Функция добавления дней к дате (точно как у конкурента)
export function addDays(date: Date, days: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

// Функция форматирования даты (точно как у конкурента)
export function formatDate(date: Date): string {
  return date.toISOString().split('T')[0];
}

// Функция получения детальных данных по артикулам кампаний (логика конкурента)
export async function fetchCampaignSkus(apiKey: string, campaigns: CampaignInfo[]): Promise<Map<number, string>> {
  const skusMap = new Map<number, string>();
  
  try {
    console.log(`📊 Запрос SKU для ${campaigns.length} кампаний...`);
    
    const campaignIds = campaigns.map(c => c.advertId);
    const batchSize = 50; // Максимум 50 ID в запросе

    for (let i = 0; i < campaignIds.length; i += batchSize) {
      const batchIds = campaignIds.slice(i, i + batchSize);
      
      try {
        const response = await fetch('https://advert-api.wildberries.ru/adv/v1/promotion/adverts', {
          method: 'POST',
          headers: {
            'Authorization': apiKey,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(batchIds)
        });

        if (response.ok) {
          const campaignsData = await response.json();
          if (Array.isArray(campaignsData)) {
            campaignsData.forEach(campaignData => {
              if (campaignData && campaignData.advertId) {
                const skus: (number | string)[] = [];

                // Для автоматических кампаний (type 8)
                if (campaignData.type === 8 && campaignData.autoParams && Array.isArray(campaignData.autoParams.nms)) {
                  skus.push(...campaignData.autoParams.nms);
                }

                // Для аукционных кампаний (type 9)
                if (campaignData.type === 9 && Array.isArray(campaignData.auction_multibids)) {
                  const auctionSkus = campaignData.auction_multibids
                    .map((bid: { nm: number }) => bid.nm)
                    .filter(Boolean);
                  skus.push(...auctionSkus);
                }
                
                // Общий параметр для разных типов кампаний
                if (Array.isArray(campaignData.unitedParams)) {
                  const unitedSkus = campaignData.unitedParams
                    .flatMap((p: { nms?: {nm: number}[] }) => p.nms || [])
                    .map((nm: {nm: number}) => nm.nm)
                    .filter(Boolean);
                  skus.push(...unitedSkus);
                }

                const uniqueSkus = [...new Set(skus)];
                const skusString = uniqueSkus.join(', ');
                skusMap.set(campaignData.advertId, skusString || 'Нет SKU');
              }
            });
          }
        }
      } catch (error) {
        console.error(`❌ Ошибка при запросе SKU для пакета кампаний:`, error);
        batchIds.forEach(id => skusMap.set(id, 'Ошибка запроса'));
      }
      
      // Пауза между пакетами для соблюдения лимита
      if (campaignIds.length > i + batchSize) {
        await new Promise(r => setTimeout(r, 250));
      }
    }

    return skusMap;
  } catch (error) {
    console.error('❌ Ошибка при получении SKU кампаний:', error);
    return skusMap;
  }
}

// Основная функция получения финансовых данных с логикой буферных дней (точно как у конкурента)
export async function fetchFinancialData(apiKey: string, startDate: string, endDate: string): Promise<FinancialData[]> {
  try {
    console.log(`📊 Получение финансовых данных: ${startDate} - ${endDate}`);
    
    // Добавляем буферные дни
    const originalStart = new Date(startDate);
    const originalEnd = new Date(endDate);
    const bufferStart = addDays(originalStart, -1);
    const bufferEnd = addDays(originalEnd, 1);
    
    const adjustedStartDate = formatDate(bufferStart);
    const adjustedEndDate = formatDate(bufferEnd);
    
    console.log(`📅 Расширенный период с буферными днями: ${adjustedStartDate} - ${adjustedEndDate}`);
    
    const response = await fetch(`https://advert-api.wildberries.ru/adv/v1/upd?from=${adjustedStartDate}&to=${adjustedEndDate}`, {
      method: 'GET',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      console.error(`❌ Ошибка получения финансовых данных: ${response.status} ${response.statusText}`);
      return [];
    }

    const data = await response.json();
    
    // Преобразуем данные в нужный формат
    const financialData: FinancialData[] = data.map((record: any) => ({
      advertId: record.advertId,
      date: record.updTime ? new Date(record.updTime).toISOString().split('T')[0] : '',
      sum: record.updSum || 0,
      bill: record.paymentType === 'Счет' ? 1 : 0,
      type: record.type || 'Неизвестно',
      docNumber: record.updNum || '',
      campName: record.campName || 'Неизвестная кампания'
    }));
    
    // Применяем логику буферных дней
    const filteredData = applyBufferDayLogic(financialData, originalStart, originalEnd);
    
    return filteredData;
  } catch (error) {
    console.error('❌ Ошибка при получении финансовых данных:', error);
    return [];
  }
}

// Логика фильтрации буферных дней (точно как у конкурента)
export function applyBufferDayLogic(data: FinancialData[], originalStart: Date, originalEnd: Date): FinancialData[] {
  const originalStartStr = formatDate(originalStart);
  const originalEndStr = formatDate(originalEnd);
  const bufferStartStr = formatDate(addDays(originalStart, -1));
  const bufferEndStr = formatDate(addDays(originalEnd, 1));
  
  // Разделяем данные по периодам
  const mainPeriodData = data.filter(record => 
    record.date >= originalStartStr && record.date <= originalEndStr
  );
  
  const previousBufferData = data.filter(record => record.date === bufferStartStr);
  const nextBufferData = data.filter(record => record.date === bufferEndStr);
  
  // Получаем номера документов из основного периода
  const mainDocNumbers = new Set(mainPeriodData.map(record => record.docNumber));
  
  // Исключаем из основного периода записи с номерами документов из следующего буферного дня
  const nextBufferDocNumbers = new Set(nextBufferData.map(record => record.docNumber));
  const filteredMainData = mainPeriodData.filter(record => !nextBufferDocNumbers.has(record.docNumber));
  
  // Добавляем записи из предыдущего буферного дня, если есть совпадения по номерам документов
  const previousBufferToAdd = previousBufferData.filter(record => 
    mainDocNumbers.has(record.docNumber) && previousBufferData.filter(r => r.docNumber === record.docNumber).length >= 2
  );
  
  const result = [...filteredMainData, ...previousBufferToAdd];
  
  return result;
}

