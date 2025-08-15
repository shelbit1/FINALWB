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

function addDays(date: Date, days: number): Date {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function iso(date: Date): string {
  return date.toISOString().split('T')[0];
}

export async function fetchCampaignSkus(apiKey: string, campaigns: CampaignInfo[]): Promise<Map<number,string>> {
  const map = new Map<number,string>();
  const ids = campaigns.map(c => c.advertId);
  const batch = 50;
  for (let i=0; i<ids.length; i+=batch) {
    const part = ids.slice(i, i+batch);
    try {
      const res = await fetch('https://advert-api.wildberries.ru/adv/v1/promotion/adverts', {
        method: 'POST',
        headers: { Authorization: apiKey, 'Content-Type': 'application/json' },
        body: JSON.stringify(part)
      });
      if (!res.ok) { part.forEach(id => map.set(id, 'Ошибка получения SKU')); continue; }
      const data = await res.json();
      if (Array.isArray(data)) {
        for (const c of data) {
          const skus: Array<number|string> = [];
          if (c?.type === 8 && Array.isArray(c?.autoParams?.nms)) skus.push(...c.autoParams.nms);
          if (c?.type === 9 && Array.isArray(c?.auction_multibids)) skus.push(...c.auction_multibids.map((b: {nm:number})=>b.nm).filter(Boolean));
          if (Array.isArray(c?.unitedParams)) skus.push(...c.unitedParams.flatMap((p: {nms?:{nm:number}[]})=>p?.nms||[]).map((x:{nm:number})=>x.nm).filter(Boolean));
          if (Array.isArray(c?.params)) skus.push(...c.params.flatMap((p: {nms?:{nm:number}[]})=>p?.nms||[]).map((x:{nm:number})=>x.nm).filter(Boolean));
          const unique = [...new Set(skus)];
          map.set(c.advertId, unique.length ? unique.join(', ') : 'Нет SKU');
        }
      }
    } catch {
      part.forEach(id => map.set(id, 'Ошибка запроса'));
    }
    if (i+batch<ids.length) await new Promise(r=>setTimeout(r,250));
  }
  ids.forEach(id => { if (!map.has(id)) map.set(id, 'Нет данных SKU'); });
  return map;
}

export async function fetchFinancialData(apiKey: string, startDate: string, endDate: string): Promise<FinancialData[]> {
  const origStart = new Date(startDate);
  const origEnd = new Date(endDate);
  const from = iso(addDays(origStart, -1));
  const to = iso(addDays(origEnd, 1));
  const url = `https://advert-api.wildberries.ru/adv/v1/upd?from=${from}&to=${to}`;
  const res = await fetch(url, { headers: { Authorization: apiKey, 'Content-Type': 'application/json' } });
  if (!res.ok) return [];
  const arr = await res.json();
  if (!Array.isArray(arr)) return [];
  const raw: FinancialData[] = arr.map((r: any) => ({
    advertId: Number(r?.advertId) || 0,
    date: r?.updTime ? new Date(r.updTime).toISOString().split('T')[0] : '',
    sum: Number(r?.updSum) || 0,
    bill: r?.paymentType === 'Счет' ? 1 : 0,
    type: r?.type || 'Неизвестно',
    docNumber: String(r?.updNum || ''),
    campName: r?.campName || 'Неизвестная кампания'
  })).filter(x => x.advertId && x.date);
  return applyBufferDayLogic(raw, origStart, origEnd);
}

export function applyBufferDayLogic(data: FinancialData[], originalStart: Date, originalEnd: Date): FinancialData[] {
  const mainFrom = iso(originalStart);
  const mainTo = iso(originalEnd);
  const prevBuf = iso(addDays(originalStart, -1));
  const nextBuf = iso(addDays(originalEnd, 1));

  const main = data.filter(d => d.date >= mainFrom && d.date <= mainTo);
  const prev = data.filter(d => d.date === prevBuf);
  const next = data.filter(d => d.date === nextBuf);

  const mainDoc = new Set(main.map(d => d.docNumber).filter(Boolean));
  const nextDoc = new Set(next.map(d => d.docNumber).filter(Boolean));
  const filteredMain = main.filter(d => !d.docNumber || !nextDoc.has(d.docNumber));

  const addFromPrev = prev.filter(d => {
    if (!d.docNumber) return false;
    const count = prev.filter(x => x.docNumber === d.docNumber).length;
    return mainDoc.has(d.docNumber) && count >= 2;
  });

  return [...filteredMain, ...addFromPrev];
}

