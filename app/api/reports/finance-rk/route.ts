import { NextRequest, NextResponse } from 'next/server';
import { generateFinanceRKXlsx } from '@/createFinanceRKSheet';

export const runtime = 'nodejs';

export async function POST(req: NextRequest) {
  try {
    const { apiKey, startDate, endDate } = await req.json();
    if (!apiKey || !startDate || !endDate) {
      return NextResponse.json({ error: 'apiKey, startDate, endDate обязательны' }, { status: 400 });
    }

    const buffer = await generateFinanceRKXlsx(apiKey, startDate, endDate);

    return new NextResponse(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="Финансы РК - ${startDate}–${endDate}.xlsx"`,
      },
    });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    return NextResponse.json({ error: message || 'Internal error' }, { status: 500 });
  }
}

