"use client";

import { useState } from "react";
import * as XLSX from "xlsx";

export default function Home() {
  const [token, setToken] = useState(
    "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjUwNTIwdjEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTc2NjExMDk4NCwiaWQiOiIwMTk3ODg5Mi02M2U3LTczOWYtYTEyMC02MjU3ZGUxZmM1YjciLCJpaWQiOjI5MzkxMDIxLCJvaWQiOjU5NjI1LCJzIjoxMDczNzQ5NzU4LCJzaWQiOiJmMTEwN2UwOS1iMGNiLTVjYTctYTU0Mi03M2IxYzZhNjQ0N2UiLCJ0IjpmYWxzZSwidWlkIjoyOTM5MTAyMX0.sW33A2YFcxWhuVEilgGTSsSc2TASz1MyeLPN9G4x-lnSgM2yAu7O7QvZcomXbnNFZpUhsSA2LRj5YjMALs7xHw"
  );
  // Функция для получения последней полностью завершенной недели (понедельник-воскресенье)
  const getLastCompletedWeek = () => {
    const today = new Date();
    const dayOfWeek = today.getDay(); // 0 = воскресенье, 1 = понедельник, ..., 6 = суббота
    
    // Находим последнее воскресенье (конец недели)
    let lastSunday = new Date(today);
    
    if (dayOfWeek === 0) {
      // Если сегодня воскресенье, берем вчерашнее воскресенье (неделю назад)
      lastSunday.setDate(today.getDate() - 7);
    } else {
      // Иначе вычитаем количество дней до прошлого воскресенья
      lastSunday.setDate(today.getDate() - dayOfWeek);
    }
    
    // Находим понедельник этой недели (6 дней назад от воскресенья)
    const mondayOfWeek = new Date(lastSunday);
    mondayOfWeek.setDate(lastSunday.getDate() - 6);
    
    console.log('Сегодня:', today.toISOString().split('T')[0]);
    console.log('День недели:', dayOfWeek);
    console.log('Последнее воскресенье:', lastSunday.toISOString().split('T')[0]);
    console.log('Понедельник недели:', mondayOfWeek.toISOString().split('T')[0]);
    
    return {
      monday: mondayOfWeek.toISOString().split('T')[0],
      sunday: lastSunday.toISOString().split('T')[0]
    };
  };
  
  const getDefaultWeek = getLastCompletedWeek();
  const [selectedMonday, setSelectedMonday] = useState(getDefaultWeek.monday);
  const [periodA, setPeriodA] = useState(getDefaultWeek.monday);
  const [periodB, setPeriodB] = useState(getDefaultWeek.sunday);
  const [isLoadingReport, setIsLoadingReport] = useState(false);

  // Функция для обработки выбора понедельника
  const handleMondayChange = (mondayDate: string) => {
    const monday = new Date(mondayDate);
    
    // Проверяем, что выбран именно понедельник
    if (monday.getDay() !== 1) {
      alert("Можно выбрать только понедельник!");
      return;
    }
    
    // Проверяем, что неделя полностью завершена
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    if (sunday >= today) {
      alert("Можно выбрать только полностью завершенные недели!");
      return;
    }
    
    setSelectedMonday(mondayDate);
    setPeriodA(mondayDate);
    setPeriodB(sunday.toISOString().split('T')[0]);
  };

  const handleDownload = async () => {
    try {
      setIsLoadingReport(true);
      
      // Валидация данных перед отправкой
      if (!token.trim()) {
        throw new Error("Введите API токен Wildberries");
      }
      
      if (!periodA || !periodB) {
        throw new Error("Выберите период для выгрузки данных");
      }
      
      const dateFrom = new Date(periodA);
      const dateTo = new Date(periodB);
      const today = new Date();
      today.setHours(23, 59, 59, 999); // Конец текущего дня
      
      if (dateFrom > dateTo) {
        throw new Error("Дата начала периода не может быть позже даты окончания");
      }
      
      if (dateTo > today) {
        throw new Error("Дата окончания периода не может быть в будущем");
      }
      
      // Проверка на слишком большой диапазон (больше 31 дня)
      const daysDiff = Math.ceil((dateTo.getTime() - dateFrom.getTime()) / (1000 * 60 * 60 * 24));
      if (daysDiff > 31) {
        throw new Error("Максимальный период выгрузки: 31 день. Выберите меньший диапазон дат");
      }
      
      if (daysDiff < 1) {
        throw new Error("Минимальный период выгрузки: 1 день");
      }
      
      const payload = { token, dateFrom: periodA, dateTo: periodB };

      console.log("Отправляем запросы к API...", payload);

      let resReport, resPaid, resAcceptance, resFinanceRK, resNomenclature;
      
      try {
        // Делаем запросы не все сразу, а с небольшими задержками для соблюдения лимитов API
        console.log("📊 Запуск основного отчета...");
        resReport = await fetch("/api/wb/report", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        }).catch(err => {
          console.error("Ошибка fetch для отчета:", err);
          throw new Error(`Ошибка сетевого запроса для отчета: ${err.message}`);
        });

        // Небольшая задержка перед следующим запросом
        await new Promise(resolve => setTimeout(resolve, 500));
        
        console.log("📊 Запуск отчета платного хранения...");
        resPaid = await fetch("/api/wb/paid-storage", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        }).catch(err => {
          console.error("Ошибка fetch для платного хранения:", err);
          throw new Error(`Ошибка сетевого запроса для платного хранения: ${err.message}`);
        });

        // Задержка перед запросом платной приемки (у неё лимит 1 запрос в минуту)
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        console.log("📊 Запуск отчета платной приемки...");
        resAcceptance = await fetch("/api/wb/acceptance-report", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        }).catch(err => {
          console.error("Ошибка fetch для платной приемки:", err);
          // Возвращаем пустые данные вместо ошибки
          return {
            ok: true,
            json: async () => ({
              fields: [
                'Кол-во',
                'Дата создания GI',
                'Income ID',
                'Артикул WB',
                'Дата создания ШК',
                'Предмет',
                'Сумма (руб)',
                'Дата отчета',
                'Номер отчета'
              ],
              rows: []
            })
          };
        });

        // Задержка перед финансами РК
        await new Promise(resolve => setTimeout(resolve, 500));
        
        console.log("📊 Запуск отчета финансов РК...");
        resFinanceRK = await fetch("/api/reports/finance-rk-data", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ token, startDate: periodA, endDate: periodB }),
        }).catch(err => {
          console.error("Ошибка fetch для финансов РК:", err);
          // Возвращаем пустые данные вместо ошибки
          return {
            ok: true,
            json: async () => ({
              fields: [
                'ID кампании',
                'Название кампании', 
                'Дата',
                'Сумма',
                'Источник списания',
                'Тип операции',
                'Номер документа',
                'SKU ID',
                'Период отчета'
              ],
              rows: []
            })
          };
        });

        // Задержка перед номенклатурой
        await new Promise(resolve => setTimeout(resolve, 500));
        
        console.log("📊 Запуск отчета номенклатуры...");
        resNomenclature = await fetch("/api/wb/nomenclature", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ token }),
        }).catch(err => {
          console.error("Ошибка fetch для номенклатуры:", err);
          // Возвращаем пустые данные вместо ошибки
          return {
            ok: true,
            json: async () => ({
              fields: [
                "ID товара",
                "ID предмета", 
                "Артикул продавца",
                "Бренд",
                "Наименование",
                "Предмет",
                "Длина (см)",
                "Ширина (см)",
                "Высота (см)",
                "Объем (л)",
                "Дата создания",
                "Дата обновления",
                "Запрещен",
                "Статус",
                "ID характеристики",
                "Технический размер",
                "Размер WB",
                "SKU",
                "Дата выгрузки"
              ],
              rows: []
            })
          };
        });
      } catch (fetchError) {
        console.error("Promise.all fetch error:", fetchError);
        throw fetchError;
      }

      console.log("Получены ответы:", { 
        reportStatus: resReport.status, 
        paidStatus: resPaid.status,
        acceptanceStatus: resAcceptance.ok ? 'success' : 'fallback',
        financeRKStatus: resFinanceRK.ok ? 'success' : 'fallback',
        nomenclatureStatus: resNomenclature.ok ? 'success' : 'fallback'
      });

      if (!resReport.ok) {
        const err = await resReport.json().catch(() => ({}));
        throw new Error(err.error || `Ошибка отчёта: ${resReport.status}`);
      }
      if (!resPaid.ok) {
        const err = await resPaid.json().catch(() => ({}));
        throw new Error(err.error || `Ошибка платного хранения: ${resPaid.status}`);
      }
      // resAcceptance и resFinanceRK всегда возвращают ok: true (либо данные, либо пустой fallback)

      const report: { fields: string[]; rows: Record<string, unknown>[] } = await resReport.json();
      const paid: { fields: string[]; rows: Record<string, unknown>[] } = await resPaid.json();
      const acceptance: { fields: string[]; rows: Record<string, unknown>[] } = await resAcceptance.json();
      const financeRK: { fields: string[]; rows: Record<string, unknown>[] } = await resFinanceRK.json();
      const nomenclature: { fields: string[]; rows: Record<string, unknown>[] } = await resNomenclature.json();

      // Сопоставляем номера отчетов из еженедельного отчета с платной приемкой
      // ВАЖНО: данные платной приемки уже отфильтрованы по "Дата создания ШК" (период -1 день)
      if (acceptance.rows.length > 0 && report.rows.length > 0) {
        // Создаем карту соответствия nmId -> номер отчета из еженедельного отчета
        const nmIdToReportNumber = new Map<number, string>();
        
        report.rows.forEach(row => {
          const nmId = Number(row.nm_id);
          const reportNumber = String(row.realizationreport_id || '');
          if (nmId && reportNumber) {
            nmIdToReportNumber.set(nmId, reportNumber);
          }
        });

        // Обновляем данные платной приемки с номерами отчетов
        acceptance.rows = acceptance.rows.map(row => ({
          ...row,
          "Номер отчета": nmIdToReportNumber.get(Number(row["Артикул WB"])) || "Не найден"
        }));

        console.log(`📊 Сопоставлено номеров отчетов: ${nmIdToReportNumber.size} уникальных артикулов`);
      }

      const reportHeader = report.fields;
      const reportRows = report.rows.map((row) => reportHeader.map((key) => row[key] ?? ""));
      const reportSheet = XLSX.utils.aoa_to_sheet([reportHeader, ...reportRows]);

      const paidHeader = paid.fields;
      const paidRows = paid.rows.map((row) => paidHeader.map((key) => row[key] ?? ""));
      const paidSheet = XLSX.utils.aoa_to_sheet([paidHeader, ...paidRows]);

      const acceptanceHeader = acceptance.fields;
      const acceptanceRows = acceptance.rows.map((row) => acceptanceHeader.map((key) => row[key] ?? ""));
      const acceptanceSheet = XLSX.utils.aoa_to_sheet([acceptanceHeader, ...acceptanceRows]);

      const financeRKHeader = financeRK.fields;
      const financeRKRows = financeRK.rows.map((row) => {
        return financeRKHeader.map((key) => {
          const value = row[key] ?? "";
          // Специальная обработка для колонки "Сумма" - сохраняем как число
          if (key === 'Сумма') {
            if (typeof value === 'number') {
              return value; // Уже число
            } else if (typeof value === 'string') {
              const numValue = parseFloat(String(value).replace(/[^\d.]/g, ''));
              return isNaN(numValue) ? 0 : numValue;
            }
            return 0;
          }
          return value;
        });
      });
      
      const financeRKSheet = XLSX.utils.aoa_to_sheet([financeRKHeader, ...financeRKRows]);

      // Применяем российское форматирование чисел для колонки "Сумма"
      if (financeRKRows.length > 0) {
        const sumColumnIndex = financeRKHeader.indexOf('Сумма');
        if (sumColumnIndex !== -1) {
          // Сначала устанавливаем формат для всей колонки
          const range = XLSX.utils.decode_range(financeRKSheet['!ref'] || 'A1');
          for (let row = 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: sumColumnIndex });
            if (financeRKSheet[cellAddress]) {
              // Российский формат чисел: пробелы для тысяч, запятая для десятичных
              financeRKSheet[cellAddress].z = '# ### ##0,##';
              financeRKSheet[cellAddress].t = 'n'; // Убеждаемся что тип - число
            }
          }
        }
      }

      const workbook = XLSX.utils.book_new();
      // Создаем лист номенклатуры
      const nomenclatureHeader = nomenclature.fields;
      const nomenclatureRows = nomenclature.rows.map((row) => nomenclatureHeader.map((key) => row[key] ?? ""));
      const nomenclatureSheet = XLSX.utils.aoa_to_sheet([nomenclatureHeader, ...nomenclatureRows]);
      
      // Устанавливаем ширину колонок для номенклатуры
      const nomenclatureColWidths = [
        { wch: 12 }, // ID товара
        { wch: 12 }, // ID предмета
        { wch: 20 }, // Артикул продавца
        { wch: 15 }, // Бренд
        { wch: 30 }, // Наименование
        { wch: 15 }, // Предмет
        { wch: 12 }, // Длина (см)
        { wch: 12 }, // Ширина (см)
        { wch: 12 }, // Высота (см)
        { wch: 12 }, // Объем (л)
        { wch: 16 }, // Дата создания
        { wch: 16 }, // Дата обновления
        { wch: 10 }, // Запрещен
        { wch: 8 },  // Статус
        { wch: 15 }, // ID характеристики
        { wch: 15 }, // Технический размер
        { wch: 12 }, // Размер WB
        { wch: 20 }, // SKU
        { wch: 12 }, // Дата выгрузки
        { wch: 15 }  // Себестоимость
      ];
      nomenclatureSheet["!cols"] = nomenclatureColWidths;

      // Создаем лист "Аналитика" из номенклатуры, сгруппированный по артикулу
      const createProductsSheet = () => {
        // Группируем товары по артикулу продавца
        const groupedProducts = new Map<string, any[]>();
        
        nomenclature.rows.forEach((row: any) => {
          const vendorCode = row["Артикул продавца"] || "Без артикула";
          if (!groupedProducts.has(vendorCode)) {
            groupedProducts.set(vendorCode, []);
          }
          groupedProducts.get(vendorCode)?.push(row);
        });

        // Создаем заголовки для листа "Аналитика" с пустыми строками сверху
        const productsHeaders = ["Артикул", "Размер", "Штрихкод", "Артикул WB", "Бренд"];
        const productsData = [
          [], // Пустая строка 1
          [], // Пустая строка 2
          productsHeaders // Заголовки в строке 3
        ];

        // Добавляем данные, сгруппированные по артикулу
        Array.from(groupedProducts.entries())
          .sort(([a], [b]) => a.localeCompare(b)) // Сортируем по артикулу
          .forEach(([vendorCode, products]) => {
            products.forEach((product: any) => {
              productsData.push([
                vendorCode, // Артикул
                product["Технический размер"] || "", // Размер - технический
                product["SKU"] || "", // Штрихкод (используем SKU как штрихкод)
                product["ID товара"] || "", // Артикул WB (nmID)
                product["Бренд"] || "" // Бренд
              ]);
            });
          });

        return XLSX.utils.aoa_to_sheet(productsData);
      };

      const productsSheet = createProductsSheet();
      
      // Устанавливаем ширину колонок для листа "Аналитика"
      productsSheet["!cols"] = [
        { wch: 20 }, // Артикул
        { wch: 15 }, // Размер
        { wch: 25 }, // Штрихкод
        { wch: 15 }, // Артикул WB
        { wch: 20 }  // Бренд
      ];

      XLSX.utils.book_append_sheet(workbook, productsSheet, "Аналитика");
      XLSX.utils.book_append_sheet(workbook, reportSheet, "Еженед отчет");
      XLSX.utils.book_append_sheet(workbook, paidSheet, "Платное хранение");
      XLSX.utils.book_append_sheet(workbook, acceptanceSheet, "Платная приемка");
      XLSX.utils.book_append_sheet(workbook, financeRKSheet, "Финансы РК");
      XLSX.utils.book_append_sheet(workbook, nomenclatureSheet, "Номенклатура");
      const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([arrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "Отчеты_WB.xlsx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error("Полная ошибка:", e);
      console.error("Стек ошибки:", (e as Error)?.stack);
      
      const errorMessage = (e as Error).message || "Не удалось сформировать файл";
      
      // Более понятные сообщения об ошибках
      let userFriendlyMessage = errorMessage;
      if (errorMessage.includes("there are no companies with correct intervals")) {
        userFriendlyMessage = "Ошибка: нет данных для указанного периода. Проверьте:\n" +
          "• Правильность выбранных дат\n" +
          "• Наличие активных рекламных кампаний в этот период\n" +
          "• Корректность API токена";
      } else if (errorMessage.includes("401") || errorMessage.includes("Unauthorized")) {
        userFriendlyMessage = "Ошибка авторизации: проверьте корректность API токена Wildberries";
      } else if (errorMessage.includes("403") || errorMessage.includes("Forbidden")) {
        userFriendlyMessage = "Доступ запрещен: убедитесь, что у токена есть необходимые права доступа";
      } else if (errorMessage.includes("429") || errorMessage.includes("Too Many Requests") || errorMessage.includes("too many requests")) {
        userFriendlyMessage = "Превышен лимит запросов к API Wildberries.\n\n" +
          "Рекомендации:\n" +
          "• Подождите 1-2 минуты перед повторной попыткой\n" +
          "• Не запускайте несколько отчетов одновременно\n" +
          "• Попробуйте скачать отчеты по отдельности (используйте отдельные кнопки)\n" +
          "• Проверьте лимиты API на https://dev.wildberries.ru/openapi/api-information";
      } else if (errorMessage.includes("500") || errorMessage.includes("Internal Server Error")) {
        userFriendlyMessage = "Внутренняя ошибка сервера Wildberries. Попробуйте позже";
      } else if (errorMessage.includes("Failed to fetch") || errorMessage.includes("fetch failed")) {
        userFriendlyMessage = "Ошибка сетевого подключения. Проверьте:\n" +
          "• Интернет-соединение\n" +
          "• Работу сервера разработки (npm run dev)\n" +
          "• Отсутствие блокировки запросов антивирусом или файрволлом";
      } else if (errorMessage.includes("Ошибка сетевого запроса")) {
        userFriendlyMessage = errorMessage + "\n\nВозможные причины:\n" +
          "• Проблемы с интернет-соединением\n" +
          "• Сервер разработки не запущен\n" +
          "• Блокировка запросов браузером или антивирусом";
      }
      
      alert(userFriendlyMessage);
    } finally {
      setIsLoadingReport(false);
    }
  };

  return (
    <div className="min-h-screen w-full flex items-center justify-center p-6">
      <div className="w-full max-w-xl rounded-2xl border border-black/[.08] dark:border-white/[.145] p-6 bg-white dark:bg-[#0f0f0f]">
        <div className="flex items-center justify-between mb-6">
          <h1 className="text-xl font-semibold">Выгрузка данных</h1>
        </div>

          <div className="flex flex-col gap-5">
            <div className="flex flex-col gap-2">
              <label htmlFor="wb-token" className="text-sm font-medium">
                АПИ ВБ:
              </label>
              <input
                id="wb-token"
                type="password"
                value={token}
                onChange={(e) => setToken(e.target.value)}
                placeholder="Введите токен ВБ"
                className="w-full h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6]"
              />
            </div>

            <div className="flex flex-col gap-2">
            <span className="text-sm font-medium">Выбор недели:</span>
            <div className="flex flex-col gap-3">
                <div className="flex flex-col gap-1">
                <label htmlFor="monday-select" className="text-xs text-black/60 dark:text-white/70">
                  Понедельник (начало недели) - только завершенные недели
                  </label>
                  <input
                  id="monday-select"
                    type="date"
                  value={selectedMonday}
                  onChange={(e) => handleMondayChange(e.target.value)}
                    className="h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6]"
                  />
                </div>
              <div className="bg-gray-50 dark:bg-gray-800 rounded-lg p-3">
                <div className="text-xs text-gray-600 dark:text-gray-400 mb-1">
                  Автоматически выбранный период:
                </div>
                <div className="text-sm font-medium">
                  {new Date(periodA).toLocaleDateString('ru-RU')} - {new Date(periodB).toLocaleDateString('ru-RU')}
                </div>
                <div className="text-xs text-gray-500 dark:text-gray-500 mt-1">
                  (Понедельник - Воскресенье)
                </div>
              </div>
              </div>
            </div>

            <div className="pt-2 flex flex-col gap-3">
              <button
                type="button"
                onClick={handleDownload}
                disabled={isLoadingReport}
                className={`w-full h-11 rounded-lg bg-black text-white dark:bg-white dark:text-black font-medium transition-opacity ${
                  isLoadingReport ? "opacity-60 cursor-not-allowed" : "hover:opacity-90"
                }`}
              >
                {isLoadingReport ? (
                  <span className="inline-flex items-center gap-2">
                    <span className="inline-block h-4 w-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
                    Загрузка...
                  </span>
                ) : (
                  "Скачать"
                )}
              </button>
            </div>
          </div>
      </div>
    </div>
  );
}
