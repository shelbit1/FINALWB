"use client";

import { useState } from "react";
import * as XLSX from "xlsx";

export default function Home() {
  const [token, setToken] = useState(
    "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjUwNTIwdjEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTc2NjExMDk4NCwiaWQiOiIwMTk3ODg5Mi02M2U3LTczOWYtYTEyMC02MjU3ZGUxZmM1YjciLCJpaWQiOjI5MzkxMDIxLCJvaWQiOjU5NjI1LCJzIjoxMDczNzQ5NzU4LCJzaWQiOiJmMTEwN2UwOS1iMGNiLTVjYTctYTU0Mi03M2IxYzZhNjQ0N2UiLCJ0IjpmYWxzZSwidWlkIjoyOTM5MTAyMX0.sW33A2YFcxWhuVEilgGTSsSc2TASz1MyeLPN9G4x-lnSgM2yAu7O7QvZcomXbnNFZpUhsSA2LRj5YjMALs7xHw"
  );
  
  // Состояния для модального окна себестоимости
  const [showCostModal, setShowCostModal] = useState(false);
  const [groupedProducts, setGroupedProducts] = useState<Array<{vendorCode: string; brand: string; items: Array<Record<string, unknown>>}>>([]);
  const [skuCosts, setSkuCosts] = useState<{[sku: string]: string}>({});
  const [bulkCost, setBulkCost] = useState<string>("");
  const [isLoadingCosts, setIsLoadingCosts] = useState(false);

  // Функции для работы с localStorage
  const saveCostsToStorage = (costs: {[sku: string]: string}) => {
    try {
      localStorage.setItem('wb_sku_costs', JSON.stringify(costs));
    } catch (error) {
      console.error('Ошибка сохранения в localStorage:', error);
    }
  };

  const loadCostsFromStorage = (): {[sku: string]: string} => {
    try {
      const stored = localStorage.getItem('wb_sku_costs');
      return stored ? JSON.parse(stored) : {};
    } catch (error) {
      console.error('Ошибка загрузки из localStorage:', error);
      return {};
    }
  };
  // Функция для получения последней полностью завершенной недели (понедельник-воскресенье)
  const getLastCompletedWeek = () => {
    const today = new Date();
    const dayOfWeek = today.getDay(); // 0 = воскресенье, 1 = понедельник, ..., 6 = суббота
    
    // Находим последнее воскресенье (конец недели)
    const lastSunday = new Date(today);
    
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

  // Функция для загрузки и группировки товаров по артикулу
  const handleLoadCosts = async () => {
    try {
      setIsLoadingCosts(true);
      
      if (!token.trim()) {
        alert("Введите API токен Wildberries");
        return;
      }

      console.log("📊 Загрузка номенклатуры для себестоимости...");
      
      const resNomenclature = await fetch("/api/wb/nomenclature", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token }),
      });

      if (!resNomenclature.ok) {
        const err = await resNomenclature.json().catch(() => ({}));
        throw new Error(err.error || `Ошибка загрузки номенклатуры: ${resNomenclature.status}`);
      }

      const nomenclature: { fields: string[]; rows: Record<string, unknown>[] } = await resNomenclature.json();
      
      // Группируем товары по артикулу продавца
      const grouped = new Map<string, {vendorCode: string; brand: string; items: Array<Record<string, unknown>>}>();
      
      nomenclature.rows.forEach((row: Record<string, unknown>) => {
        const vendorCode = String(row["Артикул продавца"] || "Без артикула");
        const skus = String(row["SKU"] || "");
        
        if (!grouped.has(vendorCode)) {
          grouped.set(vendorCode, {
            vendorCode,
            brand: String(row["Бренд"] || ""),
            items: []
          });
        }
        
        // Разбиваем штрихкоды по символам ;\n если их несколько
        const skuList = skus.split(';\n').filter((sku: string) => sku.trim() !== '');
        
        // Если штрихкодов нет, добавляем один элемент без SKU
        if (skuList.length === 0) {
          grouped.get(vendorCode)?.items.push({
            nmId: String(row["ID товара"] || ""),
            size: String(row["Технический размер"] || ""),
            sku: "",
            title: String(row["Наименование"] || ""),
            uniqueKey: `${row["ID товара"]}_${row["Технический размер"]}_no_sku`
          });
        } else {
          // Для каждого штрихкода создаем отдельный элемент
          skuList.forEach((sku: string) => {
            grouped.get(vendorCode)?.items.push({
              nmId: String(row["ID товара"] || ""),
              size: String(row["Технический размер"] || ""),
              sku: sku.trim(),
              title: String(row["Наименование"] || ""),
              uniqueKey: `${row["ID товара"]}_${row["Технический размер"]}_${sku.trim()}`
            });
          });
        }
      });

      // Преобразуем в массив и сортируем по артикулу
      const groupedArray = Array.from(grouped.values())
        .sort((a, b) => a.vendorCode.localeCompare(b.vendorCode));

      setGroupedProducts(groupedArray);
      
      // Загружаем сохраненные данные себестоимости
      const savedCosts = loadCostsFromStorage();
      setSkuCosts(savedCosts);
      
      setShowCostModal(true);

    } catch (error) {
      console.error("Ошибка загрузки номенклатуры:", error);
      alert((error as Error).message || "Не удалось загрузить номенклатуру");
    } finally {
      setIsLoadingCosts(false);
    }
  };

  // Функция для массового применения себестоимости
  const handleApplyBulkCost = () => {
    if (!bulkCost.trim()) {
      alert("Введите себестоимость для массового применения");
      return;
    }

    const newSkuCosts: {[key: string]: string} = {};
    
    // Применяем к каждому SKU в каждой группе
    groupedProducts.forEach(product => {
      product.items.forEach((item: Record<string, unknown>) => {
        const key = String(item.sku || item.uniqueKey);
        newSkuCosts[key] = bulkCost;
      });
    });

    setSkuCosts(newSkuCosts);
    saveCostsToStorage(newSkuCosts); // Сохраняем в localStorage
    setBulkCost(""); // Очищаем поле после применения
    alert(`Себестоимость ${bulkCost} ₽ применена ко всем товарам`);
  };

  // Функция для очистки всех значений себестоимости
  const handleClearAllCosts = () => {
    if (confirm("Очистить все введенные значения себестоимости?")) {
      setSkuCosts({});
      setBulkCost("");
      saveCostsToStorage({}); // Очищаем localStorage
    }
  };

  // Функция для сохранения себестоимости
  const handleSaveCosts = async () => {
    try {
      setIsLoadingCosts(true);
      
      // Загружаем номенклатуру заново с обновленными данными
      const resNomenclature = await fetch("/api/wb/nomenclature", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token }),
      });

      if (!resNomenclature.ok) {
        throw new Error("Не удалось загрузить номенклатуру");
      }

      const nomenclature: { fields: string[]; rows: Record<string, unknown>[] } = await resNomenclature.json();
      
      // Обновляем данные себестоимости
      const updatedRows = nomenclature.rows.map((row: Record<string, unknown>) => {
        const skus = String(row["SKU"] || "");
        let cost = "";
        
        // Если есть штрихкоды, ищем себестоимость для первого найденного SKU
        if (skus) {
          const skuList = skus.split(';\n').filter((sku: string) => sku.trim() !== '');
          for (const sku of skuList) {
            const trimmedSku = sku.trim();
            if (skuCosts[trimmedSku]) {
              cost = skuCosts[trimmedSku];
              break; // Берем первую найденную себестоимость для строки
            }
          }
        }
        
        return {
          ...row,
          "Себестоимость": cost
        };
      });

      // Создаем Excel файл только с листом "Номенклатура"
      const nomenclatureHeader = nomenclature.fields;
      const nomenclatureRowsData = updatedRows.map((row) => nomenclatureHeader.map((key) => (row as Record<string, unknown>)[key] ?? ""));
      const nomenclatureSheet = XLSX.utils.aoa_to_sheet([nomenclatureHeader, ...nomenclatureRowsData]);
      
      // Устанавливаем ширину колонок
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

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, nomenclatureSheet, "Номенклатура");
      
      const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([arrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "Номенклатура_с_себестоимостью.xlsx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
      
      setShowCostModal(false);
      alert("Файл с себестоимостью успешно сохранен!");
      
    } catch (error) {
      console.error("Ошибка сохранения:", error);
      alert((error as Error).message || "Не удалось сохранить файл");
    } finally {
      setIsLoadingCosts(false);
    }
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

      let resReport, resPaid, resAcceptance, resFinanceRK, resNomenclature, resWarehouseRemains;
      
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

        // Задержка перед остатками на складах (лимит 1 запрос в минуту)
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        console.log("📊 Запуск отчета остатков на складах...");
        resWarehouseRemains = await fetch("/api/wb/warehouse-remains", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ token }),
        }).catch(err => {
          console.error("Ошибка fetch для остатков на складах:", err);
          // Возвращаем пустые данные вместо ошибки
          return {
            ok: true,
            json: async () => ({
              fields: [
                "Бренд",
                "Предмет",
                "Артикул продавца",
                "Артикул WB",
                "Штрихкод",
                "Размер",
                "Объем (л)",
                "Название склада",
                "ID склада",
                "Количество",
                "В пути к клиенту",
                "В пути от клиента",
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
        nomenclatureStatus: resNomenclature.ok ? 'success' : 'fallback',
        warehouseRemainsStatus: resWarehouseRemains.ok ? 'success' : 'fallback'
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
      const warehouseRemains: { fields: string[]; rows: Record<string, unknown>[] } = await resWarehouseRemains.json();

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

      // Создаем лист "Остатки"
      const warehouseRemainsHeader = warehouseRemains.fields;
      const warehouseRemainsRows = warehouseRemains.rows.map((row) => warehouseRemainsHeader.map((key) => row[key] ?? ""));
      const warehouseRemainsSheet = XLSX.utils.aoa_to_sheet([warehouseRemainsHeader, ...warehouseRemainsRows]);
      
      // Устанавливаем ширину колонок для листа "Остатки"
      const warehouseRemainsColWidths = [
        { wch: 15 }, // Бренд
        { wch: 20 }, // Предмет
        { wch: 20 }, // Артикул продавца
        { wch: 12 }, // Артикул WB
        { wch: 20 }, // Штрихкод
        { wch: 10 }, // Размер
        { wch: 12 }, // Объем (л)
        { wch: 25 }, // Название склада
        { wch: 10 }, // ID склада
        { wch: 12 }, // Количество
        { wch: 15 }, // В пути к клиенту
        { wch: 15 }, // В пути от клиента
        { wch: 15 }  // Дата выгрузки
      ];
      warehouseRemainsSheet["!cols"] = warehouseRemainsColWidths;

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
      // Создаем лист номенклатуры с интеграцией сохраненной себестоимости
      const savedCosts = loadCostsFromStorage();
      
      // Обновляем данные номенклатуры с себестоимостью
      const updatedNomenclatureRows = nomenclature.rows.map((row: Record<string, unknown>) => {
        const skus = String(row["SKU"] || "");
        let cost = "";
        
        // Если есть штрихкоды, ищем себестоимость для первого найденного SKU
        if (skus) {
          const skuList = skus.split(';\n').filter((sku: string) => sku.trim() !== '');
          for (const sku of skuList) {
            const trimmedSku = sku.trim();
            if (savedCosts[trimmedSku]) {
              cost = savedCosts[trimmedSku];
              break; // Берем первую найденную себестоимость для строки
            }
          }
        }
        
        return {
          ...row,
          "Себестоимость": cost
        };
      });
      
      const nomenclatureHeader = nomenclature.fields;
      const nomenclatureRows = updatedNomenclatureRows.map((row) => nomenclatureHeader.map((key) => (row as Record<string, unknown>)[key] ?? ""));
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
        const groupedProducts = new Map<string, Array<Record<string, unknown>>>();
        
        nomenclature.rows.forEach((row: Record<string, unknown>) => {
          const vendorCode = String(row["Артикул продавца"] || "Без артикула");
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
            products.forEach((product: Record<string, unknown>) => {
              productsData.push([
                vendorCode, // Артикул
                String(product["Технический размер"] || ""), // Размер - технический
                String(product["SKU"] || ""), // Штрихкод (используем SKU как штрихкод)
                String(product["ID товара"] || ""), // Артикул WB (nmID)
                String(product["Бренд"] || "") // Бренд
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
      XLSX.utils.book_append_sheet(workbook, warehouseRemainsSheet, "Остатки");
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
              
              <button
                type="button"
                onClick={handleLoadCosts}
                disabled={isLoadingCosts}
                className={`w-full h-11 rounded-lg bg-blue-600 text-white dark:bg-blue-500 dark:text-white font-medium transition-opacity ${
                  isLoadingCosts ? "opacity-60 cursor-not-allowed" : "hover:opacity-90"
                }`}
              >
                {isLoadingCosts ? (
                  <span className="inline-flex items-center gap-2">
                    <span className="inline-block h-4 w-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
                    Загрузка...
                  </span>
                ) : (
                  "Себестоимость"
                )}
              </button>
            </div>
          </div>
      </div>
      
      {/* Модальное окно для ввода себестоимости */}
      {showCostModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
          <div className="bg-white dark:bg-gray-800 rounded-2xl p-6 w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-semibold">Себестоимость товаров</h2>
              <button
                onClick={() => setShowCostModal(false)}
                className="text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200"
              >
                <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>
            
            <div className="text-sm text-gray-600 dark:text-gray-400 mb-4">
              Товары сгруппированы по артикулу. Введите себестоимость для каждого штрихкода отдельно.
            </div>
            
            {/* Блок массового применения себестоимости */}
            <div className="bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-lg p-4 mb-4">
              <div className="flex items-center justify-between">
                <div className="flex-1">
                  <h4 className="font-medium text-blue-900 dark:text-blue-100 mb-1">Массовое применение</h4>
                  <p className="text-sm text-blue-700 dark:text-blue-300">Установить одинаковую себестоимость для всех товаров</p>
                </div>
                <div className="flex items-center gap-3 ml-4">
                  <input
                    type="number"
                    placeholder="Себестоимость"
                    value={bulkCost}
                    onChange={(e) => setBulkCost(e.target.value)}
                    className="w-32 h-9 rounded border border-blue-300 dark:border-blue-600 px-3 text-sm bg-white dark:bg-gray-700"
                  />
                  <span className="text-sm text-blue-600 dark:text-blue-400">₽</span>
                  <button
                    onClick={handleApplyBulkCost}
                    className="px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded hover:bg-blue-700 transition-colors"
                  >
                    Применить ко всем
                  </button>
                  <button
                    onClick={handleClearAllCosts}
                    className="px-3 py-2 bg-gray-500 text-white text-sm font-medium rounded hover:bg-gray-600 transition-colors"
                  >
                    Очистить все
                  </button>
                </div>
              </div>
            </div>
            
            <div className="flex-1 overflow-y-auto mb-4">
              <div className="space-y-4">
                {groupedProducts.map((product, index) => (
                  <div key={index} className="border border-gray-200 dark:border-gray-600 rounded-lg p-4">
                    <div className="mb-4">
                      <h3 className="font-medium text-lg">{product.vendorCode}</h3>
                      <p className="text-sm text-gray-600 dark:text-gray-400">Бренд: {product.brand}</p>
                    </div>
                    
                    <div className="space-y-3">
                      <div className="text-xs font-medium text-gray-500 dark:text-gray-400 uppercase tracking-wide">
                        Штрихкоды товаров:
                      </div>
                      {product.items.map((item: Record<string, unknown>, itemIndex: number) => (
                        <div key={itemIndex} className="flex items-center justify-between bg-gray-50 dark:bg-gray-700 rounded-lg p-3">
                          <div className="flex-1 min-w-0">
                            <div className="font-medium text-sm truncate">{String(item.title || "")}</div>
                            <div className="text-xs text-gray-600 dark:text-gray-400 mt-1">
                              <span className="inline-block mr-3">Размер: <span className="font-medium">{String(item.size || "—")}</span></span>
                              <span className="inline-block mr-3">ШК: <span className="font-medium">{String(item.sku || "Нет ШК")}</span></span>
                              <span className="inline-block">WB: <span className="font-medium">{String(item.nmId || "")}</span></span>
                            </div>
                          </div>
                          <div className="flex items-center gap-2 ml-4">
                            <input
                              type="number"
                              placeholder="0"
                              value={skuCosts[String(item.sku || item.uniqueKey)] || ""}
                              onChange={(e) => {
                                const newCosts = {
                                  ...skuCosts,
                                  [String(item.sku || item.uniqueKey)]: e.target.value
                                };
                                setSkuCosts(newCosts);
                                saveCostsToStorage(newCosts); // Сохраняем при каждом изменении
                              }}
                              className="w-24 h-8 rounded border border-gray-300 dark:border-gray-600 px-2 text-sm bg-white dark:bg-gray-700 text-right"
                            />
                            <span className="text-xs text-gray-500">₽</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
            
            <div className="flex gap-3 pt-4 border-t border-gray-200 dark:border-gray-600">
              <button
                onClick={() => setShowCostModal(false)}
                className="flex-1 h-11 rounded-lg border border-gray-300 dark:border-gray-600 text-gray-700 dark:text-gray-300 font-medium hover:bg-gray-50 dark:hover:bg-gray-700 transition-colors"
              >
                Отмена
              </button>
              <button
                onClick={handleSaveCosts}
                disabled={isLoadingCosts}
                className={`flex-1 h-11 rounded-lg bg-green-600 text-white font-medium transition-opacity ${
                  isLoadingCosts ? "opacity-60 cursor-not-allowed" : "hover:opacity-90"
                }`}
              >
                {isLoadingCosts ? (
                  <span className="inline-flex items-center gap-2">
                    <span className="inline-block h-4 w-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
                    Сохранение...
                  </span>
                ) : (
                  "Сохранить файл"
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
