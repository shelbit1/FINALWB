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

  // Состояния для РНП (один день)
  const [rnpDate, setRnpDate] = useState("");
  const [isLoadingRnp, setIsLoadingRnp] = useState(false);
  const [isLoadingRemains, setIsLoadingRemains] = useState(false);
  const [isLoadingRemainsRnp, setIsLoadingRemainsRnp] = useState(false);
  const [isLoadingAnalysis, setIsLoadingAnalysis] = useState(false);
  
  // Состояния для выгрузки РК
  const [rkDateFrom, setRkDateFrom] = useState("");
  const [rkDateTo, setRkDateTo] = useState("");
  const [isLoadingRk, setIsLoadingRk] = useState(false);
  
  // Состояния для параметров остатков
  const [deliveryDays, setDeliveryDays] = useState("");
  const [stockDays, setStockDays] = useState("");
  const [coefficient, setCoefficient] = useState("");

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

  // Функция для скачивания РНП с полным набором листов
  const handleRnpDownload = async () => {
    try {
      setIsLoadingRnp(true);
      
      // Валидация данных
      if (!token.trim()) {
        alert("Введите API токен Wildberries");
        return;
      }
      
      if (!rnpDate) {
        alert("Выберите дату для выгрузки РНП");
        return;
      }
      
      const selectedDate = new Date(rnpDate);
      const today = new Date();
      today.setHours(23, 59, 59, 999);
      
      if (selectedDate > today) {
        alert("Дата не может быть в будущем");
        return;
      }
      
      // Используем одну и ту же дату для начала и конца периода
      const payload = { token, dateFrom: rnpDate, dateTo: rnpDate };

      console.log("📊 Запуск расширенного отчета РНП с дополнительными листами...", payload);

      let resRnp, resPaid, resAcceptance, resFinanceRK, resNomenclature, resWarehouseRemains;
      
      try {
        // Основной РНП отчет
        console.log("📊 Запуск РНП (ежедневные отчеты)...");
        resRnp = await fetch("/api/wb/rnp", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        }).catch(err => {
          console.error("Ошибка fetch для РНП:", err);
          throw new Error(`Ошибка сетевого запроса для РНП: ${err.message}`);
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
          body: JSON.stringify({ token, startDate: rnpDate, endDate: rnpDate }),
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
                "Дата выгрузки",
                "Себестоимость"
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
        rnpStatus: resRnp.status, 
        paidStatus: resPaid.status,
        acceptanceStatus: resAcceptance.ok ? 'success' : 'fallback',
        financeRKStatus: resFinanceRK.ok ? 'success' : 'fallback',
        nomenclatureStatus: resNomenclature.ok ? 'success' : 'fallback',
        warehouseRemainsStatus: resWarehouseRemains.ok ? 'success' : 'fallback'
      });

      if (!resRnp.ok) {
        const err = await resRnp.json().catch(() => ({}));
        throw new Error(err.error || `Ошибка РНП: ${resRnp.status}`);
      }
      if (!resPaid.ok) {
        const err = await resPaid.json().catch(() => ({}));
        throw new Error(err.error || `Ошибка платного хранения: ${resPaid.status}`);
      }

      const rnp: { fields: string[]; rows: Record<string, unknown>[] } = await resRnp.json();
      const paid: { fields: string[]; rows: Record<string, unknown>[] } = await resPaid.json();
      const acceptance: { fields: string[]; rows: Record<string, unknown>[] } = await resAcceptance.json();
      const financeRK: { fields: string[]; rows: Record<string, unknown>[] } = await resFinanceRK.json();
      const nomenclature: { fields: string[]; rows: Record<string, unknown>[] } = await resNomenclature.json();
      const warehouseRemains: { fields: string[]; rows: Record<string, unknown>[] } = await resWarehouseRemains.json();

      // Создаем Excel файл с множественными листами
      const workbook = XLSX.utils.book_new();

      // Лист РНП (основной)
      const rnpHeader = rnp.fields;
      const rnpRows = rnp.rows.map((row) => rnpHeader.map((key) => row[key] ?? ""));
      const rnpSheet = XLSX.utils.aoa_to_sheet([rnpHeader, ...rnpRows]);
      
      // Устанавливаем ширину колонок для РНП (сокращенные для экономии места)
      const rnpColWidths = Array(rnpHeader.length).fill({ wch: 12 });
      rnpSheet["!cols"] = rnpColWidths;

      // Лист платного хранения
      const paidHeader = paid.fields;
      const paidRows = paid.rows.map((row) => paidHeader.map((key) => row[key] ?? ""));
      const paidSheet = XLSX.utils.aoa_to_sheet([paidHeader, ...paidRows]);

      // Лист платной приемки
      const acceptanceHeader = acceptance.fields;
      const acceptanceRows = acceptance.rows.map((row) => acceptanceHeader.map((key) => row[key] ?? ""));
      const acceptanceSheet = XLSX.utils.aoa_to_sheet([acceptanceHeader, ...acceptanceRows]);

      // Лист финансов РК
      const financeRKHeader = financeRK.fields;
      const financeRKRows = financeRK.rows.map((row) => {
        return financeRKHeader.map((key) => {
          const value = row[key] ?? "";
          // Специальная обработка для колонки "Сумма" - сохраняем как число
          if (key === 'Сумма') {
            if (typeof value === 'number') {
              return value;
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

      // Лист остатков на складах
      const warehouseRemainsHeader = warehouseRemains.fields;
      const warehouseRemainsRows = warehouseRemains.rows.map((row) => warehouseRemainsHeader.map((key) => row[key] ?? ""));
      const warehouseRemainsSheet = XLSX.utils.aoa_to_sheet([warehouseRemainsHeader, ...warehouseRemainsRows]);

      // Лист номенклатуры с интеграцией сохраненной себестоимости
      const savedCosts = loadCostsFromStorage();
      
      const updatedNomenclatureRows = nomenclature.rows.map((row: Record<string, unknown>) => {
        const skus = String(row["SKU"] || "");
        let cost = "";
        
        if (skus) {
          const skuList = skus.split(';\n').filter((sku: string) => sku.trim() !== '');
          for (const sku of skuList) {
            const trimmedSku = sku.trim();
            if (savedCosts[trimmedSku]) {
              cost = savedCosts[trimmedSku];
              break;
            }
          }
        }
        
        return {
          ...row,
          "Себестоимость": cost
        };
      });
      
      // Убеждаемся, что поле "Себестоимость" включено в заголовки для РНП
      const nomenclatureHeader = nomenclature.fields.includes("Себестоимость") 
        ? nomenclature.fields 
        : [...nomenclature.fields, "Себестоимость"];
      
      const nomenclatureRows = updatedNomenclatureRows.map((row) => nomenclatureHeader.map((key) => (row as Record<string, unknown>)[key] ?? ""));
      const nomenclatureSheet = XLSX.utils.aoa_to_sheet([nomenclatureHeader, ...nomenclatureRows]);
      
      // Устанавливаем ширину колонок для номенклатуры в РНП
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

      // Создаем лист "Аналитика по товарам" из номенклатуры, сгруппированный по артикулу
      const createProductAnalyticsSheet = () => {
        // Группируем товары по артикулу продавца
        const groupedProducts = new Map<string, Array<Record<string, unknown>>>();
        
        nomenclature.rows.forEach((row: Record<string, unknown>) => {
          const vendorCode = String(row["Артикул продавца"] || "Без артикула");
          if (!groupedProducts.has(vendorCode)) {
            groupedProducts.set(vendorCode, []);
          }
          groupedProducts.get(vendorCode)?.push(row);
        });

        // Создаем заголовки для листа "Аналитика по товарам" с пустыми строками сверху
        const analyticsHeaders = ["Артикул", "Размер", "Штрихкод", "Артикул WB", "Бренд"];
        const analyticsData = [
          [], // Пустая строка 1
          [], // Пустая строка 2
          analyticsHeaders // Заголовки в строке 3
        ];

        // Добавляем данные, сгруппированные по артикулу
        Array.from(groupedProducts.entries())
          .sort(([a], [b]) => a.localeCompare(b)) // Сортируем по артикулу
          .forEach(([vendorCode, products]) => {
            products.forEach((product: Record<string, unknown>) => {
              analyticsData.push([
                vendorCode, // Артикул
                String(product["Технический размер"] || ""), // Размер - технический
                String(product["SKU"] || ""), // Штрихкод (используем SKU как штрихкод)
                String(product["ID товара"] || ""), // Артикул WB (nmID)
                String(product["Бренд"] || "") // Бренд
              ]);
            });
          });

        return XLSX.utils.aoa_to_sheet(analyticsData);
      };

      const productAnalyticsSheet = createProductAnalyticsSheet();
      
      // Устанавливаем ширину колонок для листа "Аналитика по товарам"
      productAnalyticsSheet["!cols"] = [
        { wch: 20 }, // Артикул
        { wch: 15 }, // Размер
        { wch: 25 }, // Штрихкод
        { wch: 15 }, // Артикул WB
        { wch: 20 }  // Бренд
      ];

      // Добавляем все листы в книгу (Аналитика по товарам идет первой)
      XLSX.utils.book_append_sheet(workbook, productAnalyticsSheet, "Аналитика по товарам");
      XLSX.utils.book_append_sheet(workbook, rnpSheet, "РНП");
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
      link.download = `РНП_Полный_${rnpDate}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
      
    } catch (error) {
      console.error("Ошибка РНП:", error);
      const errorMessage = (error as Error).message || "Не удалось сформировать РНП";
      
      // Более понятные сообщения об ошибках для РНП
      let userFriendlyMessage = errorMessage;
      if (errorMessage.includes("401") || errorMessage.includes("Unauthorized")) {
        userFriendlyMessage = "Ошибка авторизации: проверьте корректность API токена Wildberries";
      } else if (errorMessage.includes("403") || errorMessage.includes("Forbidden")) {
        userFriendlyMessage = "Доступ запрещен: убедитесь, что у токена есть необходимые права доступа";
      } else if (errorMessage.includes("429") || errorMessage.includes("Too Many Requests")) {
        userFriendlyMessage = "Превышен лимит запросов к API Wildberries. Подождите 1-2 минуты перед повторной попыткой";
      } else if (errorMessage.includes("500") || errorMessage.includes("Internal Server Error")) {
        userFriendlyMessage = "Внутренняя ошибка сервера Wildberries. Попробуйте позже";
      }
      
      alert(userFriendlyMessage);
    } finally {
      setIsLoadingRnp(false);
    }
  };

  const handleRemainsDownload = async () => {
    try {
      setIsLoadingRemains(true);
      
      // Валидация токена
      if (!token.trim()) {
        alert("Введите API токен Wildberries");
        return;
      }
      
      console.log("📊 Запуск отчета остатков на складах...");
      
      // Запрос остатков на складах
      const resWarehouseRemains = await fetch("/api/wb/warehouse-remains", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token }),
      });

      if (!resWarehouseRemains.ok) {
        const errorData = await resWarehouseRemains.json().catch(() => ({}));
        throw new Error(errorData.error || "Ошибка при получении остатков на складах");
      }

      const warehouseRemains = await resWarehouseRemains.json();
      console.log("✅ Остатки на складах получены:", warehouseRemains.rows?.length || 0, "строк");

      // Создание Excel файла
      const workbook = XLSX.utils.book_new();
      
      // Добавляем лист "Остатки"
      const remainsHeader = warehouseRemains.fields || [];
      const remainsRows = (warehouseRemains.rows || []).map((row: Record<string, unknown>) => 
        remainsHeader.map((key: string) => row[key] ?? "")
      );
      const remainsSheet = XLSX.utils.aoa_to_sheet([remainsHeader, ...remainsRows]);
      
      // Настройка ширины колонок для остатков
      const remainsColWidths = [
        { wch: 20 }, // Бренд
        { wch: 20 }, // Предмет
        { wch: 20 }, // Артикул продавца
        { wch: 12 }, // Артикул WB
        { wch: 15 }, // Штрихкод
        { wch: 10 }, // Размер
        { wch: 10 }, // Объем (л)
        { wch: 25 }, // Название склада
        { wch: 12 }, // ID склада
        { wch: 12 }, // Количество
        { wch: 15 }, // В пути к клиенту
        { wch: 15 }, // В пути от клиента
        { wch: 12 }  // Дата выгрузки
      ];
      remainsSheet["!cols"] = remainsColWidths;

      XLSX.utils.book_append_sheet(workbook, remainsSheet, "Остатки");

      // Генерация и скачивание файла
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      
      const currentDate = new Date().toISOString().split('T')[0];
      link.download = `Остатки_на_складах_${currentDate}.xlsx`;
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      console.log("✅ Файл остатков успешно скачан");
    } catch (error) {
      console.error("❌ Ошибка при скачивании остатков:", error);
      
      let userFriendlyMessage = "Произошла ошибка при загрузке остатков";
      
      if (error instanceof Error) {
        if (error.message.includes("Failed to fetch") || error.message.includes("NetworkError")) {
          userFriendlyMessage = "Ошибка сети. Проверьте подключение к интернету";
        } else if (error.message.includes("401") || error.message.includes("авторизац")) {
          userFriendlyMessage = "Ошибка авторизации. Проверьте правильность токена";
        } else if (error.message.includes("429")) {
          userFriendlyMessage = "Превышен лимит запросов. Попробуйте через минуту";
        } else if (error.message.includes("500")) {
          userFriendlyMessage = "Внутренняя ошибка сервера Wildberries. Попробуйте позже";
        } else {
          userFriendlyMessage = error.message;
        }
      }
      
      alert(userFriendlyMessage);
    } finally {
      setIsLoadingRemains(false);
    }
  };

  const handleRemainsRnpDownload = async () => {
    try {
      setIsLoadingRemainsRnp(true);
      
      // Валидация токена
      if (!token.trim()) {
        alert("Введите API токен Wildberries");
        return;
      }
      
      // Валидация полей срока поставки и запаса
      const delivery = parseFloat(deliveryDays);
      const stock = parseFloat(stockDays);
      
      if (!deliveryDays || isNaN(delivery) || delivery <= 0) {
        alert("Введите корректный срок поставки (больше 0)");
        return;
      }
      
      if (!stockDays || isNaN(stock) || stock <= 0) {
        alert("Введите корректный запас (больше 0)");
        return;
      }
      
      // Рассчитываем период: (Срок поставки + Запас) дней, заканчивая вчерашним днем
      const totalDays = Math.ceil(delivery + stock); // Округляем вверх для дробных значений
      
      // Вчерашний день (конец периода)
      const yesterday = new Date();
      yesterday.setDate(yesterday.getDate() - 1);
      yesterday.setHours(0, 0, 0, 0);
      
      // Начало периода (вчера минус totalDays дней)
      const startDate = new Date(yesterday);
      startDate.setDate(yesterday.getDate() - totalDays + 1); // +1 потому что включаем вчерашний день
      
      const dateFrom = startDate.toISOString().split('T')[0];
      const dateTo = yesterday.toISOString().split('T')[0];
      
      console.log(`📊 Запуск РНП за период ${totalDays} дней: ${dateFrom} - ${dateTo}`);
      
      const payload = { token, dateFrom, dateTo };
      
      // Запрос РНП данных
      const resRnp = await fetch("/api/wb/rnp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!resRnp.ok) {
        const errorData = await resRnp.json().catch(() => ({}));
        throw new Error(errorData.error || "Ошибка при получении данных РНП");
      }

      const rnpData = await resRnp.json();
      console.log("✅ Данные РНП получены:", rnpData.rows?.length || 0, "строк");

      // Создание Excel файла
      const workbook = XLSX.utils.book_new();
      
      // Добавляем лист "РНП"
      const rnpHeader = rnpData.fields || [];
      const rnpRows = (rnpData.rows || []).map((row: Record<string, unknown>) => 
        rnpHeader.map((key: string) => row[key] ?? "")
      );
      const rnpSheet = XLSX.utils.aoa_to_sheet([rnpHeader, ...rnpRows]);
      
      // Настройка ширины колонок для РНП
      const rnpColWidths = [
        { wch: 12 },  // ID отчета
        { wch: 12 },  // Дата начала
        { wch: 12 },  // Дата окончания
        { wch: 12 },  // Дата создания
        { wch: 10 },  // Валюта
        { wch: 15 },  // Договор
        { wch: 10 },  // Номер отчета
        { wch: 12 },  // Старый ID отчета
        { wch: 20 },  // Артикул продавца
        { wch: 12 },  // Размер
        { wch: 15 },  // Штрихкод
        { wch: 12 },  // Всего
        { wch: 12 },  // Количество доставок
        { wch: 12 },  // Количество возвратов
        { wch: 15 },  // Цена розничная
        { wch: 15 },  // Скидка продавца
        { wch: 15 },  // Скидка WB
        { wch: 12 },  // Промокод
        { wch: 15 },  // Цена со скидкой
        { wch: 15 },  // Комиссия WB
        { wch: 12 },  // Оплата продавцу
        { wch: 15 },  // К перечислению
        { wch: 12 },  // Дата продажи
        { wch: 12 },  // ГП
        { wch: 12 },  // Номер поставки
        { wch: 15 },  // Страна
        { wch: 15 },  // Область
        { wch: 12 },  // Артикул WB
        { wch: 12 },  // Тип документа
        { wch: 12 },  // Номер заказа
        { wch: 20 },  // Наименование
        { wch: 15 }   // Офис
      ];
      rnpSheet["!cols"] = rnpColWidths;

      XLSX.utils.book_append_sheet(workbook, rnpSheet, "РНП");

      // Генерация и скачивание файла
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      
      link.download = `РНП_${totalDays}дн_${dateFrom}_${dateTo}.xlsx`;
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      console.log("✅ Файл РНП успешно скачан");
    } catch (error) {
      console.error("❌ Ошибка при скачивании РНП:", error);
      
      let userFriendlyMessage = "Произошла ошибка при загрузке РНП";
      
      if (error instanceof Error) {
        if (error.message.includes("Failed to fetch") || error.message.includes("NetworkError")) {
          userFriendlyMessage = "Ошибка сети. Проверьте подключение к интернету";
        } else if (error.message.includes("401") || error.message.includes("авторизац")) {
          userFriendlyMessage = "Ошибка авторизации. Проверьте правильность токена";
        } else if (error.message.includes("429")) {
          userFriendlyMessage = "Превышен лимит запросов. Попробуйте через минуту";
        } else if (error.message.includes("500")) {
          userFriendlyMessage = "Внутренняя ошибка сервера Wildberries. Попробуйте позже";
        } else {
          userFriendlyMessage = error.message;
        }
      }
      
      alert(userFriendlyMessage);
    } finally {
      setIsLoadingRemainsRnp(false);
    }
  };

  const handleSupplyAnalysisDownload = async () => {
    try {
      setIsLoadingAnalysis(true);
      
      // Валидация токена
      if (!token.trim()) {
        alert("Введите API токен Wildberries");
        return;
      }
      
      // Валидация полей срока поставки и запаса
      const delivery = parseFloat(deliveryDays);
      const stock = parseFloat(stockDays);
      
      if (!deliveryDays || isNaN(delivery) || delivery <= 0) {
        alert("Введите корректный срок поставки (больше 0)");
        return;
      }
      
      if (!stockDays || isNaN(stock) || stock <= 0) {
        alert("Введите корректный запас (больше 0)");
        return;
      }
      
      console.log("📊 Запуск анализа поставок (Остатки + РНП + Заказы)...");
      
      // Рассчитываем период для РНП
      const totalDays = Math.ceil(delivery + stock);
      const yesterday = new Date();
      yesterday.setDate(yesterday.getDate() - 1);
      yesterday.setHours(0, 0, 0, 0);
      
      const startDate = new Date(yesterday);
      startDate.setDate(yesterday.getDate() - totalDays + 1);
      
      const dateFrom = startDate.toISOString().split('T')[0];
      const dateTo = yesterday.toISOString().split('T')[0];
      
      const payload = { token, dateFrom, dateTo };
      
      // Последовательный запрос с задержками для соблюдения лимитов API
      console.log("📊 Загрузка данных с задержками между запросами...");
      
      console.log("📊 Загрузка остатков...");
      const resWarehouseRemains = await fetch("/api/wb/warehouse-remains", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token }),
      });
      
      // Задержка 2 секунды
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      console.log("📊 Загрузка РНП...");
      const resRnp = await fetch("/api/wb/rnp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      
      // Задержка 2 секунды
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      console.log("📊 Загрузка номенклатуры...");
      const resNomenclature = await fetch("/api/wb/nomenclature", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token }),
      });
      
      // Задержка 2 секунды перед заказами
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      console.log("📊 Загрузка заказов...");
      const resOrders = await fetch("/api/wb/orders", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token, dateFrom: dateFrom, dateTo: dateTo }),
      });

      // Проверка ответов
      if (!resWarehouseRemains.ok) {
        const errorData = await resWarehouseRemains.json().catch(() => ({}));
        throw new Error(errorData.error || "Ошибка при получении остатков на складах");
      }

      if (!resRnp.ok) {
        const errorData = await resRnp.json().catch(() => ({}));
        throw new Error(errorData.error || "Ошибка при получении данных РНП");
      }

      if (!resNomenclature.ok) {
        const errorData = await resNomenclature.json().catch(() => ({}));
        throw new Error(errorData.error || "Ошибка при получении номенклатуры");
      }

      if (!resOrders.ok) {
        const errorData = await resOrders.json().catch(() => ({}));
        throw new Error(errorData.error || "Ошибка при получении заказов");
      }

      const warehouseRemains = await resWarehouseRemains.json();
      const rnpData = await resRnp.json();
      const nomenclature = await resNomenclature.json();
      const ordersData = await resOrders.json();
      
      console.log("✅ Остатки получены:", warehouseRemains.rows?.length || 0, "строк");
      console.log("✅ РНП получены:", rnpData.rows?.length || 0, "строк");
      console.log("✅ Номенклатура получена:", nomenclature.rows?.length || 0, "строк");
      console.log("✅ Заказы получены:", ordersData.rows?.length || 0, "строк");

      // Группировка товаров по артикулу продавца
      const groupedProducts: Record<string, {
        vendorCode: string;
        brand: string;
        sizes: Array<{
          size: string;
          barcode: string;
          nmId: string;
        }>;
      }> = {};

      (nomenclature.rows || []).forEach((item: Record<string, unknown>) => {
        const vendorCode = String(item["Артикул продавца"] || "");
        const brand = String(item["Бренд"] || "");
        const techSize = String(item["Технический размер"] || "");
        const skus = String(item["SKU"] || "");
        const nmId = String(item["ID товара"] || "");
        
        if (!vendorCode) return;
        
        if (!groupedProducts[vendorCode]) {
          groupedProducts[vendorCode] = {
            vendorCode,
            brand,
            sizes: []
          };
        }
        
        // SKU содержит штрихкоды, разделенные ;\n
        // Технический размер уже в отдельном поле
        // ID товара (nmId) тоже в отдельном поле
        if (skus && skus.trim() !== '') {
          const barcodes = skus.split(';\n').filter((barcode: string) => barcode.trim() !== '');
          
          // Если есть штрихкоды, добавляем запись для каждого штрихкода
          if (barcodes.length > 0) {
            barcodes.forEach((barcode: string) => {
              groupedProducts[vendorCode].sizes.push({
                size: techSize,
                barcode: barcode.trim(),
                nmId: nmId
              });
            });
          } else {
            // Если штрихкодов нет, но есть размер
            groupedProducts[vendorCode].sizes.push({
              size: techSize,
              barcode: "",
              nmId: nmId
            });
          }
        } else if (techSize) {
          // Если SKU пусто, но есть размер
          groupedProducts[vendorCode].sizes.push({
            size: techSize,
            barcode: "",
            nmId: nmId
          });
        }
      });

      // Создание Excel файла с листами
      const workbook = XLSX.utils.book_new();
      
      // Лист 1: Аналитика (группированные товары)
      const analyticsData: unknown[][] = [];
      
      // Заголовки (строка 3)
      analyticsData.push([]); // Строка 1 (пустая)
      analyticsData.push([]); // Строка 2 (пустая)
      analyticsData.push(["Артикул", "Размер", "Штрихкод", "Артикул WB", "Бренд"]); // Строка 3 - заголовки
      
      // Данные по товарам
      Object.values(groupedProducts).forEach((product) => {
        if (product.sizes.length === 0) {
          // Если нет размеров, добавляем одну строку с артикулом и брендом
          analyticsData.push([
            product.vendorCode,
            "",
            "",
            "",
            product.brand
          ]);
        } else {
          // Для каждого размера добавляем отдельную строку
          product.sizes.forEach((size) => {
            analyticsData.push([
              product.vendorCode,
              size.size,
              size.barcode,
              size.nmId,
              product.brand
            ]);
          });
        }
      });
      
      const analyticsSheet = XLSX.utils.aoa_to_sheet(analyticsData);
      
      // Настройка ширины колонок для аналитики
      const analyticsColWidths = [
        { wch: 20 }, // Артикул
        { wch: 12 }, // Размер
        { wch: 15 }, // Штрихкод
        { wch: 12 }, // Артикул WB
        { wch: 20 }  // Бренд
      ];
      analyticsSheet["!cols"] = analyticsColWidths;
      
      XLSX.utils.book_append_sheet(workbook, analyticsSheet, "Аналитика");
      
      // Лист 2: Остатки
      const remainsHeader = warehouseRemains.fields || [];
      const remainsRows = (warehouseRemains.rows || []).map((row: Record<string, unknown>) => 
        remainsHeader.map((key: string) => row[key] ?? "")
      );
      const remainsSheet = XLSX.utils.aoa_to_sheet([remainsHeader, ...remainsRows]);
      
      const remainsColWidths = [
        { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 12 }, { wch: 15 },
        { wch: 10 }, { wch: 10 }, { wch: 25 }, { wch: 12 }, { wch: 12 },
        { wch: 15 }, { wch: 15 }, { wch: 12 }
      ];
      remainsSheet["!cols"] = remainsColWidths;
      XLSX.utils.book_append_sheet(workbook, remainsSheet, "Остатки");

      // Лист 3: РНП
      const rnpHeader = rnpData.fields || [];
      const rnpRows = (rnpData.rows || []).map((row: Record<string, unknown>) => 
        rnpHeader.map((key: string) => row[key] ?? "")
      );
      const rnpSheet = XLSX.utils.aoa_to_sheet([rnpHeader, ...rnpRows]);
      
      const rnpColWidths = [
        { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 },
        { wch: 15 }, { wch: 10 }, { wch: 12 }, { wch: 20 }, { wch: 12 },
        { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 },
        { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 15 }, { wch: 15 },
        { wch: 12 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
        { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
        { wch: 20 }, { wch: 15 }
      ];
      rnpSheet["!cols"] = rnpColWidths;
      XLSX.utils.book_append_sheet(workbook, rnpSheet, "РНП");

      // Лист 4: Номенклатура с себестоимостью
      const savedCosts = loadCostsFromStorage(); // Загружаем сохраненные себестоимости
      
      // Обновляем строки номенклатуры с себестоимостью
      const updatedNomenclatureRows = (nomenclature.rows || []).map((row: Record<string, unknown>) => {
        const skus = String(row["SKU"] || "");
        let cost = "";
        
        if (skus) {
          const skuList = skus.split(';\n').filter((sku: string) => sku.trim() !== '');
          for (const sku of skuList) {
            const trimmedSku = sku.trim();
            if (savedCosts[trimmedSku]) {
              cost = savedCosts[trimmedSku];
              break;
            }
          }
        }
        
        return { ...row, "Себестоимость": cost };
      });
      
      // Убеждаемся, что "Себестоимость" есть в заголовках
      const nomenclatureHeader = nomenclature.fields.includes("Себестоимость") 
        ? nomenclature.fields 
        : [...nomenclature.fields, "Себестоимость"];
      
      const nomenclatureRows = updatedNomenclatureRows.map((row: Record<string, unknown>) => 
        nomenclatureHeader.map((key: string) => row[key] ?? "")
      );
      const nomenclatureSheet = XLSX.utils.aoa_to_sheet([nomenclatureHeader, ...nomenclatureRows]);
      
      // Настройка ширины колонок для номенклатуры
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
      
      XLSX.utils.book_append_sheet(workbook, nomenclatureSheet, "Номенклатура");

      // Лист 5: Заказы
      const ordersHeader = ordersData.fields || [];
      const ordersRows = (ordersData.rows || []).map((row: Record<string, unknown>) => 
        ordersHeader.map((key: string) => row[key] ?? "")
      );
      const ordersSheet = XLSX.utils.aoa_to_sheet([ordersHeader, ...ordersRows]);
      
      // Настройка ширины колонок для заказов
      const ordersColWidths = [
        { wch: 12 }, // Дата
        { wch: 18 }, // Дата изменения
        { wch: 20 }, // Склад
        { wch: 15 }, // Тип склада
        { wch: 15 }, // Страна
        { wch: 20 }, // Округ
        { wch: 15 }, // Регион
        { wch: 20 }, // Артикул продавца
        { wch: 12 }, // Артикул WB
        { wch: 15 }, // Штрихкод
        { wch: 20 }, // Категория
        { wch: 20 }, // Предмет
        { wch: 15 }, // Бренд
        { wch: 12 }, // Размер
        { wch: 12 }, // ID поставки
        { wch: 10 }, // Поставка
        { wch: 12 }, // Реализация
        { wch: 15 }, // Цена без скидки
        { wch: 10 }, // Скидка %
        { wch: 10 }, // СПП
        { wch: 18 }, // Цена после всех скидок
        { wch: 15 }, // Цена со скидкой
        { wch: 10 }, // Отменен
        { wch: 18 }, // Дата отмены
        { wch: 15 }, // Тип заказа
        { wch: 15 }, // Стикер
        { wch: 20 }, // Номер заказа
        { wch: 35 }, // SRID
        { wch: 12 }  // Количество
      ];
      ordersSheet["!cols"] = ordersColWidths;
      XLSX.utils.book_append_sheet(workbook, ordersSheet, "Заказы");

      // Лист 6: Значения (параметры) - последний лист
      const coeffValue = coefficient ? parseFloat(coefficient) : 0;
      const deliveryValue = delivery; // Срок поставки уже распарсен выше
      const stockValue = stock; // Запас уже распарсен выше
      
      const valuesData: unknown[][] = [
        ["Параметр", "Значение"],
        ["Срок поставки (дн.)", deliveryValue],
        ["Запас (дн.)", stockValue],
        ["Коэффициент", coeffValue],
        [], // Строка 5 (пустая)
        [], // Строка 6 (пустая)
        [], // Строка 7 (пустая)
        [], // Строка 8 (пустая)
        [], // Строка 9 (пустая)
        ["Центральный"], // Строка 10
        ["Пушкино"],
        ["Вёшки"],
        ["Иваново"],
        ["Подольск 3"],
        ["Радумля 1"],
        ["Подольск 4"],
        ["Обухово 2"],
        ["Чашниково"],
        ["Истра"],
        ["Коледино: Горючее"],
        ["Обухово СГТ"],
        ["Голицыно СГТ"],
        ["Радумля СГТ"],
        ["Софьино СГТ"],
        ["Софьино СГТ"],
        ["Ярославль СГТ"],
        ["Цифровой склад"],
        ["Рязань (Тюшевское)"],
        ["Сабурово"],
        ["Владимир"],
        ["Тула"],
        ["Котовск"],
        ["Электросталь"],
        ["Воронеж"],
        ["Обухово"],
        ["Коледино"],
        ["Белая дача"],
        ["Подольск"],
        ["Щербинка"],
        ["Чехов 1"],
        ["Чехов 2"],
        ["Белые Столбы"],
        [], // Строка 43 (пустая)
        ["Екб"], // Строка 44
        ["Екатеринбург - Испытателей 14г"],
        ["Екатеринбург - Перспективный 12/2"],
        [], // Строка 47 (пустая)
        ["Приволжский"], // Строка 48
        ["СЦ Ижевск"],
        ["СЦ Кузнецк"],
        ["Пенза СГТ"],
        ["Кузнецк СГТ"],
        ["Пенза"],
        ["Самара (Новосемейкино)"],
        ["Сарапул"],
        ["Казань"],
        [], // Строка 57 (пустая)
        ["Юг + Кавказ"], // Строка 58
        ["Крыловская"],
        ["Волгоград"],
        ["Невинномысск"],
        ["Краснодар"]
      ];
      
      const valuesSheet = XLSX.utils.aoa_to_sheet(valuesData);
      
      // Настройка ширины колонок для значений
      const valuesColWidths = [
        { wch: 25 }, // Параметр
        { wch: 15 }  // Значение
      ];
      valuesSheet["!cols"] = valuesColWidths;
      
      XLSX.utils.book_append_sheet(workbook, valuesSheet, "Значения");

      // Генерация и скачивание файла
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      
      const currentDate = new Date().toISOString().split('T')[0];
      link.download = `Анализ_поставок_${totalDays}дн_${currentDate}.xlsx`;
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      console.log("✅ Файл анализа поставок успешно скачан");
    } catch (error) {
      console.error("❌ Ошибка при скачивании анализа поставок:", error);
      
      let userFriendlyMessage = "Произошла ошибка при загрузке анализа поставок";
      
      if (error instanceof Error) {
        if (error.message.includes("Failed to fetch") || error.message.includes("NetworkError")) {
          userFriendlyMessage = "Ошибка сети. Проверьте подключение к интернету";
        } else if (error.message.includes("401") || error.message.includes("авторизац")) {
          userFriendlyMessage = "Ошибка авторизации. Проверьте правильность токена";
        } else if (error.message.includes("429")) {
          userFriendlyMessage = "Превышен лимит запросов. Попробуйте через минуту";
        } else if (error.message.includes("500")) {
          userFriendlyMessage = "Внутренняя ошибка сервера Wildberries. Попробуйте позже";
        } else {
          userFriendlyMessage = error.message;
        }
      }
      
      alert(userFriendlyMessage);
    } finally {
      setIsLoadingAnalysis(false);
    }
  };

  const handleRkDownload = async () => {
    try {
      setIsLoadingRk(true);
      
      // Валидация данных
      if (!token.trim()) {
        alert("Введите API токен Wildberries");
        return;
      }
      
      if (!rkDateFrom || !rkDateTo) {
        alert("Выберите период для выгрузки РК (от даты и до даты)");
        return;
      }
      
      const dateFrom = new Date(rkDateFrom);
      const dateTo = new Date(rkDateTo);
      const today = new Date();
      today.setHours(23, 59, 59, 999);
      
      if (dateFrom > dateTo) {
        alert("Дата начала периода не может быть позже даты окончания");
        return;
      }
      
      if (dateTo > today) {
        alert("Дата окончания периода не может быть в будущем");
        return;
      }
      
      console.log("📊 Загрузка данных рекламных кампаний за период:", rkDateFrom, "-", rkDateTo);
      
      // Запрос данных рекламных кампаний
      const resCampaigns = await fetch("/api/wb/campaigns", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token }),
      });

      if (!resCampaigns.ok) {
        const errorData = await resCampaigns.json().catch(() => ({}));
        throw new Error(errorData.error || "Ошибка при получении данных рекламных кампаний");
      }

      const campaignsData: { 
        fields: string[]; 
        rows: Record<string, unknown>[]; 
        detailedCampaigns?: Array<Record<string, unknown>>
      } = await resCampaigns.json();
      console.log("✅ Данные рекламных кампаний получены:", campaignsData.rows?.length || 0, "строк");
      
      // Создаем Excel файл
      const workbook = XLSX.utils.book_new();
      
      // Лист 1: "РК" - основная информация
      const rkHeader = campaignsData.fields || [];
      const rkRows = (campaignsData.rows || []).map((row: Record<string, unknown>) => 
        rkHeader.map((key: string) => row[key] ?? "")
      );
      const rkSheet = XLSX.utils.aoa_to_sheet([rkHeader, ...rkRows]);
      
      // Настройка ширины колонок для листа "РК"
      const rkColWidths = [
        { wch: 12 }, // ID кампании
        { wch: 30 }, // Название кампании
        { wch: 25 }, // Тип
        { wch: 18 }, // Статус
        { wch: 20 }, // Дата создания
        { wch: 20 }, // Дата изменения
        { wch: 20 }, // Дата начала
        { wch: 20 }, // Дата окончания
        { wch: 15 }  // Дневной бюджет
      ];
      rkSheet["!cols"] = rkColWidths;
      
      XLSX.utils.book_append_sheet(workbook, rkSheet, "РК");

      // Лист 2: "Информация о компаниях" - детальная информация
      if (campaignsData.detailedCampaigns && campaignsData.detailedCampaigns.length > 0) {
        console.log("📊 Создание листа с детальной информацией о кампаниях...");
        
        // Функция для безопасного преобразования значений с ограничением длины
        const safeStringify = (value: unknown): string => {
          const MAX_EXCEL_CELL_LENGTH = 32767; // Максимальная длина текста в ячейке Excel
          
          if (value === null || value === undefined) return '';
          
          let result = '';
          if (typeof value === 'object') {
            try {
              result = JSON.stringify(value, null, 0);
            } catch {
              result = String(value);
            }
          } else {
            result = String(value);
          }
          
          // Обрезаем текст, если он слишком длинный
          if (result.length > MAX_EXCEL_CELL_LENGTH) {
            return result.substring(0, MAX_EXCEL_CELL_LENGTH - 20) + '... (обрезано)';
          }
          
          return result;
        };

        // Собираем все возможные ключи из всех кампаний
        const allKeys = new Set<string>();
        campaignsData.detailedCampaigns.forEach((campaign: Record<string, unknown>) => {
          Object.keys(campaign).forEach(key => allKeys.add(key));
        });

        // Создаем заголовки
        const detailedHeaders = Array.from(allKeys).sort();
        
        // Создаем строки данных
        const detailedRows = campaignsData.detailedCampaigns.map((campaign: Record<string, unknown>) => 
          detailedHeaders.map(key => safeStringify(campaign[key]))
        );

        const detailedSheet = XLSX.utils.aoa_to_sheet([detailedHeaders, ...detailedRows]);
        
        // Настройка ширины колонок для детального листа
        const detailedColWidths = detailedHeaders.map(() => ({ wch: 20 }));
        detailedSheet["!cols"] = detailedColWidths;
        
        XLSX.utils.book_append_sheet(workbook, detailedSheet, "Информация о компаниях");
        
        console.log(`✅ Лист "Информация о компаниях" создан с ${detailedRows.length} записями`);
      }
      
      // Лист 3: "Статистика компаний" - по методу /adv/v3/fullstats
      // ВСЕГДА создаем этот лист, даже если данных нет
      console.log("📊 Начинаем загрузку статистики компаний...");
      
      let statsHeader: string[] = ['ID кампании', 'Тип', 'Дата', 'SKU ID', 'Примечание'];
      let statsRows: (string | number)[][] = [];
      
      try {
        // Собираем ID кампаний ТОЛЬКО со статусами 7, 9, 11 (требование API WB)
        // Статусы: 7 = завершена, 9 = активна, 11 = на паузе
        const allowedStatuses = [7, 9, 11];
        let ids: number[] = [];
        
        if (campaignsData.detailedCampaigns && campaignsData.detailedCampaigns.length > 0) {
          ids = campaignsData.detailedCampaigns
            .filter((c: Record<string, unknown>) => {
              const status = Number(c.status);
              return allowedStatuses.includes(status);
            })
            .map((c: Record<string, unknown>) => Number(c.advertId))
            .filter((id: number) => Number.isFinite(id));
        } else {
          console.warn("⚠️ Нет детальной информации о кампаниях, невозможно определить статусы");
        }
        ids = Array.from(new Set(ids));
        
        console.log(`📊 Найдено ${ids.length} кампаний со статусами 7/9/11 для статистики:`, ids.slice(0, 10));

        if (ids.length > 0) {
          console.log("📊 Отправляем запрос на статистику...");
          
          const resStats = await fetch('/api/wb/fullstats', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ token, dateFrom: rkDateFrom, dateTo: rkDateTo, ids }),
          });

          console.log("📊 Ответ от fullstats:", resStats.status, resStats.ok);

          if (resStats.ok) {
            const statsData: { fields: string[]; rows: Record<string, unknown>[] } = await resStats.json();
            console.log(`✅ Получена статистика: ${statsData.rows?.length || 0} строк`);
            
            if (statsData.fields && statsData.fields.length > 0) {
              statsHeader = statsData.fields;
            }
            if (statsData.rows && statsData.rows.length > 0) {
              statsRows = statsData.rows.map((row: Record<string, unknown>) => 
                statsHeader.map((key: string) => {
                  const value = row[key];
                  // Безопасная конвертация значений для Excel
                  if (value === null || value === undefined) return '';
                  if (typeof value === 'string' || typeof value === 'number') {
                    return value;
                  }
                  if (typeof value === 'boolean') {
                    return value ? 'true' : 'false';
                  }
                  if (typeof value === 'object') {
                    // Если это объект или массив, конвертируем в JSON строку
                    try {
                      const jsonStr = JSON.stringify(value);
                      // Если JSON слишком длинный, обрезаем
                      return jsonStr.length > 32000 ? jsonStr.substring(0, 31980) + '... (обрезано)' : jsonStr;
                    } catch {
                      return String(value);
                    }
                  }
                  return String(value);
                })
              );
            } else {
              statsRows = [['Нет данных за выбранный период', '', '', '', '']];
            }
          } else {
            const err = await resStats.json().catch(() => ({} as any));
            console.error('❌ Ошибка получения статистики:', err?.error || resStats.status);
            statsRows = [[`Ошибка загрузки: ${err?.error || resStats.status}`, '', '', '', '']];
          }
        } else {
          console.warn('⚠️ Нет кампаний со статусами 7/9/11 для загрузки статистики');
          statsRows = [['Нет кампаний в статусах: Завершена (7), Активна (9), На паузе (11)', '', '', '', '']];
        }
      } catch (statsErr) {
        console.error('❌ Критическая ошибка при построении листа "Статистика компаний":', statsErr);
        statsRows = [[`Ошибка: ${(statsErr as Error).message || 'Неизвестная ошибка'}`, '', '', '', '']];
      }
      
      // Собираем уникальные SKU ID из statsRows для листов аналитики
      const uniqueSkus = new Set<string>();
      
      if (statsRows && statsRows.length > 0) {
        const skuColumnIndex = statsHeader.indexOf('SKU ID');
        
        if (skuColumnIndex !== -1) {
          statsRows.forEach((row: (string | number)[]) => {
            const skuValue = row[skuColumnIndex];
            if (skuValue && String(skuValue).trim() !== '') {
              // Если SKU ID содержит несколько артикулов через запятую, разделяем их
              const skus = String(skuValue).split(',').map(s => s.trim()).filter(Boolean);
              skus.forEach(sku => uniqueSkus.add(sku));
            }
          });
        }
      }
      
      const uniqueSkuArray = Array.from(uniqueSkus).sort((a, b) => a.localeCompare(b));
      console.log(`📊 Найдено ${uniqueSkuArray.length} уникальных артикулов для аналитики`);
      
      // Создаем лист ВСЕГДА
      const statsSheet = XLSX.utils.aoa_to_sheet([statsHeader, ...statsRows]);
      statsSheet['!cols'] = statsHeader.map(() => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(workbook, statsSheet, 'Статистика компаний');
      console.log('✅ Лист "Статистика компаний" добавлен в книгу');

      // Лист 4: "Статистика кампании с единой ставкой по кластерам фраз"
      // ВСЕГДА создаем этот лист, даже если данных нет
      console.log("📊 Начинаем загрузку статистики по кластерам фраз...");
      
      let clusterHeader: string[] = ['ID кампании', 'Примечание'];
      let clusterRows: (string | number)[][] = [];
      
      try {
        // Собираем ID кампаний с типом 8 (единая ставка)
        let campaignType8Ids: number[] = [];
        
        if (campaignsData.detailedCampaigns && campaignsData.detailedCampaigns.length > 0) {
          campaignType8Ids = campaignsData.detailedCampaigns
            .filter((c: Record<string, unknown>) => Number(c.type) === 8)
            .map((c: Record<string, unknown>) => Number(c.advertId))
            .filter((id: number) => Number.isFinite(id));
        }
        campaignType8Ids = Array.from(new Set(campaignType8Ids));
        
        console.log(`📊 Найдено ${campaignType8Ids.length} кампаний с единой ставкой (тип 8)`);

        if (campaignType8Ids.length > 0) {
          console.log("📊 Отправляем запрос на статистику по кластерам...");
          
          const resCluster = await fetch('/api/wb/stat-words', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ token, ids: campaignType8Ids }),
          });

          console.log("📊 Ответ от stat-words:", resCluster.status, resCluster.ok);

          if (resCluster.ok) {
            const clusterData: { fields: string[]; rows: Record<string, unknown>[] } = await resCluster.json();
            console.log(`✅ Получена статистика по кластерам: ${clusterData.rows?.length || 0} строк`);
            
            if (clusterData.fields && clusterData.fields.length > 0) {
              clusterHeader = clusterData.fields;
            }
            if (clusterData.rows && clusterData.rows.length > 0) {
              clusterRows = clusterData.rows.map((row: Record<string, unknown>) => 
                clusterHeader.map((key: string) => {
                  const value = row[key];
                  // Безопасная конвертация значений для Excel
                  if (value === null || value === undefined) return '';
                  if (typeof value === 'string' || typeof value === 'number') {
                    return value;
                  }
                  if (typeof value === 'boolean') {
                    return value ? 'true' : 'false';
                  }
                  if (typeof value === 'object') {
                    // Если это объект или массив, конвертируем в JSON строку
                    try {
                      const jsonStr = JSON.stringify(value);
                      // Если JSON слишком длинный, обрезаем
                      return jsonStr.length > 32000 ? jsonStr.substring(0, 31980) + '... (обрезано)' : jsonStr;
                    } catch {
                      return String(value);
                    }
                  }
                  return String(value);
                })
              );
            } else {
              clusterRows = [['Нет данных по кластерам фраз', '']];
            }
          } else {
            const err = await resCluster.json().catch(() => ({} as any));
            console.error('❌ Ошибка получения статистики по кластерам:', err?.error || resCluster.status);
            clusterRows = [[`Ошибка загрузки: ${err?.error || resCluster.status}`, '']];
          }
        } else {
          console.warn('⚠️ Нет кампаний с единой ставкой (тип 8)');
          clusterRows = [['Нет кампаний с единой ставкой (тип 8)', '']];
        }
      } catch (clusterErr) {
        console.error('❌ Критическая ошибка при построении листа кластеров:', clusterErr);
        clusterRows = [[`Ошибка: ${(clusterErr as Error).message || 'Неизвестная ошибка'}`, '']];
      }
      
      // Создаем лист ВСЕГДА
      const clusterSheet = XLSX.utils.aoa_to_sheet([clusterHeader, ...clusterRows]);
      clusterSheet['!cols'] = clusterHeader.map(() => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(workbook, clusterSheet, 'Кластеры фраз');
      console.log('✅ Лист "Кластеры фраз" добавлен в книгу');

      // Лист 5: "Статистика поисковых кластеров" - по методу /adv/v0/normquery/stats
      // ВСЕГДА создаем этот лист, даже если данных нет
      console.log("📊 Начинаем загрузку статистики поисковых кластеров...");
      
      let normQueryHeader: string[] = ['ID кампании', 'Примечание'];
      let normQueryRows: (string | number)[][] = [];
      
      try {
        // Собираем ID кампаний со статусами 7, 9, 11 (активные, завершенные, на паузе)
        // Метод /adv/v0/normquery/stats работает только для кампаний CPM
        let cpmCampaignIds: number[] = [];
        
        if (campaignsData.detailedCampaigns && campaignsData.detailedCampaigns.length > 0) {
          cpmCampaignIds = campaignsData.detailedCampaigns
            .filter((c: Record<string, unknown>) => {
              const status = Number(c.status);
              return [7, 9, 11].includes(status);
            })
            .map((c: Record<string, unknown>) => Number(c.advertId))
            .filter((id: number) => Number.isFinite(id));
        }
        cpmCampaignIds = Array.from(new Set(cpmCampaignIds));
        
        console.log(`📊 Найдено ${cpmCampaignIds.length} кампаний для статистики поисковых кластеров`);

        if (cpmCampaignIds.length > 0) {
          console.log("📊 Отправляем запрос на статистику поисковых кластеров...");
          
          const resNormQuery = await fetch('/api/wb/normquery-stats', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ token, dateFrom: rkDateFrom, dateTo: rkDateTo, ids: cpmCampaignIds }),
          });

          console.log("📊 Ответ от normquery-stats:", resNormQuery.status, resNormQuery.ok);

          if (resNormQuery.ok) {
            const normQueryData: { fields: string[]; rows: Record<string, unknown>[] } = await resNormQuery.json();
            console.log(`✅ Получена статистика поисковых кластеров: ${normQueryData.rows?.length || 0} строк`);
            
            if (normQueryData.fields && normQueryData.fields.length > 0) {
              normQueryHeader = normQueryData.fields;
            }
            if (normQueryData.rows && normQueryData.rows.length > 0) {
              normQueryRows = normQueryData.rows.map((row: Record<string, unknown>) => 
                normQueryHeader.map((key: string) => {
                  const value = row[key];
                  // Безопасная конвертация значений для Excel
                  if (value === null || value === undefined) return '';
                  if (typeof value === 'string' || typeof value === 'number') {
                    return value;
                  }
                  if (typeof value === 'boolean') {
                    return value ? 'true' : 'false';
                  }
                  if (typeof value === 'object') {
                    // Если это объект или массив, конвертируем в JSON строку
                    try {
                      const jsonStr = JSON.stringify(value);
                      // Если JSON слишком длинный, обрезаем
                      return jsonStr.length > 32000 ? jsonStr.substring(0, 31980) + '... (обрезано)' : jsonStr;
                    } catch {
                      return String(value);
                    }
                  }
                  return String(value);
                })
              );
            } else {
              normQueryRows = [['Нет данных по поисковым кластерам', '']];
            }
          } else {
            const err = await resNormQuery.json().catch(() => ({} as any));
            console.error('❌ Ошибка получения статистики поисковых кластеров:', err?.error || resNormQuery.status);
            normQueryRows = [[`Ошибка загрузки: ${err?.error || resNormQuery.status}`, '']];
          }
        } else {
          console.warn('⚠️ Нет кампаний для загрузки статистики поисковых кластеров');
          normQueryRows = [['Нет активных кампаний для статистики', '']];
        }
      } catch (normQueryErr) {
        console.error('❌ Критическая ошибка при построении листа "Статистика поисковых кластеров":', normQueryErr);
        normQueryRows = [[`Ошибка: ${(normQueryErr as Error).message || 'Неизвестная ошибка'}`, '']];
      }
      
      // Создаем лист ВСЕГДА
      const normQuerySheet = XLSX.utils.aoa_to_sheet([normQueryHeader, ...normQueryRows]);
      normQuerySheet['!cols'] = normQueryHeader.map(() => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(workbook, normQuerySheet, 'Статистика поисковых кластеров');
      console.log('✅ Лист "Статистика поисковых кластеров" добавлен в книгу');

      // Создаем листы аналитики
      console.log('📊 Формирование листов аналитики...');
      
      // Функция для создания листа аналитики
      const createAnalyticsSheet = () => {
        const analyticsData: (string | number)[][] = [
          [rkDateFrom], // A1 - дата начала
          [rkDateTo],   // A2 - дата конца
          [],           // A3 - пустая строка
          ['Артикул WB'], // A4 - заголовок
        ];
        
        // Добавляем уникальные SKU ID начиная с A5
        uniqueSkuArray.forEach(sku => {
          analyticsData.push([sku]);
        });
        
        const sheet = XLSX.utils.aoa_to_sheet(analyticsData);
        sheet['!cols'] = [{ wch: 20 }]; // Ширина столбца A
        
        return sheet;
      };
      
      const analyticsGeneralSheet = createAnalyticsSheet();
      const analyticsAutoSheet = createAnalyticsSheet();
      const analyticsManualSheet = createAnalyticsSheet();
      
      // Пересобираем workbook с правильным порядком листов
      // Создаем новый workbook и добавляем листы в нужном порядке
      const finalWorkbook = XLSX.utils.book_new();
      
      // Сначала добавляем листы аналитики
      XLSX.utils.book_append_sheet(finalWorkbook, analyticsGeneralSheet, 'Аналитика общая (ЕД+РУЧ)');
      console.log('✅ Лист "Аналитика общая (ЕДИНАЯ + РУЧНАЯ)" добавлен в книгу');
      
      XLSX.utils.book_append_sheet(finalWorkbook, analyticsAutoSheet, 'Аналитика ЕДИНАЯ');
      console.log('✅ Лист "Аналитика ЕДИНАЯ" добавлен в книгу');
      
      XLSX.utils.book_append_sheet(finalWorkbook, analyticsManualSheet, 'Аналитика РУЧНАЯ');
      console.log('✅ Лист "Аналитика РУЧНАЯ" добавлен в книгу');
      
      // Затем добавляем все остальные листы из старого workbook
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        XLSX.utils.book_append_sheet(finalWorkbook, sheet, sheetName);
      });
      
      // Генерация и скачивание файла
      const arrayBuffer = XLSX.write(finalWorkbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([arrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `РК_${rkDateFrom}_${rkDateTo}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
      
      console.log("✅ Файл РК успешно скачан");
    } catch (error) {
      console.error("❌ Ошибка при выгрузке РК:", error);
      
      let userFriendlyMessage = "Произошла ошибка при загрузке данных РК";
      
      if (error instanceof Error) {
        if (error.message.includes("Failed to fetch") || error.message.includes("NetworkError")) {
          userFriendlyMessage = "Ошибка сети. Проверьте подключение к интернету";
        } else if (error.message.includes("401") || error.message.includes("авторизац")) {
          userFriendlyMessage = "Ошибка авторизации. Проверьте правильность токена";
        } else if (error.message.includes("429")) {
          userFriendlyMessage = "Превышен лимит запросов. Попробуйте через минуту";
        } else if (error.message.includes("500")) {
          userFriendlyMessage = "Внутренняя ошибка сервера Wildberries. Попробуйте позже";
        } else {
          userFriendlyMessage = error.message;
        }
      }
      
      alert(userFriendlyMessage);
    } finally {
      setIsLoadingRk(false);
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

            {/* Секция РНП */}
            <div className="pt-4 border-t border-black/[.08] dark:border-white/[.145]">
              <div className="flex flex-col gap-3">
                <h3 className="text-lg font-semibold">РНП (Полный отчет за день)</h3>
                <div className="text-sm text-gray-600 dark:text-gray-400 mb-2">
                  Включает: Аналитика по товарам, РНП, Платное хранение, Платная приемка, Финансы РК, Остатки, Номенклатура
                </div>
                
                <div className="flex flex-col gap-2">
                  <span className="text-sm font-medium">Дата для РНП:</span>
                  <div>
                    <label htmlFor="rnp-date" className="text-xs text-black/60 dark:text-white/70 block mb-1">
                      Выберите дату
                    </label>
                    <input
                      id="rnp-date"
                      type="date"
                      value={rnpDate}
                      onChange={(e) => setRnpDate(e.target.value)}
                      className="w-full h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6]"
                    />
                  </div>
                  <div className="text-xs text-gray-500 dark:text-gray-400">
                    Полный отчет с ежедневными данными реализации и всеми дополнительными листами за выбранный день
                  </div>
                </div>

                <button
                  type="button"
                  onClick={handleRnpDownload}
                  disabled={isLoadingRnp}
                  className={`w-full h-11 rounded-lg bg-green-600 text-white dark:bg-green-500 dark:text-white font-medium transition-opacity ${
                    isLoadingRnp ? "opacity-60 cursor-not-allowed" : "hover:opacity-90"
                  }`}
                >
                  {isLoadingRnp ? (
                    <span className="inline-flex items-center gap-2">
                      <span className="inline-block h-4 w-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
                      Загрузка полного отчета...
                    </span>
                  ) : (
                    "Скачать полный РНП"
                  )}
                </button>
              </div>
            </div>

            {/* Секция Смарт поставка */}
            <div className="pt-4 border-t border-black/[.08] dark:border-white/[.145]">
              <div className="flex flex-col gap-3">
                <h3 className="text-lg font-semibold">Смарт поставка</h3>
                <div className="text-sm text-gray-600 dark:text-gray-400 mb-2">
                  Распределение по регионам
                </div>

                {/* Параметры для расчета остатков */}
                <div className="grid grid-cols-3 gap-3 mt-2">
                  <div>
                    <label htmlFor="delivery-days" className="text-xs text-black/60 dark:text-white/70 block mb-1">
                      Срок поставки (дн.):
                    </label>
                    <input
                      id="delivery-days"
                      type="text"
                      inputMode="decimal"
                      value={deliveryDays}
                      onChange={(e) => {
                        const value = e.target.value.replace(',', '.');
                        if (value === '' || /^\d*\.?\d*$/.test(value)) {
                          setDeliveryDays(value);
                        }
                      }}
                      placeholder="0"
                      className="w-full h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6] text-center"
                    />
                  </div>
                  <div>
                    <label htmlFor="stock-days" className="text-xs text-black/60 dark:text-white/70 block mb-1">
                      Запас (дн.):
                    </label>
                    <input
                      id="stock-days"
                      type="text"
                      inputMode="decimal"
                      value={stockDays}
                      onChange={(e) => {
                        const value = e.target.value.replace(',', '.');
                        if (value === '' || /^\d*\.?\d*$/.test(value)) {
                          setStockDays(value);
                        }
                      }}
                      placeholder="0"
                      className="w-full h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6] text-center"
                    />
                  </div>
                  <div>
                    <label htmlFor="coefficient" className="text-xs text-black/60 dark:text-white/70 block mb-1">
                      Коэффициент:
                    </label>
                    <input
                      id="coefficient"
                      type="text"
                      inputMode="decimal"
                      value={coefficient}
                      onChange={(e) => {
                        const value = e.target.value.replace(',', '.');
                        if (value === '' || /^\d*\.?\d*$/.test(value)) {
                          setCoefficient(value);
                        }
                      }}
                      placeholder="0"
                      className="w-full h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6] text-center"
                    />
                  </div>
                </div>

                <button
                  type="button"
                  onClick={handleSupplyAnalysisDownload}
                  disabled={isLoadingAnalysis}
                  className={`w-full h-11 rounded-lg bg-emerald-600 text-white dark:bg-emerald-500 dark:text-white font-medium transition-opacity ${
                    isLoadingAnalysis ? "opacity-60 cursor-not-allowed" : "hover:opacity-90"
                  }`}
                >
                  {isLoadingAnalysis ? (
                    <span className="inline-flex items-center gap-2">
                      <span className="inline-block h-4 w-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
                      Загрузка анализа...
                    </span>
                  ) : (
                    "Анализ поставок"
                  )}
                </button>
              </div>
            </div>

            {/* Секция Выгрузка РК */}
            <div className="pt-4 border-t border-black/[.08] dark:border-white/[.145]">
              <div className="flex flex-col gap-3">
                <h3 className="text-lg font-semibold">Выгрузка РК</h3>
                <div className="text-sm text-gray-600 dark:text-gray-400 mb-2">
                  Отчет по рекламным кампаниям
                </div>
                
                <div className="flex flex-col gap-2">
                  <span className="text-sm font-medium">Период для выгрузки РК:</span>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label htmlFor="rk-date-from" className="text-xs text-black/60 dark:text-white/70 block mb-1">
                        От даты
                      </label>
                      <input
                        id="rk-date-from"
                        type="date"
                        value={rkDateFrom}
                        onChange={(e) => setRkDateFrom(e.target.value)}
                        className="w-full h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6]"
                      />
                    </div>
                    <div>
                      <label htmlFor="rk-date-to" className="text-xs text-black/60 dark:text-white/70 block mb-1">
                        До даты
                      </label>
                      <input
                        id="rk-date-to"
                        type="date"
                        value={rkDateTo}
                        onChange={(e) => setRkDateTo(e.target.value)}
                        className="w-full h-11 rounded-lg border border-black/[.12] dark:border-white/[.18] bg-transparent px-3 outline-none focus:ring-2 focus:ring-[#3b82f6]"
                      />
                    </div>
                  </div>
                  <div className="text-xs text-gray-500 dark:text-gray-400">
                    Выгрузка данных рекламных кампаний за выбранный период
                  </div>
                </div>

                <button
                  type="button"
                  onClick={handleRkDownload}
                  disabled={isLoadingRk}
                  className={`w-full h-11 rounded-lg bg-purple-600 text-white dark:bg-purple-500 dark:text-white font-medium transition-opacity ${
                    isLoadingRk ? "opacity-60 cursor-not-allowed" : "hover:opacity-90"
                  }`}
                >
                  {isLoadingRk ? (
                    <span className="inline-flex items-center gap-2">
                      <span className="inline-block h-4 w-4 rounded-full border-2 border-white border-t-transparent animate-spin" />
                      Загрузка...
                    </span>
                  ) : (
                    "Выгрузка РК"
                  )}
                </button>
              </div>
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
