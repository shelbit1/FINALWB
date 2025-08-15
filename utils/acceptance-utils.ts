// Интерфейс для данных платной приемки (адаптировано под архитектуру проекта)
export interface AcceptanceData {
  count: number;
  giCreateDate: string;
  incomeId: number;
  nmID: number;
  shkCreateDate: string;
  subjectName: string;
  total: number;
}

// Интерфейс для ответа API создания задачи
interface CreateTaskResponse {
  data: {
    taskId: string;
  };
}

// Интерфейс для ответа API статуса задачи
interface TaskStatusResponse {
  data: {
    status: 'pending' | 'done' | 'error';
  };
}

/**
 * Основная функция получения данных платной приемки
 * Использует асинхронный API Wildberries с созданием задачи и ожиданием результата
 */
export async function fetchAcceptanceData(apiKey: string, dateFrom: string, dateTo: string): Promise<AcceptanceData[]> {
  try {
    console.log(`🚀 Начало получения данных платной приемки: ${dateFrom} - ${dateTo}`);
    
    // Шаг 1: Создаем задачу на генерацию отчета
    const taskId = await createAcceptanceTask(apiKey, dateFrom, dateTo);
    console.log(`📋 Задача создана с ID: ${taskId}`);
    
    // Шаг 2: Ожидаем готовности отчета
    await waitForTaskCompletion(apiKey, taskId);
    console.log(`✅ Задача ${taskId} завершена, скачиваем данные`);
    
    // Шаг 3: Скачиваем готовый отчет
    const acceptanceData = await downloadAcceptanceReport(apiKey, taskId);
    console.log(`📦 Получено записей платной приемки: ${acceptanceData.length}`);
    
    return acceptanceData;
    
  } catch (error) {
    console.error('❌ Ошибка при получении данных платной приемки:', error);
    throw error;
  }
}

/**
 * Создает задачу на генерацию отчета платной приемки
 */
async function createAcceptanceTask(apiKey: string, dateFrom: string, dateTo: string): Promise<string> {
  const createUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report?dateFrom=${dateFrom}&dateTo=${dateTo}`;
  
  console.log(`📋 Создание задачи на генерацию отчета...`);
  console.log(`🌐 URL: ${createUrl}`);
  
  const response = await fetch(createUrl, {
    method: 'GET',
    headers: {
      'Authorization': apiKey,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error(`❌ Ошибка создания задачи: ${response.status}`, errorText);
    
    if (response.status === 401) {
      throw new Error('Неверный или недействительный API токен');
    } else if (response.status === 400) {
      throw new Error('Некорректные параметры запроса. Проверьте период дат.');
    }
    
    throw new Error(`Ошибка создания задачи: ${response.status} ${response.statusText}`);
  }

  const responseData: CreateTaskResponse = await response.json();
  const taskId = responseData.data?.taskId;
  
  if (!taskId) {
    throw new Error('Не получен ID задачи из ответа API');
  }

  return taskId;
}

/**
 * Ожидает завершения выполнения задачи
 */
async function waitForTaskCompletion(apiKey: string, taskId: string): Promise<void> {
  console.log('⏳ Ожидание готовности отчета...');
  
  let attempts = 0;
  const maxAttempts = 30; // 30 попыток по 5 секунд = 2.5 минуты максимум
  let isReady = false;

  while (!isReady && attempts < maxAttempts) {
    attempts++;
    console.log(`🔄 Проверка статуса, попытка ${attempts}/${maxAttempts}...`);
    
    const statusUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/${taskId}/status`;
    
    try {
      const statusResponse = await fetch(statusUrl, {
        method: 'GET',
        headers: {
          'Authorization': apiKey,
          'Content-Type': 'application/json'
        }
      });

      if (statusResponse.ok) {
        const statusData: TaskStatusResponse = await statusResponse.json();
        const status = statusData.data?.status;
        console.log(`📊 Статус задачи: ${status}`);
        
        if (status === 'done') {
          isReady = true;
          break;
        } else if (status === 'error') {
          throw new Error('Ошибка при генерации отчета на сервере Wildberries');
        }
        // Если статус 'pending', продолжаем ожидание
      } else {
        console.warn(`⚠️ Ошибка проверки статуса: ${statusResponse.status}`);
      }
    } catch (error) {
      console.warn(`⚠️ Ошибка при проверке статуса (попытка ${attempts}):`, error);
    }

    if (!isReady && attempts < maxAttempts) {
      console.log('⏳ Ожидание 5 секунд перед следующей проверкой...');
      await new Promise(resolve => setTimeout(resolve, 5000)); // Ждем 5 секунд
    }
  }

  if (!isReady) {
    throw new Error('Превышено время ожидания генерации отчета (2.5 минуты). Попробуйте позже.');
  }
}

/**
 * Скачивает готовый отчет платной приемки
 */
async function downloadAcceptanceReport(apiKey: string, taskId: string): Promise<AcceptanceData[]> {
  console.log('📥 Скачивание готового отчета...');
  
  const downloadUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/${taskId}/download`;
  
  const downloadResponse = await fetch(downloadUrl, {
    method: 'GET',
    headers: {
      'Authorization': apiKey,
      'Content-Type': 'application/json'
    }
  });

  if (!downloadResponse.ok) {
    const errorText = await downloadResponse.text();
    console.error(`❌ Ошибка скачивания отчета: ${downloadResponse.status}`, errorText);
    throw new Error(`Ошибка скачивания отчета: ${downloadResponse.status} ${downloadResponse.statusText}`);
  }

  const acceptanceData: AcceptanceData[] = await downloadResponse.json();
  
  // Валидация данных
  if (!Array.isArray(acceptanceData)) {
    throw new Error('Некорректный формат данных в ответе API');
  }

  // Фильтруем и валидируем записи
  const validData = acceptanceData.filter((item: any) => {
    return (
      typeof item === 'object' &&
      item !== null &&
      typeof item.count === 'number' &&
      typeof item.nmID === 'number' &&
      typeof item.total === 'number'
    );
  });

  if (validData.length !== acceptanceData.length) {
    console.warn(`⚠️ Отфильтровано ${acceptanceData.length - validData.length} некорректных записей`);
  }

  return validData;
}

/**
 * Преобразует данные платной приемки в формат для Excel
 */
export function transformAcceptanceDataToExcel(acceptanceData: AcceptanceData[], startDate: string, endDate: string) {
  const reportPeriod = `${new Date(startDate).toLocaleDateString('ru-RU')} - ${new Date(endDate).toLocaleDateString('ru-RU')}`;
  
  return acceptanceData.map((item) => ({
    "Кол-во": item.count || 0,
    "Дата создания GI": item.giCreateDate ? new Date(item.giCreateDate).toLocaleDateString('ru-RU') : '',
    "Income ID": item.incomeId || '',
    "Артикул WB": item.nmID || '',
    "Дата создания ШК": item.shkCreateDate ? new Date(item.shkCreateDate).toLocaleDateString('ru-RU') : '',
    "Предмет": item.subjectName || '',
    "Сумма (руб)": item.total || 0,
    "Дата отчета": reportPeriod
  }));
}

/**
 * Функция для валидации периода дат
 */
export function validateDateRange(startDate: string, endDate: string): { isValid: boolean; error?: string } {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const now = new Date();
  
  // Проверяем, что даты корректные
  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    return { isValid: false, error: 'Некорректный формат дат' };
  }
  
  // Проверяем, что начальная дата не больше конечной
  if (start > end) {
    return { isValid: false, error: 'Дата начала не может быть больше даты окончания' };
  }
  
  // Проверяем, что период не больше 31 дня
  const diffTime = Math.abs(end.getTime() - start.getTime());
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  if (diffDays > 31) {
    return { isValid: false, error: 'Период не может превышать 31 день' };
  }
  
  // Проверяем, что даты не в будущем
  if (start > now || end > now) {
    return { isValid: false, error: 'Даты не могут быть в будущем' };
  }
  
  return { isValid: true };
}

/**
 * Получение порядка полей для вывода в Excel (в соответствии с архитектурой проекта)
 */
export const ACCEPTANCE_FIELD_ORDER: string[] = [
  "Кол-во",
  "Дата создания GI", 
  "Income ID",
  "Артикул WB",
  "Дата создания ШК",
  "Предмет",
  "Сумма (руб)",
  "Дата отчета"
];

