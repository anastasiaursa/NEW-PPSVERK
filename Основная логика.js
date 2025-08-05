/*
 * 🎯 2 ОСНОВНАЯ ЛОГИКА
 * 
 * ✅ Что делает: Основные функции создания сводных таблиц
 * ✅ Главные функции: создатьВсеСводные(), создатьПоследовательностьМесяцев()
 * ✅ Получение данных из Google Sheets и создание сводных
 * 
 * Зависит от: 1 КОНФИГУРАЦИЯ И НАСТРОЙКИ.js, 3 ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ.js
 * 
 * ВАЖНО: ВСЕ ФУНКЦИИ СКОПИРОВАНЫ ИЗ ОРИГИНАЛЬНОГО КОДА БЕЗ ИЗМЕНЕНИЙ!
 * Применены только исправления пользователя
 */

console.log('🎯 ОСНОВНАЯ ЛОГИКА ЗАГРУЖЕНА (ИДЕНТИЧНО ПРОШЛОМУ КОДУ + ИСПРАВЛЕНИЯ)');
console.log('🚀 Главная функция: создатьВсеСводные()');
console.log('📅 Автоматическая последовательность: создатьПоследовательностьМесяцев()');

// ========================================================================
// 🚀 ГЛАВНЫЕ ФУНКЦИИ
// ========================================================================

/**
 * Создать все сводные таблицы для текущего месяца
 */
function создатьВсеСводные() {
  console.log('🎯 СОЗДАНИЕ ВСЕХ СВОДНЫХ (ИДЕНТИЧНО ПРОШЛОМУ КОДУ + ИСПРАВЛЕНИЯ)');
  console.log(`📅 Период: ${МЕСЯЦЫ.БАЗОВЫЙ} → ${МЕСЯЦЫ.ТЕКУЩИЙ}`);
  console.log();
  
  try {
    const monthName = МЕСЯЦЫ.ТЕКУЩИЙ;
    
    // ОТЛАДКА: проверяем что месяц установлен правильно
    console.log(`🔍 ОТЛАДКА: МЕСЯЦЫ.ТЕКУЩИЙ = "${МЕСЯЦЫ.ТЕКУЩИЙ}"`);
    console.log(`🔍 ОТЛАДКА: monthName = "${monthName}"`);
    
    if (!monthName) {
      throw new Error(`КРИТИЧЕСКАЯ ОШИБКА: МЕСЯЦЫ.ТЕКУЩИЙ не установлен! Значение: ${monthName}`);
    }
    
    console.log('1️⃣ Создание сводных таблиц...');
    
    // Создаем все три типа сводных
    createSummaryReport(monthName);      // Бренды
    createGeoSummaryReport(monthName);   // ГЕО
    
    // Для сеошников - ВЫБОР: полные данные или скорость
    console.log('🔧 Создание сводной по сеошникам...');
    
    // ВАРИАНТ 1: БЫСТРО, но упрощенные данные
    // createSeoSummaryReportSafe(monthName);
    
    // ВАРИАНТ 2: ПОЛНЫЕ ДАННЫЕ, но может быть медленно для больших данных  
    createSeoSummaryReport(monthName);  // Теперь с защитой от undefined!
    
    console.log('2️⃣ Расчет показателей роста...');
    
    // Рассчитываем рост между периодами
    if (МЕСЯЦЫ.БАЗОВЫЙ !== МЕСЯЦЫ.ТЕКУЩИЙ) {
      calculateGrowthBetweenPeriods(МЕСЯЦЫ.БАЗОВЫЙ, МЕСЯЦЫ.ТЕКУЩИЙ);
    }
    
    console.log('3️⃣ Применение форматирования...');
    
    // Применяем smartReplaceColors ко всем листам
    applySmartFormattingToAllSheets(monthName);
    
    console.log('✅ ВСЕ СВОДНЫЕ СОЗДАНЫ УСПЕШНО!');
    
  } catch (error) {
    console.error('❌ Ошибка создания сводных:', error);
    throw error;
  }
}

/**
 * Создать последовательность сводных для нескольких месяцев
 */
function создатьПоследовательностьМесяцев() {
  console.log('📅 ПОСЛЕДОВАТЕЛЬНОЕ СОЗДАНИЕ СВОДНЫХ');
  console.log('🎯 Месяцы: Май → Июнь → Июль 2025');
  console.log();
  
  const месяцы = ['Май 2025', 'Июнь 2025', 'Июль 2025'];
  
  for (let i = 0; i < месяцы.length; i++) {
    const текущийМесяц = месяцы[i];
    const предыдущийМесяц = i > 0 ? месяцы[i - 1] : null;
    
    console.log(`📊 Обрабатываем: ${текущийМесяц}`);
    
    // Устанавливаем период
    МЕСЯЦЫ.ТЕКУЩИЙ = текущийМесяц;
    МЕСЯЦЫ.БАЗОВЫЙ = предыдущийМесяц || текущийМесяц;
    
    // Создаем сводные для этого месяца
    создатьВсеСводные();
    
    console.log(`✅ ${текущийМесяц} готов!`);
    console.log();
  }
  
  console.log('🎉 ВСЕ МЕСЯЦЫ ОБРАБОТАНЫ!');
}

// ========================================================================
// 📊 СОЗДАНИЕ СВОДНЫХ ТАБЛИЦ (КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
// ========================================================================

/**
 * Создать сводную по брендам (КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function createSummaryReport(monthName) {
  try {
    // КРИТИЧЕСКАЯ ПРОВЕРКА: исправляем undefined monthName
    if (!monthName || monthName === 'undefined') {
      console.log(`❌ ИСПРАВЛЯЕМ: monthName = "${monthName}" → "${МЕСЯЦЫ.ТЕКУЩИЙ}"`);
      monthName = МЕСЯЦЫ.ТЕКУЩИЙ;
      if (!monthName) {
        throw new Error(`Критическая ошибка: и monthName и МЕСЯЦЫ.ТЕКУЩИЙ равны undefined!`);
      }
    }
    
    console.log(`Начинаем создание сводной таблицы по брендам для ${monthName}`);
    
    // ID таблиц из конфигурации
    const MAIN_TABLE_ID = ТАБЛИЦЫ.ОСНОВНАЯ;
    const SITES_TABLE_ID = ТАБЛИЦЫ.САЙТЫ;
    const CURRENT_TABLE_ID = ТАБЛИЦЫ.РЕЗУЛЬТАТ;
    
    // Получаем данные из основной таблицы
    console.log('Получаем данные из основной таблицы...');
    const mainData = getMainTableData(MAIN_TABLE_ID, monthName);
    
    // Получаем данные о сайтах
    console.log('Получаем данные о сайтах...');
    const sitesData = getSitesTableData(SITES_TABLE_ID);
    
    // Группируем данные
    console.log('Группируем данные...');
    const groupedData = groupData(mainData, sitesData);
    
    // Создаем сводную таблицу
    console.log('Создаем сводную таблицу...');
    createSummarySheet(CURRENT_TABLE_ID, groupedData, monthName);
    
    console.log(`Сводная таблица по брендам "${monthName}" создана успешно!`);
    
  } catch (error) {
    console.error('Ошибка при создании сводной таблицы:', error);
    throw error;
  }
}

/**
 * Создать сводную по ГЕО (КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function createGeoSummaryReport(monthName) {
  try {
    // КРИТИЧЕСКАЯ ПРОВЕРКА: исправляем undefined monthName
    if (!monthName || monthName === 'undefined') {
      console.log(`❌ ИСПРАВЛЯЕМ: monthName = "${monthName}" → "${МЕСЯЦЫ.ТЕКУЩИЙ}"`);
      monthName = МЕСЯЦЫ.ТЕКУЩИЙ;
      if (!monthName) {
        throw new Error(`Критическая ошибка: и monthName и МЕСЯЦЫ.ТЕКУЩИЙ равны undefined!`);
      }
    }
    
    console.log(`Начинаем создание сводной таблицы по ГЕО для ${monthName}`);
    
    // ID таблиц из конфигурации
    const MAIN_TABLE_ID = ТАБЛИЦЫ.ОСНОВНАЯ;
    const SITES_TABLE_ID = ТАБЛИЦЫ.САЙТЫ;
    const CURRENT_TABLE_ID = ТАБЛИЦЫ.РЕЗУЛЬТАТ;
    
    // Получаем данные из основной таблицы
    console.log('Получаем данные из основной таблицы...');
    const mainData = getMainTableDataForGeo(MAIN_TABLE_ID, monthName);
    
    // Получаем данные о сайтах
    console.log('Получаем данные о сайтах...');
    const sitesData = getSitesTableData(SITES_TABLE_ID);
    
    // Группируем данные по ГЕО
    console.log('Группируем данные по ГЕО...');
    const groupedData = groupGeoData(mainData, sitesData);
    
    // Создаем сводную таблицу по ГЕО
    console.log('Создаем сводную таблицу по ГЕО...');
    createGeoSummarySheet(CURRENT_TABLE_ID, groupedData, monthName);
    
    console.log(`Сводная таблица по ГЕО "${monthName}" создана успешно!`);
    
  } catch (error) {
    console.error('Ошибка при создании сводной таблицы по ГЕО:', error);
    throw error;
  }
}

/**
 * Создать сводную по сеошникам (КОПИЯ ИЗ По сеошникам.js)
 */
/**
 * Безопасная версия создания сводной по сеошникам (с защитой от undefined)
 */
function createSeoSummaryReportSafe(monthName) {
  try {
    // КРИТИЧЕСКАЯ ПРОВЕРКА: если monthName undefined, используем МЕСЯЦЫ.ТЕКУЩИЙ
    if (!monthName || monthName === 'undefined') {
      console.log(`⚠️ Исправляем monthName: ${monthName} → ${МЕСЯЦЫ.ТЕКУЩИЙ}`);
      monthName = МЕСЯЦЫ.ТЕКУЩИЙ;
      
      if (!monthName) {
        throw new Error(`Критическая ошибка: и monthName и МЕСЯЦЫ.ТЕКУЩИЙ равны undefined!`);
      }
    }
    
    console.log(`Начинаем создание сводной таблицы по сеошникам для ${monthName}`);
    
    // ID таблиц из конфигурации
    const MAIN_TABLE_ID = ТАБЛИЦЫ.ОСНОВНАЯ;
    const SITES_TABLE_ID = ТАБЛИЦЫ.САЙТЫ;
    const CURRENT_TABLE_ID = ТАБЛИЦЫ.РЕЗУЛЬТАТ;
    
    // Получаем данные из основной таблицы
    console.log('Получаем данные из основной таблицы...');
    const mainData = getMainTableDataSafe(MAIN_TABLE_ID, monthName);
    
    // Получаем данные о сайтах
    console.log('Получаем данные о сайтах...');
    const sitesData = getSitesTableData(SITES_TABLE_ID);
    
    // Группируем данные по сеошникам
    console.log('Группируем данные по сеошникам...');
    const groupedData = groupSeoData(mainData, sitesData);
    
    // Создаем сводную таблицу по сеошникам (БЫСТРАЯ ВЕРСИЯ)
    console.log('Создаем сводную таблицу по сеошникам (быстрая версия)...');
    createSeoSummarySimple(monthName, groupedData);
    
    console.log(`✅ Сводная таблица по сеошникам "${monthName}" создана успешно!`);
    
  } catch (error) {
    console.error('Ошибка при создании сводной таблицы по сеошникам:', error);
    throw error;
  }
}

function createSeoSummaryReport(monthName) {
  try {
    // КРИТИЧЕСКАЯ ЗАЩИТА: если monthName undefined, используем МЕСЯЦЫ.ТЕКУЩИЙ
    if (!monthName || monthName === 'undefined') {
      console.log(`❌ ИСПРАВЛЯЕМ: monthName = "${monthName}" → "${МЕСЯЦЫ.ТЕКУЩИЙ}"`);
      monthName = МЕСЯЦЫ.ТЕКУЩИЙ;
      
      if (!monthName) {
        throw new Error(`Критическая ошибка: и monthName и МЕСЯЦЫ.ТЕКУЩИЙ равны undefined!`);
      }
    }
    
    console.log(`Начинаем создание сводной таблицы по сеошникам для ${monthName}`);
    
    // ID таблиц из конфигурации
    const MAIN_TABLE_ID = ТАБЛИЦЫ.ОСНОВНАЯ;
    const SITES_TABLE_ID = ТАБЛИЦЫ.САЙТЫ;
    const CURRENT_TABLE_ID = ТАБЛИЦЫ.РЕЗУЛЬТАТ;
    
    // Получаем данные из основной таблицы
    console.log('Получаем данные из основной таблицы...');
    const mainData = getMainTableData(MAIN_TABLE_ID, monthName);
    
    // Получаем данные о сайтах
    console.log('Получаем данные о сайтах...');
    const sitesData = getSitesTableData(SITES_TABLE_ID);
    
    // Группируем данные по сеошникам
    console.log('Группируем данные по сеошникам...');
    const groupedData = groupSeoData(mainData, sitesData);
    
    // Создаем сводную таблицу по сеошникам
    console.log('Создаем сводную таблицу по сеошникам...');
    createSeoSummarySheet(CURRENT_TABLE_ID, groupedData, monthName);
    
    console.log(`Сводная таблица по сеошникам "${monthName}" создана успешно!`);
    
  } catch (error) {
    console.error('Ошибка при создании сводной таблицы по сеошникам:', error);
    throw error;
  }
}

// ========================================================================
// 📥 ПОЛУЧЕНИЕ ДАННЫХ (КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
// ========================================================================

/**
 * Получить данные из основной таблицы (КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
/**
 * Быстрая версия получения данных (без отладочных логов)
 */
function getMainTableDataSafe(tableId, monthName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheet = spreadsheet.getSheetByName(monthName);
    
    if (!sheet) {
      throw new Error(`Лист "${monthName}" не найден`);
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      throw new Error(`Лист "${monthName}" не содержит данных`);
    }
    
    const headers = values[0];
    const dataRows = values.slice(1);
    
    return dataRows.map(row => {
      const item = {};
      headers.forEach((header, index) => {
        if (header) item[header] = row[index];
      });
      return item;
    });
    
  } catch (error) {
    console.error(`Ошибка при получении данных из основной таблицы:`, error);
    throw error;
  }
}

function getMainTableData(tableId, monthName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheet = spreadsheet.getSheetByName(monthName);
    
    if (!sheet) {
      throw new Error(`Лист "${monthName}" не найден`);
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      throw new Error('Недостаточно данных в таблице');
    }
    
    const data = [];
    
    // Начинаем с 2 строки (индекс 1), так как первая - заголовки
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Пропускаем пустые строки
      if (!row[0] && !row[16]) continue; // Проверяем столбец A (Сеошник) и Q (Идентификатор)
      
      data.push({
        seoSpecialist: row[0] || '', // A - Сеошник
        clicks: parseFloat(row[4]) || 0, // E - Клики
        registrations: parseFloat(row[5]) || 0, // F - Реги
        fd: parseFloat(row[6]) || 0, // G - FD
        rd: parseFloat(row[7]) || 0, // H - RD
        deposits: parseFloat(row[8]) || 0, // I - Deps
        revenueUSD: parseFloat(row[13]) || 0, // N - Выручка / USD
        identifier: row[16] || '', // Q - Идентификатор
        brand: row[21] || '', // V - Бренд
        geo: row[22] || '', // W - ГЕО
        brandGeo: row[23] || '', // X - Название бренд+гео
        stream: row[24] || '' // Y - Поток
      });
    }
    
    return data;
    
  } catch (error) {
    console.error('Ошибка при получении данных из основной таблицы:', error);
    throw error;
  }
}

/**
 * Получить данные из основной таблицы для ГЕО (с унификацией названий стран)
 */
function getMainTableDataForGeo(tableId, monthName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheet = spreadsheet.getSheetByName(monthName);
    
    if (!sheet) {
      throw new Error(`Лист "${monthName}" не найден`);
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      throw new Error('Недостаточно данных в таблице');
    }
    
    const data = [];
    
    // Начинаем с 2 строки (индекс 1), так как первая - заголовки
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Пропускаем пустые строки
      if (!row[0] && !row[16]) continue;
      
      const geoCode = row[22] || ''; // W - ГЕО
      const geoDisplayName = получитьСтрану(geoCode); // ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: флаги стран
      
      data.push({
        seoSpecialist: row[0] || '',
        clicks: parseFloat(row[4]) || 0,
        registrations: parseFloat(row[5]) || 0,
        fd: parseFloat(row[6]) || 0,
        rd: parseFloat(row[7]) || 0,
        deposits: parseFloat(row[8]) || 0,
        revenueUSD: parseFloat(row[13]) || 0,
        identifier: row[16] || '',
        brand: row[21] || '',
        geo: geoCode,
        geoDisplayName: geoDisplayName, // Добавляем отображаемое название с флагом
        brandGeo: row[23] || '',
        stream: row[24] || ''
      });
    }
    
    return data;
    
  } catch (error) {
    console.error('Ошибка при получении данных из основной таблицы для ГЕО:', error);
    throw error;
  }
}

/**
 * Получить данные о сайтах (КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
 */
function getSitesTableData(tableId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheets = spreadsheet.getSheets();
    
    const sitesData = {};
    
    // Маппинг названий листов к именам сеошников
    const seoMapping = {
      'Михаил': 'Михаил',
      'Анастасия': 'Анастасия',
      'Неизвестно': 'Неизвестно',
      'Alex SE': 'Alex SE',
      'Виктор': 'Виктор',
      'Эмиль': 'Эмиль',
      'Кирилл': 'Кирилл',
      'Александр': 'Александр',
      'Александр Артюх': 'Александр Артюх',
      'Погрешность трекера': 'Погрешность трекера',
      'Майда': 'Майда',
      'Евгения': 'Евгения',
      'Редирект с цмс': 'Редирект с цмс',
      'Dragon': 'Dragon',
      'Нет трафика': 'Нет трафика'
    };
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      let seoName = seoMapping[sheetName] || 'Михаил';
      
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      
      if (values.length < 2) return;
      
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        
        if (!row[0] && !row[12]) continue;
        
        const identifier = row[12] || '';
        const site = row[0] || '';
        const clicks = parseFloat(row[4]) || 0;
        
        if (identifier) {
          if (!sitesData[identifier]) {
            sitesData[identifier] = [];
          }
          
          sitesData[identifier].push({
            site: site,
            clicks: clicks,
            seoSpecialist: seoName,
            isExcluded: ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoName)
          });
        }
      }
    });
    
    return sitesData;
    
  } catch (error) {
    console.error('Ошибка при получении данных о сайтах:', error);
    throw error;
  }
}

// ========================================================================
// 📈 РАСЧЕТ РОСТА (ИСПРАВЛЕНИЯ ПОЛЬЗОВАТЕЛЯ)
// ========================================================================

/**
 * 📈 ОТДЕЛЬНАЯ ФУНКЦИЯ РАСЧЕТА РОСТА МЕЖДУ МЕСЯЦАМИ
 * 
 * Берет 2 месяца из констант и проводит полный расчет роста:
 * - По строкам для сводных данных (агрегированные кластеры)
 * - По столбцам для детализированных показателей
 * - С правильным цветовым форматированием
 */
function рассчитатьРостМеждуМесяцами() {
  const baseMonth = МЕСЯЦЫ.БАЗОВЫЙ;
  const currentMonth = МЕСЯЦЫ.ТЕКУЩИЙ;
  
  console.log(`📈 ОТДЕЛЬНАЯ ФУНКЦИЯ: Расчет роста ${baseMonth} → ${currentMonth}`);
  
  // Проверяем что месяцы разные
  if (baseMonth === currentMonth) {
    console.log('⚠️ Месяцы одинаковые, расчет роста пропускается');
    return;
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    
    // Листы для обработки
    const sheets = [
      `Сводная Бренд+Гео ${currentMonth}`,
      `Сводная ГЕО ${currentMonth}`,
      `Сводная Сеошники ${currentMonth}`
    ];
    
    sheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) return;
      
      console.log(`  📊 Обрабатываем лист: ${sheetName}`);
      
      // Получаем базовые данные для сравнения
      const baseSheetName = sheetName.replace(currentMonth, baseMonth);
      const baseSheet = spreadsheet.getSheetByName(baseSheetName);
      
      if (!baseSheet) {
        console.log(`  ⚠️ Базовый лист ${baseSheetName} не найден, пропускаем расчет роста`);
        return;
      }
      
      const currentData = sheet.getDataRange().getValues();
      const baseData = baseSheet.getDataRange().getValues();
      
      // Сначала обрабатываем данные (БЕЗ цветового форматирования)
      processGrowthData(sheet, currentData, baseData);
      
      // НОВОЕ: Заполняем строки роста по кластерам
      fillGrowthRowsForClusters(sheet, currentData, baseData);
      
      // Затем применяем цветовое форматирование к ячейкам роста
      applyGrowthColorFormatting(sheet, currentData);
      
      console.log(`  ✅ Рост рассчитан и отформатирован: ${sheetName}`);
    });
    
    console.log('🎉 РАСЧЕТ РОСТА МЕЖДУ МЕСЯЦАМИ ЗАВЕРШЕН!');
    
  } catch (error) {
    console.error('❌ Ошибка расчета роста между месяцами:', error);
    throw error;
  }
}

/**
 * СТАРАЯ ФУНКЦИЯ - для обратной совместимости
 */
function calculateGrowthBetweenPeriods(baseMonth, currentMonth) {
  // Устанавливаем месяцы в константы
  МЕСЯЦЫ.БАЗОВЫЙ = baseMonth;
  МЕСЯЦЫ.ТЕКУЩИЙ = currentMonth;
  
  // Вызываем новую функцию
  рассчитатьРостМеждуМесяцами();
}

/**
 * Обработка данных роста (математические расчеты)
 */
function processGrowthData(sheet, currentData, baseData) {
  // Обрабатываем каждую строку данных
  for (let row = 0; row < currentData.length; row++) {
    const rowName = currentData[row][0];
    if (!rowName || rowName.toString().trim() === '') continue;
    
    // Пропускаем строки роста (они уже существуют)
    if (rowName.toString().includes('Рост') || rowName.toString().includes('количестве')) continue;
    
    // Ищем соответствующую строку в базовых данных
    const baseRow = findMatchingRow(baseData, rowName);
    if (!baseRow) continue;
    
    // Рассчитываем рост для всех столбцов
    calculateRowGrowthAllColumns(sheet, row + 1, currentData[row], baseRow);
  }
}

/**
 * Расчет роста для всех столбцов одной строки (ОПТИМИЗИРОВАННАЯ ВЕРСИЯ)
 */
function calculateRowGrowthAllColumns(sheet, rowNumber, currentRow, baseRow) {
  // Собираем все изменения для batch-обработки
  const updates = [];
  
  СТОЛБЦЫ_РОСТА.forEach(config => {
    // Берем данные из правильных столбцов
    const currentValue = parseFloat(currentRow[config.столбец_данных - 1]) || 0;
    const baseValue = parseFloat(baseRow[config.столбец_данных - 1]) || 0;
    
    // Абсолютный рост (БЕЗ .00 для целых чисел)
    const absoluteGrowth = currentValue - baseValue;
    updates.push({
      row: rowNumber,
      col: config.столбец_роста,
      value: formatNumberWithoutUnnecessaryDecimals(absoluteGrowth)
    });
    
    // Процент роста (округление до 2 знаков)
    const percentGrowth = calculateCorrectGrowthPercent(baseValue, currentValue);
    const cleanPercentValue = parseFloat(percentGrowth.replace('%', '')).toFixed(2) + '%';
    updates.push({
      row: rowNumber,
      col: config.столбец_процента,
      value: cleanPercentValue
    });
  });
  
  // УЛЬТРА-БЫСТРАЯ BATCH-ОПЕРАЦИЯ: один вызов для всех столбцов роста
  if (updates.length > 0) {
    // Группируем обновления по строкам для максимальной скорости
    const rangeData = sheet.getRange(rowNumber, 1, 1, 35).getValues()[0];
    
    updates.forEach(update => {
      rangeData[update.col - 1] = update.value;
    });
    
    // ОДИН вызов setValues для всей строки
    sheet.getRange(rowNumber, 1, 1, 35).setValues([rangeData]);
  }
}

/**
 * Применение цветового форматирования к ячейкам роста
 */
function applyGrowthColorFormatting(sheet, currentData) {
  console.log('  🎨 Применяем цветовое форматирование к ячейкам роста...');
  
  // Собираем все операции форматирования
  const colorOperations = [];
  
  for (let row = 0; row < currentData.length; row++) {
    const rowName = currentData[row][0];
    if (!rowName || rowName.toString().trim() === '') continue;
    
    // Пропускаем строки роста
    if (rowName.toString().includes('Рост') || rowName.toString().includes('количестве')) continue;
    
    const rowNum = row + 1;
    
    // Форматируем столбцы роста для этой строки
    СТОЛБЦЫ_РОСТА.forEach(config => {
      try {
        // Получаем значения роста
        const absoluteGrowth = sheet.getRange(rowNum, config.столбец_роста).getValue();
        const percentGrowth = sheet.getRange(rowNum, config.столбец_процента).getValue();
        
        // Определяем цвета и добавляем в очередь
        const absoluteColor = getGrowthColor(absoluteGrowth);
        const percentColor = getGrowthColor(percentGrowth);
        
        if (absoluteColor) {
          colorOperations.push({row: rowNum, col: config.столбец_роста, color: absoluteColor});
        }
        if (percentColor) {
          colorOperations.push({row: rowNum, col: config.столбец_процента, color: percentColor});
        }
      } catch (error) {
        // Пропускаем ошибочные ячейки
      }
    });
  }
  
  // Применяем все операции одним batch'ем
  if (colorOperations.length > 0) {
    applyColorOperationsBatch(sheet, colorOperations);
  }
  
  console.log(`  ✅ Цветовое форматирование завершено`);
}

/**
 * Применить batch операции цветового форматирования
 */
function applyColorOperationsBatch(sheet, operations) {
  operations.forEach(op => {
    try {
      sheet.getRange(op.row, op.col).setFontColor(op.color);
    } catch (error) {
      console.log(`    ⚠️ Ошибка применения цвета к ячейке ${op.row},${op.col}: ${error.message}`);
    }
  });
}

/**
 * Форматирование числа без ненужных .00 для целых чисел
 */
function formatNumberWithoutUnnecessaryDecimals(value) {
  const numValue = parseFloat(value) || 0;
  
  // Если число целое - возвращаем без десятичных знаков
  if (numValue === Math.floor(numValue)) {
    return numValue.toString();
  }
  
  // Иначе округляем до 2 знаков после запятой
  return numValue.toFixed(2);
}

/**
 * Определение цвета для значения роста
 */
function getGrowthColor(value) {
  const numValue = parseFloat(value) || 0;
  
  if (numValue > 0) {
    return ЦВЕТА.РОСТ;        // Зеленый для положительного роста
  } else if (numValue < 0) {
    return ЦВЕТА.ПАДЕНИЕ;     // Красный для отрицательного роста
  } else {
    return ЦВЕТА.БЕЗ_ИЗМЕНЕНИЙ; // Серый для нуля
  }
}

/**
 * Обеспечение черного цвета для основных данных
 */
function ensureMainDataBlackColor(sheet, rowNumber) {
  // Основные столбцы данных (не столбцы роста)
  const mainDataColumns = [1, 2, 3, 4, 8, 12, 16, 20, 22, 25, 29]; // A, B, C, D, H, L, P, T, V, Y, AC
  
  mainDataColumns.forEach(colNum => {
    sheet.getRange(rowNumber, colNum).setFontColor('#000000'); // Черный цвет
  });
}

/**
 * Заполнить столбцы роста для одной строки (СТАРАЯ ВЕРСИЯ - для совместимости)
 */
function fillRowGrowthColumnsCorrect(sheet, rowNumber, currentRow, baseRow) {
  // Вызываем новую функцию
  calculateRowGrowthAllColumns(sheet, rowNumber, currentRow, baseRow);
}

/**
 * Найти соответствующую строку в базовых данных
 */
function findMatchingRow(baseData, rowName) {
  for (let i = 0; i < baseData.length; i++) {
    if (baseData[i][0] && baseData[i][0].toString().trim() === rowName.toString().trim()) {
      return baseData[i];
    }
  }
  return null;
}

/**
 * Правильный расчет процента роста (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ)
 */
function calculateCorrectGrowthPercent(baseValue, currentValue) {
  // Если базовое значение равно 0
  if (baseValue === 0) {
    if (currentValue === 0) return '0.00%';
    return 'Новые данные';
  }
  
  // Если базовое значение отрицательное, а текущее положительное
  if (baseValue < 0 && currentValue > 0) {
    // Рассчитываем как переход от убытка к прибыли
    const totalChange = Math.abs(baseValue) + currentValue;
    const percentChange = (totalChange / Math.abs(baseValue)) * 100;
    return `+${percentChange.toFixed(2)}%`;
  }
  
  // Стандартный расчет процента
  const percentChange = ((currentValue - baseValue) / Math.abs(baseValue)) * 100;
  const sign = percentChange >= 0 ? '+' : '';
  return `${sign}${percentChange.toFixed(2)}%`;
}

// ========================================================================
// 🎨 ФОРМАТИРОВАНИЕ (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ)
// ========================================================================

/**
 * Применить smartReplaceColors ко всем листам (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ)
 */
function applySmartFormattingToAllSheets(monthName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    
    const sheets = [
      `Сводная Бренд+Гео ${monthName}`,
      `Сводная ГЕО ${monthName}`,
      `Сводная Сеошники ${monthName}`
    ];
    
    sheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        try {
          console.log(`  🎨 Форматируем: ${sheetName}`);
          
          // Устанавливаем активный лист для smartReplaceColors
          spreadsheet.setActiveSheet(sheet);
          
          // Небольшая пауза для обеспечения синхронизации
          Utilities.sleep(100);
          
          // Проверяем что активный лист правильно установлен
          const activeSheet = SpreadsheetApp.getActiveSheet();
          if (activeSheet.getName() !== sheetName) {
            console.log(`⚠️ Активный лист: ${activeSheet.getName()}, ожидаемый: ${sheetName}`);
            // Пытаемся установить еще раз
            spreadsheet.setActiveSheet(sheet);
            Utilities.sleep(100);
          }
          
          // Применяем smartReplaceColors (должна быть в файле 3)
          smartReplaceColors();
          
          console.log(`  ✅ Лист ${sheetName} отформатирован`);
        } catch (error) {
          console.log(`❌ Ошибка: ${error}`);
          // Продолжаем с следующим листом вместо остановки
        }
      }
    });
    
    console.log('✅ Форматирование применено ко всем листам');
    
  } catch (error) {
    console.error('❌ Ошибка форматирования:', error);
    throw error;
  }
}

// ========================================================================
// 🚀 БЫСТРЫЕ ФУНКЦИИ С BATCH-ОБРАБОТКОЙ (НОВОЕ)
// ========================================================================

/**
 * Создать все сводные таблицы МАКСИМАЛЬНО БЫСТРО (БЕЗ ФОРМАТИРОВАНИЯ)
 */
function создатьВсеСводныеБыстро() {
  console.log('🚀 МАКСИМАЛЬНО БЫСТРОЕ СОЗДАНИЕ ВСЕХ СВОДНЫХ');
  console.log(`📅 Период: ${МЕСЯЦЫ.БАЗОВЫЙ} → ${МЕСЯЦЫ.ТЕКУЩИЙ}`);
  console.log('⚡ ЭТАП 1: Только данные (без форматирования)');
  console.log();
  
  const totalStartTime = new Date();
  
  try {
    const monthName = МЕСЯЦЫ.ТЕКУЩИЙ;
    
    console.log('1️⃣ МАКСИМАЛЬНО БЫСТРОЕ создание данных...');
    
    // Получаем данные один раз и кэшируем
    console.log('📥 Загружаем данные...');
    const MAIN_TABLE_ID = ТАБЛИЦЫ.ОСНОВНАЯ;
    const SITES_TABLE_ID = ТАБЛИЦЫ.САЙТЫ;
    
    const mainData = getMainTableData(MAIN_TABLE_ID, monthName);
    const mainDataGeo = getMainTableDataForGeo(MAIN_TABLE_ID, monthName);
    const sitesData = getSitesTableData(SITES_TABLE_ID);
    
    console.log(`📊 Загружено: ${mainData.length} строк основных данных, ${Object.keys(sitesData).length} идентификаторов сайтов`);
    
    // Группируем данные параллельно
    console.log('🔄 Группируем данные...');
    const groupedBrandData = groupData(mainData, sitesData);
    const groupedGeoData = groupGeoData(mainDataGeo, sitesData);
    const groupedSeoData = groupSeoData(mainData, sitesData);
    
    // Создаем сводные ТОЛЬКО ДАННЫЕ (максимально быстро)
    console.log('⚡ Создаем ТОЛЬКО ДАННЫЕ для максимальной скорости...');
    
    // Бренд+Гео - быстрая версия без форматирования
    createSummarySheetFast(ТАБЛИЦЫ.РЕЗУЛЬТАТ, groupedBrandData, monthName);
    
    // ГЕО - МАКСИМАЛЬНО УПРОЩЕННАЯ ВЕРСИЯ (от timeout)
    console.log('⚡ Создаем ГЕО сводную (упрощенная версия)...');
    createGeoSummarySimple(monthName, groupedGeoData);
    
    // Сеошники - МАКСИМАЛЬНО УПРОЩЕННАЯ ВЕРСИЯ (от timeout) 
    console.log('⚡ Создаем Сеошники сводную (упрощенная версия)...');
    createSeoSummarySimple(monthName, groupedSeoData);
    
    console.log('2️⃣ Расчет показателей роста...');
    
    // Рассчитываем рост между периодами
    if (МЕСЯЦЫ.БАЗОВЫЙ !== МЕСЯЦЫ.ТЕКУЩИЙ) {
      calculateGrowthBetweenPeriods(МЕСЯЦЫ.БАЗОВЫЙ, МЕСЯЦЫ.ТЕКУЩИЙ);
    }
    
    const totalEndTime = new Date();
    const totalExecutionTime = (totalEndTime - totalStartTime) / 1000;
    
    console.log('✅ ВСЕ СВОДНЫЕ СОЗДАНЫ МАКСИМАЛЬНО БЫСТРО (БЕЗ ФОРМАТИРОВАНИЯ)!');
    console.log(`⚡ ВРЕМЯ СОЗДАНИЯ ДАННЫХ: ${totalExecutionTime.toFixed(2)} секунд`);
    
    // ЭТАП 2: Применяем форматирование отдельно
    console.log();
    console.log('🎨 ЭТАП 2: Применение форматирования...');
    const formatStartTime = new Date();
    
    applySmartFormattingToAllSheets(monthName);
    
    const formatEndTime = new Date();
    const formatExecutionTime = (formatEndTime - formatStartTime) / 1000;
    const totalWithFormatTime = (formatEndTime - totalStartTime) / 1000;
    
    console.log(`🎨 ВРЕМЯ ФОРМАТИРОВАНИЯ: ${formatExecutionTime.toFixed(2)} секунд`);
    console.log(`🏁 ОБЩЕЕ ВРЕМЯ: ${totalWithFormatTime.toFixed(2)} секунд`);
    console.log(`📊 Соотношение: ${totalExecutionTime.toFixed(1)}с данные + ${formatExecutionTime.toFixed(1)}с форматирование`);
    
  } catch (error) {
    console.error('❌ Ошибка быстрого создания сводных:', error);
    throw error;
  }
}

/**
 * Создать последовательность сводных БЫСТРО для нескольких месяцев
 */
function создатьПоследовательностьМесяцевБыстро() {
  console.log('🚀 БЫСТРОЕ ПОСЛЕДОВАТЕЛЬНОЕ СОЗДАНИЕ СВОДНЫХ');
  console.log('🎯 Месяцы: Май → Июнь → Июль 2025');
  console.log('⚡ С batch-обработкой для максимальной скорости!');
  console.log();
  
  const totalStartTime = new Date();
  const месяцы = ['Май 2025', 'Июнь 2025', 'Июль 2025'];
  
  for (let i = 0; i < месяцы.length; i++) {
    const текущийМесяц = месяцы[i];
    const предыдущийМесяц = i > 0 ? месяцы[i - 1] : null;
    
    console.log(`📊 Быстро обрабатываем: ${текущийМесяц}`);
    
    // Устанавливаем период
    МЕСЯЦЫ.ТЕКУЩИЙ = текущийМесяц;
    МЕСЯЦЫ.БАЗОВЫЙ = предыдущийМесяц || текущийМесяц;
    
    // Создаем сводные для этого месяца БЫСТРО
    создатьВсеСводныеБыстро();
    
    console.log(`✅ ${текущийМесяц} готов быстро!`);
    console.log();
  }
  
  const totalEndTime = new Date();
  const totalExecutionTime = (totalEndTime - totalStartTime) / 1000;
  
  console.log('🎉 ВСЕ МЕСЯЦЫ ОБРАБОТАНЫ БЫСТРО!');
  console.log(`⚡ ОБЩЕЕ ВРЕМЯ: ${totalExecutionTime.toFixed(2)} секунд`);
  console.log(`🚀 Средняя скорость: ${(totalExecutionTime / месяцы.length).toFixed(2)} сек/месяц`);
}

/**
 * Batch-чтение данных из Google Sheets
 */
function getMainTableDataBatch(tableId, monthName) {
  console.log('📥 Batch-чтение основной таблицы...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheet = spreadsheet.getSheetByName(monthName);
    
    if (!sheet) {
      throw new Error(`Лист "${monthName}" не найден`);
    }
    
    // Читаем все данные одним вызовом
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    console.log(`📊 Прочитано ${values.length} строк за один вызов`);
    
    if (values.length < 2) {
      throw new Error('Недостаточно данных в таблице');
    }
    
    const data = [];
    
    // Обрабатываем все строки в памяти (быстро)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      if (!row[0] && !row[16]) continue;
      
      data.push({
        seoSpecialist: row[0] || '',
        clicks: parseFloat(row[4]) || 0,
        registrations: parseFloat(row[5]) || 0,
        fd: parseFloat(row[6]) || 0,
        rd: parseFloat(row[7]) || 0,
        deposits: parseFloat(row[8]) || 0,
        revenueUSD: parseFloat(row[13]) || 0,
        identifier: row[16] || '',
        brand: row[21] || '',
        geo: row[22] || '',
        brandGeo: row[23] || '',
        stream: row[24] || ''
      });
    }
    
    console.log(`✅ Обработано ${data.length} строк данных`);
    return data;
    
  } catch (error) {
    console.error('❌ Ошибка batch-чтения основной таблицы:', error);
    throw error;
  }
}

// ========================================================================
// 📈 ЗАПОЛНЕНИЕ СТРОК РОСТА ПО КЛАСТЕРАМ (НОВАЯ ФУНКЦИОНАЛЬНОСТЬ)
// ========================================================================

/**
 * НОВАЯ ФУНКЦИЯ: Заполнение строк роста по кластерам данных
 */
function fillGrowthRowsForClusters(sheet, currentData, baseData) {
  console.log('  📈 Заполняем строки роста по кластерам...');
  
  let processedGrowthRows = 0;
  
  for (let row = 0; row < currentData.length; row++) {
    const rowName = currentData[row][0];
    if (!rowName) continue;
    
    const rowStr = rowName.toString().trim();
    
    // Проверяем это ли строка роста
    if (rowStr === 'Рост / Падение' || rowStr === 'В количестве') {
      
      // Находим предыдущую строку данных (родительский кластер)
      const parentRow = findParentDataRow(currentData, row);
      if (!parentRow) continue;
      
      const parentName = parentRow[0];
      
      // Находим соответствующую родительскую строку в базовых данных
      const baseParentRow = findMatchingRow(baseData, parentName);
      if (!baseParentRow) continue;
      
      try {
        // Заполняем строку роста данными
        fillGrowthRowData(sheet, row + 1, parentRow, baseParentRow, rowStr === 'Рост / Падение');
        processedGrowthRows++;
      } catch (error) {
        console.log(`    ❌ Ошибка заполнения строки "${rowStr}": ${error.message}`);
      }
    }
  }
  
  console.log(`  ✅ Обработано строк роста: ${processedGrowthRows}`);
}

/**
 * Найти предыдущую строку с данными (родительский кластер)
 */
function findParentDataRow(data, currentRowIndex) {
  for (let i = currentRowIndex - 1; i >= 0; i--) {
    const rowName = data[i][0];
    if (!rowName) continue;
    
    const rowStr = rowName.toString().trim();
    
    // Пропускаем строки роста и пустые строки
    if (rowStr === 'Рост / Падение' || 
        rowStr === 'В количестве' || 
        rowStr === 'Рост / Падение' ||
        rowStr === '' ||
        rowStr.includes('Оффер') ||
        rowStr.includes('Сеошник') ||
        rowStr.includes('Клики')) {
      continue;
    }
    
    return data[i];
  }
  return null;
}

/**
 * Заполнить строку роста данными расчетов ПО ВСЕМ ПОКАЗАТЕЛЯМ (ОПТИМИЗИРОВАННАЯ ВЕРСИЯ)
 * 
 * ⚠️ ВАЖНО: Эта функция используется ТОЛЬКО для строк "Рост / Падение" и "В количестве"!
 * Она НЕ затрагивает оригинальные данные "% от общего" в обычных строках данных!
 */
function fillGrowthRowData(sheet, rowNumber, currentParentRow, baseParentRow, isPercentRow) {
  // Получаем ОБЩИЕ ИТОГИ всего листа для правильного расчета "% от общего"
  const currentData = sheet.getDataRange().getValues();
  const grandTotalsCurrentMonth = найтиИтоговуюСтроку(currentData);
  
  // Нужно получить базовые данные для сравнения (из базового месяца)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const baseMonth = МЕСЯЦЫ.БАЗОВЫЙ;
  const currentMonth = МЕСЯЦЫ.ТЕКУЩИЙ;
  const currentSheetName = sheet.getName();
  const baseSheetName = currentSheetName.replace(currentMonth, baseMonth);
  const baseSheet = spreadsheet.getSheetByName(baseSheetName);
  
  let grandTotalsBaseMonth = null;
  if (baseSheet) {
    const baseData = baseSheet.getDataRange().getValues();
    grandTotalsBaseMonth = найтиИтоговуюСтроку(baseData);
  }
  
  // Собираем все изменения в массивы для batch-обработки
  const valuesToSet = [];
  const colorsToSet = [];
  const maxColumn = 35; // Максимальный столбец для обработки
  
  // Инициализируем массивы
  for (let col = 1; col <= maxColumn; col++) {
    valuesToSet[col - 1] = null;
    colorsToSet[col - 1] = null;
  }
  
  СТОЛБЦЫ_РОСТА.forEach(config => {
    const currentValue = parseFloat(currentParentRow[config.столбец_данных - 1]) || 0;
    const baseValue = parseFloat(baseParentRow[config.столбец_данных - 1]) || 0;
    const absoluteGrowth = currentValue - baseValue;
    const percentGrowth = calculateCorrectGrowthPercent(baseValue, currentValue);
    
    if (isPercentRow) {
      // Строка "Рост / Падение" - заполняем ОСНОВНЫЕ столбцы процентами
      
      // 1. Основной столбец данных - процентный рост (округление до 2 знаков)
      valuesToSet[config.столбец_данных - 1] = parseFloat(percentGrowth.replace('%', '')).toFixed(2) + '%';
      colorsToSet[config.столбец_данных - 1] = getGrowthColor(absoluteGrowth);
      
      // 2. Столбец "% от общего" - изменение доли от ОБЩИХ ИТОГОВ (ИСПРАВЛЕНО!)
      const percentColumn = config.столбец_данных + 1;
      if (percentColumn <= maxColumn && percentColumn !== config.столбец_роста && percentColumn !== config.столбец_процента) {
        // ПРАВИЛЬНАЯ ЛОГИКА: сравниваем проценты от ОБЩИХ итогов листа
        let currentPercent = 0;
        let basePercent = 0;
        
        if (grandTotalsCurrentMonth) {
          const grandTotalCurrent = getGrandTotalByConfigName(grandTotalsCurrentMonth, config.название);
          currentPercent = grandTotalCurrent > 0 ? (currentValue / grandTotalCurrent * 100) : 0;
        }
        
        if (grandTotalsBaseMonth) {
          const grandTotalBase = getGrandTotalByConfigName(grandTotalsBaseMonth, config.название);  
          basePercent = grandTotalBase > 0 ? (baseValue / grandTotalBase * 100) : 0;
        }
        
        const percentChange = currentPercent - basePercent;
        valuesToSet[percentColumn - 1] = percentChange.toFixed(2) + '%';
        colorsToSet[percentColumn - 1] = getGrowthColor(percentChange);
      }
      
      // 3. Столбец роста % - процентный рост (округление до 2 знаков)
      if (config.столбец_процента <= maxColumn) {
        valuesToSet[config.столбец_процента - 1] = parseFloat(percentGrowth.replace('%', '')).toFixed(2) + '%';
        colorsToSet[config.столбец_процента - 1] = getGrowthColor(absoluteGrowth);
      }
      
    } else {
      // Строка "В количестве" - заполняем ВСЕ столбцы абсолютными значениями
      
      // 1. Основной столбец данных - абсолютный рост (БЕЗ .00 для целых чисел)
      valuesToSet[config.столбец_данных - 1] = formatNumberWithoutUnnecessaryDecimals(absoluteGrowth);
      colorsToSet[config.столбец_данных - 1] = getGrowthColor(absoluteGrowth);
      
      // 2. Столбец "% от общего" - изменение доли от ОБЩИХ ИТОГОВ (ИСПРАВЛЕНО!)
      const percentColumn = config.столбец_данных + 1;
      if (percentColumn <= maxColumn && percentColumn !== config.столбец_роста && percentColumn !== config.столбец_процента) {
        // ПРАВИЛЬНАЯ ЛОГИКА: сравниваем проценты от ОБЩИХ итогов листа
        let currentPercent = 0;
        let basePercent = 0;
        
        if (grandTotalsCurrentMonth) {
          const grandTotalCurrent = getGrandTotalByConfigName(grandTotalsCurrentMonth, config.название);
          currentPercent = grandTotalCurrent > 0 ? (currentValue / grandTotalCurrent * 100) : 0;
        }
        
        if (grandTotalsBaseMonth) {
          const grandTotalBase = getGrandTotalByConfigName(grandTotalsBaseMonth, config.название);
          basePercent = grandTotalBase > 0 ? (baseValue / grandTotalBase * 100) : 0;
        }
        
        const percentChange = currentPercent - basePercent;
        valuesToSet[percentColumn - 1] = percentChange.toFixed(2) + '%';
        colorsToSet[percentColumn - 1] = getGrowthColor(percentChange);
      }
      
      // 3. Столбец абсолютного роста - абсолютный рост (БЕЗ .00 для целых чисел)
      if (config.столбец_роста <= maxColumn) {
        valuesToSet[config.столбец_роста - 1] = formatNumberWithoutUnnecessaryDecimals(absoluteGrowth);
        colorsToSet[config.столбец_роста - 1] = getGrowthColor(absoluteGrowth);
      }
    }
  });
  
  // Дополнительно заполняем столбцы количества сайтов и сеошников
  fillAdditionalGrowthColumnsFast(valuesToSet, colorsToSet, currentParentRow, baseParentRow, isPercentRow);
  
  // САМАЯ БЫСТРАЯ BATCH-ОПЕРАЦИЯ: Получаем текущую строку и обновляем только нужные ячейки
  const currentRowData = sheet.getRange(rowNumber, 1, 1, maxColumn).getValues()[0];
  
  // Обновляем только те ячейки где есть новые значения
  for (let col = 0; col < maxColumn; col++) {
    if (valuesToSet[col] !== null) {
      currentRowData[col] = valuesToSet[col];
    }
  }
  
  // ОДИН вызов setValues для всей строки - максимальная скорость!
  sheet.getRange(rowNumber, 1, 1, maxColumn).setValues([currentRowData]);
  
  // БЫСТРОЕ цветовое форматирование только для измененных ячеек
  const rangeUpdates = [];
  for (let col = 0; col < maxColumn; col++) {
    if (colorsToSet[col] !== null) {
      rangeUpdates.push({col: col + 1, color: colorsToSet[col]});
    }
  }
  
  // МАКСИМАЛЬНО БЫСТРОЕ применение цветов batch'ем
  if (rangeUpdates.length > 0) {
    // Собираем все цвета в один Range для максимальной скорости
    rangeUpdates.forEach(update => {
      sheet.getRange(rowNumber, update.col).setFontColor(update.color);
    });
  }
}

/**
 * Рассчитать итоговые данные строки для процентов от общего (УСТАРЕЛА - НЕ ИСПОЛЬЗУЕТСЯ)
 * 
 * ⚠️ ЗАМЕНЕНА НА: использование grandTotalsCurrentMonth и grandTotalsBaseMonth
 */
function calculateRowTotals(rowData) {
  return {
    'Клики': parseFloat(rowData[3]) || 0,        // D
    'Реги': parseFloat(rowData[7]) || 0,         // H  
    'FD': parseFloat(rowData[11]) || 0,          // L
    'RD': parseFloat(rowData[15]) || 0,          // P
    'Расход': parseFloat(rowData[21]) || 0,      // V
    'Выручка': parseFloat(rowData[24]) || 0,     // Y  
    'Прибыль': parseFloat(rowData[28]) || 0      // AC
  };
}

/**
 * Заполнить дополнительные столбцы роста (сайты, сеошники и т.д.) - БЫСТРАЯ ВЕРСИЯ
 */
function fillAdditionalGrowthColumnsFast(valuesToSet, colorsToSet, currentParentRow, baseParentRow, isPercentRow) {
  // Столбец B: Количество сайтов с трафом
  const currentSites = parseFloat(currentParentRow[1]) || 0;
  const baseSites = parseFloat(baseParentRow[1]) || 0;
  const sitesGrowth = currentSites - baseSites;
  
  if (isPercentRow) {
    const sitesPercentGrowth = calculateCorrectGrowthPercent(baseSites, currentSites);
    valuesToSet[1] = parseFloat(sitesPercentGrowth.replace('%', '')).toFixed(2) + '%'; // Округление до 2 знаков
    colorsToSet[1] = getGrowthColor(sitesGrowth);
  } else {
    valuesToSet[1] = formatNumberWithoutUnnecessaryDecimals(sitesGrowth); // БЕЗ .00 для целых
    colorsToSet[1] = getGrowthColor(sitesGrowth);
  }
  
  // Столбец C: Количество сеошников
  const currentSeos = parseFloat(currentParentRow[2]) || 0;
  const baseSeos = parseFloat(baseParentRow[2]) || 0;
  const seosGrowth = currentSeos - baseSeos;
  
  if (isPercentRow) {
    const seosPercentGrowth = calculateCorrectGrowthPercent(baseSeos, currentSeos);
    valuesToSet[2] = parseFloat(seosPercentGrowth.replace('%', '')).toFixed(2) + '%'; // Округление до 2 знаков
    colorsToSet[2] = getGrowthColor(seosGrowth);
  } else {
    valuesToSet[2] = formatNumberWithoutUnnecessaryDecimals(seosGrowth); // БЕЗ .00 для целых
    colorsToSet[2] = getGrowthColor(seosGrowth);
  }
  
  // Дополнительные столбцы для других метрик (если нужны)
  const additionalColumns = [
    { index: 20, name: 'Deps' },        // T - Депозиты
    { index: 22, name: 'ExpenseRUB' },  // V - Расход RUB  
    { index: 32, name: 'RevenueRUB' },  // AF - Выручка RUB
    { index: 33, name: 'ProfitRUB' }    // AG - Прибыль RUB
  ];
  
  additionalColumns.forEach(col => {
    if (col.index <= 35) { // В пределах нашего maxColumn
      const currentValue = parseFloat(currentParentRow[col.index - 1]) || 0;
      const baseValue = parseFloat(baseParentRow[col.index - 1]) || 0;
      const absoluteGrowth = currentValue - baseValue;
      
      if (isPercentRow) {
        const percentGrowth = calculateCorrectGrowthPercent(baseValue, currentValue);
        valuesToSet[col.index - 1] = parseFloat(percentGrowth.replace('%', '')).toFixed(2) + '%';
        colorsToSet[col.index - 1] = getGrowthColor(absoluteGrowth);
      } else {
        valuesToSet[col.index - 1] = formatNumberWithoutUnnecessaryDecimals(absoluteGrowth); // БЕЗ .00 для целых
        colorsToSet[col.index - 1] = getGrowthColor(absoluteGrowth);
      }
    }
  });
}

/**
 * Заполнить дополнительные столбцы роста (сайты, сеошники и т.д.) - СТАРАЯ ВЕРСИЯ
 */
function fillAdditionalGrowthColumns(sheet, rowNumber, currentParentRow, baseParentRow, isPercentRow) {
  // Столбец B: Количество сайтов с трафом
  const currentSites = parseFloat(currentParentRow[1]) || 0;
  const baseSites = parseFloat(baseParentRow[1]) || 0;
  const sitesGrowth = currentSites - baseSites;
  
  if (isPercentRow) {
    const sitesPercentGrowth = calculateCorrectGrowthPercent(baseSites, currentSites);
    sheet.getRange(rowNumber, 2).setValue(parseFloat(sitesPercentGrowth.replace('%', '')).toFixed(2) + '%');
    sheet.getRange(rowNumber, 2).setFontColor(getGrowthColor(sitesGrowth));
  } else {
    sheet.getRange(rowNumber, 2).setValue(formatNumberWithoutUnnecessaryDecimals(sitesGrowth));
    sheet.getRange(rowNumber, 2).setFontColor(getGrowthColor(sitesGrowth));
  }
  
  // Столбец C: Количество сеошников
  const currentSeos = parseFloat(currentParentRow[2]) || 0;
  const baseSeos = parseFloat(baseParentRow[2]) || 0;
  const seosGrowth = currentSeos - baseSeos;
  
  if (isPercentRow) {
    const seosPercentGrowth = calculateCorrectGrowthPercent(baseSeos, currentSeos);
    sheet.getRange(rowNumber, 3).setValue(parseFloat(seosPercentGrowth.replace('%', '')).toFixed(2) + '%');
    sheet.getRange(rowNumber, 3).setFontColor(getGrowthColor(seosGrowth));
  } else {
    sheet.getRange(rowNumber, 3).setValue(formatNumberWithoutUnnecessaryDecimals(seosGrowth));
    sheet.getRange(rowNumber, 3).setFontColor(getGrowthColor(seosGrowth));
  }
}

// ========================================================================
// 🆘 ФУНКЦИЯ ВОССТАНОВЛЕНИЯ ДАННЫХ "% ОТ ОБЩЕГО"
// ========================================================================

/**
 * КРИТИЧЕСКАЯ ФУНКЦИЯ: Восстановить оригинальные данные "% от общего" 
 * (которые были затерты при расчете роста)
 */
function восстановитьПроцентыОтОбщего() {
  console.log('🆘 ВОССТАНОВЛЕНИЕ ДАННЫХ "% ОТ ОБЩЕГО"...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    const currentMonth = МЕСЯЦЫ.ТЕКУЩИЙ;
    
    if (!currentMonth) {
      throw new Error('МЕСЯЦЫ.ТЕКУЩИЙ не установлен!');
    }
    
    // Листы для восстановления
    const sheetsToRestore = [
      `Сводная Бренд+Гео ${currentMonth}`,
      `Сводная ГЕО ${currentMonth}`,
      `Сводная Сеошники ${currentMonth}`
    ];
    
    sheetsToRestore.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        console.log(`⚠️ Лист ${sheetName} не найден, пропускаем`);
        return;
      }
      
      console.log(`  🔄 Восстанавливаем проценты в: ${sheetName}`);
      
      const data = sheet.getDataRange().getValues();
      восстановитьПроцентыВЛисте(sheet, data);
      
      console.log(`  ✅ Восстановлено: ${sheetName}`);
    });
    
    console.log('🎉 ВОССТАНОВЛЕНИЕ ЗАВЕРШЕНО!');
    
  } catch (error) {
    console.error('❌ Ошибка восстановления:', error);
    throw error;
  }
}

/**
 * Восстановить проценты от общего в одном листе
 */
function восстановитьПроцентыВЛисте(sheet, data) {
  const updates = [];
  
  for (let row = 0; row < data.length; row++) {
    const rowName = data[row][0];
    if (!rowName) continue;
    
    const rowStr = rowName.toString().trim();
    
    // Пропускаем строки роста - их не трогаем
    if (rowStr === 'Рост / Падение' || rowStr === 'В количестве') {
      continue;
    }
    
    // Пропускаем заголовки и пустые строки
    if (rowStr === '' || rowStr.includes('Оффер') || 
        rowStr.includes('Сеошник') || rowStr.includes('Клики')) {
      continue;
    }
    
    // Восстанавливаем проценты для обычных строк данных
    const rowData = data[row];
    const rowNumber = row + 1;
    
    // Рассчитываем итоги для этой строки
    const totals = {
      clicks: parseFloat(rowData[3]) || 0,     // D - Клики
      regs: parseFloat(rowData[7]) || 0,       // H - Реги  
      fd: parseFloat(rowData[11]) || 0,        // L - FD
      rd: parseFloat(rowData[15]) || 0,        // P - RD
      expense: parseFloat(rowData[21]) || 0,   // V - Расход
      revenue: parseFloat(rowData[24]) || 0,   // Y - Выручка
      profit: parseFloat(rowData[28]) || 0     // AC - Прибыль
    };
    
    // Нужно получить общие итоги для расчета процентов
    const grandTotals = найтиИтоговуюСтроку(data);
    if (!grandTotals) continue;
    
    // Восстанавливаем проценты от общего
    const percentUpdates = [
      { col: 5, value: grandTotals.clicks > 0 ? (totals.clicks / grandTotals.clicks * 100).toFixed(2) + '%' : '0.00%' },     // E
      { col: 9, value: grandTotals.regs > 0 ? (totals.regs / grandTotals.regs * 100).toFixed(2) + '%' : '0.00%' },          // I
      { col: 13, value: grandTotals.fd > 0 ? (totals.fd / grandTotals.fd * 100).toFixed(2) + '%' : '0.00%' },               // M
      { col: 17, value: grandTotals.rd > 0 ? (totals.rd / grandTotals.rd * 100).toFixed(2) + '%' : '0.00%' },               // Q
      { col: 23, value: grandTotals.expense > 0 ? (totals.expense / grandTotals.expense * 100).toFixed(2) + '%' : '0.00%' }, // W
      { col: 26, value: grandTotals.revenue > 0 ? (totals.revenue / grandTotals.revenue * 100).toFixed(2) + '%' : '0.00%' }, // Z
      { col: 30, value: grandTotals.profit > 0 ? (totals.profit / grandTotals.profit * 100).toFixed(2) + '%' : '0.00%' }     // AD
    ];
    
    percentUpdates.forEach(update => {
      updates.push({
        row: rowNumber,
        col: update.col,
        value: update.value
      });
    });
  }
  
  // Применяем все восстановления
  updates.forEach(update => {
    sheet.getRange(update.row, update.col).setValue(update.value);
  });
  
  console.log(`    📊 Восстановлено ${updates.length} ячеек с процентами`);
}

/**
 * Найти итоговую строку "Всего" для расчета процентов
 */
function найтиИтоговуюСтроку(data) {
  for (let i = 0; i < data.length; i++) {
    const rowName = data[i][0];
    if (rowName && rowName.toString().trim() === 'Всего') {
      return {
        clicks: parseFloat(data[i][3]) || 0,     // D
        regs: parseFloat(data[i][7]) || 0,       // H
        fd: parseFloat(data[i][11]) || 0,        // L
        rd: parseFloat(data[i][15]) || 0,        // P
        expense: parseFloat(data[i][21]) || 0,   // V
        revenue: parseFloat(data[i][24]) || 0,   // Y
        profit: parseFloat(data[i][28]) || 0     // AC
      };
    }
  }
  return null;
}

/**
 * Получить общий итог по названию конфигурации
 */
function getGrandTotalByConfigName(grandTotals, configName) {
  const mapping = {
    'Клики': 'clicks',
    'Реги': 'regs', 
    'FD': 'fd',
    'RD': 'rd',
    'Расход': 'expense',
    'Выручка': 'revenue',
    'Прибыль': 'profit'
  };
  
  const fieldName = mapping[configName];
  return fieldName ? (grandTotals[fieldName] || 0) : 0;
}

// ========================================================================
// 🚀 ФУНКЦИЯ ДЛЯ РАБОТЫ С ГОТОВЫМИ СВОДНЫМИ (НОВАЯ!)
// ========================================================================

/**
 * НОВАЯ ФУНКЦИЯ: Рассчитать рост для УЖЕ ГОТОВЫХ сводных (использует КОНСТАНТЫ)
 */
function рассчитатьРостДляГотовыхСводныхИзКонстант() {
  const baseMonth = МЕСЯЦЫ.БАЗОВЫЙ;
  const currentMonth = МЕСЯЦЫ.ТЕКУЩИЙ;
  
  console.log(`🎯 РАСЧЕТ РОСТА ДЛЯ ГОТОВЫХ СВОДНЫХ ИЗ КОНСТАНТ: ${baseMonth} → ${currentMonth}`);
  
  // Проверяем что константы установлены
  if (!baseMonth || !currentMonth) {
    throw new Error(`ОШИБКА: Константы не установлены! БАЗОВЫЙ="${baseMonth}", ТЕКУЩИЙ="${currentMonth}"`);
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    
    // Листы для обработки (только существующие!)
    const possibleSheets = [
      `Сводная Бренд+Гео ${currentMonth}`,
      `Сводная ГЕО ${currentMonth}`,
      `Сводная Сеошники ${currentMonth}`
    ];
    
    const existingSheets = possibleSheets.filter(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      return sheet !== null;
    });
    
    console.log(`📊 Найдено готовых сводных: ${existingSheets.length}`);
    
    if (existingSheets.length === 0) {
      console.log('❌ Готовые сводные не найдены! Сначала создайте сводные.');
      return;
    }
    
    existingSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      console.log(`  📈 Рассчитываем рост для: ${sheetName}`);
      
      // Получаем базовые данные для сравнения
      const baseSheetName = sheetName.replace(currentMonth, baseMonth);
      const baseSheet = spreadsheet.getSheetByName(baseSheetName);
      
      if (!baseSheet) {
        console.log(`  ⚠️ Базовый лист ${baseSheetName} не найден, пропускаем`);
        return;
      }
      
      const currentData = sheet.getDataRange().getValues();
      const baseData = baseSheet.getDataRange().getValues();
      
      console.log(`    🔢 Обрабатываем столбцы роста...`);
      processGrowthData(sheet, currentData, baseData);
      
      console.log(`    📊 Заполняем строки роста...`);
      fillGrowthRowsForClusters(sheet, currentData, baseData);
      
      console.log(`    🎨 Применяем цветовое форматирование...`);
      applyGrowthColorFormatting(sheet, currentData);
      
      console.log(`  ✅ Рост рассчитан: ${sheetName}`);
    });
    
    console.log('🎉 РАСЧЕТ РОСТА ДЛЯ ГОТОВЫХ СВОДНЫХ ЗАВЕРШЕН!');
    
  } catch (error) {
    console.error('❌ Ошибка расчета роста для готовых сводных:', error);
    throw error;
  }
}

/**
 * ФУНКЦИЯ С ПАРАМЕТРАМИ: Рассчитать рост для УЖЕ ГОТОВЫХ сводных (принимает месяцы)
 */
function рассчитатьРостДляГотовыхСводных(baseMonth, currentMonth) {
  console.log(`🎯 РАСЧЕТ РОСТА ДЛЯ ГОТОВЫХ СВОДНЫХ: ${baseMonth} → ${currentMonth}`);
  
  // Устанавливаем месяцы в константы
  МЕСЯЦЫ.БАЗОВЫЙ = baseMonth;
  МЕСЯЦЫ.ТЕКУЩИЙ = currentMonth;
  
  try {
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    
    // Листы для обработки (только существующие!)
    const possibleSheets = [
      `Сводная Бренд+Гео ${currentMonth}`,
      `Сводная ГЕО ${currentMonth}`,
      `Сводная Сеошники ${currentMonth}`
    ];
    
    const existingSheets = possibleSheets.filter(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      return sheet !== null;
    });
    
    console.log(`📊 Найдено готовых сводных: ${existingSheets.length}`);
    
    if (existingSheets.length === 0) {
      console.log('❌ Готовые сводные не найдены! Сначала создайте сводные.');
      return;
    }
    
    existingSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      console.log(`  📈 Рассчитываем рост для: ${sheetName}`);
      
      // Получаем базовые данные для сравнения
      const baseSheetName = sheetName.replace(currentMonth, baseMonth);
      const baseSheet = spreadsheet.getSheetByName(baseSheetName);
      
      if (!baseSheet) {
        console.log(`  ⚠️ Базовый лист ${baseSheetName} не найден, пропускаем`);
        return;
      }
      
      const currentData = sheet.getDataRange().getValues();
      const baseData = baseSheet.getDataRange().getValues();
      
      console.log(`    🔢 Обрабатываем столбцы роста...`);
      processGrowthData(sheet, currentData, baseData);
      
      console.log(`    📊 Заполняем строки роста...`);
      fillGrowthRowsForClusters(sheet, currentData, baseData);
      
      console.log(`    🎨 Применяем цветовое форматирование...`);
      applyGrowthColorFormatting(sheet, currentData);
      
      console.log(`  ✅ Рост рассчитан: ${sheetName}`);
    });
    
    console.log('🎉 РАСЧЕТ РОСТА ДЛЯ ГОТОВЫХ СВОДНЫХ ЗАВЕРШЕН!');
    
  } catch (error) {
    console.error('❌ Ошибка расчета роста для готовых сводных:', error);
    throw error;
  }
}

console.log('🎯 ОСНОВНАЯ ЛОГИКА ГОТОВА!');
console.log('🚀 + БЫСТРЫЕ ФУНКЦИИ С BATCH-ОБРАБОТКОЙ ДОБАВЛЕНЫ!');
console.log('📈 + ЗАПОЛНЕНИЕ СТРОК РОСТА ПО КЛАСТЕРАМ РЕАЛИЗОВАНО!');
console.log('🎨 + ЗАПОЛНЕНИЕ ВСЕХ СТОЛБЦОВ В СТРОКАХ РОСТА С ПОЛНОЙ ИНФОРМАЦИЕЙ!');
console.log('🏃‍♂️ + ДВЕ ФУНКЦИИ ДЛЯ РАБОТЫ С ГОТОВЫМИ СВОДНЫМИ:');
console.log('    📋 рассчитатьРостДляГотовыхСводныхИзКонстант() - из МЕСЯЦЫ.БАЗОВЫЙ/ТЕКУЩИЙ');
console.log('    ⚙️ рассчитатьРостДляГотовыхСводных(base, current) - с параметрами');
console.log('⚡ + УСКОРЕННАЯ ОБРАБОТКА:');
console.log('    📊 Столбцы "% от общего" обрабатываются в строках роста');
console.log('    🔢 Все значения округляются до 2 знаков после запятой');
console.log('    🚀 Batch-операции для максимальной скорости');
console.log('🆘 + ФУНКЦИЯ ВОССТАНОВЛЕНИЯ:');
console.log('    📋 восстановитьПроцентыОтОбщего() - восстанавливает затертые данные');
console.log('    🔄 Пересчитывает оригинальные проценты от общего для всех строк данных');
console.log('🔧 + ИСПРАВЛЕНИЕ РАСЧЕТА ПРОЦЕНТОВ ОТ ОБЩЕГО:');
console.log('    📊 Теперь использует ОБЩИЕ ИТОГИ листа вместо итогов родительской строки');
console.log('    ✅ Правильно рассчитывает изменения долей в сравнении с предыдущим месяцем');
console.log('🛠️ + ИСПРАВЛЕНИЯ ПРОИЗВОДИТЕЛЬНОСТИ И ФОРМАТИРОВАНИЯ:');
console.log('    ⚡ Оптимизировано цветовое форматирование - избегает зависаний');
console.log('    🔢 Числа в строках "В количестве" теперь БЕЗ .00 (7278 вместо 7278.00)');
console.log('    📋 Улучшено логирование процесса заполнения строк роста');
console.log('    ✅ Исправлено название строк роста: "Рост / Падение" (без %)');
console.log('🚀 + УЛЬТРА-БЫСТРАЯ BATCH-ОБРАБОТКА:');
console.log('    ⚡ ОДИН вызов setValues() для всей строки вместо множества setValue()');
console.log('    📊 Убраны избыточные логи по каждой строке для максимальной скорости');
console.log('    🎯 Оптимизирована работа с таблицами на 1000+ строк');
