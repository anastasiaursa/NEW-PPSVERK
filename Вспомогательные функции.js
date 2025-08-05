/*
 * 🔧 3 ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
 * 
 * ✅ Что делает: ВСЕ вспомогательные функции из оригинального кода
 * ✅ smartReplaceColors, группировка данных, создание листов, форматирование
 * ✅ Функции расчетов, заголовков, строк, группировки
 * 
 * Зависит от: 1 КОНФИГУРАЦИЯ И НАСТРОЙКИ.js
 * 
 * ВАЖНО: ВСЕ ФУНКЦИИ СКОПИРОВАНЫ ИЗ ОРИГИНАЛЬНОГО КОДА!
 * Применены только исправления пользователя
 */

console.log('🔧 ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ЗАГРУЖЕНЫ ПОЛНОСТЬЮ');
console.log('🎨 smartReplaceColors доступна (КАК ТРЕБОВАЛ ПОЛЬЗОВАТЕЛЬ)');
console.log('📊 Все функции из оригинального кода скопированы для ИДЕНТИЧНОГО результата');

// ========================================================================
// 🎨 SMARTREPLACECOLORS (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ)
// ========================================================================

/**
 * ОСНОВНАЯ ФУНКЦИЯ ФОРМАТИРОВАНИЯ (ТОЧНАЯ КОПИЯ - ПОЛЬЗОВАТЕЛЬ ТРЕБУЕТ ИМЕННО ЭТУ ЛОГИКУ)
 */
function smartReplaceColors() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    console.log(`Обрабатываем лист: ${sheet.getName()}`);
    console.log(`Строк: ${lastRow}, Столбцов: ${lastCol}`);
    
    if (lastRow <= 1 || lastCol <= 0) {
      console.log('Нет данных для обработки');
      return;
    }
    
    // Получаем все данные
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const values = range.getValues();
    
    // Обрабатываем каждую строку
    for (let row = 0; row < lastRow; row++) {
      const cellA = values[row][0];
      if (!cellA) continue;
      
      const cellText = cellA.toString().trim();
      let color = getColorForText(cellText);
      
      if (color) {
        // Специальная обработка для сводных Бренд+Гео
        if (isBrandGeoSummary(cellText)) {
          console.log(`🎨 Применяем сиреневый градиент для Бренд+Гео: "${cellText}"`);
          // Столбец A (бренд) темнее
          sheet.getRange(row + 1, 1).setBackground('#c9b3d9'); // Темный сиреневый
          // Остальные столбцы (бренд+гео) светлее
          if (lastCol > 1) {
            sheet.getRange(row + 1, 2, 1, lastCol - 1).setBackground('#f0ecf5'); // Светлый сиреневый
          }
          // Черный текст для сиреневых фонов
          sheet.getRange(row + 1, 1, 1, lastCol).setFontColor('#000000');
        }
        // Проверяем, нужен ли градиент
        else if (needsGradient(cellText)) {
          console.log(`🎨 Применяем обычный градиент для: "${cellText}"`);
          // Столбец A темнее
          sheet.getRange(row + 1, 1).setBackground(getDarkColor(color));
          // Остальные светлее
          if (lastCol > 1) {
            sheet.getRange(row + 1, 2, 1, lastCol - 1).setBackground(getLightColor(color));
          }
          
          // Ставим черный текст для светлых фонов
          if (color === '#ffffff' || color === '#d9d2e9' || color === '#f4e6ed') {
            sheet.getRange(row + 1, 1, 1, lastCol).setFontColor('#000000');
          }
        } else {
          console.log(`🎨 Применяем одинаковый цвет для всей строки: "${cellText}" - ${color}`);
          // Вся строка одинаковая
          sheet.getRange(row + 1, 1, 1, lastCol).setBackground(color);
          
          // Для белого фона ставим черный текст
          if (color === '#ffffff') {
            sheet.getRange(row + 1, 1, 1, lastCol).setFontColor('#000000');
          }
        }
      }
    }
    
    // Окрашиваем текст в столбцах роста
    colorizeGrowthText(sheet, values, lastRow, lastCol);
    
    console.log('✅ Готово!');
    
  } catch (error) {
    console.error('❌ Ошибка:', error);
    SpreadsheetApp.getUi().alert('Ошибка', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Определяет цвет для текста
 */
function getColorForText(text) {
  const lower = text.toLowerCase();
  
  console.log(`🔍 Определяем цвет для: "${text}"`);
  
  // 1. Заголовки месяцев (содержат год)
  if (/\d{4}/.test(text) && (text.includes('2024') || text.includes('2025') || text.includes('2026'))) {
    console.log(`📅 Заголовок месяца: ${text}`);
    return '#7f9cb9'; // Синий
  }
  
  // 2. Строки "Всего"
  if (/^(всего|итого|total)$/i.test(lower)) {
    console.log(`📊 Строка "Всего": ${text}`);
    return '#d9e1f2'; // Светло-голубой
  }
  
  // 3. Строки роста и "В количестве" - БЕЛЫЕ!
  if (lower === 'рост / падение' || 
      lower === 'рост / падение %' || 
      lower === 'в количестве' ||
      /^в\s+количестве/i.test(text) ||
      lower.includes('рост') || 
      lower.includes('падение')) {
    console.log(`📈 Строка роста: ${text}`);
    return '#ffffff'; // Белый
  }
  
  // 4. Заголовки таблиц
  if (isTableHeader(lower)) {
    console.log(`📋 Заголовок таблицы: ${text}`);
    return '#b8cce4'; // Средне-голубой
  }
  
  // 5. Проверяем сначала комбинации Бренд+Гео
  if (isBrandGeoSummary(text)) {
    console.log(`🎯 Бренд+Гео комбинация: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 6. Бренды (только отдельные)
  if (isBrand(lower)) {
    console.log(`🏢 Бренд: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 7. ГЕО строки
  if (isGeo(text)) {
            console.log(`ГЕО строка: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 8. Сеошники
  if (isSeo(text)) {
    console.log(`👤 Сеошник: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 9. Подгруппы (сводные)
  if (lower.includes('сводная') || text.includes('/')) {
    console.log(`📁 Сводная/подгруппа: ${text}`);
    return '#f4e6ed'; // Светлый розовый
  }
  
  // 10. Все остальное - белое
  console.log(`⚪ Белый фон по умолчанию: ${text}`);
  return '#ffffff';
}

/**
 * Проверяет, нужен ли градиент (темнее A, светлее остальные)
 */
function needsGradient(text) {
  const lower = text.toLowerCase();
  
  // Градиент НЕ нужен для сводных Бренд+Гео (они обрабатываются отдельно)
  if (isBrandGeoSummary(text)) {
    return false;
  }
  
  // Градиент для отдельных брендов, гео и сеошников
  return isBrand(lower) || isGeo(text) || isSeo(text);
}

/**
 * Проверяет, является ли сводной строкой Бренд+Гео
 * Поддерживает форматы: "ГЕО БРЕНД", "ГЕО / ГЕО БРЕНД", "Бренд+ГЕО"
 */
function isBrandGeoSummary(text) {
  const cleanText = text.trim();
  
  console.log(`Проверяем строку на Бренд+Гео: "${cleanText}"`);
  
  // Формат "Бренд+ГЕО" (оригинальная логика)
  if (cleanText.includes('+')) {
    const parts = cleanText.split('+');
    if (parts.length === 2) {
      const part1 = parts[0].trim().toLowerCase();
      const part2 = parts[1].trim().toLowerCase();
      
      if ((isBrand(part1) && isGeoCode(part2)) || 
          (isGeoCode(part1) && isBrand(part2))) {
        console.log(`✅ Найден формат "бренд+гео": ${cleanText}`);
        return true;
      }
    }
  }
  
  // Формат "ГЕО / ГЕО БРЕНД" (например "IN / UZ PinUp")
  if (cleanText.includes(' / ') && cleanText.includes(' ')) {
    const parts = cleanText.split(' ');
    if (parts.length >= 4) { // минимум "XX / XX Brand"
      const geoPart1 = parts[0].toLowerCase();
      const geoPart2 = parts[2].toLowerCase();
      const brandPart = parts.slice(3).join(' ').toLowerCase();
      
      if (isGeoCode(geoPart1) && isGeoCode(geoPart2) && isBrand(brandPart)) {
        console.log(`✅ Найден формат "гео / гео бренд": ${cleanText}`);
        return true;
      }
    }
  }
  
  // Формат "ГЕО БРЕНД" (например "HU Mostbet", "AZ 1win")
  if (cleanText.includes(' ')) {
    const parts = cleanText.split(' ');
    
    if (parts.length >= 2) {
      const geoPart = parts[0].toLowerCase();
      const brandPart = parts.slice(1).join(' ').toLowerCase();
      
      if (isGeoCode(geoPart) && isBrand(brandPart)) {
        console.log(`✅ Найден формат "гео бренд": ${cleanText}`);
        return true;
      }
    }
  }
  
  console.log(`❌ Не является Бренд+Гео: ${cleanText}`);
  return false;
}

/**
 * Проверяет коды стран для сводных (ОБНОВЛЕННЫЙ СПИСОК)
 */
function isGeoCode(text) {
  const geoCodes = [
    'az', 'ru', 'uz', 'kz', 'by', 'ua', 'ge', 'am', 'md', 'tj', 'kg', 'tm', 'hu',
    'tr', 'bd', 'np', 'in', 'ar', 'fr', 'kr', 'ww', 'ug', 'ec', 'ga', 'es', 'br',
    'cm', 'ke', 'sn', 'pe', 'jp', 'ca', 'bg', 'ch', 'co', 'pl', 'cl', 'eg', 'it',
    'gr', 'de', 'sk', 'lk', 'pt', 'ng', 'mx', 'vn', 'tz', 'za', 'mz', 'pk', 'cz'
  ];
  const result = geoCodes.includes(text.toLowerCase());
  if (result) {
    console.log(`✅ Распознан гео-код: ${text}`);
  }
  return result;
}

/**
 * Проверяет, является ли заголовком таблицы
 */
function isTableHeader(lower) {
  const keywords = ['оффер', 'сайты', 'сеошник', 'клики', 'реги', 'deps', 'выручка', 'количество', 'fd', 'rd'];
  return keywords.some(word => lower.includes(word)) && 
         !lower.includes('рост') && 
         !lower.includes('падение');
}

/**
 * Проверяет, является ли брендом (ОБНОВЛЕННЫЙ СПИСОК)
 */
function isBrand(lower) {
  const brands = [
    'glory', 'profitpoint', 'pinup', 'pin-up', 'banzai', '1win', '22bet', 'hellspin', 
    'ivibet', 'elon', '888starz', 'weiss', 'linebet', 'betmatch', 'пари',
    'фонбет', 'бетсити', '1xbet', 'advertise', 'joycasino', 'casinox',
    'melbet', 'aslan', 'casino', '7slots.casino', 'ice casino', 'mostbet',
    'bilbet', 'bons', 'betwinner', 'platinum casino', 'leon', 'booi casino',
    'aslan casino', 'aslan сasino'
  ];
  
  // Точное совпадение или бренд + 2-3 буквы
  for (const brand of brands) {
    if (lower === brand) {
      console.log(`✅ Распознан бренд: ${lower}`);
      return true;
    }
    if (lower.startsWith(brand + ' ') && /^[a-zа-я]{2,3}$/.test(lower.substring(brand.length + 1))) {
      console.log(`✅ Распознан бренд с суффиксом: ${lower}`);
      return true;
    }
    if (lower.startsWith(brand) && /^[a-zа-я]{2,3}$/.test(lower.substring(brand.length))) {
      console.log(`✅ Распознан бренд с суффиксом: ${lower}`);
      return true;
    }
  }
  
  return false;
}

/**
 * Проверяет, является ли ГЕО строкой
 */
function isGeo(text) {
  // ОБНОВЛЕНО: поддержка одиночных и составных ГЕО кодов
  const geoPatterns = [
    // Одиночные ГЕО коды
    /^AZ$/i, /^RU$/i, /^UZ$/i, /^KZ$/i, /^BY$/i,
    /^UA$/i, /^GE$/i, /^AM$/i, /^MD$/i, /^TJ$/i,
    /^KG$/i, /^TM$/i, /^HU$/i, /^[A-Z]{2}$/i,
    
    // Составные ГЕО коды типа "UZ / AZ" или "КR / JP"
    /^[A-Z]{2}\s*\/\s*[A-Z]{2}$/i,        // XX / YY с пробелами или без
    /^[А-Я]{2}\s*\/\s*[A-Z]{2}$/i,        // КR / JP (кириллица + латиница)
    /^[A-Z]{2}\s*\/\s*[А-Я]{2}$/i,        // XX / УК (латиница + кириллица)
    /^[А-Я]{2}\s*\/\s*[А-Я]{2}$/i         // УК / РУ (кириллица + кириллица)
  ];
  
  return geoPatterns.some(pattern => pattern.test(text.trim()));
}

/**
 * Проверяет, является ли сеошником
 */
function isSeo(text) {
  if (/^👤\s*/.test(text) || /^⚠️\s*/.test(text)) return true;
  
  const seoNames = [
    'михаил', 'анастасия', 'alex se', 'виктор', 'эмиль', 'кирилл', 
    'александр', 'александр артюх', 'майда', 'евгения'
  ];
  
  const clean = text.toLowerCase().replace(/^(👤|⚠️)\s*/, '').replace(/\s*\(исключен\)$/, '');
  return seoNames.includes(clean);
}

/**
 * Возвращает темный вариант цвета
 */
function getDarkColor(color) {
  const map = {
    '#d9d2e9': '#c9b3d9', // Сиреневый -> темнее
    '#f4e6ed': '#e6c9d6', // Розовый -> темнее
    '#ffffff': '#f0f0f0'  // Белый -> серый
  };
  return map[color] || color;
}

/**
 * Возвращает светлый вариант цвета
 */
function getLightColor(color) {
  const map = {
    '#d9d2e9': '#f0ecf5', // Сиреневый -> светлее
    '#f4e6ed': '#faf3f7', // Розовый -> светлее
    '#ffffff': '#ffffff'  // Белый остается белым
  };
  return map[color] || color;
}

/**
 * Окрашивает текст в столбцах роста
 */
function colorizeGrowthText(sheet, values, lastRow, lastCol) {
  // Находим столбцы роста по заголовкам
  const headerRow = values[0];
  const growthColumns = [];
  
  for (let col = 0; col < lastCol; col++) {
    const header = headerRow[col] ? headerRow[col].toString().toLowerCase().trim() : '';
    
    if (header.includes('рост') && (header.includes('падение') || header.includes('%'))) {
      growthColumns.push(col);
    }
  }
  
  // Окрашиваем текст в найденных столбцах
  for (const col of growthColumns) {
    for (let row = 1; row < lastRow; row++) { // Начинаем с 1, пропускаем заголовок
      const cellValue = values[row][col];
      
      if (cellValue === null || cellValue === undefined || cellValue === '') {
        continue;
      }
      
      // Парсим число из текста
      const numericValue = parseFloat(cellValue.toString().replace(/[^\-\d.,]/g, '').replace(',', '.'));
      
      if (isNaN(numericValue)) {
        continue;
      }
      
      const cell = sheet.getRange(row + 1, col + 1);
      
      if (numericValue > 0) {
        cell.setFontColor('#00AA00'); // Зеленый для роста
      } else if (numericValue < 0) {
        cell.setFontColor('#FF0000'); // Красный для падения
      } else {
        cell.setFontColor('#000000'); // Черный для нуля
      }
    }
  }
  
  if (growthColumns.length > 0) {
    console.log(`Окрашен текст в столбцах роста: ${growthColumns.length}`);
  }
}

// ========================================================================
// 📊 ГРУППИРОВКА ДАННЫХ (КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
// ========================================================================

/**
 * Группировка данных по брендам (ТОЧНАЯ КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function groupData(mainData, sitesData) {
  const grouped = {};
  
  mainData.forEach(item => {
    const brandKey = item.brand || 'Неизвестно';
    const brandGeoKey = item.brandGeo || 'Неизвестно';
    const streamKey = item.stream || 'Неизвестно';
    const seoKey = item.seoSpecialist || 'Неизвестно';
    
    if (!grouped[brandKey]) {
      grouped[brandKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        brandGeos: {},
        seoSpecialists: new Set(),
        allSites: new Set()
      };
    }
    
    if (!grouped[brandKey].brandGeos[brandGeoKey]) {
      grouped[brandKey].brandGeos[brandGeoKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        streams: {},
        seoSpecialists: new Set(),
        allSites: new Set()
      };
    }
    
    if (!grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey]) {
      grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        seos: {},
        sites: [],
        seoSpecialists: new Set(),
        allSites: new Set()
      };
    }
    
    if (!grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].seos[seoKey]) {
      grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].seos[seoKey] = {
        clicks: 0,
        registrations: 0,
        fd: 0,
        rd: 0,
        deposits: 0,
        revenueUSD: 0,
        sitesCount: 0,
        seoCount: 1, // Один сеошник в этой группе
        allSites: new Set()
      };
    }
    
    // Добавляем сеошника в множества (исключаем определенных из подсчета)
    if (!ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoKey)) {
      grouped[brandKey].seoSpecialists.add(seoKey);
      grouped[brandKey].brandGeos[brandGeoKey].seoSpecialists.add(seoKey);
      grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].seoSpecialists.add(seoKey);
    }
    
    // Суммируем данные на всех уровнях
    const levels = [
      grouped[brandKey].data,
      grouped[brandKey].brandGeos[brandGeoKey].data,
      grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].data,
      grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].seos[seoKey]
    ];
    
    levels.forEach(level => {
      level.clicks += item.clicks;
      level.registrations += item.registrations;
      level.fd += item.fd;
      level.rd += item.rd;
      level.deposits += item.deposits;
      level.revenueUSD += item.revenueUSD;
    });
    
    // Добавляем информацию о сайтах (убираем дублирование по идентификатору)
    if (item.identifier && sitesData[item.identifier]) {
      const uniqueSites = new Map();
      
      sitesData[item.identifier].forEach(site => {
        const siteKey = `${site.site}-${item.identifier}`;
        if (!uniqueSites.has(siteKey)) {
          uniqueSites.set(siteKey, {
            stream: streamKey,
            site: site.site,
            seoSpecialist: site.seoSpecialist,
            identifier: item.identifier,
            clicks: site.clicks
          });
        }
      });
      
      // ИСПРАВЛЕНИЕ: проверяем что сайты еще не добавлены в этот поток
      if (!grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed) {
        grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed = new Set();
      }
      
      uniqueSites.forEach(siteInfo => {
        const siteStreamKey = `${siteInfo.site}-${item.identifier}-${streamKey}`;
        
        // Добавляем сайт только если он еще не был добавлен в этот поток
        if (!grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed.has(siteStreamKey)) {
          grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].sites.push(siteInfo);
          grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed.add(siteStreamKey);
          
          // Добавляем сайты в множества для подсчета
          const siteKey = `${siteInfo.site}-${item.identifier}`;
          grouped[brandKey].allSites.add(siteKey);
          grouped[brandKey].brandGeos[brandGeoKey].allSites.add(siteKey);
          grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].allSites.add(siteKey);
          
          // Добавляем сайт к конкретному сеошнику
          if (siteInfo.seoSpecialist === seoKey) {
            grouped[brandKey].brandGeos[brandGeoKey].streams[streamKey].seos[seoKey].allSites.add(siteKey);
          }
        }
      });
    }
  });
  
  // Устанавливаем количество сеошников и сайтов
  Object.values(grouped).forEach(brand => {
    brand.data.seoCount = brand.seoSpecialists.size;
    brand.data.sitesCount = brand.allSites.size;
    
    Object.values(brand.brandGeos).forEach(brandGeo => {
      brandGeo.data.seoCount = brandGeo.seoSpecialists.size;
      brandGeo.data.sitesCount = brandGeo.allSites.size;
      
      Object.values(brandGeo.streams).forEach(stream => {
        stream.data.seoCount = stream.seoSpecialists.size;
        stream.data.sitesCount = stream.allSites.size;
        
        // Устанавливаем количество сайтов для каждого сеошника
        Object.values(stream.seos).forEach(seo => {
          seo.sitesCount = seo.allSites ? seo.allSites.size : 0;
        });
      });
    });
  });
  
  return grouped;
}

/**
 * Группировка данных по ГЕО (ТОЧНАЯ КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function groupGeoData(mainData, sitesData) {
  const grouped = {};
  
  mainData.forEach(item => {
    const geoKey = item.geo || 'Неизвестно';
    const geoBrandKey = `${geoKey}_${item.brand}`;
    const streamKey = item.stream || 'Неизвестно';
    const seoKey = item.seoSpecialist || 'Неизвестно';
    
    // Инициализируем структуру ГЕО
    if (!grouped[geoKey]) {
      grouped[geoKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        geoBrands: {},
        seoSpecialists: new Set(),
        allSites: new Set()
      };
    }
    
    if (!grouped[geoKey].geoBrands[geoBrandKey]) {
      grouped[geoKey].geoBrands[geoBrandKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        streams: {},
        seoSpecialists: new Set(),
        allSites: new Set()
      };
    }
    
    if (!grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey]) {
      grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        seos: {},
        sites: [],
        seoSpecialists: new Set(),
        allSites: new Set()
      };
    }
    
    if (!grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].seos[seoKey]) {
      grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].seos[seoKey] = {
        clicks: 0,
        registrations: 0,
        fd: 0,
        rd: 0,
        deposits: 0,
        revenueUSD: 0,
        sitesCount: 0,
        seoCount: 1,
        allSites: new Set()
      };
    }
    
    // Добавляем сеошника в множества
    if (!ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoKey)) {
      grouped[geoKey].seoSpecialists.add(seoKey);
      grouped[geoKey].geoBrands[geoBrandKey].seoSpecialists.add(seoKey);
      grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].seoSpecialists.add(seoKey);
    }
    
    // Суммируем данные на всех уровнях
    const levels = [
      grouped[geoKey].data,
      grouped[geoKey].geoBrands[geoBrandKey].data,
      grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].data,
      grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].seos[seoKey]
    ];
    
    levels.forEach(level => {
      level.clicks += item.clicks;
      level.registrations += item.registrations;
      level.fd += item.fd;
      level.rd += item.rd;
      level.deposits += item.deposits;
      level.revenueUSD += item.revenueUSD;
    });
    
    // Добавляем информацию о сайтах
    if (item.identifier && sitesData[item.identifier]) {
      const uniqueSites = new Map();
      
      sitesData[item.identifier].forEach(site => {
        const siteKey = `${site.site}-${item.identifier}`;
        if (!uniqueSites.has(siteKey)) {
          uniqueSites.set(siteKey, {
            stream: streamKey,
            site: site.site,
            seoSpecialist: site.seoSpecialist,
            identifier: item.identifier,
            clicks: site.clicks
          });
        }
      });
      
      // ИСПРАВЛЕНИЕ: проверяем что сайты еще не добавлены в этот поток
      if (!grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].sitesProcessed) {
        grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].sitesProcessed = new Set();
      }
      
      uniqueSites.forEach(siteInfo => {
        const siteStreamKey = `${siteInfo.site}-${item.identifier}-${streamKey}`;
        
        // Добавляем сайт только если он еще не был добавлен в этот поток
        if (!grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].sitesProcessed.has(siteStreamKey)) {
          grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].sites.push(siteInfo);
          grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].sitesProcessed.add(siteStreamKey);
          
          const siteKey = `${siteInfo.site}-${item.identifier}`;
          grouped[geoKey].allSites.add(siteKey);
          grouped[geoKey].geoBrands[geoBrandKey].allSites.add(siteKey);
          grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].allSites.add(siteKey);
          
          if (siteInfo.seoSpecialist === seoKey) {
            grouped[geoKey].geoBrands[geoBrandKey].streams[streamKey].seos[seoKey].allSites.add(siteKey);
          }
        }
      });
    }
  });
  
  // Устанавливаем количество сеошников и сайтов
  Object.values(grouped).forEach(geo => {
    geo.data.seoCount = geo.seoSpecialists.size;
    geo.data.sitesCount = geo.allSites.size;
    
    Object.values(geo.geoBrands).forEach(geoBrand => {
      geoBrand.data.seoCount = geoBrand.seoSpecialists.size;
      geoBrand.data.sitesCount = geoBrand.allSites.size;
      
      Object.values(geoBrand.streams).forEach(stream => {
        stream.data.seoCount = stream.seoSpecialists.size;
        stream.data.sitesCount = stream.allSites.size;
        
        Object.values(stream.seos).forEach(seo => {
          seo.sitesCount = seo.allSites ? seo.allSites.size : 0;
        });
      });
    });
  });
  
  return grouped;
}

/**
 * Группировка данных по сеошникам (ТОЧНАЯ КОПИЯ ИЗ По сеошникам.js)
 */
function groupSeoData(mainData, sitesData) {
  const grouped = {};
  
  mainData.forEach(item => {
    const seoKey = item.seoSpecialist || 'Неизвестно';
    const brandGeoKey = item.brandGeo || 'Неизвестно';
    const streamKey = item.stream || 'Неизвестно';
    
    // Инициализируем структуру сеошников
    if (!grouped[seoKey]) {
      grouped[seoKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        isExcluded: ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoKey),
        brandGeos: {}
      };
    }
    
    if (!grouped[seoKey].brandGeos[brandGeoKey]) {
      grouped[seoKey].brandGeos[brandGeoKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        streams: {}
      };
    }
    
    if (!grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey]) {
      grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey] = {
        data: {
          clicks: 0,
          registrations: 0,
          fd: 0,
          rd: 0,
          deposits: 0,
          revenueUSD: 0,
          sitesCount: 0,
          seoCount: 0
        },
        sites: []
      };
    }
    
    // Добавляем данные
    const levels = [
      grouped[seoKey].data,
      grouped[seoKey].brandGeos[brandGeoKey].data,
      grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey].data
    ];
    
    levels.forEach(level => {
      level.clicks += item.clicks;
      level.registrations += item.registrations;
      level.fd += item.fd;
      level.rd += item.rd;
      level.deposits += item.deposits;
      level.revenueUSD += item.revenueUSD;
    });
    
    // Добавляем сайты
    if (item.identifier && sitesData[item.identifier]) {
      // ИСПРАВЛЕНИЕ: проверяем что сайты еще не добавлены в этот поток
      if (!grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed) {
        grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed = new Set();
      }
      
      sitesData[item.identifier].forEach(site => {
        const siteStreamKey = `${site.site}-${item.identifier}-${streamKey}`;
        
        // Добавляем сайт только если он еще не был добавлен в этот поток
        if (!grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed.has(siteStreamKey)) {
          grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey].sites.push({
            site: site.site,
            seoSpecialist: site.seoSpecialist,
            clicks: site.clicks
          });
          grouped[seoKey].brandGeos[brandGeoKey].streams[streamKey].sitesProcessed.add(siteStreamKey);
        }
      });
    }
  });
  
  return grouped;
}

// ========================================================================
// 📋 СОЗДАНИЕ ЛИСТОВ (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
// ========================================================================

/**
 * Создать лист сводной по брендам (ТОЧНАЯ КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function createSummarySheet(tableId, groupedData, monthName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheetName = `Сводная Бренд+Гео ${monthName}`;
    
    // Удаляем лист если он уже существует
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Создаем заголовки
    createHeaders(sheet, monthName);
    
    let currentRow = 3;
    
    // Добавляем строку "Всего"
    const totalData = calculateTotalData(groupedData);
    addTotalRow(sheet, currentRow, totalData);
    currentRow += 3; // Пропускаем строки для роста
    
    // Сортируем бренды по доходу
    const sortedBrands = Object.entries(groupedData)
      .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
    
    sortedBrands.forEach(([brandKey, brandData]) => {
      // Добавляем строку бренда (всегда видимая)
      addBrandRow(sheet, currentRow, brandKey, brandData.data, totalData);
      currentRow += 3; // Пропускаем строки для роста
      
      // Запоминаем начало содержимого бренда (ПОСЛЕ строки бренда)
      const brandContentStart = currentRow;
      
      // Сортируем и обрабатываем Бренд+Гео
      const sortedBrandGeos = Object.entries(brandData.brandGeos)
        .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
      
      sortedBrandGeos.forEach(([brandGeoKey, brandGeoData]) => {
        // Добавляем строку Бренд+Гео (всегда видимая)
        addBrandGeoRow(sheet, currentRow, brandGeoKey, brandGeoData.data, brandData.data);
        currentRow += 3; // Пропускаем строки для роста
        
        // Запоминаем начало содержимого бренд+гео (ПОСЛЕ строки бренд+гео)  
        const brandGeoContentStart = currentRow;
        
        // Сортируем и добавляем детализацию
        const sortedStreams = Object.entries(brandGeoData.streams)
          .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
        
        sortedStreams.forEach(([streamKey, streamData]) => {
          addStreamRow(sheet, currentRow, streamKey, streamData.data, brandGeoData.data);
          currentRow++;
          
          // Сортируем сеошников
          const sortedSeos = Object.entries(streamData.seos)
            .sort(([,a], [,b]) => b.revenueUSD - a.revenueUSD);
          
          sortedSeos.forEach(([seoKey, seoData]) => {
            addSeoRow(sheet, currentRow, seoKey, seoData, streamData.data);
            currentRow++;
          });
          
          // Сортируем сайты
          const sortedSites = streamData.sites
            .sort((a, b) => b.clicks - a.clicks);
          
          sortedSites.forEach(site => {
            addSiteRow(sheet, currentRow, site, streamData.data);
            currentRow++;
          });
        });
        
        // Группируем содержимое Бренд+Гео (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
        if (brandGeoContentStart <= currentRow - 1) {
          groupBrandGeoContent(sheet, brandGeoContentStart, currentRow - 1);
        }
      });
      
      // Группируем содержимое бренда (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
      if (brandContentStart <= currentRow - 1) {
        groupBrandContent(sheet, brandContentStart, currentRow - 1);
      }
    });
    
    // Применяем окончательное форматирование
    finalizeSummarySheet(sheet);
    
  } catch (error) {
    console.error('Ошибка при создании сводной таблицы:', error);
    throw error;
  }
}

/**
 * Создать лист сводной по ГЕО (ТОЧНАЯ КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function createGeoSummarySheet(tableId, groupedData, monthName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheetName = `Сводная ГЕО ${monthName}`;
    
    // Удаляем существующий лист
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Создаем заголовки
    createHeaders(sheet, monthName);
    
    let currentRow = 3;
    
    // Считаем общие итоги
    const totalData = calculateTotalGeoData(groupedData);
    
    // Добавляем строку "Всего"
    addTotalRow(sheet, currentRow, totalData);
    currentRow += 3;
    
    // Обрабатываем каждое ГЕО
    Object.entries(groupedData)
      .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD)
      .forEach(([geoKey, geoData]) => {
        // Добавляем строку ГЕО (всегда видимая, ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: с флагами стран)
        const geoDisplayName = получитьСтрану(geoKey);
        addGeoRow(sheet, currentRow, geoDisplayName, geoData.data, totalData);
        currentRow += 3; // Пропускаем строки для роста
        
        // Запоминаем начало содержимого ГЕО (ПОСЛЕ строки ГЕО)
        const geoContentStart = currentRow;
        
        // Обрабатываем ГЕО+Бренд
        Object.entries(geoData.geoBrands)
          .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD)
          .forEach(([geoBrandKey, geoBrandData]) => {
            // Создаем отображаемое название ГЕО+Бренд
            const geoBrandDisplayName = `${geoDisplayName} ${geoBrandKey.split('_')[1] || ''}`;
            addGeoBrandRow(sheet, currentRow, geoBrandDisplayName, geoBrandData.data, geoData.data);
            currentRow += 3; // Пропускаем строки для роста
            
            // Запоминаем начало содержимого ГЕО+Бренд (ПОСЛЕ строки ГЕО+Бренд)
            const geoBrandContentStart = currentRow;
            
            // Добавляем детализацию
            Object.entries(geoBrandData.streams)
              .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD)
              .forEach(([streamKey, streamData]) => {
                addStreamRow(sheet, currentRow, streamKey, streamData.data, geoBrandData.data);
                currentRow++;
                
                Object.entries(streamData.seos)
                  .sort(([,a], [,b]) => b.revenueUSD - a.revenueUSD)
                  .forEach(([seoKey, seoData]) => {
                    addSeoRow(sheet, currentRow, seoKey, seoData, streamData.data);
                    currentRow++;
                  });
                
                streamData.sites
                  .sort((a, b) => b.clicks - a.clicks)
                  .forEach(site => {
                    addSiteRow(sheet, currentRow, site, streamData.data);
                    currentRow++;
                  });
              });
            
            // Группируем содержимое ГЕО+Бренд (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
            if (geoBrandContentStart <= currentRow - 1) {
              groupBrandGeoContent(sheet, geoBrandContentStart, currentRow - 1);
            }
          });
        
        // Группируем содержимое ГЕО (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
        if (geoContentStart <= currentRow - 1) {
          groupBrandContent(sheet, geoContentStart, currentRow - 1);
        }
      });
    
    // Применяем окончательное форматирование
    finalizeSummarySheet(sheet);
    
  } catch (error) {
    console.error('Ошибка при создании сводной таблицы по ГЕО:', error);
    throw error;
  }
}

/**
 * Создать лист сводной по сеошникам (ТОЧНАЯ КОПИЯ ИЗ По сеошникам.js)
 */
function createSeoSummarySheet(tableId, groupedData, monthName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheetName = `Сводная Сеошники ${monthName}`;
    
    // Удаляем существующий лист
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Создаем заголовки
    createHeaders(sheet, monthName);
    
    let currentRow = 3;
    
    // Считаем общие итоги (ВКЛЮЧАЯ ВСЕХ сеошников для единообразия с другими сводными)
    const totalData = calculateTotalData(groupedData);
    
    // Добавляем строку "Всего"
    addTotalRow(sheet, currentRow, totalData);
    currentRow += 3;
    
    // Обрабатываем сеошников
    Object.entries(groupedData)
      .sort(([,a], [,b]) => {
        // Исключенные сеошники в конец
        if (a.isExcluded && !b.isExcluded) return 1;
        if (!a.isExcluded && b.isExcluded) return -1;
        return b.data.revenueUSD - a.data.revenueUSD;
      })
      .forEach(([seoKey, seoData]) => {
        // Добавляем строку сеошника (всегда видимая)
        addSeoMainRow(sheet, currentRow, seoKey, seoData.data, totalData, seoData.isExcluded);
        currentRow += 3; // Пропускаем строки для роста
        
        // Запоминаем начало содержимого сеошника (ПОСЛЕ строки сеошника)
        const seoContentStart = currentRow;
        
        // Добавляем детализацию
        Object.entries(seoData.brandGeos)
          .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD)
          .forEach(([brandGeoKey, brandGeoData]) => {
            // Добавляем строку Бренд+Гео (всегда видимая)
            addBrandGeoRow(sheet, currentRow, brandGeoKey, brandGeoData.data, seoData.data);
            currentRow += 3; // Пропускаем строки для роста
            
            // Запоминаем начало содержимого Бренд+Гео (ПОСЛЕ строки Бренд+Гео)
            const brandGeoContentStart = currentRow;
            
            Object.entries(brandGeoData.streams)
              .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD)
              .forEach(([streamKey, streamData]) => {
                addStreamRow(sheet, currentRow, streamKey, streamData.data, brandGeoData.data);
                currentRow++;
                
                streamData.sites
                  .sort((a, b) => b.clicks - a.clicks)
                  .forEach(site => {
                    addSiteRow(sheet, currentRow, site, streamData.data);
                    currentRow++;
                  });
              });
            
            // Группируем содержимое Бренд+Гео (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
            if (brandGeoContentStart <= currentRow - 1) {
              groupBrandGeoContent(sheet, brandGeoContentStart, currentRow - 1);
            }
          });
        
        // Группируем содержимое сеошника (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
        if (seoContentStart <= currentRow - 1) {
          groupBrandContent(sheet, seoContentStart, currentRow - 1);
        }
      });
    
    // Применяем окончательное форматирование
    finalizeSummarySheet(sheet);
    
  } catch (error) {
    console.error('Ошибка при создании сводной таблицы по сеошникам:', error);
    throw error;
  }
}

// ========================================================================
// 📋 СОЗДАНИЕ ЗАГОЛОВКОВ И СТРОК (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
// ========================================================================

/**
 * Создать заголовки сводной таблицы (ТОЧНАЯ КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function createHeaders(sheet, monthName) {
  // Строка 1 - Название месяца
  sheet.getRange(1, 1).setValue(monthName);
  
  // Строка 2 - Заголовки столбцов
  const headers = [
    'Оффер', 'Количество сайтов с трафом', 'Сеошник', 'Клики', '% от общего', 'Рост', 'Рост %',
    'Реги', '% от общего', 'Рост', 'Рост %', 'FD', '% от общего', 'Рост', 'Рост %', 'RD', '% от общего', 'Рост', 'Рост %', 'Deps', '% от общего',
    'Расход / USD', 'Рост', 'Рост %', 'Выручка / USD', '% от общего', 'Рост', 'Рост %',
    'Прибыль / USD', 'Рост', 'Рост %', 'Расход / RUB', 'Выручка / RUB', 'Прибыль / RUB'
  ];
  
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(2, i + 1).setValue(headers[i]);
  }
  
  // Форматируем заголовки
  const headerRange = sheet.getRange(1, 1, 2, headers.length);
  headerRange.setBackground('#4472C4')
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
}

/**
 * Заполнить строку данными (ТОЧНАЯ КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function fillDataRow(sheet, row, data, parentData) {
  sheet.getRange(row, 2).setValue(data.sitesCount || 0); // B - Количество сайтов
  sheet.getRange(row, 3).setValue(data.seoCount || 0);   // C - Количество сеошников
  sheet.getRange(row, 4).setValue(data.clicks || 0);     // D - Клики
  
  // E - % от общего для кликов
  if (parentData && parentData.clicks > 0) {
    const percentage = (data.clicks / parentData.clicks) * 100;
    sheet.getRange(row, 5).setValue(percentage / 100);
  }
  
  sheet.getRange(row, 8).setValue(data.registrations || 0); // H - Реги
  
  // I - % от общего для реги
  if (parentData && parentData.registrations > 0) {
    const percentage = (data.registrations / parentData.registrations) * 100;
    sheet.getRange(row, 9).setValue(percentage / 100);
  }
  
  sheet.getRange(row, 12).setValue(data.fd || 0); // L - FD
  
  // M - % от общего для FD
  if (parentData && parentData.fd > 0) {
    const percentage = (data.fd / parentData.fd) * 100;
    sheet.getRange(row, 13).setValue(percentage / 100);
  }
  
  sheet.getRange(row, 16).setValue(data.rd || 0); // P - RD
  
  // Q - % от общего для RD
  if (parentData && parentData.rd > 0) {
    const percentage = (data.rd / parentData.rd) * 100;
    sheet.getRange(row, 17).setValue(percentage / 100);
  }
  
  sheet.getRange(row, 20).setValue(data.deposits || 0); // T - Deps
  
  // U - % от общего для Deps
  if (parentData && parentData.deposits > 0) {
    const percentage = (data.deposits / parentData.deposits) * 100;
    sheet.getRange(row, 21).setValue(percentage / 100);
  }
  
  sheet.getRange(row, 25).setValue(data.revenueUSD || 0); // Y - Выручка USD
  
  // Z - % от общего для выручки
  if (parentData && parentData.revenueUSD > 0) {
    const percentage = (data.revenueUSD / parentData.revenueUSD) * 100;
    sheet.getRange(row, 26).setValue(percentage / 100);
  }
  
  // Форматирование числовых значений
  const numericCells = [
    sheet.getRange(row, 2),  // Количество сайтов
    sheet.getRange(row, 3),  // Количество сеошников
    sheet.getRange(row, 4),  // Клики
    sheet.getRange(row, 8),  // Реги
    sheet.getRange(row, 12), // FD
    sheet.getRange(row, 16), // RD
    sheet.getRange(row, 20), // Deps
    sheet.getRange(row, 25)  // Выручка USD
  ];
  
  numericCells.forEach(cell => {
    cell.setNumberFormat('#,##0');
  });
  
  // Форматирование процентов
  const percentCells = [
    sheet.getRange(row, 5),  // % от общего для кликов
    sheet.getRange(row, 9),  // % от общего для реги
    sheet.getRange(row, 13), // % от общего для FD
    sheet.getRange(row, 17), // % от общего для RD
    sheet.getRange(row, 21), // % от общего для Deps
    sheet.getRange(row, 26)  // % от общего для выручки
  ];
  
  percentCells.forEach(cell => {
    cell.setNumberFormat('0.00%');
  });
  
  // Специальное форматирование для денежных значений
  sheet.getRange(row, 25).setNumberFormat('$#,##0.00'); // Выручка USD
}

/**
 * Добавить строку "Всего" (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА + ЦВЕТА ИЗ КОНСТАНТ)
 */
function addTotalRow(sheet, row, data) {
  sheet.getRange(row, 1).setValue('Всего');
  fillDataRow(sheet, row, data);
  
  // Цветовое форматирование для строки "Всего" (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const totalRange = sheet.getRange(row, 1, 1, 34);
  totalRange.setBackground(ЦВЕТА.ВСЕГО).setFontWeight('bold');
  
  sheet.getRange(row + 1, 1).setValue('Рост / Падение %');
  sheet.getRange(row + 2, 1).setValue('В количестве');
  
  // Форматирование строк роста
  sheet.getRange(row + 1, 1, 2, 34).setBackground('#f8f9fa').setFontStyle('italic');
}

/**
 * Добавить строку бренда (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА + ЦВЕТА ИЗ КОНСТАНТ)
 */
function addBrandRow(sheet, row, brandName, data, totalData) {
  sheet.getRange(row, 1).setValue(brandName);
  fillDataRow(sheet, row, data, totalData);
  
  // Цветовое форматирование для строки бренда (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const brandRange = sheet.getRange(row, 1, 1, 34);
  brandRange.setBackground(ЦВЕТА.БРЕНДЫ).setFontWeight('bold');
  
  sheet.getRange(row + 1, 1).setValue('Рост / Падение');
  sheet.getRange(row + 2, 1).setValue('В количестве');
  
  // Форматирование строк роста
  sheet.getRange(row + 1, 1, 2, 34).setBackground('#f8f9fa').setFontStyle('italic');
}

/**
 * Добавить строку бренд+гео (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА + ЦВЕТА ИЗ КОНСТАНТ)
 */
function addBrandGeoRow(sheet, row, brandGeoName, data, parentData) {
  sheet.getRange(row, 1).setValue(brandGeoName);
  fillDataRow(sheet, row, data, parentData);
  
  // Цветовое форматирование для строки бренд+гео (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const brandGeoRange = sheet.getRange(row, 1, 1, 34);
  brandGeoRange.setBackground(ЦВЕТА.БРЕНД_ГЕО).setFontWeight('bold');
  
  sheet.getRange(row + 1, 1).setValue('Рост / Падение');
  sheet.getRange(row + 2, 1).setValue('В количестве');
  
  // Форматирование строк роста
  sheet.getRange(row + 1, 1, 2, 34).setBackground('#f8f9fa').setFontStyle('italic');
}

/**
 * Добавить строку ГЕО (ЦВЕТА ИЗ КОНСТАНТ)
 */
function addGeoRow(sheet, row, geoName, data, totalData) {
  sheet.getRange(row, 1).setValue(geoName); // Уже содержит флаг
  fillDataRow(sheet, row, data, totalData);
  
  // Цветовое форматирование для строки ГЕО (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const geoRange = sheet.getRange(row, 1, 1, 34);
  geoRange.setBackground(ЦВЕТА.ГЕО).setFontWeight('bold');
  
  sheet.getRange(row + 1, 1).setValue('Рост / Падение');
  sheet.getRange(row + 2, 1).setValue('В количестве');
  
  // Форматирование строк роста
  sheet.getRange(row + 1, 1, 2, 34).setBackground('#f8f9fa').setFontStyle('italic');
}

/**
 * Добавить строку гео+бренд (ЦВЕТА ИЗ КОНСТАНТ)
 */
function addGeoBrandRow(sheet, row, geoBrandName, data, parentData) {
  sheet.getRange(row, 1).setValue(geoBrandName);
  fillDataRow(sheet, row, data, parentData);
  
  // Цветовое форматирование для строки гео+бренд (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const geoBrandRange = sheet.getRange(row, 1, 1, 34);
  geoBrandRange.setBackground(ЦВЕТА.ГЕО_БРЕНД).setFontWeight('bold');
  
  sheet.getRange(row + 1, 1).setValue('Рост / Падение');
  sheet.getRange(row + 2, 1).setValue('В количестве');
  
  // Форматирование строк роста
  sheet.getRange(row + 1, 1, 2, 34).setBackground('#f8f9fa').setFontStyle('italic');
}

/**
 * Добавить строку потока (ЦВЕТА ИЗ КОНСТАНТ)
 */
function addStreamRow(sheet, row, streamName, data, parentData) {
  sheet.getRange(row, 1).setValue(streamName);
  fillDataRow(sheet, row, data, parentData);
  
  // Цветовое форматирование для строки потока (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const streamRange = sheet.getRange(row, 1, 1, 34);
  streamRange.setBackground(ЦВЕТА.ПОТОКИ);
}

/**
 * Добавить строку сеошника (ЦВЕТА ИЗ КОНСТАНТ)
 */
function addSeoRow(sheet, row, seoName, data, parentData) {
  sheet.getRange(row, 1).setValue(seoName);
  fillDataRow(sheet, row, data, parentData);
  
  // Цветовое форматирование для строки сеошника (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const seoRange = sheet.getRange(row, 1, 1, 34);
  if (ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoName)) {
    seoRange.setBackground(ЦВЕТА.ИСКЛЮЧЕННЫЕ).setFontStyle('italic');
  } else {
    seoRange.setBackground(ЦВЕТА.СЕОШНИКИ);
  }
}

/**
 * Добавить главную строку сеошника (ЦВЕТА ИЗ КОНСТАНТ)
 */
function addSeoMainRow(sheet, row, seoName, data, totalData, isExcluded) {
  sheet.getRange(row, 1).setValue(seoName);
  fillDataRow(sheet, row, data, totalData);
  
  // Цветовое форматирование (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const seoRange = sheet.getRange(row, 1, 1, 34);
  if (isExcluded) {
    seoRange.setBackground(ЦВЕТА.ИСКЛЮЧЕННЫЕ).setFontStyle('italic');
  } else {
    seoRange.setBackground(ЦВЕТА.СЕОШНИКИ).setFontWeight('bold');
  }
  
  sheet.getRange(row + 1, 1).setValue('Рост / Падение');
  sheet.getRange(row + 2, 1).setValue('В количестве');
  
  // Форматирование строк роста
  sheet.getRange(row + 1, 1, 2, 34).setBackground('#f8f9fa').setFontStyle('italic');
}

/**
 * Добавить строку сайта (ЦВЕТА ИЗ КОНСТАНТ)
 */
function addSiteRow(sheet, row, site, parentData) {
  // ИСПРАВЛЕНО: точная копия из оригинального кода
  sheet.getRange(row, 1).setValue(site.stream);        // A - Поток
  sheet.getRange(row, 2).setValue(site.site);          // B - Сайт
  sheet.getRange(row, 3).setValue(site.seoSpecialist); // C - Сеошник
  sheet.getRange(row, 4).setValue(site.clicks);        // D - Клики
  
  // Расчитываем % от сеошника для кликов
  if (parentData && parentData.clicks > 0) {
    const percentage = (site.clicks / parentData.clicks) * 100;
    sheet.getRange(row, 5).setValue(percentage / 100); // E - % от общего (в формате процента)
  }
  
  sheet.getRange(row, 6).setValue(site.identifier);   // F - Идентификатор (в сгруппированном столбце)
  
  // Цветовое форматирование для строки сайта (ИСПОЛЬЗУЕМ КОНСТАНТЫ)
  const siteRange = sheet.getRange(row, 1, 1, 34);
  siteRange.setBackground(ЦВЕТА.САЙТЫ);
  
  // Форматирование кликов как числа
  sheet.getRange(row, 4).setNumberFormat('#,##0');
  // Форматирование процента
  sheet.getRange(row, 5).setNumberFormat('0.00%');
}

// ========================================================================
// 🔧 ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ РАСЧЕТОВ (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛА)
// ========================================================================

/**
 * Рассчитать общие данные для сводной по сеошникам (ИСКЛЮЧАЕТ исключенных сеошников)
 */
function calculateSeoTotalDataFiltered(groupedData) {
  const total = {
    clicks: 0,
    registrations: 0,
    fd: 0,
    rd: 0,
    deposits: 0,
    revenueUSD: 0,
    sitesCount: 0,
    seoCount: 0
  };
  
  const allSeoSpecialists = new Set();
  const allSites = new Set();
  
  Object.entries(groupedData).forEach(([seoName, group]) => {
    // ИСКЛЮЧАЕМ исключенных сеошников из расчета итогов
    if (!ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoName)) {
      total.clicks += group.data.clicks;
      total.registrations += group.data.registrations;
      total.fd += group.data.fd;
      total.rd += group.data.rd;
      total.deposits += group.data.deposits;
      total.revenueUSD += group.data.revenueUSD;
      
      // Подсчет сайтов и сеошников только для неисключенных
      if (group.brandGeos) {
        Object.values(group.brandGeos).forEach(brandGeo => {
          if (brandGeo.allSites) {
            brandGeo.allSites.forEach(site => allSites.add(site));
          }
        });
      }
      allSeoSpecialists.add(seoName);
    }
  });
  
  total.sitesCount = allSites.size;
  total.seoCount = allSeoSpecialists.size;
  
  return total;
}

/**
 * Рассчитать общие данные (ТОЧНАЯ КОПИЯ ИЗ ГЕО + бренд+гео.js)
 */
function calculateTotalData(groupedData) {
  const total = {
    clicks: 0,
    registrations: 0,
    fd: 0,
    rd: 0,
    deposits: 0,
    revenueUSD: 0,
    sitesCount: 0,
    seoCount: 0
  };
  
  const allSeoSpecialists = new Set();
  const allSites = new Set();
  
  Object.values(groupedData).forEach(group => {
    total.clicks += group.data.clicks;
    total.registrations += group.data.registrations;
    total.fd += group.data.fd;
    total.rd += group.data.rd;
    total.deposits += group.data.deposits;
    total.revenueUSD += group.data.revenueUSD;
    
    // Собираем все уникальные сеошники и сайты (исключаем определенных сеошников из подсчета)
    if (group.seoSpecialists) {
      group.seoSpecialists.forEach(seo => {
        if (!ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seo)) {
          allSeoSpecialists.add(seo);
        }
      });
    }
    if (group.allSites) {
      group.allSites.forEach(site => allSites.add(site));
    }
  });
  
  total.seoCount = allSeoSpecialists.size;
  total.sitesCount = allSites.size;
  
  return total;
}

/**
 * Рассчитать общие данные для ГЕО
 */
function calculateTotalGeoData(groupedData) {
  const total = {
    clicks: 0,
    registrations: 0,
    fd: 0,
    rd: 0,
    deposits: 0,
    revenueUSD: 0,
    sitesCount: 0,
    seoCount: 0
  };
  
  const allSeoSpecialists = new Set();
  const allSites = new Set();
  
  Object.values(groupedData).forEach(group => {
    total.clicks += group.data.clicks;
    total.registrations += group.data.registrations;
    total.fd += group.data.fd;
    total.rd += group.data.rd;
    total.deposits += group.data.deposits;
    total.revenueUSD += group.data.revenueUSD;
    
    if (group.seoSpecialists) {
      group.seoSpecialists.forEach(seo => {
        if (!ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seo)) {
          allSeoSpecialists.add(seo);
        }
      });
    }
    if (group.allSites) {
      group.allSites.forEach(site => allSites.add(site));
    }
  });
  
  total.seoCount = allSeoSpecialists.size;
  total.sitesCount = allSites.size;
  
  return total;
}

/**
 * Рассчитать общие данные для сеошников (исключая исключенных)
 */
function calculateTotalSeoData(groupedData) {
  const total = {
    clicks: 0,
    registrations: 0,
    fd: 0,
    rd: 0,
    deposits: 0,
    revenueUSD: 0,
    sitesCount: 0,
    seoCount: 0
  };
  
  Object.values(groupedData).forEach(seoData => {
    if (!seoData.isExcluded) {
      total.clicks += seoData.data.clicks;
      total.registrations += seoData.data.registrations;
      total.fd += seoData.data.fd;
      total.rd += seoData.data.rd;
      total.deposits += seoData.data.deposits;
      total.revenueUSD += seoData.data.revenueUSD;
    }
  });
  
  return total;
}

// ========================================================================
// 📊 ГРУППИРОВКА СТРОК И СТОЛБЦОВ (ТОЧНАЯ КОПИЯ + ИСПРАВЛЕНИЯ ПОЛЬЗОВАТЕЛЯ)
// ========================================================================

/**
 * Группировать содержимое бренда (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
 */
function groupBrandContent(sheet, startRow, endRow) {
  try {
    if (startRow <= endRow) { // ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: было <, стало <=
      const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 34);
      range.shiftRowGroupDepth(1);
      
      // Сворачиваем группу по умолчанию
      const group = sheet.getRowGroup(startRow, 1);
      if (group) {
        group.collapse();
      }
    }
  } catch (error) {
    console.log('Ошибка группировки содержимого бренда:', error);
  }
}

/**
 * Группировать содержимое бренд+гео (ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: даже 1 строку)
 */
function groupBrandGeoContent(sheet, startRow, endRow) {
  try {
    if (startRow <= endRow) { // ИСПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ: было <, стало <=
      const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 34);
      range.shiftRowGroupDepth(1);
      
      // Сворачиваем группу по умолчанию
      const group = sheet.getRowGroup(startRow, 1);
      if (group) {
        group.collapse();
      }
    }
  } catch (error) {
    console.log('Ошибка группировки содержимого бренд+гео:', error);
  }
}

/**
 * Группировать столбцы роста (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
 */
function groupGrowthColumns(sheet) {
  try {
    // Группируем столбцы по парам согласно ИСПРАВЛЕННОЙ конфигурации
    const growthColumnGroups = [
      [6, 7],   // F-G (рост кликов)
      [10, 11], // J-K (рост регов)
      [14, 15], // N-O (рост FD)
      [18, 19], // R-S (рост RD) ✅ ИСПРАВЛЕНО
      [23, 24], // W-X (рост расходов) ✅ ДОБАВЛЕНО
      [27, 28], // AA-AB (рост выручки) ✅ ИСПРАВЛЕНО 
      [30, 31]  // AD-AE (рост прибыли) ✅ ИСПРАВЛЕНО
    ];
    
    growthColumnGroups.forEach(([start, end]) => {
      const range = sheet.getRange(1, start, sheet.getLastRow(), end - start + 1);
      range.shiftColumnGroupDepth(1);
      
      try {
        sheet.getColumnGroup(start, 1).collapse();
      } catch (e) {
        // Игнорируем ошибки группировки столбцов
      }
    });
    
  } catch (error) {
    console.error('Ошибка группировки столбцов роста:', error);
  }
}

/**
 * Установить компактные размеры строк (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
 */
function setCompactRowSizes(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow > 0) {
      sheet.setRowHeights(1, lastRow, 21); // Компактная высота строк
    }
  } catch (error) {
    console.error('Ошибка установки размеров строк:', error);
  }
}

/**
 * Общее форматирование листа (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
 */
function formatSheet(sheet) {
  try {
    // Замораживаем первую строку и первый столбец
    sheet.setFrozenRows(2); // 2 строки для заголовков
    sheet.setFrozenColumns(1);
    
    // Устанавливаем ширину столбцов
    sheet.setColumnWidth(1, 250); // Название - широкий столбец
    
    // Остальные столбцы стандартной ширины
    for (let col = 2; col <= sheet.getLastColumn(); col++) {
      sheet.setColumnWidth(col, 80);
    }
    
  } catch (error) {
    console.error('Ошибка общего форматирования листа:', error);
  }
}

/**
 * Финализировать форматирование листа (ТОЧНАЯ КОПИЯ ИЗ ОРИГИНАЛЬНОГО КОДА)
 */
function finalizeSummarySheet(sheet) {
  // Группируем столбцы роста
  groupGrowthColumns(sheet);
  
  // Устанавливаем компактные размеры строк
  setCompactRowSizes(sheet);
  
  // Форматируем таблицу
  formatSheet(sheet);
}

// ========================================================================
// 🚀 BATCH-ОБРАБОТКА ДЛЯ УСКОРЕНИЯ (НОВОЕ)
// ========================================================================

/**
 * Batch-запись данных в лист (ускоряет в 10-50 раз)
 */
class BatchWriter {
  constructor(sheet) {
    this.sheet = sheet;
    this.data = []; // Двумерный массив для setValues()
    this.formats = []; // Массив форматов
    this.startRow = 1;
    this.startCol = 1;
    this.maxRow = 0;
    this.maxCol = 0;
  }
  
  /**
   * Установить значение для записи в batch
   */
  setValue(row, col, value) {
    // Убеждаемся что массив data достаточно большой
    while (this.data.length < row) {
      this.data.push([]);
    }
    while (this.data[row - 1].length < col) {
      this.data[row - 1].push('');
    }
    
    this.data[row - 1][col - 1] = value;
    this.maxRow = Math.max(this.maxRow, row);
    this.maxCol = Math.max(this.maxCol, col);
  }
  
  /**
   * Установить значения для целой строки
   */
  setRowValues(row, values) {
    for (let i = 0; i < values.length; i++) {
      this.setValue(row, i + 1, values[i]);
    }
  }
  
  /**
   * Записать все накопленные данные одним вызовом
   */
  flush() {
    if (this.maxRow > 0 && this.maxCol > 0) {
      console.log(`📝 Batch-запись: ${this.maxRow} строк x ${this.maxCol} столбцов`);
      
      // Выравниваем все строки по максимальной длине
      for (let i = 0; i < this.data.length; i++) {
        while (this.data[i].length < this.maxCol) {
          this.data[i].push('');
        }
      }
      
      // Записываем все данные одним вызовом
      const range = this.sheet.getRange(1, 1, this.maxRow, this.maxCol);
      range.setValues(this.data);
    }
  }
  
  /**
   * Очистить накопленные данные
   */
  clear() {
    this.data = [];
    this.formats = [];
    this.maxRow = 0;
    this.maxCol = 0;
  }
}

/**
 * Batch-форматирование (ускоряет форматирование в разы)
 */
class BatchFormatter {
  constructor(sheet) {
    this.sheet = sheet;
    this.operations = []; // Массив операций форматирования
  }
  
  /**
   * Добавить операцию форматирования строки
   */
  formatRow(row, backgroundColor, fontColor = '#000000', fontWeight = 'normal') {
    this.operations.push({
      type: 'row',
      row: row,
      backgroundColor: backgroundColor,
      fontColor: fontColor,
      fontWeight: fontWeight
    });
  }
  
  /**
   * Добавить операцию форматирования диапазона
   */
  formatRange(startRow, startCol, numRows, numCols, backgroundColor, fontColor = '#000000') {
    this.operations.push({
      type: 'range',
      startRow: startRow,
      startCol: startCol,
      numRows: numRows,
      numCols: numCols,
      backgroundColor: backgroundColor,
      fontColor: fontColor
    });
  }
  
  /**
   * Применить все накопленные операции форматирования
   */
  flush() {
    console.log(`🎨 Batch-форматирование: ${this.operations.length} операций`);
    
    // Группируем операции по типу для оптимизации
    const rowOps = this.operations.filter(op => op.type === 'row');
    const rangeOps = this.operations.filter(op => op.type === 'range');
    
    // Применяем форматирование строк
    rowOps.forEach(op => {
      const range = this.sheet.getRange(op.row, 1, 1, this.sheet.getLastColumn());
      range.setBackground(op.backgroundColor);
      if (op.fontColor !== '#000000') range.setFontColor(op.fontColor);
      if (op.fontWeight !== 'normal') range.setFontWeight(op.fontWeight);
    });
    
    // Применяем форматирование диапазонов
    rangeOps.forEach(op => {
      const range = this.sheet.getRange(op.startRow, op.startCol, op.numRows, op.numCols);
      range.setBackground(op.backgroundColor);
      if (op.fontColor !== '#000000') range.setFontColor(op.fontColor);
    });
  }
  
  /**
   * Очистить накопленные операции
   */
  clear() {
    this.operations = [];
  }
}

/**
 * ПОЛНАЯ ВЕРСИЯ БЕЗ ОГРАНИЧЕНИЙ - ОДИН ВЫЗОВ setValues() для всех данных (ВСЯ ИНФОРМАЦИЯ)
 */
function createSummarySheetFast(tableId, groupedData, monthName) {
  console.log('📊 ПОЛНАЯ ВЕРСИЯ БЕЗ ОГРАНИЧЕНИЙ: ВСЯ ИНФОРМАЦИЯ + ОДИН ВЫЗОВ setValues()');
  
  return executeWithRetry(() => {
    const startTime = new Date();
    const spreadsheet = SpreadsheetApp.openById(tableId);
    const sheetName = `Сводная Бренд+Гео ${monthName}`;
    
    // Удаляем существующий лист
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // СОБИРАЕМ ВСЕ ДАННЫЕ В ОДИН МАССИВ для записи одним вызовом
    const allData = [];
    
    // 1. Заголовки
    allData.push([monthName, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    allData.push(['Оффер', 'Количество сайтов с трафом', 'Сеошник', 'Клики', '% от общего', 'Рост', 'Рост %',
      'Реги', '% от общего', 'Рост', 'Рост %', 'FD', '% от общего', 'Рост', 'Рост %', 'RD', '% от общего', 'Рост', 'Рост %', 'Deps', '% от общего',
      'Расход / USD', 'Рост', 'Рост %', 'Выручка / USD', '% от общего', 'Рост', 'Рост %',
      'Прибыль / USD', 'Рост', 'Рост %', 'Расход / RUB', 'Выручка / RUB', 'Прибыль / RUB']);
    
    // 2. Строка "Всего" (ВКЛЮЧАЕТ ВСЕХ сеошников для единообразия с другими сводными)
    const totalData = calculateTotalData(groupedData);
    allData.push(['Всего', totalData.sitesCount || 0, totalData.seoCount || 0, totalData.clicks || 0, 0, 0, 0,
      totalData.registrations || 0, 0, 0, 0, totalData.fd || 0, 0, 0, 0, totalData.rd || 0, 0, 0, 0, totalData.deposits || 0, 0,
      0, 0, 0, totalData.revenueUSD || 0, 0, 0, 0, (totalData.revenueUSD || 0) - (totalData.expenseUSD || 0), 0, 0,
      (totalData.expenseUSD || 0) * 95, (totalData.revenueUSD || 0) * 95, ((totalData.revenueUSD || 0) - (totalData.expenseUSD || 0)) * 95]);
    allData.push(['Рост / Падение %', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // 3. ВСЕ ДАННЫЕ БЕЗ ОГРАНИЧЕНИЙ - ПОЛНАЯ КАРТИНА
    const sortedBrands = Object.entries(groupedData)
      .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
      // УБРАЛ ВСЕ ОГРАНИЧЕНИЯ - показываем ВСЕХ брендов
    
    sortedBrands.forEach(([brandName, brandData]) => {
      // Строка бренда
      allData.push([brandName, brandData.data.sitesCount || 0, brandData.data.seoCount || 0, brandData.data.clicks || 0,
        0, 0, 0, brandData.data.registrations || 0, 0, 0, 0, brandData.data.fd || 0, 0, 0, 0, brandData.data.rd || 0, 0, 0, 0,
        brandData.data.deposits || 0, 0, brandData.data.expenseUSD || 0, 0, 0, brandData.data.revenueUSD || 0, 0, 0, 0,
        (brandData.data.revenueUSD || 0) - (brandData.data.expenseUSD || 0), 0, 0,
        (brandData.data.expenseUSD || 0) * 95, (brandData.data.revenueUSD || 0) * 95, ((brandData.data.revenueUSD || 0) - (brandData.data.expenseUSD || 0)) * 95]);
      allData.push(['Рост / Падение', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      
      // ВСЕ Бренд+Гео БЕЗ ОГРАНИЧЕНИЙ
      const sortedBrandGeos = Object.entries(brandData.brandGeos)
        .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
        // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ гео
      
      sortedBrandGeos.forEach(([brandGeoName, brandGeoData]) => {
        // Строка Бренд+Гео
        allData.push([brandGeoName, brandGeoData.data.sitesCount || 0, brandGeoData.data.seoCount || 0, brandGeoData.data.clicks || 0,
          0, 0, 0, brandGeoData.data.registrations || 0, 0, 0, 0, brandGeoData.data.fd || 0, 0, 0, 0, brandGeoData.data.rd || 0, 0, 0, 0,
          brandGeoData.data.deposits || 0, 0, brandGeoData.data.expenseUSD || 0, 0, 0, brandGeoData.data.revenueUSD || 0, 0, 0, 0,
          (brandGeoData.data.revenueUSD || 0) - (brandGeoData.data.expenseUSD || 0), 0, 0,
          (brandGeoData.data.expenseUSD || 0) * 95, (brandGeoData.data.revenueUSD || 0) * 95, ((brandGeoData.data.revenueUSD || 0) - (brandGeoData.data.expenseUSD || 0)) * 95]);
        allData.push(['Рост / Падение', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        // ВСЕ ПОТОКИ БЕЗ ОГРАНИЧЕНИЙ
        const sortedStreams = Object.entries(brandGeoData.streams || {})
          .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
          // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ потоки
          
        sortedStreams.forEach(([streamName, streamData]) => {
          allData.push([`  🔸 ${streamName}`, streamData.data.sitesCount || 0, streamData.data.seoCount || 0, streamData.data.clicks || 0,
            0, 0, 0, streamData.data.registrations || 0, 0, 0, 0, streamData.data.fd || 0, 0, 0, 0, streamData.data.rd || 0, 0, 0, 0,
            streamData.data.deposits || 0, 0, streamData.data.expenseUSD || 0, 0, 0, streamData.data.revenueUSD || 0, 0, 0, 0,
            (streamData.data.revenueUSD || 0) - (streamData.data.expenseUSD || 0), 0, 0,
            (streamData.data.expenseUSD || 0) * 95, (streamData.data.revenueUSD || 0) * 95, ((streamData.data.revenueUSD || 0) - (streamData.data.expenseUSD || 0)) * 95]);
          
          // ВСЕ СЕОШНИКИ БЕЗ ОГРАНИЧЕНИЙ
          const sortedSeos = Object.entries(streamData.seos || {})
            .sort(([,a], [,b]) => b.revenueUSD - a.revenueUSD);
            // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕХ сеошников
            
          sortedSeos.forEach(([seoName, seoData]) => {
            if (!ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoName)) {
              allData.push([`    👤 ${seoName}`, 1, 1, seoData.clicks || 0,
                0, 0, 0, seoData.registrations || 0, 0, 0, 0, seoData.fd || 0, 0, 0, 0, seoData.rd || 0, 0, 0, 0,
                seoData.deposits || 0, 0, seoData.expenseUSD || 0, 0, 0, seoData.revenueUSD || 0, 0, 0, 0,
                (seoData.revenueUSD || 0) - (seoData.expenseUSD || 0), 0, 0,
                (seoData.expenseUSD || 0) * 95, (seoData.revenueUSD || 0) * 95, ((seoData.revenueUSD || 0) - (seoData.expenseUSD || 0)) * 95]);
            }
          });
          
          // ВСЕ САЙТЫ БЕЗ ОГРАНИЧЕНИЙ
          if (streamData.sites && streamData.sites.length > 0) {
            const topSites = streamData.sites
              .sort((a, b) => (b.clicks || 0) - (a.clicks || 0));
              // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ сайты
              
            topSites.forEach(site => {
              allData.push([`      🌐 ${site.domain}`, 1, 1, site.clicks || 0,
                0, 0, 0, site.registrations || 0, 0, 0, 0, site.fd || 0, 0, 0, 0, site.rd || 0, 0, 0, 0,
                site.deposits || 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
            });
          }
        });
      });
    });
    
    // ЗАПИСЫВАЕМ ВСЕ ДАННЫЕ ОДНИМ ВЫЗОВОМ (максимально быстро)
    if (allData.length > 0) {
      sheet.getRange(1, 1, allData.length, 34).setValues(allData);
    }
    
    // Минимальное структурное форматирование
    sheet.setFrozenRows(2);
    sheet.setColumnWidth(1, 200);
    
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    console.log(`✅ ПОЛНАЯ Бренд+Гео сводная создана за ${executionTime.toFixed(2)} секунд (${allData.length} строк)`);
    console.log('⚡ ОДИН ВЫЗОВ setValues(): максимальная скорость + минимум API calls');
    console.log('📊 ВСЯ ИНФОРМАЦИЯ БЕЗ ОГРАНИЧЕНИЙ: ВСЕ бренды × ВСЕ гео × ВСЕ потоки × ВСЕ сеошники × ВСЕ сайты');
  });
}

/**
 * Быстрые функции добавления строк с batch-записью
 */
function addTotalRowFast(writer, formatter, row, data) {
  const values = ['Всего', data.sitesCount || 0, data.seoCount || 0, data.clicks || 0];
  writer.setRowValues(row, values);
  formatter.formatRow(row, '#d9e1f2', '#000000', 'bold');
}

function addBrandRowFast(writer, formatter, row, brandName, data, totalData) {
  const values = [brandName, data.sitesCount || 0, data.seoCount || 0, data.clicks || 0];
  writer.setRowValues(row, values);
  formatter.formatRow(row, '#d9d2e9', '#000000', 'bold');
}

function addBrandGeoRowFast(writer, formatter, row, brandGeoName, data, parentData) {
  const values = [brandGeoName, data.sitesCount || 0, data.seoCount || 0, data.clicks || 0];
  writer.setRowValues(row, values);
  formatter.formatRow(row, '#f0ecf5', '#000000', 'bold');
}

function addStreamRowFast(writer, formatter, row, streamName, data, parentData) {
  const values = [`Сводная / ${streamName}`, data.sitesCount || 0, data.seoCount || 0, data.clicks || 0];
  writer.setRowValues(row, values);
  formatter.formatRow(row, '#f4e6ed');
}

function addSeoRowFast(writer, formatter, row, seoName, data, parentData) {
  const values = [seoName, data.sitesCount || 0, 1, data.clicks || 0];
  writer.setRowValues(row, values);
  
  const bgColor = ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoName) ? '#E8E8E8' : '#D9DAD6';
  formatter.formatRow(row, bgColor);
}

function addSiteRowFast(writer, formatter, row, site, parentData) {
  const values = [site.stream, site.site, site.seoSpecialist, site.clicks];
  writer.setRowValues(row, values);
  formatter.formatRow(row, '#F3E3EB');
}

/**
 * Быстрое создание сводной ГЕО (БЕЗ ФОРМАТИРОВАНИЯ)
 */
function createGeoSummaryReportFast(monthName, groupedGeoData) {
  console.log('⚡ Быстрое создание сводной ГЕО (только данные)...');
  const startTime = new Date();
  
  try {
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    const sheetName = `Сводная ГЕО ${monthName}`;
    
    // Удаляем существующий лист
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Используем batch-writer только для данных
    const writer = new BatchWriter(sheet);
    
    let currentRow = 1;
    
    // Заголовки
    const headers = [
      'ГЕО', 'Количество сайтов с трафом', 'Сеошник', 'Клики', '% от общего', 'Рост', 'Рост %',
      'Реги', '% от общего', 'Рост', 'Рост %', 'FD', '% от общего', 'Рост', 'Рост %', 'RD', '% от общего', 'Рост', 'Рост %', 'Deps', '% от общего',
      'Расход / USD', 'Рост', 'Рост %', 'Выручка / USD', '% от общего', 'Рост', 'Рост %',
      'Прибыль / USD', 'Рост', 'Рост %', 'Расход / RUB', 'Выручка / RUB', 'Прибыль / RUB'
    ];
    
    writer.setValue(1, 1, monthName);
    writer.setRowValues(2, headers);
    currentRow = 3;
    
    // Добавляем данные (без детального форматирования)
    const sortedGeos = Object.entries(groupedGeoData || {})
      .sort(([,a], [,b]) => (b.data?.revenueUSD || 0) - (a.data?.revenueUSD || 0));
    
    sortedGeos.forEach(([geoKey, geoData]) => {
      // Добавляем только основные данные без сложного форматирования
      const geoName = получитьСтрану(geoKey);
      const values = [
        geoName,
        geoData.data?.sitesCount || 0,
        geoData.data?.seoCount || 0,
        geoData.data?.clicks || 0,
        0, // % от общего - заполним позже если нужно
        0, // Рост
        0, // Рост %
        geoData.data?.registrations || 0
        // ... можно добавить больше полей по необходимости
      ];
      writer.setRowValues(currentRow, values);
      currentRow++;
    });
    
    // Записываем все данные одним batch
    writer.flush();
    
    // Только структурное форматирование
    finalizeSummarySheetStructureOnly(sheet);
    
    const endTime = new Date();
    console.log(`✅ ГЕО сводная создана за ${(endTime - startTime) / 1000} секунд (только данные)`);
    
  } catch (error) {
    console.error('❌ Ошибка быстрого создания ГЕО сводной:', error);
    throw error;
  }
}

/**
 * Быстрое создание сводной Сеошники (БЕЗ ФОРМАТИРОВАНИЯ)
 */
function createSeoSummaryReportFast(monthName, groupedSeoData) {
  console.log('⚡ Быстрое создание сводной Сеошники (только данные)...');
  const startTime = new Date();
  
  try {
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    const sheetName = `Сводная Сеошники ${monthName}`;
    
    // Удаляем существующий лист
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Используем batch-writer только для данных
    const writer = new BatchWriter(sheet);
    
    let currentRow = 1;
    
    // Заголовки
    const headers = [
      'Сеошник', 'Количество сайтов с трафом', 'Сеошник', 'Клики', '% от общего', 'Рост', 'Рост %',
      'Реги', '% от общего', 'Рост', 'Рост %', 'FD', '% от общего', 'Рост', 'Рост %', 'RD', '% от общего', 'Рост', 'Рост %', 'Deps', '% от общего',
      'Расход / USD', 'Рост', 'Рост %', 'Выручка / USD', '% от общего', 'Рост', 'Рост %',
      'Прибыль / USD', 'Рост', 'Рост %', 'Расход / RUB', 'Выручка / RUB', 'Прибыль / RUB'
    ];
    
    writer.setValue(1, 1, monthName);
    writer.setRowValues(2, headers);
    currentRow = 3;
    
    // Добавляем данные SEO (без детального форматирования)
    const sortedSeos = Object.entries(groupedSeoData || {})
      .sort(([,a], [,b]) => (b.revenueUSD || 0) - (a.revenueUSD || 0));
    
    sortedSeos.forEach(([seoKey, seoData]) => {
      const seoIcon = ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoKey) ? '⚠️' : '👤';
      const values = [
        `${seoIcon} ${seoKey}`,
        seoData.sitesCount || 0,
        1, // количество сеошников = 1
        seoData.clicks || 0,
        0, // % от общего
        0, // Рост
        0, // Рост %
        seoData.registrations || 0
        // ... можно добавить больше полей
      ];
      writer.setRowValues(currentRow, values);
      currentRow++;
    });
    
    // Записываем все данные одним batch
    writer.flush();
    
    // Только структурное форматирование
    finalizeSummarySheetStructureOnly(sheet);
    
    const endTime = new Date();
    console.log(`✅ Сеошники сводная создана за ${(endTime - startTime) / 1000} секунд (только данные)`);
    
  } catch (error) {
    console.error('❌ Ошибка быстрого создания сводной Сеошники:', error);
    throw error;
  }
}

/**
 * Retry-функция для защиты от timeout
 */
function executeWithRetry(func, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`⚡ Попытка ${attempt}/${maxRetries}...`);
      return func();
    } catch (error) {
      console.error(`❌ Попытка ${attempt} неудачна:`, error.toString());
      
      if (attempt === maxRetries) {
        throw error;
      }
      
      // Пауза между попытками
      console.log('⏳ Пауза 2 секунды перед повтором...');
      Utilities.sleep(2000);
    }
  }
}

/**
 * ПОЛНАЯ ГЕО БЕЗ ОГРАНИЧЕНИЙ - ОДИН ВЫЗОВ setValues() (ВСЯ ИНФОРМАЦИЯ)
 */
function createGeoSummarySimple(monthName, groupedGeoData) {
  console.log('📊 ПОЛНАЯ ГЕО БЕЗ ОГРАНИЧЕНИЙ: ВСЯ ИНФОРМАЦИЯ + ОДИН ВЫЗОВ setValues()');
  
  return executeWithRetry(() => {
    const startTime = new Date();
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    const sheetName = `Сводная ГЕО ${monthName}`;
    
    // Удаляем старый лист если есть
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // СОБИРАЕМ ВСЕ ДАННЫЕ В ОДИН МАССИВ для записи одним вызовом
    const allData = [];
    
    // 1. Заголовки
    allData.push([monthName, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    allData.push(['ГЕО', 'Количество сайтов с трафом', 'Сеошник', 'Клики', '% от общего', 'Рост', 'Рост %',
      'Реги', '% от общего', 'Рост', 'Рост %', 'FD', '% от общего', 'Рост', 'Рост %', 'RD', '% от общего', 'Рост', 'Рост %', 'Deps', '% от общего',
      'Расход / USD', 'Рост', 'Рост %', 'Выручка / USD', '% от общего', 'Рост', 'Рост %',
      'Прибыль / USD', 'Рост', 'Рост %', 'Расход / RUB', 'Выручка / RUB', 'Прибыль / RUB']);
    
    // 2. Строка "Всего"
    const totalData = calculateTotalGeoData(groupedGeoData);
    allData.push(['Всего', totalData.sitesCount || 0, totalData.seoCount || 0, totalData.clicks || 0, 0, 0, 0,
      totalData.registrations || 0, 0, 0, 0, totalData.fd || 0, 0, 0, 0, totalData.rd || 0, 0, 0, 0, totalData.deposits || 0, 0,
      0, 0, 0, totalData.revenueUSD || 0, 0, 0, 0, (totalData.revenueUSD || 0) - (totalData.expenseUSD || 0), 0, 0,
      (totalData.expenseUSD || 0) * 95, (totalData.revenueUSD || 0) * 95, ((totalData.revenueUSD || 0) - (totalData.expenseUSD || 0)) * 95]);
    allData.push(['Рост / Падение %', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // 3. ВСЕ ГЕО ДАННЫЕ БЕЗ ОГРАНИЧЕНИЙ - ПОЛНАЯ КАРТИНА
    const sortedGeos = Object.entries(groupedGeoData)
      .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
      // УБРАЛ ВСЕ ОГРАНИЧЕНИЯ - показываем ВСЕ ГЕО
    
    sortedGeos.forEach(([geoKey, geoData]) => {
      const geoDisplayName = получитьСтрану(geoKey);
      
      // Строка ГЕО
      allData.push([geoDisplayName, geoData.data.sitesCount || 0, geoData.data.seoCount || 0, geoData.data.clicks || 0,
        0, 0, 0, geoData.data.registrations || 0, 0, 0, 0, geoData.data.fd || 0, 0, 0, 0, geoData.data.rd || 0, 0, 0, 0,
        geoData.data.deposits || 0, 0, geoData.data.expenseUSD || 0, 0, 0, geoData.data.revenueUSD || 0, 0, 0, 0,
        (geoData.data.revenueUSD || 0) - (geoData.data.expenseUSD || 0), 0, 0,
        (geoData.data.expenseUSD || 0) * 95, (geoData.data.revenueUSD || 0) * 95, ((geoData.data.revenueUSD || 0) - (geoData.data.expenseUSD || 0)) * 95]);
      allData.push(['Рост / Падение', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      
      // ВСЕ ГЕО+БРЕНДЫ БЕЗ ОГРАНИЧЕНИЙ
      const sortedGeoBrands = Object.entries(geoData.geoBrands || {})
        .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
        // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕХ брендов на ГЕО
        
      sortedGeoBrands.forEach(([geoBrandKey, geoBrandData]) => {
        const brandName = geoBrandKey.split('_')[1] || '';
        const geoBrandDisplayName = `  🔸 ${geoDisplayName} ${brandName}`;
        
        allData.push([geoBrandDisplayName, geoBrandData.data.sitesCount || 0, geoBrandData.data.seoCount || 0, geoBrandData.data.clicks || 0,
          0, 0, 0, geoBrandData.data.registrations || 0, 0, 0, 0, geoBrandData.data.fd || 0, 0, 0, 0, geoBrandData.data.rd || 0, 0, 0, 0,
          geoBrandData.data.deposits || 0, 0, geoBrandData.data.expenseUSD || 0, 0, 0, geoBrandData.data.revenueUSD || 0, 0, 0, 0,
          (geoBrandData.data.revenueUSD || 0) - (geoBrandData.data.expenseUSD || 0), 0, 0,
          (geoBrandData.data.expenseUSD || 0) * 95, (geoBrandData.data.revenueUSD || 0) * 95, ((geoBrandData.data.revenueUSD || 0) - (geoBrandData.data.expenseUSD || 0)) * 95]);
        allData.push(['Рост / Падение', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        // ВСЕ ПОТОКИ БЕЗ ОГРАНИЧЕНИЙ
        const sortedStreams = Object.entries(geoBrandData.streams || {})
          .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
          // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ потоки
          
        sortedStreams.forEach(([streamName, streamData]) => {
          allData.push([`    🔸 ${streamName}`, streamData.data.sitesCount || 0, streamData.data.seoCount || 0, streamData.data.clicks || 0,
            0, 0, 0, streamData.data.registrations || 0, 0, 0, 0, streamData.data.fd || 0, 0, 0, 0, streamData.data.rd || 0, 0, 0, 0,
            streamData.data.deposits || 0, 0, streamData.data.expenseUSD || 0, 0, 0, streamData.data.revenueUSD || 0, 0, 0, 0,
            (streamData.data.revenueUSD || 0) - (streamData.data.expenseUSD || 0), 0, 0,
            (streamData.data.expenseUSD || 0) * 95, (streamData.data.revenueUSD || 0) * 95, ((streamData.data.revenueUSD || 0) - (streamData.data.expenseUSD || 0)) * 95]);
          
          // ВСЕ СЕОШНИКИ БЕЗ ОГРАНИЧЕНИЙ
          const sortedSeos = Object.entries(streamData.seos || {})
            .sort(([,a], [,b]) => b.revenueUSD - a.revenueUSD);
            // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕХ сеошников
            
          sortedSeos.forEach(([seoName, seoData]) => {
            if (!ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoName)) {
              allData.push([`      👤 ${seoName}`, 1, 1, seoData.clicks || 0,
                0, 0, 0, seoData.registrations || 0, 0, 0, 0, seoData.fd || 0, 0, 0, 0, seoData.rd || 0, 0, 0, 0,
                seoData.deposits || 0, 0, seoData.expenseUSD || 0, 0, 0, seoData.revenueUSD || 0, 0, 0, 0,
                (seoData.revenueUSD || 0) - (seoData.expenseUSD || 0), 0, 0,
                (seoData.expenseUSD || 0) * 95, (seoData.revenueUSD || 0) * 95, ((seoData.revenueUSD || 0) - (seoData.expenseUSD || 0)) * 95]);
            }
          });
          
          // ВСЕ САЙТЫ БЕЗ ОГРАНИЧЕНИЙ
          if (streamData.sites && streamData.sites.length > 0) {
            const topSites = streamData.sites
              .sort((a, b) => (b.clicks || 0) - (a.clicks || 0));
              // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ сайты
              
            topSites.forEach(site => {
              allData.push([`        🌐 ${site.domain}`, 1, 1, site.clicks || 0,
                0, 0, 0, site.registrations || 0, 0, 0, 0, site.fd || 0, 0, 0, 0, site.rd || 0, 0, 0, 0,
                site.deposits || 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
            });
          }
        });
      });
    });
    
    // ЗАПИСЫВАЕМ ВСЕ ДАННЫЕ ОДНИМ ВЫЗОВОМ (максимально быстро)
    if (allData.length > 0) {
      sheet.getRange(1, 1, allData.length, 34).setValues(allData);
    }
    
    // Минимальное структурное форматирование
    sheet.setFrozenRows(2);
    sheet.setColumnWidth(1, 150);
    
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    console.log(`✅ ПОЛНАЯ ГЕО сводная создана за ${executionTime.toFixed(2)} секунд (${allData.length} строк)`);
    console.log('⚡ ОДИН ВЫЗОВ setValues(): максимальная скорость + минимум API calls');
    console.log('📊 ВСЯ ИНФОРМАЦИЯ БЕЗ ОГРАНИЧЕНИЙ: ВСЕ ГЕО × ВСЕ бренды × ВСЕ потоки × ВСЕ сеошники × ВСЕ сайты');
  });
}

/**
 * ПОЛНЫЕ СЕОШНИКИ БЕЗ ОГРАНИЧЕНИЙ - ОДИН ВЫЗОВ setValues() (ВСЯ ИНФОРМАЦИЯ)
 */
function createSeoSummarySimple(monthName, groupedSeoData) {
  console.log('📊 ПОЛНЫЕ СЕОШНИКИ БЕЗ ОГРАНИЧЕНИЙ: ВСЯ ИНФОРМАЦИЯ + ОДИН ВЫЗОВ setValues()');
  
  return executeWithRetry(() => {
    const startTime = new Date();
    const spreadsheet = SpreadsheetApp.openById(ТАБЛИЦЫ.РЕЗУЛЬТАТ);
    const sheetName = `Сводная Сеошники ${monthName}`;
    
    // Удаляем старый лист если есть
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // СОБИРАЕМ ВСЕ ДАННЫЕ В ОДИН МАССИВ для записи одним вызовом
    const allData = [];
    
    // 1. Заголовки
    allData.push([monthName, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    allData.push(['Сеошник', 'Количество сайтов с трафом', 'Сеошник', 'Клики', '% от общего', 'Рост', 'Рост %',
      'Реги', '% от общего', 'Рост', 'Рост %', 'FD', '% от общего', 'Рост', 'Рост %', 'RD', '% от общего', 'Рост', 'Рост %', 'Deps', '% от общего',
      'Расход / USD', 'Рост', 'Рост %', 'Выручка / USD', '% от общего', 'Рост', 'Рост %',
      'Прибыль / USD', 'Рост', 'Рост %', 'Расход / RUB', 'Выручка / RUB', 'Прибыль / RUB']);
    
    // 2. Строка "Всего" (ВКЛЮЧАЯ ВСЕХ сеошников для единообразия с другими сводными)
    const totalData = calculateTotalData(groupedSeoData);
    allData.push(['Всего', totalData.sitesCount || 0, totalData.seoCount || 0, totalData.clicks || 0, 0, 0, 0,
      totalData.registrations || 0, 0, 0, 0, totalData.fd || 0, 0, 0, 0, totalData.rd || 0, 0, 0, 0, totalData.deposits || 0, 0,
      0, 0, 0, totalData.revenueUSD || 0, 0, 0, 0, (totalData.revenueUSD || 0) - (totalData.expenseUSD || 0), 0, 0,
      (totalData.expenseUSD || 0) * 95, (totalData.revenueUSD || 0) * 95, ((totalData.revenueUSD || 0) - (totalData.expenseUSD || 0)) * 95]);
    allData.push(['Рост / Падение %', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // 3. ВСЕ СЕОШНИКИ БЕЗ ОГРАНИЧЕНИЙ - ПОЛНАЯ КАРТИНА
    const sortedSeos = Object.entries(groupedSeoData)
      .sort(([,a], [,b]) => {
        // Исключенные сеошники в конец
        if (a.isExcluded && !b.isExcluded) return 1;
        if (!a.isExcluded && b.isExcluded) return -1;
        return b.data.revenueUSD - a.data.revenueUSD;
      });
      // УБРАЛ ВСЕ ОГРАНИЧЕНИЯ - показываем ВСЕХ сеошников
    
    sortedSeos.forEach(([seoKey, seoData]) => {
      const seoIcon = ИСКЛЮЧЕННЫЕ_СЕОШНИКИ.includes(seoKey) ? '⚠️' : '👤';
      const seoDisplayName = `${seoIcon} ${seoKey}`;
      
      // Строка сеошника
      allData.push([seoDisplayName, seoData.data.sitesCount || 0, 1, seoData.data.clicks || 0,
        0, 0, 0, seoData.data.registrations || 0, 0, 0, 0, seoData.data.fd || 0, 0, 0, 0, seoData.data.rd || 0, 0, 0, 0,
        seoData.data.deposits || 0, 0, seoData.data.expenseUSD || 0, 0, 0, seoData.data.revenueUSD || 0, 0, 0, 0,
        (seoData.data.revenueUSD || 0) - (seoData.data.expenseUSD || 0), 0, 0,
        (seoData.data.expenseUSD || 0) * 95, (seoData.data.revenueUSD || 0) * 95, ((seoData.data.revenueUSD || 0) - (seoData.data.expenseUSD || 0)) * 95]);
      allData.push(['Рост / Падение', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      
      // ВСЕ БРЕНД+ГЕО БЕЗ ОГРАНИЧЕНИЙ
      const sortedBrandGeos = Object.entries(seoData.brandGeos || {})
        .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
        // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ Бренд+Гео
        
      sortedBrandGeos.forEach(([brandGeoKey, brandGeoData]) => {
        allData.push([`  🔸 ${brandGeoKey}`, brandGeoData.data.sitesCount || 0, brandGeoData.data.seoCount || 0, brandGeoData.data.clicks || 0,
          0, 0, 0, brandGeoData.data.registrations || 0, 0, 0, 0, brandGeoData.data.fd || 0, 0, 0, 0, brandGeoData.data.rd || 0, 0, 0, 0,
          brandGeoData.data.deposits || 0, 0, brandGeoData.data.expenseUSD || 0, 0, 0, brandGeoData.data.revenueUSD || 0, 0, 0, 0,
          (brandGeoData.data.revenueUSD || 0) - (brandGeoData.data.expenseUSD || 0), 0, 0,
          (brandGeoData.data.expenseUSD || 0) * 95, (brandGeoData.data.revenueUSD || 0) * 95, ((brandGeoData.data.revenueUSD || 0) - (brandGeoData.data.expenseUSD || 0)) * 95]);
        allData.push(['Рост / Падение', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        allData.push(['В количестве', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        
        // ВСЕ ПОТОКИ БЕЗ ОГРАНИЧЕНИЙ
        const sortedStreams = Object.entries(brandGeoData.streams || {})
          .sort(([,a], [,b]) => b.data.revenueUSD - a.data.revenueUSD);
          // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ потоки
          
        sortedStreams.forEach(([streamName, streamData]) => {
          allData.push([`    🔸 ${streamName}`, streamData.data.sitesCount || 0, streamData.data.seoCount || 0, streamData.data.clicks || 0,
            0, 0, 0, streamData.data.registrations || 0, 0, 0, 0, streamData.data.fd || 0, 0, 0, 0, streamData.data.rd || 0, 0, 0, 0,
            streamData.data.deposits || 0, 0, streamData.data.expenseUSD || 0, 0, 0, streamData.data.revenueUSD || 0, 0, 0, 0,
            (streamData.data.revenueUSD || 0) - (streamData.data.expenseUSD || 0), 0, 0,
            (streamData.data.expenseUSD || 0) * 95, (streamData.data.revenueUSD || 0) * 95, ((streamData.data.revenueUSD || 0) - (streamData.data.expenseUSD || 0)) * 95]);
          
          // ВСЕ САЙТЫ БЕЗ ОГРАНИЧЕНИЙ
          if (streamData.sites && streamData.sites.length > 0) {
            const topSites = streamData.sites
              .sort((a, b) => (b.clicks || 0) - (a.clicks || 0));
              // УБРАЛ ОГРАНИЧЕНИЯ - показываем ВСЕ сайты
              
            topSites.forEach(site => {
              allData.push([`      🌐 ${site.domain}`, 1, 1, site.clicks || 0,
                0, 0, 0, site.registrations || 0, 0, 0, 0, site.fd || 0, 0, 0, 0, site.rd || 0, 0, 0, 0,
                site.deposits || 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
            });
          }
        });
      });
    });
    
    // ЗАПИСЫВАЕМ ВСЕ ДАННЫЕ ОДНИМ ВЫЗОВОМ (максимально быстро)
    if (allData.length > 0) {
      sheet.getRange(1, 1, allData.length, 34).setValues(allData);
    }
    
    // Минимальное структурное форматирование
    sheet.setFrozenRows(2);
    sheet.setColumnWidth(1, 150);
    
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    console.log(`✅ ПОЛНАЯ Сеошники сводная создана за ${executionTime.toFixed(2)} секунд (${allData.length} строк)`);
    console.log('⚡ ОДИН ВЫЗОВ setValues(): максимальная скорость + минимум API calls');
    console.log('📊 ВСЯ ИНФОРМАЦИЯ БЕЗ ОГРАНИЧЕНИЙ: ВСЕ сеошники × ВСЕ Бренд+Гео × ВСЕ потоки × ВСЕ сайты');
  });
}

/**
 * Базовое структурное форматирование БЕЗ ЦВЕТОВ (для максимальной скорости)
 */
function finalizeSummarySheetStructureOnly(sheet) {
  console.log('⚡ Применяем только структурное форматирование...');
  
  try {
    // Только самое необходимое для структуры
    
    // 1. Заморозить панели (строка 2)
    sheet.setFrozenRows(2);
    
    // 2. Базовые размеры столбцов (без детального форматирования)
    sheet.setColumnWidth(1, 200); // Оффер
    sheet.setColumnWidth(2, 80);  // Количество сайтов
    sheet.setColumnWidth(3, 80);  // Сеошник
    
    // 3. Высота строк для заголовков
    sheet.setRowHeight(1, 25);
    sheet.setRowHeight(2, 25);
    
    console.log('✅ Структурное форматирование применено быстро');
    
  } catch (error) {
    console.error('❌ Ошибка структурного форматирования:', error);
  }
}

console.log('🔧 ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ГОТОВЫ - ТОЧНЫЕ КОПИИ ИЗ ОРИГИНАЛА!');
console.log('🚀 + BATCH-ОБРАБОТКА ДОБАВЛЕНА ДЛЯ УСКОРЕНИЯ!');
console.log('⚡ + РАЗДЕЛЕНИЕ ДАННЫХ И ФОРМАТИРОВАНИЯ ДЛЯ МАКСИМАЛЬНОЙ СКОРОСТИ!');
