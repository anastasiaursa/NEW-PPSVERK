/**
 * 🎨 РАБОЧАЯ ВЕРСИЯ УМНОЙ ЗАМЕНЫ ЦВЕТОВ
 * Простая и надежная функция для активного листа
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
    
    // Определяем тип сводной таблицы
    const summaryType = detectSummaryType(values, lastRow);
    console.log(`🔍 Тип сводной: ${summaryType}`);
    
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
        else if (needsGradient(cellText, summaryType)) {
          console.log(`🎨 Применяем градиент для: "${cellText}" (тип: ${summaryType})`);
          
          // Определяем, какой элемент должен быть темнее
          const shouldBeDarker = shouldElementBeDarker(cellText, summaryType);
          
          if (shouldBeDarker) {
            // Столбец A темнее
            sheet.getRange(row + 1, 1).setBackground(getDarkColor(color));
            // Остальные светлее
            if (lastCol > 1) {
              sheet.getRange(row + 1, 2, 1, lastCol - 1).setBackground(getLightColor(color));
            }
          } else {
            // Столбец A светлее
            sheet.getRange(row + 1, 1).setBackground(getLightColor(color));
            // Остальные темнее
            if (lastCol > 1) {
              sheet.getRange(row + 1, 2, 1, lastCol - 1).setBackground(getDarkColor(color));
            }
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
 * Определяет тип сводной таблицы по содержимому
 */
function detectSummaryType(values, lastRow) {
  let seoCount = 0;
  let geoCount = 0;
  let brandGeoCount = 0;
  let pureGeoCount = 0;
  
  // Анализируем первые 20 строк (или все, если меньше)
  const rowsToCheck = Math.min(20, lastRow);
  
  for (let row = 1; row < rowsToCheck; row++) { // Пропускаем заголовок
    const cellA = values[row][0];
    if (!cellA) continue;
    
    const cellText = cellA.toString().trim();
    
    if (isSeo(cellText)) {
      seoCount++;
    } else if (isBrandGeoSummary(cellText)) {
      brandGeoCount++;
    } else if (isPureGeoCode(cellText)) {
      pureGeoCount++;
    } else if (isGeo(cellText)) {
      geoCount++;
    }
  }
  
  console.log(`Статистика: SEO=${seoCount}, GEO=${geoCount}, PureGEO=${pureGeoCount}, BrandGEO=${brandGeoCount}`);
  
  // Определяем тип по преобладающим элементам
  if (seoCount > Math.max(geoCount, brandGeoCount, pureGeoCount)) {
    return 'seo_summary';
  } else if (brandGeoCount > Math.max(seoCount, geoCount, pureGeoCount)) {
    return 'brand_geo_summary';
  } else if (pureGeoCount > Math.max(seoCount, geoCount, brandGeoCount)) {
    return 'geo_summary';
  } else if (geoCount > 0) {
    return 'geo_summary';
  }
  
  return 'mixed';
}

/**
 * Проверяет, является ли текст чистым ГЕО кодом (2 буквы)
 */
function isPureGeoCode(text) {
  const cleanText = text.trim().toUpperCase();
  
  // Проверяем, что это ровно 2 буквы и это известный ГЕО код
  if (cleanText.length === 2 && /^[A-Z]{2}$/.test(cleanText)) {
    return isGeoCode(cleanText.toLowerCase());
  }
  
  return false;
}

/**
 * Определяет, должен ли элемент в столбце A быть темнее
 */
function shouldElementBeDarker(text, summaryType) {
  switch (summaryType) {
    case 'seo_summary':
      // В сводной по сеошникам - сеошники должны быть темнее
      return isSeo(text);
      
    case 'geo_summary':
      // В сводной по ГЕО - ГЕО должны быть темнее
      return isGeo(text) || isPureGeoCode(text);
      
    case 'brand_geo_summary':
      // В сводной по бренд+ГЕО - бренды должны быть темнее (всегда столбец A)
      return true;
      
    default:
      // По умолчанию столбец A темнее
      return true;
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
  
  // 6. Чистые ГЕО коды (2 буквы)
  if (isPureGeoCode(text)) {
    console.log(`🌍 Чистый ГЕО код: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 7. Бренды (только отдельные)
  if (isBrand(lower)) {
    console.log(`🏢 Бренд: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 8. ГЕО строки
  if (isGeo(text)) {
    console.log(`🌍 ГЕО строка: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 9. Сеошники
  if (isSeo(text)) {
    console.log(`👤 Сеошник: ${text}`);
    return '#d9d2e9'; // Сиреневый
  }
  
  // 10. Подгруппы (сводные)
  if (lower.includes('сводная') || text.includes('/')) {
    console.log(`📁 Сводная/подгруппа: ${text}`);
    return '#f4e6ed'; // Светлый розовый
  }
  
  // 11. Все остальное - белое
  console.log(`⚪ Белый фон по умолчанию: ${text}`);
  return '#ffffff';
}

/**
 * Проверяет, нужен ли градиент (темнее A, светлее остальные)
 */
function needsGradient(text, summaryType) {
  const lower = text.toLowerCase();
  
  // Градиент НЕ нужен для сводных Бренд+Гео (они обрабатываются отдельно)
  if (isBrandGeoSummary(text)) {
    return false;
  }
  
  // Градиент для элементов в зависимости от типа сводной
  switch (summaryType) {
    case 'seo_summary':
      return isSeo(text);
      
    case 'geo_summary':
      return isGeo(text) || isPureGeoCode(text);
      
    case 'brand_geo_summary':
      return false; // Обрабатывается отдельно
      
    default:
      // Стандартная логика для смешанных таблиц
      return isBrand(lower) || isGeo(text) || isSeo(text) || isPureGeoCode(text);
  }
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
 * Проверяет коды стран для сводных (АКТУАЛЬНЫЙ СПИСОК - только 2-буквенные коды)
 */
function isGeoCode(text) {
  const geoCodes = [
    'az', 'ru', 'uz', 'kz', 'by', 'ua', 'ge', 'am', 'md', 'tj', 'kg', 'tm',
    'hu', 'tr', 'fr', 'de', 'it', 'es', 'pt', 'gr', 'pl', 'cz', 'sk', 'bg', 'ch',
    'bd', 'np', 'in', 'pk', 'lk', 'vn', 'kr', 'jp',
    'ar', 'eg', 'ng', 'za', 'ke', 'tz', 'ug', 'sn', 'cm', 'ga', 'mz',
    'br', 'mx', 'co', 'cl', 'pe', 'ec', 'ca',
    'ww'
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
  const geoPatterns = [
    /^🇦🇿\s*AZ/i, /^🇷🇺\s*RU/i, /^🇺🇿\s*UZ/i, /^🇰🇿\s*KZ/i, /^🇧🇾\s*BY/i,
    /^🇺🇦\s*UA/i, /^🇬🇪\s*GE/i, /^🇦🇲\s*AM/i, /^🇲🇩\s*MD/i, /^🇹🇯\s*TJ/i,
    /^🇰🇬\s*KG/i, /^🇹🇲\s*TM/i, /^🇭🇺\s*HU/i, /^🌍\s*[A-Z]{2}/i
  ];
  
  return geoPatterns.some(pattern => pattern.test(text));
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
