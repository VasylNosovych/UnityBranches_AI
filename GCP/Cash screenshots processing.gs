// === НАЛАШТУВАННЯ ===
const CONFIG = {
  FOLDER_NEW_ID: '1cN-idgOL-HGTkue-9VGfYjyigefWsAis',
  FOLDER_PROCESSED_ID: '1151TG3dkdOoOuVjStl8mAoKl6lfRs8mW', // Moving processed files prevents duplicate OCR runs
  SHEET_NAME: 'Колективні внески'
};
// ====================

// ====== ENUMS ======
const CATEGORIES = {
  RELIGIOUS_INSTITUTION: "Храм",
  BUSINESS: "Бізнес",
  COMMUNITY: "Громада",
  ESTABLISHMENT: "Установа"
};

const KNOWN_DONORS = {
  // Бізнес
  HLIANETS_CHARITY: 'Благодійна акція "5 грн. на ЗСУ" у магазині «Глянець»',
  PLANETA_ROSLYN: '"Планета Рослин"',
  DYMYDS_CHARITY: 'Магазин Димидів: акція',
  PAVLIV_HALYNA_KALNA_SHOP: 'Магазин Галини Павлів у с.Кальна',

  // Храм
  CHURCH_CATHOLIC_VYTVYTSIA: 'УГКЦ Храм св. І.Богослова с.Витвиця',
  CHURCH_CATHOLIC_ROZTOCHKY: 'УГКЦ Храм св. Василія Великого с.Розточки',
  CHURCH_ORTHODOX_ROZTOCHKY: 'ПЦУ Храм Різдва Пресвятої Богородиці с.Розточки',
  CHURCH_ORTHODOX_VYTVYTSIA: 'ПЦУ Храм св. А.Первозванного с.Витвиця',
  CHURCH_CATHOLIC_STANKIVTSI: 'УГКЦ Храм св. Архистратига Михаїла с.Станківці',
  CHURCH_CATHOLIC_KALNA: 'УГКЦ Храм св. Архистратига Михаїла с.Кальна',
  CHURCH_ORTHODOX_KALNA: 'ПЦУ Храм св. Архистратига Михаїла с.Кальна',
  CHURCH_ORTHODIX_SLOBODA: 'ПЦУ Храм св. Богоявлення с.Слобода-Болехівська',
  CHURCH_CATHOLIC_MADRID: 'ГКЦ м.Мадрид, Іспанія',
  CHURCH_CATHOLIC_GETAFE: 'ГКЦ м.Гетафе, Іспанія',

  // Установа
  LYCEUM_VYTVYTSIA: 'Витвицький ліцей',
  LYCEUM_KALNA: 'Кальнянський ліцей ім. М. Дяка',
  GYMNASIUM_ROZTOCHKY: 'Розточківська гімназія',
  SCHOOL_STANKIVTSI: 'Станківецька початкова школа ім.Василя Дорошенка',
  COUNCIL_VYTVYTSIA: 'Витвицька сільрада',
  CULTURAL_CENTER_VYTVYTSIA: 'Витвицький БК',
  CULTURAL_CENTER_SLOBODA: 'БК Слобода-Болехівська',

  // Громада
  HLIANETS_FUNDRASING: 'Збір на потреби ЗСУ в магазині «Глянець»',
  BOYKO_MARKET_FUNDRASING: 'Збір на потреби ЗСУ у магазині «Boyko market»',
  COMMUNITY_STANKIVTSI: 'Громада Станківців'
};
// ===================

function processNewScreenshots() {
  const folderNew = DriveApp.getFolderById(CONFIG.FOLDER_NEW_ID);
  const folderProcessed = DriveApp.getFolderById(CONFIG.FOLDER_PROCESSED_ID);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  
  const files = folderNew.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    
    // Перевіряємо, чи файл є зображенням
    if (mimeType.indexOf('image') === -1) continue; 
    
    // 1. Отримуємо текст із зображення
    const text = extractTextFromImage(file);
    
    if (text) {
      // 2. Аналізуємо текст за встановленими правилами
      const rowData = parseReceiptData(text);
      
      // 3. Додаємо новий рядок до таблиці
      sheet.appendRow(rowData.values);
      
      // 4. Якщо система має сумніви (не знайдено точну суму або дату) — виділяємо рядок жовтим кольором
      if (rowData.needsReview) {
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).setBackground('#FFFF00');
      }
      
      // 5. Переміщуємо файл до папки "Оброблені знімки"
      file.moveTo(folderProcessed);
    }
  }
}

function extractTextFromImage(imageFile) {
  const imageBytes = imageFile.getBlob().getBytes();
  const encodedImageContent = Utilities.base64Encode(imageBytes);
  
  const ocrUrl = 'https://vision.googleapis.com/v1/images:annotate';
  const ocrPayload = {
    "requests": [
      {
        "image": { "content": encodedImageContent },
        "features": [{ "type": "DOCUMENT_TEXT_DETECTION" }]
      }
    ]
  };
  
  const ocrRequestOptions = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(ocrPayload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(ocrUrl, ocrRequestOptions);
  const status = response.getResponseCode();
  const json = JSON.parse(response.getContentText());
  
  if (status !== 200) {
    throw new Error(`Vision API error (${status}): ${JSON.stringify(json.error)}`);
  }

  if (json.responses && json.responses[0].textAnnotations) {
    return json.responses[0].textAnnotations[0].description;
  }
  return null;
}

function parseReceiptData(text) {
  let sender = "";
  let amount = "";
  let category = "";
  let date = "";
  let target = "";
  let needsReview = false;
  
  const lines = text.split('\n');
  const fullText = text.toLowerCase();

  // --- ПРАВИЛА ДЛЯ КОЛОНОК "Від кого" і "Категорія" ---
  if (fullText.includes('5 грн. на зсу')) {
    sender = KNOWN_DONORS.HLIANETS_CHARITY;
    category = CATEGORIES.BUSINESS;
  }
  else if (fullText.includes('зі збору у магазині «глянець»')) {
    sender = KNOWN_DONORS.HLIANETS_FUNDRASING;
    category = CATEGORIES.COMMUNITY;
  }
  else if (fullText.includes('святого івана богослова')) {
    sender = KNOWN_DONORS.CHURCH_CATHOLIC_VYTVYTSIA;
    category = CATEGORIES.RELIGIOUS_INSTITUTION;
  }
  else if (fullText.includes('святого василія великого')) {
    sender = KNOWN_DONORS.CHURCH_CATHOLIC_ROZTOCHKY;
    category = CATEGORIES.RELIGIOUS_INSTITUTION;
  }
  else if (fullText.includes('пресвятої богородиці')) {
    sender = KNOWN_DONORS.CHURCH_ORTHODOX_ROZTOCHKY;
    category = CATEGORIES.RELIGIOUS_INSTITUTION;
  }
  else if (fullText.includes('святого андрія первозванного')) {
    sender = KNOWN_DONORS.CHURCH_ORTHODOX_VYTVYTSIA;
    category = CATEGORIES.RELIGIOUS_INSTITUTION;
  }
  else if (fullText.includes('від витвицького ліцею')) {
    sender = KNOWN_DONORS.LYCEUM_VYTVYTSIA;
    category = CATEGORIES.ESTABLISHMENT;
  }
  else {
    // Якщо точного збігу немає, беремо перший рядок тексту як відправника (потребує ручної перевірки)
    sender = lines[0]; 
    needsReview = true;
  }

  // --- ПРАВИЛА ДЛЯ КОЛОНКИ "КАТЕГОРІЯ" ---
  if (category !== "") {
    const sLower = sender.toLowerCase();
    if (sLower.match(/(магазин|глянець|бойко маркет|boyko_market|silver|пленет|plenet|аптека|планета рослин|planeta_roslyn)/)) {
      category = CATEGORIES.BUSINESS;
    }
    else if (sLower.match(/(храм|церква|парафія|парафіяни)/)) {
      category = CATEGORIES.RELIGIOUS_INSTITUTION;
    }
    else if (sLower.match(/(ліцей|школа|бк|турнір|гімназія|сільрада)/)) {
      category = CATEGORIES.ESTABLISHMENT;
    }
    else if (sLower.match(/(громада|від дітей|вертеп|коляда|щедрівка|вигодівські|волонтер|захід|ярмарок|аквагрим|збір)/)) {
      category = CATEGORIES.COMMUNITY;
    }
  }
  

  // --- ПРАВИЛА ДЛЯ КОЛОНКИ "СУМА" ---
  // Шукаємо число, біля якого є "грн" або "uah", ігноруючи слово "залишок"
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].toLowerCase();
    if (line.includes('залишок')) continue;
    
    let match = line.match(/(\d+[.,]?\d*)\s*(грн|uah)/);
    if (match) {
      amount = match[1].replace(',', '.'); // Нормалізуємо роздільник до крапки
      break;
    }
  }
  if (!amount) needsReview = true;

  // --- ПРАВИЛА ДЛЯ КОЛОНКИ "ДАТА" ---
  // Шукаємо дату у форматі ДД.ММ.РРРР або РРРР-ММ-ДД
  let dateMatch = text.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (dateMatch) {
    date = `${dateMatch[3]}-${dateMatch[2]}-${dateMatch[1]}`; // Конвертуємо у формат РРРР-ММ-ДД
  } else {
    let dateMatchISO = text.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (dateMatchISO) {
      date = dateMatchISO[0];
    } else {
      needsReview = true;
    }
  }

  // --- ПРАВИЛА ДЛЯ КОЛОНКИ "КУДИ СКЕРОВАНО" ---
  if (fullText.includes('на матеріали для маскувальних сіт')) {
    target = 'На матеріали для маскувальних сіт';
  }
  else if (fullText.includes('залишок готівки')) {
    target = 'В готівковий залишок на наступні потреби із фронту';
  }

  // Формуємо масив даних для нового рядка: 
  // [Від кого, Сума, Дод.внески, Дод.внески грн, Категорія, Дата, Куди скеровано, Коментар]
  const values = [sender, amount, "", "", category, date, target, ""];
  
  return { values: values, needsReview: needsReview };
}
