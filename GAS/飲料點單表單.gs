/**
 * ============================================================
 *  飲料點單 — Google 表單產生器（GAS / Apps Script）
 *  ------------------------------------------------------------
 *  版本：v3（新增手機欄位，兼容 V2 回應）
 *  功能：
 *    1. createDrinkOrderForm()  → 建立「飲料點單 V3」Google 表單
 *       （姓名 / 手機 / 飲料 / 甜度 / 冰量 / 其他備註，全部單頁）
 *    2. onFormSubmit(e)         → 訂單送出後自動補上「類別、飲料名稱、單價」
 *                                 並寫入綁定的 Google 試算表
 *
 *  使用步驟：
 *    在試算表 → 擴充功能 → Apps Script → 貼上本檔 → 執行 createDrinkOrderForm()
 * ============================================================
 */

// ========== 菜單資料（25 項，與「飲料點單表格.md」同步） ==========
var DRINK_MENU = [
  { category: '原萃醇茶', name: '鮮萃綠茶',     price: 30 },
  { category: '原萃醇茶', name: '阿薩姆紅茶',   price: 30 },
  { category: '原萃醇茶', name: '台灣文青',     price: 50 },
  { category: '原萃醇茶', name: '蜂蜜文青',     price: 65 },
  { category: '原萃醇茶', name: '蜂蜜紅茶',     price: 45 },
  { category: '原萃醇茶', name: '蜂蜜綠茶',     price: 45 },
  { category: '原萃醇茶', name: '珍珠紅茶',     price: 35 },
  { category: '原萃醇茶', name: '珍珠綠茶',     price: 35 },
  { category: '醇品奶茶', name: '奶茶',         price: 35 },
  { category: '醇品奶茶', name: '奶綠',         price: 35 },
  { category: '醇品奶茶', name: '珍珠奶茶',     price: 40 },
  { category: '醇品奶茶', name: '珍珠奶綠',     price: 40 },
  { category: '醇品奶茶', name: '文青奶茶',     price: 50 },
  { category: '醇品奶茶', name: '文青珍奶',     price: 55 },
  { category: '鮮奶安心喝', name: '紅茶鮮奶茶', price: 65 },
  { category: '鮮奶安心喝', name: '綠茶鮮奶茶', price: 65 },
  { category: '鮮奶安心喝', name: '文青鮮奶茶', price: 65 },
  { category: '鮮奶安心喝', name: '黑糖珍珠鮮奶', price: 65 },
  { category: '友善喝咖啡', name: '美式咖啡',   price: 50 },
  { category: '友善喝咖啡', name: '拿鐵咖啡',   price: 65 },
  { category: '友善喝咖啡', name: '焦糖拿鐵',   price: 75 },
  { category: '友善喝咖啡', name: '香草拿鐵',   price: 75 },
  { category: '友善喝咖啡', name: '榛果拿鐵',   price: 75 },
  { category: '風味飲品', name: '蜂蜜檸檬汁',   price: 65 },
  { category: '風味飲品', name: '檸檬汁',       price: 65 }
];

var FORM_TITLE = '飲料點單 V3';
var FORM_DESCRIPTION = '請填寫您的姓名、手機與訂單內容，送出後我們就會收到囉！';

// V3 新表單欄位順序為：時間戳記 / 姓名 / 手機 / 飲料 / 甜度 / 冰量 / 備註。
// onFormSubmit 會依照欄位標題尋找位置，因此也兼容 V2 表單與舊版訂購網頁。

/**
 * 【主要執行函數】建立 Google 訂購表單（單頁）
 * 執行方式：Apps Script 編輯器中選取 createDrinkOrderForm → 執行
 * 注意：每執行一次都會建立一張新的表單，請只執行一次。
 */
function createDrinkOrderForm() {
  // 必須從目標 Google 試算表的「擴充功能→Apps Script」開啟本程式。
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('找不到目前的 Google 試算表，請從目標試算表的「擴充功能→Apps Script」開啟。');
  }

  // 建立表單，並正式連結到目前的試算表。
  var form = FormApp.create(FORM_TITLE);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.setDescription(FORM_DESCRIPTION)
      .setAllowResponseEdits(true)
      .setCollectEmail(false)
      .setLimitOneResponsePerUser(false)
      .setConfirmationMessage('訂單已送出，感謝您的訂購！');

  // 1. 使用者姓名（必填）
  var nameItem = form.addTextItem();
  nameItem.setTitle('使用者姓名（或暱稱）')
          .setHelpText('輸入您的姓名或暱稱')
          .setRequired(true);

  // 2. 手機號碼（必填）
  var phoneValidation = FormApp.createTextValidation()
      .setHelpText('請輸入 09 開頭的 10 碼手機號碼，例如 0912345678')
      .requireTextMatchesPattern('^09[0-9]{8}$')
      .build();
  var phoneItem = form.addTextItem();
  phoneItem.setTitle('手機號碼')
           .setHelpText('請輸入 09 開頭的 10 碼手機號碼，例如 0912345678')
           .setValidation(phoneValidation)
           .setRequired(true);

  // 3. 飲料選擇（必填，下拉選單）
  // Google 表單無法做「選了類別才出現飲料」的連動，
  // 因此用一個下拉選單，選項格式：類別｜名稱｜單價。
  var drinkChoices = DRINK_MENU.map(function (d) {
    return d.category + '｜' + d.name + '｜' + d.price + ' 元';
  });
  var drinkItem = form.addListItem();
  drinkItem.setTitle('選擇飲料（類別｜名稱｜單價）')
           .setChoiceValues(drinkChoices)
           .setRequired(true);

  // 4. 甜度（選填）
  var sweetItem = form.addListItem();
  sweetItem.setTitle('甜度（選填）')
           .setHelpText('不選也可以，代表交給店家決定')
           .setChoiceValues(['全糖', '二分之一糖', '三分之一糖', '無糖'])
           .setRequired(false);

  // 5. 冰量（選填）
  var iceItem = form.addListItem();
  iceItem.setTitle('冰量（選填）')
         .setHelpText('不選也可以，代表交給店家決定')
         .setChoiceValues(['正常冰', '少冰', '微冰', '去冰'])
         .setRequired(false);

  // 6. 其他備註（選填）
  var noteItem = form.addParagraphTextItem();
  noteItem.setTitle('其他備註（選填）')
          .setHelpText('例如：想跟朋友分兩杯、送到哪個位子…')
          .setRequired(false);

  // 將新表單的網址寫入試算表，方便取得與分享。
  var sheet = ss.getSheetByName('表單網址') || ss.insertSheet('表單網址');
  sheet.clear();
  sheet.getRange(1, 1).setValue('表單網址');
  sheet.getRange(2, 1).setValue(form.getPublishedUrl());
  sheet.getRange(3, 1).setValue('表單編輯網址');
  sheet.getRange(4, 1).setValue(form.getEditUrl());
  sheet.getRange(5, 1).setValue('版本');
  sheet.getRange(6, 1).setValue('v3（含手機號碼）');
  sheet.getRange(1, 1, 6, 1).setFontWeight('bold');

  // 安裝「送出時自動補類別/名稱/單價」觸發器。
  installFormSubmitTrigger_(ss.getId());

  Logger.log('表單已建立！');
  Logger.log('表單網址（給客人）：' + form.getPublishedUrl());
  Logger.log('編輯網址（給自己）：' + form.getEditUrl());
  return form.getPublishedUrl();
}

/**
 * 安裝 onFormSubmit 觸發器。
 */
function installFormSubmitTrigger_(spreadsheetId) {
  // 先移除舊的同名觸發器，避免重複執行時越裝越多。
  ScriptApp.getProjectTriggers().forEach(function (trig) {
    if (trig.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trig);
    }
  });

  ScriptApp.newTrigger('onFormSubmit')
          .forSpreadsheet(spreadsheetId)
          .onFormSubmit()
          .create();
}

/**
 * 表單送出後的處理：拆解「類別｜名稱｜單價」並補寫到試算表。
 * 注意：此函數由觸發器自動呼叫，不要手動執行。
 */
function onFormSubmit(e) {
  try {
    // e 只會由「試算表的表單提交觸發器」提供。
    if (!e || !e.range) {
      Logger.log('onFormSubmit(e) 不需要手動執行，請開啟 V3 表單填寫並按下提交。');
      return;
    }

    var range = e.range;
    var sheet = range.getSheet();
    var row = range.getRow();

    // 依照標題尋找飲料欄，兼容 V2 與 V3 不同的欄位位置。
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var drinkColumn = findColumnByHeader_(headers, '選擇飲料') ||
                      findColumnByHeader_(headers, '飲料選擇');
    if (!drinkColumn) {
      Logger.log('找不到飲料欄位，工作表：' + sheet.getName());
      return;
    }

    // 讀取飲料原始字串，例如「原萃醇茶｜鮮萃綠茶｜30 元」。
    var drinkStr = String(sheet.getRange(row, drinkColumn).getValue());
    var parts = drinkStr.split('｜');
    if (parts.length !== 3) {
      Logger.log('第 ' + row + ' 列飲料格式無法拆解：' + drinkStr);
      return;
    }

    var category = parts[0].trim();
    var drinkName = parts[1].trim();
    var price = parseInt(parts[2].replace('元', '').trim(), 10) || '';

    // 找到或建立自動欄位，V2 與 V3 都使用同一組標題。
    var autoColumns = ensureAutoColumns_(sheet, headers);

    // 寫入類別、飲料名稱、單價。
    sheet.getRange(row, autoColumns.category, 1, 3).setValues([
      [category, drinkName, price]
    ]);
  } catch (err) {
    Logger.log('onFormSubmit 發生錯誤：' + err.message);
  }
}

/**
 * 依照欄位標題尋找欄位編號。
 */
function findColumnByHeader_(headers, keyword) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).indexOf(keyword) !== -1) {
      return i + 1;
    }
  }
  return 0;
}

/**
 * 找到或建立自動補上的三個欄位。
 */
function ensureAutoColumns_(sheet, headers) {
  var categoryColumn = findColumnByHeader_(headers, '類別（自動）');
  if (!categoryColumn) {
    categoryColumn = sheet.getLastColumn() + 1;
  }

  sheet.getRange(1, categoryColumn, 1, 3).setValues([
    ['類別（自動）', '飲料名稱（自動）', '單價（自動）']
  ]);

  return {
    category: categoryColumn,
    drinkName: categoryColumn + 1,
    price: categoryColumn + 2
  };
}
