/**
 * ============================================================
 *  飲料點單 — Google 表單產生器（GAS / Apps Script）
 *  ------------------------------------------------------------
 *  版本：v2（修正表單與試算表連結）
 *  功能：
 *    1. createDrinkOrderForm()  → 建立「飲料點單」Google 表單
 *       （姓名 / 飲料選擇 / 甜度 / 冰量 / 其他備註，全部單頁）
 *    2. onFormSubmit(e)         → 訂單送出後自動補上「類別、飲料名稱、單價」
 *                                 並寫入綁定的 Google 試算表
 *
 *  使用步驟（見檔案下方註解或向 AI 索取教學）：
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

var FORM_TITLE = '飲料點單';
var FORM_DESCRIPTION = '請填寫您的訂單，送出後我們就會收到囉！';

// 試算表「表單回應」欄位編號（第 1 欄起算，會因你建立的欄位而變）
// 1 時間戳記 | 2 姓名 | 3 飲料選擇 | 4 甜度 | 5 冰量 | 6 其他備註
// 7 類別(自動) | 8 飲料名稱(自動) | 9 單價(自動)
var COL_TIMESTAMP = 1;
var COL_NAME      = 2;
var COL_DRINK     = 3;
var COL_SWEETNESS = 4;
var COL_ICE       = 5;
var COL_NOTE      = 6;
var COL_CATEGORY  = 7;
var COL_DRINKNAME = 8;
var COL_PRICE     = 9;

/**
 * 【主要執行函數】建立 Google 訂購表單（單頁）
 * 執行方式：Apps Script 編輯器中選取 createDrinkOrderForm → 執行
 */
function createDrinkOrderForm() {
  // 1. 綁定目前的試算表（你從試算表的「擴充功能→Apps Script」開啟時，就是同一份）
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 2. 建立表單
  var form = FormApp.create(FORM_TITLE);
  // 將表單回應正式連結到目前的試算表，送出後才會寫入表單回應工作表。
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.setDescription(FORM_DESCRIPTION)
      .setAllowResponseEdits(true)
      .setCollectEmail(false)          // 不強制收集 email
      .setLimitOneResponsePerUser(false) // 允許重複下單
      .setConfirmationMessage('訂單已送出，感謝您的訂購！');

  // 3. 建立欄位（單頁、依序）
  // 3-1 姓名（必填，文字）
  var nameItem = form.addTextItem();
  nameItem.setTitle('使用者姓名（或暱稱）')
          .setHelpText('輸入您的姓名或暱稱')
          .setRequired(true);

  // 3-2 飲料選擇（必填，下拉選單）
  //     Google 表單無法做「選了類別才出現飲料」的連動，
  //     因此用一個下拉選單，選項格式：類別｜名稱｜單價
  var drinkChoices = DRINK_MENU.map(function (d) {
    return d.category + '｜' + d.name + '｜' + d.price + ' 元';
  });
  var drinkItem = form.addListItem();
  drinkItem.setTitle('選擇飲料（類別｜名稱｜單價）')
           .setChoiceValues(drinkChoices)
           .setRequired(true);

  // 3-3 甜度（選填，下拉）
  var sweetItem = form.addListItem();
  sweetItem.setTitle('甜度（選填）')
           .setHelpText('不選也可以，代表交給店家決定')
           .setChoiceValues(['全糖', '二分之一糖', '三分之一糖', '無糖'])
           .setRequired(false);

  // 3-4 冰量（選填，下拉）
  var iceItem = form.addListItem();
  iceItem.setTitle('冰量（選填）')
         .setHelpText('不選也可以，代表交給店家決定')
         .setChoiceValues(['正常冰', '少冰', '微冰', '去冰'])
         .setRequired(false);

  // 3-5 其他備註（選填，文字）
  var noteItem = form.addParagraphTextItem();
  noteItem.setTitle('其他備註（選填）')
          .setHelpText('例如：想跟朋友分兩杯、送到哪個位子…')
          .setRequired(false);

  // 4. 把表單連結記錄到試算表（新工作表「表單網址」）
  var sheet = ss.getSheetByName('表單網址') || ss.insertSheet('表單網址');
  sheet.clear();
  sheet.getRange(1, 1).setValue('表單網址');
  sheet.getRange(2, 1).setValue(form.getPublishedUrl());
  sheet.getRange(3, 1).setValue('表單編輯網址');
  sheet.getRange(4, 1).setValue(form.getEditUrl());
  sheet.getRange(1, 1, 4, 1).setFontWeight('bold');

  // 5. 安裝「送出時自動補類別/名稱/單價」觸發器
  installFormSubmitTrigger_(ss.getId());

  Logger.log('表單已建立！');
  Logger.log('表單網址（給客人）：' + form.getPublishedUrl());
  Logger.log('編輯網址（給自己）：' + form.getEditUrl());
  return form.getPublishedUrl();
}

/**
 * 安裝 onFormSubmit 觸發器
 */
function installFormSubmitTrigger_(spreadsheetId) {
  // 先移除舊的同名觸發器，避免重複執行時越裝越多
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
 * 表單送出後的處理：拆解「類別｜名稱｜單價」並補寫到試算表
 * 注意：此函數由觸發器自動呼叫，不要手動執行
 */
function onFormSubmit(e) {
  try {
    // e 只會由「試算表的表單提交觸發器」提供，不能在編輯器中直接執行。
    if (!e || !e.range) {
      Logger.log('onFormSubmit(e) 不需要手動執行，請開啟表單填寫並按下提交。');
      return;
    }

    var range = e.range;
    var sheet = range.getSheet();
    var row = range.getRow();

    // 讀取「飲料選擇」欄的原始字串，例如「原萃醇茶｜鮮萃綠茶｜30 元」
    var drinkStr = String(sheet.getRange(row, COL_DRINK).getValue());

    // 拆解
    var parts = drinkStr.split('｜');
    if (parts.length !== 3) {
      Logger.log('第 ' + row + ' 列飲料格式無法拆解：' + drinkStr);
      return;
    }
    var category = parts[0].trim();
    var drinkName = parts[1].trim();
    var price = parseInt(parts[2].replace('元', '').trim(), 10) || '';

    // 第一次收到回應時，補上自動欄位標題。
    sheet.getRange(1, COL_CATEGORY, 1, 3).setValues([
      ['類別（自動）', '飲料名稱（自動）', '單價（自動）']
    ]);

    // 寫入試算表第 7、8、9 欄
    sheet.getRange(row, COL_CATEGORY).setValue(category);
    sheet.getRange(row, COL_DRINKNAME).setValue(drinkName);
    sheet.getRange(row, COL_PRICE).setValue(price);
  } catch (err) {
    Logger.log('onFormSubmit 發生錯誤：' + err.message);
  }
}
