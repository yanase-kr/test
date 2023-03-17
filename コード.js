const LENDING_LEDER_NAME      = '貸出台帳'

const LENDING_TYPE_LENDING    = '貸出'
const LENDING_TYPE_RETURN     = '返却'
const LENDING_TYPE_INVENTORY  = '棚卸'

const STOCK_TYPE_WAREHOUSING  = '入庫'
const STOCK_TYPE_DELIVER      = '出庫'

const LENDING_COLUMN_ID           = 1
const LENDING_COLUMN_PERSON_ID    = 2
const LENDING_COLUMN_ITEM_ID      = 3
const LENDING_COLUMN_TYPE         = 4
const LENDING_COLUMN_COUNT        = 5
const LENDING_COLUMN_ACTION_TIME  = 6
const LENDING_COLUMN_INPUT_TIME   = 7

const STOCK_SHEET_START_ROW       = 2
const STOCK_SHEET_START_COLUMN    = 1

const STOCK_COLUMN_STOCK_ID       = 1
const STOCK_COLUMN_ITEM_ID        = 2
const STOCK_COLUMN_TYPE           = 3
const STOCK_COLUMN_ALL_COUNT      = 4
const STOCK_COLUMN_ACTION_TIME    = 5
const STOCK_COLUMN_INPUT_TIME     = 6

function myFunction() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫台帳');

  if (activeSheet.getName() != LENDING_LEDER_NAME) {
    Logger.log("対象じゃないですよ");
    return;
  }

  var lastLine = activeSheet.getActiveCell().getRow();

  var lendingId = activeSheet.getRange(lastLine, LENDING_COLUMN_ID).getDisplayValue();
  var itemId = activeSheet.getRange(lastLine, LENDING_COLUMN_ITEM_ID).getDisplayValue();
  var itemType = activeSheet.getRange(lastLine, LENDING_COLUMN_TYPE).getDisplayValue();
  var itemCount = activeSheet.getRange(lastLine, LENDING_COLUMN_COUNT).getDisplayValue();
  var itemActTime = activeSheet.getRange(lastLine, LENDING_COLUMN_ACTION_TIME).getDisplayValue();
  var itenInputTime = activeSheet.getRange(lastLine, LENDING_COLUMN_INPUT_TIME).getDisplayValue();

  // 現在までのSTOCK_COLUMN_ALL_COUNTをレコードから引く
  var stockArray = stockSheet.getRange(STOCK_SHEET_START_ROW, STOCK_SHEET_START_COLUMN, lastLine, STOCK_COLUMN_INPUT_TIME).getValues();
  var targetArray = stockArray.filter(record => record[2] == itemId);
  

  var allCount = 0;
  var stockSheetLastLine = stockSheet.getLastRow() + 1;

  allCount = allCount + itemCount;

  stockSheet.getRange(stockSheetLastLine, STOCK_COLUMN_STOCK_ID).setValue(lendingId);
  stockSheet.getRange(stockSheetLastLine, STOCK_COLUMN_ITEM_ID).setValue(itemId);
  stockSheet.getRange(stockSheetLastLine, STOCK_COLUMN_TYPE).setValue(itemType);
  stockSheet.getRange(stockSheetLastLine, STOCK_COLUMN_ALL_COUNT).setValue(allCount);
  stockSheet.getRange(stockSheetLastLine, STOCK_COLUMN_ACTION_TIME).setValue(itemActTime);
  stockSheet.getRange(stockSheetLastLine, STOCK_COLUMN_INPUT_TIME).setValue(itenInputTime);
}
