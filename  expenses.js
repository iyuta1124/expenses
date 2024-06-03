function doPost(e) {
  const { input1, input2, replyToken } = createWebhookItems(e);
  let response = createInitialResponse();
  switch (input1) {
    case "表示名設定":
      response = setUserName(e, input2);
      break;
    case "合計":
      response = sumAmountsByColumn(e);
      break;
    case "削除":
      response = deleteLastRecord(e);
      break;
    default:
      response = recordToSheet(e);
      break;
  }
  if (response.status === 201) {
    sendBroadCatsMessage(response.message);
  } else {
    sendReply(replyToken, response.message);
  }
  sendJsonResponse();
}

function sumAmountsByColumn(e) {
  if (validate(e, "sum")) {
    return createErrorResponse("sum");
  }
  const { input2 } = createWebhookItems(e);
  const parsedDate = input2 ? parseYearAndMonth(input2) : null;
  const targetSheetName = parsedDate
    ? convertDateToYearMonthFormat(new Date(parsedDate[0], parsedDate[1] - 1))
    : null;
  const sheet = getSheet(targetSheetName);
  const dataRange = sheet.getDataRange(); // データが存在する範囲を取得
  const values = dataRange.getValues(); // 全データを二次元配列として取得
  const { YEAR_MONTH, USER_NAME, CATEGORY, AMOUNT } = COLUMN_NUMBERS;
  const amountColumn = AMOUNT; // 金額が格納されている列番号（例: 6列目）
  // 合計結果をログに出力
  const response = createInitialResponse();
  const targetColumnNumbers = [YEAR_MONTH, USER_NAME, CATEGORY];
  const allSums = targetColumnNumbers.map((columnNum) => {
    let sums = {}; // 合計を格納するオブジェクト
    values.forEach((row, i) => {
      // 1行目はヘッダーなので2行目から処理
      if (i === 0) return;
      const key =
        columnNum === COLUMN_NUMBERS.YEAR_MONTH
          ? convertDateToYearMonthFormat(new Date(values[i][columnNum - 1]))
          : values[i][columnNum - 1];
      const amount = parseFloat(values[i][amountColumn - 1]); // 金額を取得し、数値に変換
      if (!sums[key]) {
        sums[key] = 0; // キーがまだオブジェクトに存在しない場合は初期化
      }
      sums[key] += amount; // 金額を加算
    });
    return sums;
  });
  response.message = "合計金額は下記の通りです！\n";
  allSums.forEach((sums, index) => {
    response.message += `=====================\n`;
    const columnName = COLUMN_ID_NAME_MAP[targetColumnNumbers[index]];
    response.message += `${columnName}ごとの合計金額\n`;
    if (Object.keys(sums).length === 0) {
      response.message += "該当するデータはありませんでした。\n";
    } else {
      for (const [key, value] of Object.entries(sums)) {
        response.message += key + ": " + value + "円\n";
      }
    }
  });
  return response;
}

function createWebhookItems(webhookEvent) {
  const postData = JSON.parse(webhookEvent.postData.contents);
  const replyToken = postData.events[0].replyToken;
  const userMessage = postData.events[0].message.text;
  const messageId = postData.events[0].message.id;
  const userId = postData.events[0].source.userId;
  const userName = getUserName(userId);
  // 文字列の前後の空白をトリムし、文字列内の連続する空白を1つの空白に置き換える
  const [input1, input2, input3] = userMessage
    .trim()
    .replace(/\s+/g, " ")
    .split(/ |　/);
  return { messageId, input1, input2, input3, userName, replyToken, userId };
}

function createColumnItems(webhookEvent) {
  const { messageId, input1, input2, input3, userName } =
    createWebhookItems(webhookEvent);
  const writingCategory = input3 ? input3 : "食料品";
  const inputDay = formatDateToJapaneseStandard(new Date());
  const currentMonth = convertDateToYearMonthFormat(new Date());
  const amount = Number(input2);
  const item = input1;
  return {
    messageId,
    inputDay,
    currentMonth,
    writingCategory,
    item,
    amount,
    userName,
  };
}

function validate(webhookEvent, type) {
  const { input1, input2 } = createWebhookItems(webhookEvent);
  switch (type) {
    case "record":
      return !input1 || Number.isNaN(Number(input2));
    case "sum": {
      if (input2) {
        const isYearMonthMatch = input2.match(/(\d{4})年(\d{1,2})月/);
        const isMonthMatch = input2.match(/(\d{1,2})月/);
        return !(isYearMonthMatch || isMonthMatch);
      }
      return false;
    }
    case "userName":
      return !input2;
  }
  return false;
}

function createErrorResponse(type) {
  const response = createInitialResponse();
  switch (type) {
    case "record": {
      response.status = 400;
      response.message =
        "メッセージの形式が正しくありません。「記録名 金額 カテゴリ」の順で空白区切りで入力してください。";
      break;
    }
    case "sum": {
      response.status = 400;
      response.message =
        "メッセージの形式が正しくありません。「合計 6月」または「合計 2023年6月」のように入力してください。";
      break;
    }
    case "userName": {
      response.status = 400;
      response.message = "ユーザー名を入力してください。";
      break;
    }
  }
  return response;
}

function recordToSheet(webhookEvent) {
  const response = createInitialResponse();
  if (validate(webhookEvent, "record")) {
    return createErrorResponse("record");
  }
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const columnItems = createColumnItems(webhookEvent);
  Object.values(columnItems).forEach((columnValue, index) => {
    sheet.getRange(lastRow + 1, index + 1).setValue(columnValue);
  });
  const { inputDay, writingCategory, item, amount, userName } = columnItems;
  response.status = 201;
  response.message = `${userName}さんが下記内容で記録しました！
  ==============
  ${COLUMN_NAMES.INPUT_DAY}:${inputDay}
  ${COLUMN_NAMES.CATEGORY}:${writingCategory}
  ${COLUMN_NAMES.ITEM}:${item}
  ${COLUMN_NAMES.AMOUNT}:${amount}
  ${COLUMN_NAMES.USER_NAME}:${userName}
  ==============
  `;
  return response;
}

function formatDateToJapaneseStandard(date) {
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd, HH:mm");
}

function convertDateToYearMonthFormat(date) {
  const formattedDate = Utilities.formatDate(date, "Asia/Tokyo", "yyyy年M月");
  return formattedDate;
}

function parseYearAndMonth(input) {
  let year, month;
  // 入力が「2023年6月」の形式の場合
  const yearMonthMatch = input.match(/(\d{4})年(\d{1,2})月/);
  const monthMatch = input.match(/(\d{1,2})月/);
  if (yearMonthMatch) {
    year = parseInt(yearMonthMatch[1], 10);
    month = parseInt(yearMonthMatch[2], 10);
  } else if (monthMatch) {
    year = new Date().getFullYear(); // 現在の年を取得
    month = parseInt(monthMatch[1], 10);
  }
  return [year, month];
}

function deleteLastRecord(e) {
  const response = createInitialResponse();
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
    const values = range.getValues()[0];
    sheet.deleteRow(lastRow);
    const { INPUT_DAY, CATEGORY, ITEM, AMOUNT, USER_NAME } = COLUMN_NUMBERS;
    response.status = 201;
    const { userName } = createWebhookItems(e);
    response.message = `${userName}さんが最後の記録を削除しました！
  ==============
  ${COLUMN_NAMES.INPUT_DAY}:${formatDateToJapaneseStandard(
      new Date(values[INPUT_DAY - 1])
    )}
  ${COLUMN_NAMES.CATEGORY}:${values[CATEGORY - 1]}
  ${COLUMN_NAMES.ITEM}:${values[ITEM - 1]}
  ${COLUMN_NAMES.AMOUNT}:${values[AMOUNT - 1]}
  ${COLUMN_NAMES.USER_NAME}:${values[USER_NAME - 1]}
  ==============
  `;
  } else {
    response.status = 404;
    response.message = "削除する記録がありません。";
  }
  return response;
}

function sendReply(replyToken, message) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const headers = {
    "Content-Type": "application/json; charset=UTF-8",
    Authorization: "Bearer " + getAccessToken(),
  };
  const payload = JSON.stringify({
    replyToken: replyToken,
    messages: [
      { type: "text", text: message ? message : "エラーが発生しました" },
    ],
  });
  UrlFetchApp.fetch(url, {
    method: "post",
    headers: headers,
    payload: payload,
  });
}

function setUserName(e, userName) {
  if (validate(e, "userName")) {
    return createErrorResponse("userName");
  }
  const { userId } = createWebhookItems(e);
  PropertiesService.getUserProperties().setProperty(userId, userName);
  const response = createInitialResponse();
  response.message = `表示名を${userName}に設定しました！`;
  return response;
}

function getUserName(userId) {
  const userName = PropertiesService.getUserProperties().getProperty(userId);
  return userName ? userName : "未設定";
}

function getAccessToken() {
  return PropertiesService.getScriptProperties().getProperty("ACCESS_TOKEN");
}

function sendBroadCatsMessage(message) {
  const url = "https://api.line.me/v2/bot/message/broadcast";
  const headers = {
    "Content-Type": "application/json; charset=UTF-8",
    Authorization: "Bearer " + getAccessToken(),
  };
  const payload = JSON.stringify({
    messages: [
      { type: "text", text: message ? message : "エラーが発生しました" },
    ],
  });
  UrlFetchApp.fetch(url, {
    method: "post",
    headers: headers,
    payload: payload,
  });
}

function sendJsonResponse() {
  return ContentService.createTextOutput(
    JSON.stringify({ content: "post ok" })
  ).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(sheetName) {
  if (sheetName) {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }
  return getOrCreateSheetByMonth();
}

function createInitialResponse() {
  return { status: 200, message: "", data: null };
}

function getOrCreateSheetByMonth() {
  const currentMonth = convertDateToYearMonthFormat(new Date());
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(currentMonth);
  if (!sheet) {
    // シートが存在しない場合、新しいシートを作成
    sheet = spreadsheet.insertSheet(currentMonth, 0);
    // 新しいシートにヘッダーなどの初期設定を行う
    sheet.appendRow(COLUMN_NAMES);
  }
  spreadsheet.setActiveSheet(sheet);
  return sheet;
}

const COLUMN_NUMBERS = {
  MESSAGE_ID: 1,
  INPUT_DAY: 2,
  YEAR_MONTH: 3,
  CATEGORY: 4,
  ITEM: 5,
  AMOUNT: 6,
  USER_NAME: 7,
};

const COLUMN_NAMES = {
  MESSAGE_ID: "ID",
  INPUT_DAY: "入力日時",
  YEAR_MONTH: "日付",
  CATEGORY: "カテゴリ",
  ITEM: "品目",
  AMOUNT: "価格",
  USER_NAME: "記録者",
};

const COLUMN_ID_NAME_MAP = {
  [COLUMN_NUMBERS.MESSAGE_ID]: COLUMN_NAMES.MESSAGE_ID,
  [COLUMN_NUMBERS.INPUT_DAY]: COLUMN_NAMES.INPUT_DAY,
  [COLUMN_NUMBERS.YEAR_MONTH]: COLUMN_NAMES.YEAR_MONTH,
  [COLUMN_NUMBERS.CATEGORY]: COLUMN_NAMES.CATEGORY,
  [COLUMN_NUMBERS.ITEM]: COLUMN_NAMES.ITEM,
  [COLUMN_NUMBERS.AMOUNT]: COLUMN_NAMES.AMOUNT,
  [COLUMN_NUMBERS.USER_NAME]: COLUMN_NAMES.USER_NAME,
};
