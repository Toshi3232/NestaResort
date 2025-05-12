// フルコード（進捗を画面右下にトースト表示）

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('カスタム機能')
      .addItem('期間指定分析', 'showDateRangePicker')
      .addItem('全シートをソート', 'sortAllSheets')
      .addToUi();

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('纏め');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('纏め');
    }

    sheet.getRange('A1').setValue('分析状態').setFontWeight('bold');
    sheet.getRange('B1').setValue('未実行').setFontColor('red');

    if (sheet.getRange('A2').getValue() !== '担当者') {
      var headers = [['担当者', '月', 'Time', '現状', '課題', '戦術', '行動特性', '強み']];
      sheet.getRange('A2:H2').setValues(headers)
        .setFontWeight('bold')
        .setBackground('#f3f3f3')
        .setFontColor('black')
        .setFontSize(11);
    }
  } catch (error) {
    Logger.log('onOpen エラー: ' + error);
  }
}

function updateProgress(status) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('纏め');
  sheet.getRange('B1').setValue(status);
  
  if (status === '分析完了') {
    SpreadsheetApp.getActiveSpreadsheet().toast('🎉 分析が終わりました', '完了通知', 5);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(status, '進捗', 5);
  }

  SpreadsheetApp.flush();
}


function showDateRangePicker() {
  var html = HtmlService.createHtmlOutputFromFile('DateRangePicker')
    .setWidth(300)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '日付範囲を選択');
}

function processDateRange(startDate, endDate) {
  try {
    startAnalysis(new Date(startDate), new Date(endDate));
  } catch (error) {
    Logger.log('processDateRange エラー: ' + error);
  }
}

function startAnalysis(startDate, endDate) {
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('纏め');
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function(sheet) {
    var name = sheet.getName();
    return name !== '纏め' && name !== 'ログ' && name !== 'ドラッカー';
  });

  var formattedStart = formatDate(startDate);
  var formattedEnd = formatDate(endDate);

  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    updateProgress('分析中: ' + sheetName);  // 進捗表示
    analyzeSheet(sheets[i], formattedStart, formattedEnd);
    sortSheet(sheets[i]);
  }

  sortSummarySheet(summarySheet);
  applyBordersToDataRange(summarySheet);

  // 最後に分析完了を表示
  updateProgress('分析完了');  // B1セルに「分析完了」を表示
  SpreadsheetApp.getActiveSpreadsheet().toast('🎉 分析が終わりました', '完了通知', 5);  // トースト通知
}
function analyzeSheet(sheet, startDate, endDate) {
  try {
    var data = getDateRangeData(sheet, startDate, endDate);
    if (data.length > 0) {
      var results;
      try {
        results = analyzeWithGPT4Mini(data, sheet.getName(), startDate + ' から ' + endDate);
      } catch (e) {
        logToSheet('GPT分析失敗: ' + e.message);
        results = {
          担当者: sheet.getName(),
          月: startDate + ' から ' + endDate,
          time: new Date().toLocaleTimeString(),
          現状: 'エラー', 課題: 'エラー', 戦術: 'エラー', 行動特性: 'エラー', 強み: 'エラー'
        };
      }
      logToSheet('シート ' + sheet.getName() + ' の分析結果を出力中');
      logToSheet(JSON.stringify(results, null, 2));
      appendToSummary(results);
    }
  } catch (error) {
    Logger.log('analyzeSheet エラー: ' + error);
    logToSheet('analyzeSheet エラー: ' + error);
  }
}

function getDateRangeData(sheet, startDate, endDate) {
  var data = sheet.getRange('A2:G' + sheet.getLastRow()).getValues();
  var rangeData = [];
  var start = new Date(startDate), end = new Date(endDate);
  start.setHours(0,0,0,0);
  end.setHours(0,0,0,0);

  Logger.log('処理対象シート: ' + sheet.getName());
  logToSheet('処理対象シート: ' + sheet.getName());
  Logger.log('開始日: ' + start);
  logToSheet('開始日: ' + start);
  Logger.log('終了日: ' + end);
  logToSheet('終了日: ' + end);

  for (var i = 0; i < data.length; i++) {
    var rowDate = new Date(data[i][1]);
    if (!isNaN(rowDate.getTime())) {
      rowDate.setHours(0,0,0,0);
      if (rowDate >= start && rowDate <= end) {
        rangeData.push([data[i][0], data[i][4], data[i][5], data[i][6]]);
      }
    }
  }
  Logger.log('抽出された行数: ' + rangeData.length);
  logToSheet('抽出された行数: ' + rangeData.length);
  return rangeData;
}

function logToSheet(message) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ログ');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('ログ');
    sheet.appendRow(['タイムスタンプ', 'メッセージ']);
  }
  sheet.appendRow([new Date(), message]);
}

function formatDate(date) {
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getOpenAIApiKey() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('OPENAI_API_KEY');
}

function callGPT4Mini(apiKey, prompt) {
  var maxRetries = 3;
  var retryDelay = 5000;
  for (var i = 0; i < maxRetries; i++) {
    try {
      var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
        method: "post",
        headers: {
          "Authorization": "Bearer " + apiKey,
          "Content-Type": "application/json"
        },
        payload: JSON.stringify({
          model: "gpt-4o",
          messages: [
            {"role": "system", "content": "あなたは分析を行うAIアシスタントです。回答は必ず文章を最後まで完結させ、途中で切れないようにしてください。"},
            {"role": "user", "content": prompt}
          ],
          max_tokens: 1024,
          temperature: 0.7
        }),
        muteHttpExceptions: true
      });

      var responseCode = response.getResponseCode();
      if (responseCode === 200) {
        var responseText = JSON.parse(response.getContentText()).choices[0].message.content.trim();
        return responseText;
      } else {
        Logger.log('API呼び出しエラー: ' + responseCode + ' ' + response.getContentText());
        if (i === maxRetries - 1) {
          throw new Error('API呼び出しに失敗しました: ' + responseCode);
        }
      }
    } catch (error) {
      Logger.log('エラーが発生しました: ' + error.toString());
      if (i === maxRetries - 1) {
        throw error;
      }
    }
    Utilities.sleep(retryDelay);
  }
}

function appendToSummary(results) {
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('纏め');
  if (!summarySheet) return;

  var lastRow = summarySheet.getLastRow();
  var newRow = [
    results.担当者,
    results.月,
    results.time,
    results.現状,
    results.課題,
    results.戦術,
    results.行動特性,
    results.強み
  ];
  summarySheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
  logToSheet('appendToSummary 実行完了');
}

function analyzeWithGPT4Mini(data, sheetName, dateLabel) {
  var apiKey = getOpenAIApiKey();
  var results = {
    担当者: sheetName,
    月: dateLabel,
    time: new Date().toLocaleTimeString(),
    現状: '', 課題: '', 戦術: '', 行動特性: '', 強み: ''
  };

  var E_text = data.map(row => row[1]).join('\n');
  var F_text = data.map(row => row[2]).join('\n');
  var G_text = data.map(row => row[3]).join('\n');
var prompts = [
  {
    column: "現状",
    instruction: "E列の内容から、現在の取り組みの様子を3点まとめてください。\n・100文字以内、箇条書き＋改行\n・業務内容だけでなく、入力者が考えて動いた工夫や行動の背景も拾ってください\n・特に、日常の中で見えにくい『前向きな試み』や『創意工夫』があれば見逃さないでください\n・人名を使わず、やさしい表現で"
  },
  {
    column: "課題",
    instruction: "F列を分析し、明らかな問題だけでなく『構造的な原因』『現場で見過ごされがちな不便や悩み』も見つけてください。\n■本質的な課題 ×2　■気になる兆候 ×2\n・各項目100文字以内、箇条書き＋改行\n・放置した場合の影響や、組織全体にどう関わるかも考えてください\n・ユーザーの視点から『言語化されていない不安やギャップ』も見逃さずに"
  },
  {
    column: "戦術",
    instruction: "G列から、実践された工夫や提案をもとに、イノベーションにつながる新たな仮説や施策を3点まとめてください。\n・100文字以内、箇条書き＋改行\n・入力者が『なぜそれをやったか』に注目し、背景にある発想・動機・価値観を読み取り、そこから展開してください\n・改善点の指摘ではなく、『この行動は伸ばせば強みになる』という視点で"
  },
  {
    column: "行動特性",
    instruction: "E列の内容から、その人の思考・判断・行動の傾向を3点抽出してください。\n・形式：「【タイプ】〇〇型 - ○○○」\n・100文字以内、具体的な行動ベースで\n・思考のクセや判断のパターンが現れている行動をピックアップ\n・人名は使わず、やさしい言葉で"
  },
  {
    column: "強み",
    instruction: "E列の中から、本人が無意識に行っていることや実践したことの中に、成果につながる行動（強み）を見つけてください。\n・形式：「【タイプ】〇〇型 - ○○が得意で○○に貢献しています」\n・100文字以内、箇条書き＋改行\n・目立たないが価値のある動き、小さな行動の積み重ねなど『隠れた強み』に注目してください"
  },
  {
    column: "イノベーションの種",
    instruction: "E列・F列・G列の中から、入力者の中にある『気づき』『視点』『工夫』『発想』の中に、将来的に新しい価値につながる可能性のあることを3点抽出してください。\n・100文字以内、箇条書き＋改行\n・実践や仮説、日常の中の違和感や工夫に注目し、それを「もし育てたらどんな価値になるか」という視点で書いてください\n・本人が無意識のうちにやっていること、まだ言語化されていないことを拾い上げてください"
  }
];


  for (var i = 0; i < prompts.length; i++) {
  updateProgress(sheetName + ' の ' + prompts[i].column + ' を分析中');
  var columnData = '';
  switch (prompts[i].column) {
    case '現状':
    case '行動特性':
    case '強み':
      columnData = E_text;
      break;
    case '課題':
      columnData = F_text;
      break;
    case '戦術':
      columnData = G_text;
      break;
  }

  // ここで強みの処理が完了した後に、分析完了を表示
  if (prompts[i].column === '強み') {
    // 強み分析終了後に「分析完了」を表示
    updateProgress(sheetName + '分析完了');
  }


    var promptText = prompts[i].instruction + "\n\nデータ:\n" + columnData;
    var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, promptText).join('');
    var cache = PropertiesService.getScriptProperties().getProperty(hash);
    if (cache) {
      results[prompts[i].column] = cache;
      continue;
    }
    try {
      Logger.log('API呼び出し開始: ' + promptText.slice(0, 100));
      logToSheet('API呼び出し開始: ' + promptText.slice(0, 100));
      var response = callGPT4Mini(apiKey, promptText);
      results[prompts[i].column] = response;
      PropertiesService.getScriptProperties().setProperty(hash, response);
    } catch (error) {
      var errorMsg = prompts[i].column + ' 処理でエラー発生: ' + error.message;
      Logger.log(errorMsg);
      logToSheet(errorMsg);
      results[prompts[i].column] = 'エラー: ' + error.message;
    }
  }
  results.担当者 = data[0][0];
  return results;
}

function sortSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) return; // データがない場合はスキップ

  // A列昇順、B列降順で 3行目以降をソート
  sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).sort([
    { column: 1, ascending: true },   // A列 昇順
    { column: 2, ascending: false }   // B列 降順
  ]);
}
