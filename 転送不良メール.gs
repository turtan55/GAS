/**
 * スプレッドシート開いた時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('メールチェック')
    .addItem('転送不良メールチェック実行', 'checkMissingIDsToSheet')
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * ヘルプダイアログを表示
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `
【転送不良メールチェック機能】

このツールは、テコスNAVIのお見積りメールから
ID の欠番をチェックします。

使い方：
1. メニューから「メールチェック」→「転送不良メールチェック実行」を選択
2. 結果が「抜けチェック」シートに出力されます

注意事項：
- 直近100件のメールを対象とします
- 13桁のIDから下4桁を抽出して連続性をチェックします
- 処理には数秒かかる場合があります
  `;
  
  ui.alert('ヘルプ', helpText, ui.ButtonSet.OK);
}

/**
 * 転送不良メールのIDチェック機能
 */
function checkMissingIDsToSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // 実行確認ダイアログ
  const response = ui.alert(
    '転送不良メールチェック',
    'メールチェックを実行しますか？\n（処理に数秒かかる場合があります）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return; // キャンセルされた場合は終了
  }
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("抜けチェック") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("抜けチェック");
    sheet.clear(); // 初期化
    sheet.getRange("A1").setValue("抜けID（下4桁）");

    // より柔軟な検索クエリ - 角括弧とテコスNAVIを含む件名を検索
    const query = 'subject:"[" subject:"テコスNAVI"';
    const threads = GmailApp.search(query, 0, 100); // 直近100件
    const idSet = new Set();
    // より柔軟な正規表現 - 文字列の途中でもマッチするように修正
    const pattern = /\[(\d{13})\].*テコスNAVIお見積りを承りました。/;

    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const msg of messages) {
        const subject = msg.getSubject();
        if (pattern.test(subject)) {
          const id = subject.match(pattern)[1];
          idSet.add(id);
          break; // そのスレッドから1件で十分
        }
      }
    }

    const ids = Array.from(idSet).map(id => parseInt(id.slice(-4))).sort((a, b) => a - b);
    
    if (ids.length === 0) {
      sheet.getRange("A2").setValue("対象メールが見つかりませんでした");
      ui.alert('完了', '対象メールが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    const minID = ids[0];
    const maxID = ids[ids.length - 1];
    const missing = [];

    for (let i = minID; i <= maxID; i++) {
      if (!ids.includes(i)) missing.push(i);
    }

    // スプレッドシートに出力
    if (missing.length > 0) {
      sheet.getRange(2, 1, missing.length, 1).setValues(missing.map(id => [id]));
      ui.alert('完了', `処理が完了しました。\n${missing.length}件の抜けIDが見つかりました。`, ui.ButtonSet.OK);
    } else {
      sheet.getRange("A2").setValue("抜けはありません");
      ui.alert('完了', '処理が完了しました。\n抜けIDはありませんでした。', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('エラー', `処理中にエラーが発生しました：\n${error.message}`, ui.ButtonSet.OK);
    console.error('Error in checkMissingIDsToSheet:', error);
  }
}
