
/**
 * 工具使用履歴・寿命管理システム
 * ver 1.0.0 (2026/09/10 リリース)
 * 修正履歴:
 *  - 1.0.0: 初期バージョン作成（日報入力、工具交換リセット、ログ自動蓄積機能）
 */





function onOpen() {
  // スプレッドシートを開いたときに、専用メニューを追加します
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ 工具管理システム')
      .addItem('1. 初期セットアップ（最初のみ実行）', 'setupSystem')
      .addSeparator()
      .addItem('📝 加工を記録する', 'recordProduction')
      .addItem('🔄 工具交換を記録する', 'replaceTool')
      .addToUi();
}

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. メインシートの作成
  let mainSheet = ss.getSheetByName("メイン");
  if (!mainSheet) {
    mainSheet = ss.insertSheet("メイン");
  } else {
    const res = SpreadsheetApp.getUi().alert("確認", "「メイン」シートを初期化して上書きしてもよろしいですか？", SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (res != SpreadsheetApp.getUi().Button.YES) return;
    mainSheet.clear();
  }
  
  // --- UI（入力画面）の構築 ---
  mainSheet.getRange("B2").setValue("【現在の工具ステータス】").setFontWeight("bold").setFontSize(14);
  mainSheet.getRange("C3:F3").setValues([["工具位置", "工具1", "工具2", "工具3"]]).setBackground("#d9ead3").setFontWeight("bold");
  mainSheet.getRange("C4:F4").setValues([["現在のLot", "Lot-A", "Lot-B", "Lot-C"]]);
  mainSheet.getRange("C5:F5").setValues([["累計加工数", 0, 0, 0]]);
  
  mainSheet.getRange("B8").setValue("【日々の加工記録】").setFontWeight("bold").setFontSize(14);
  mainSheet.getRange("C9:D11").setValues([
    ["日付", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd")],
    ["作業者", ""],
    ["加工個数", ""]
  ]);
  mainSheet.getRange("C9:C11").setBackground("#cfe2f3").setFontWeight("bold");
  mainSheet.getRange("E11").setValue("← 個数を入力したら「📝 加工を記録する」を実行");
  
  mainSheet.getRange("B15").setValue("【工具交換の記録】").setFontWeight("bold").setFontSize(14);
  mainSheet.getRange("C16:D18").setValues([
    ["交換する位置", "工具1"],
    ["新しいLot番号", ""],
    ["交換理由(任意)", "摩耗"]
  ]);
  mainSheet.getRange("C16:C18").setBackground("#fce5cd").setFontWeight("bold");
  mainSheet.getRange("E18").setValue("← 入力したら「🔄 工具交換を記録する」を実行");
  
  // プルダウンの設定
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(["工具1", "工具2", "工具3"]).build();
  mainSheet.getRange("D16").setDataValidation(rule);
  
  // 罫線と幅の調整
  mainSheet.getRange("C3:F5").setBorder(true, true, true, true, true, true);
  mainSheet.getRange("C9:D11").setBorder(true, true, true, true, true, true);
  mainSheet.getRange("C16:D18").setBorder(true, true, true, true, true, true);
  mainSheet.setColumnWidth(3, 120);
  mainSheet.setColumnWidth(4, 150);
  mainSheet.setColumnWidth(5, 150);
  mainSheet.setColumnWidth(6, 150);
  
  // 2. 加工ログシート（日報）の作成
  let logSheet = ss.getSheetByName("加工ログ");
  if (!logSheet) {
    logSheet = ss.insertSheet("加工ログ");
    logSheet.appendRow(["システム日時", "入力日付", "作業者", "追加した加工数", "工具1 Lot", "工具1 累計", "工具2 Lot", "工具2 累計", "工具3 Lot", "工具3 累計"]);
    logSheet.getRange("A1:J1").setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    logSheet.setFrozenRows(1);
  }
  
  // 3. 交換履歴シート（トレーサビリティ）の作成
  let exchangeSheet = ss.getSheetByName("交換履歴");
  if (!exchangeSheet) {
    exchangeSheet = ss.insertSheet("交換履歴");
    exchangeSheet.appendRow(["交換日時", "交換した位置", "古いLot番号", "最終累計(寿命)", "新しいLot番号", "交換理由"]);
    exchangeSheet.getRange("A1:F1").setBackground("#e69138").setFontColor("white").setFontWeight("bold");
    exchangeSheet.setFrozenRows(1);
  }
  
  SpreadsheetApp.getUi().alert("✅ 画面のセットアップが完了しました！\nまずは各工具の「現在のLot」と「累計加工数」を実際の数字に書き換えて運用をスタートしてください。");
}

function recordProduction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("メイン");
  const logSheet = ss.getSheetByName("加工ログ");
  
  if (!mainSheet || !logSheet) return;
  
  const dateStr = mainSheet.getRange("D9").getDisplayValue();
  const worker = mainSheet.getRange("D10").getValue();
  const count = mainSheet.getRange("D11").getValue();
  
  if (!count || isNaN(count) || count <= 0) {
    SpreadsheetApp.getUi().alert("⚠️ 加工個数に正しい数値を入力してください。");
    return;
  }
  
  // 3つの工具の現在の累計を取得して足し算
  const tool1Lot = mainSheet.getRange("D4").getValue();
  let tool1Cumul = Number(mainSheet.getRange("D5").getValue()) || 0;
  const tool2Lot = mainSheet.getRange("E4").getValue();
  let tool2Cumul = Number(mainSheet.getRange("E5").getValue()) || 0;
  const tool3Lot = mainSheet.getRange("F4").getValue();
  let tool3Cumul = Number(mainSheet.getRange("F5").getValue()) || 0;
  
  tool1Cumul += Number(count);
  tool2Cumul += Number(count);
  tool3Cumul += Number(count);
  
  // メイン画面の累計を更新
  mainSheet.getRange("D5").setValue(tool1Cumul);
  mainSheet.getRange("E5").setValue(tool2Cumul);
  mainSheet.getRange("F5").setValue(tool3Cumul);
  
  // ログに記録
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  logSheet.appendRow([timestamp, dateStr, worker, count, tool1Lot, tool1Cumul, tool2Lot, tool2Cumul, tool3Lot, tool3Cumul]);
  
  // 個数欄だけクリア（次回入力しやすくするため）
  mainSheet.getRange("D11").clearContent();
  SpreadsheetApp.getUi().alert("✅ " + count + "個の加工を記録し、3本の工具の累計を増やしました。");
}

function replaceTool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("メイン");
  const exchangeSheet = ss.getSheetByName("交換履歴");
  
  if (!mainSheet || !exchangeSheet) return;
  
  const targetPos = mainSheet.getRange("D16").getValue(); 
  const newLot = mainSheet.getRange("D17").getValue();
  const memo = mainSheet.getRange("D18").getValue();
  
  if (!targetPos || !newLot) {
    SpreadsheetApp.getUi().alert("⚠️ 「交換する位置」と「新しいLot番号」を入力してください。");
    return;
  }
  
  let oldLotRange, oldCumulRange;
  if (targetPos === "工具1") { oldLotRange = mainSheet.getRange("D4"); oldCumulRange = mainSheet.getRange("D5"); }
  else if (targetPos === "工具2") { oldLotRange = mainSheet.getRange("E4"); oldCumulRange = mainSheet.getRange("E5"); }
  else if (targetPos === "工具3") { oldLotRange = mainSheet.getRange("F4"); oldCumulRange = mainSheet.getRange("F5"); }
  else { return; }
  
  const oldLot = oldLotRange.getValue();
  const oldCumul = oldCumulRange.getValue();
  
  // 交換履歴に記録
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  exchangeSheet.appendRow([timestamp, targetPos, oldLot, oldCumul, newLot, memo]);
  
  // メイン画面のLotを書き換え、累計を0にリセット
  oldLotRange.setValue(newLot);
  oldCumulRange.setValue(0);
  
  // 入力欄をクリア
  mainSheet.getRange("D17:D18").clearContent();
  SpreadsheetApp.getUi().alert("✅ " + targetPos + " を新しいLot (" + newLot + ") に交換し、累計をリセットしました。\n（前のLot " + oldLot + " は " + oldCumul + " 個で寿命を迎えました）");
}

// =========================================
// 手順書とグラフシートを追加する機能
// =========================================
function addInstructionAndCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 手順シート（マニュアル）の追加
  let manualSheet = ss.getSheetByName("手順書");
  if (!manualSheet) {
    manualSheet = ss.insertSheet("手順書", 0); // 最初のシート（一番左）として追加
    
    // タイトル
    manualSheet.getRange("B2").setValue("【工具管理システム 操作マニュアル】").setFontWeight("bold").setFontSize(16);
    
    // 日々の作業
    manualSheet.getRange("B4").setValue("■ 日々の作業（加工を記録する）").setFontWeight("bold").setFontSize(14);
    manualSheet.getRange("B5:C8").setValues([
      ["手順1", "「メイン」シートを開きます。"],
      ["手順2", "「日々の加工記録」欄に、今日の日付、作業者名、加工個数を入力します。"],
      ["手順3", "「📝 加工を記録」ボタンを押します。"],
      ["確認", "「加工ログ」シートに記録が追加され、各工具の累計加工数が増えていることを確認します。"]
    ]);
    
    // 工具交換時の作業
    manualSheet.getRange("B10").setValue("■ 工具交換時の作業（寿命を記録してリセットする）").setFontWeight("bold").setFontSize(14);
    manualSheet.getRange("B11:C14").setValues([
      ["手順1", "「メイン」シートを開きます。"],
      ["手順2", "「工具交換の記録」欄で、交換した工具の位置をプルダウンから選びます。"],
      ["手順3", "新しい工具のLot番号を入力します（理由は任意）。"],
      ["手順4", "「🔄 工具交換を記録」ボタンを押します。"]
    ]);
    
    // その他の注意事項
    manualSheet.getRange("B16").setValue("■ 注意事項").setFontWeight("bold").setFontSize(14);
    manualSheet.getRange("B17:C19").setValues([
      ["注意1", "色のついていないセル（現在の工具ステータス等）は自動で書き換わるため、手入力しないでください。"],
      ["注意2", "間違ってボタンを押してしまった場合は、ログシートの行を直接削除し、メインシートの累計を手動で修正してください。"],
      ["", ""]
    ]);
    
    // 見栄えの調整
    manualSheet.getRange("B5:C8").setBorder(true, true, true, true, true, true);
    manualSheet.getRange("B11:C14").setBorder(true, true, true, true, true, true);
    manualSheet.getRange("B17:C18").setBorder(true, true, true, true, true, true);
    manualSheet.setColumnWidth(2, 60);
    manualSheet.setColumnWidth(3, 600);
    manualSheet.getRange("B5:B8").setBackground("#fff2cc");
    manualSheet.getRange("B11:B14").setBackground("#fce5cd");
    manualSheet.getRange("B17:B18").setBackground("#f4cccc");
  }

  // 2. グラフ（ダッシュボード）シートの準備
  let chartSheet = ss.getSheetByName("グラフ(傾向分析)");
  if (!chartSheet) {
    chartSheet = ss.insertSheet("グラフ(傾向分析)");
    
    chartSheet.getRange("B2").setValue("【工具 寿命傾向分析 ダッシュボード】").setFontWeight("bold").setFontSize(16);
    chartSheet.getRange("B4").setValue("ここに、工具が『何個で寿命を迎えたか』のグラフを配置します。");
    chartSheet.getRange("B5").setValue("データ（交換履歴）が数件〜10件ほど溜まってきたら、グラフを作成してください。");
    
    // グラフ作成のヒントを記載
    chartSheet.getRange("B7").setValue("■ おすすめのグラフの作り方").setFontWeight("bold");
    chartSheet.getRange("B8:C10").setValues([
      ["1", "「交換履歴」シートを開き、C列(古いLot)とD列(最終累計)などを選択します。"],
      ["2", "上部メニューの「挿入」＞「グラフ」をクリックします。"],
      ["3", "グラフの種類を『縦棒グラフ』などにし、このシートに移動させると見やすくなります。"]
    ]);
  }
  
  SpreadsheetApp.getUi().alert("✅ 「手順書」シートと「グラフ(傾向分析)」シートを追加しました。");
}
