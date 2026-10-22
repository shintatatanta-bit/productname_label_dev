function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('商品ラベル発行システム (A-one 31072)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * シートの初期化・入力規則の自動セットアップ
 * 単品マスタ: A列=レコード番号, B列=商品名(内部), C列=シール商品名, D列=シール商品名(英), E列=温め方
 */
function getMasterDataFromSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 単品マスタシートの取得・初期化
  let singleSheet = ss.getSheetByName('単品マスタ');
  if (!singleSheet) {
    singleSheet = ss.insertSheet('単品マスタ');
    singleSheet.appendRow(['レコード番号', '商品名', 'シール商品名', 'シール商品名（英語）', '温め方']);
    singleSheet.appendRow(['5040', '鶏のロースト 50g', '鶏のロースト', 'Roast Chicken', 'レンジ 600W で 3分温め']);
    singleSheet.appendRow(['5621', 'プチさつまいもとブロッコリー', 'プチさつまいもとブロッコリー', 'Sweet Potato & Broccoli', 'レンジ 600W で 2分温め']);
    singleSheet.setColumnWidth(1, 110);
    singleSheet.setColumnWidth(2, 280);
    singleSheet.setColumnWidth(3, 240);
    singleSheet.setColumnWidth(4, 200);
    singleSheet.setColumnWidth(5, 220);
  }

  // 2. セットマスタシートの取得・初期化
  let setSheet = ss.getSheetByName('セットマスタ');
  if (!setSheet) {
    setSheet = ss.insertSheet('セットマスタ');
    setSheet.appendRow(['セット名', '品目1（プルダウン選択）', '品目2（プルダウン選択）', '品目3（プルダウン選択）', '品目4（プルダウン選択）', '温め方指定', 'タイプ']);
    setSheet.appendRow(['定番鶏セット', '鶏のロースト', 'プチさつまいもとブロッコリー', '', '', 'レンジ 600W で 3分温め', 'warm']);
  }

  // 3. セットマスタのB2:E200に、単品マスタの「C列（シール商品名）」から選ぶデータの入力規則を付与
  const singleLabelRange = singleSheet.getRange('C2:C500');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(singleLabelRange, true)
    .setAllowInvalid(false)
    .setHelpText('単品マスタのシール商品名から選択してください。')
    .build();
  
  setSheet.getRange('B2:E200').setDataValidation(rule);

  // 4. データの読み込み
  const singleVals = singleSheet.getDataRange().getValues();
  const singleItems = [];
  const singleMapByName = {};

  for (let r = 1; r < singleVals.length; r++) {
    const row = singleVals[r];
    const recordNo = String(row[0] || '').trim();
    const internalName = String(row[1] || '').trim(); // B列: 商品名（内部）
    const labelJa = String(row[2] || '').trim();      // C列: シール商品名（印字用）
    const labelEn = String(row[3] || '').trim();      // D列: シール商品名（英語）
    const method = String(row[4] || 'レンジ 600W で 2分温め').trim();

    const finalNameJa = labelJa || internalName;
    if (!finalNameJa) continue;

    const item = {
      recordNo: recordNo,
      internalName: internalName,
      nameJa: finalNameJa,
      nameEn: labelEn,
      method: method,
      type: method.includes('冷') ? 'cold' : 'warm'
    };

    singleItems.push(item);

    // 同名のシール商品名や内部名が複数行ある場合、英語名(nameEn)が設定されている行を優先して辞書に登録
    if (!singleMapByName[item.nameJa] || (!singleMapByName[item.nameJa].nameEn && item.nameEn)) {
      singleMapByName[item.nameJa] = item;
    }
    if (internalName) {
      if (!singleMapByName[internalName] || (!singleMapByName[internalName].nameEn && item.nameEn)) {
        singleMapByName[internalName] = item;
      }
    }
  }

  // ② セットマスタ取得
  const setVals = setSheet.getDataRange().getValues();
  const setPresets = [];

  for (let r = 1; r < setVals.length; r++) {
    const row = setVals[r];
    const setName = String(row[0] || '').trim();
    if (!setName) continue;

    const prodsJa = [row[1], row[2], row[3], row[4]]
      .map(s => String(s || '').trim())
      .filter(Boolean);

    // 単品マスタ辞書から英語名を自動引き当て
    const prodsEn = prodsJa.map(ja => (singleMapByName[ja] && singleMapByName[ja].nameEn) ? singleMapByName[ja].nameEn : '');

    const method = String(row[5] || '').trim() || (prodsJa[0] && singleMapByName[prodsJa[0]] ? singleMapByName[prodsJa[0]].method : 'レンジ 600W で 2分温め');
    const type = String(row[6] || '').trim().toLowerCase() === 'cold' || method.includes('冷') ? 'cold' : 'warm';

    setPresets.push({
      setName: setName,
      prodsJa: prodsJa,
      prodsEn: prodsEn,
      methodText: method,
      methodType: type
    });
  }

// ★★★ ここから追加：シノニム設定シートの取得・初期化と読み込み ★★★
  let synonymSheet = ss.getSheetByName('シノニム設定');
  if (!synonymSheet) {
    synonymSheet = ss.insertSheet('シノニム設定');
    synonymSheet.appendRow(['代表表記', '同義語・別表記（カンマ区切り）', '備考']);
    synonymSheet.appendRow(['鶏', '鳥, とり, トリ, チキン', '鶏肉関連']);
    synonymSheet.appendRow(['豚', 'ぶた, ブタ, ポーク', '豚肉関連']);
    synonymSheet.appendRow(['牛', 'うし, ウシ, ビーフ', '牛肉関連']);
    synonymSheet.appendRow(['鮭', 'さけ, サケ, サーモン', '魚介関連']);
    synonymSheet.appendRow(['鯖', 'さば, サバ', '魚介関連']);
    synonymSheet.appendRow(['芋', 'いも, イモ, ポテト', '野菜・芋類']);
    synonymSheet.setColumnWidth(1, 120);
    synonymSheet.setColumnWidth(2, 280);
    synonymSheet.setColumnWidth(3, 180);
  }

  const synVals = synonymSheet.getDataRange().getValues();
  const synonymGroups = [];

  for (let r = 1; r < synVals.length; r++) {
    const row = synVals[r];
    const canonical = String(row[0] || '').trim();
    if (!canonical) continue;

    const wordsStr = String(row[1] || '').trim();
    // カンマ、読点、スペースで分割
    const words = wordsStr.split(/[,、\s]+/).map(w => w.trim()).filter(Boolean);
    
    // 代表語自身もグループに含め、重複を除去
    const groupSet = new Set([canonical, ...words]);
    synonymGroups.push(Array.from(groupSet));
  }
  // ★★★ 追加ここまで ★★★

// 5. 一時保存データ一覧の取得
  const savedDrafts = getDraftListFromSpreadsheet();

  return {
    singleItems: singleItems,
    setPresets: setPresets,
    savedDrafts: savedDrafts,
    synonymGroups: synonymGroups // ★ここを追加してブラウザへ渡す
  };
}

/**
 * 一時保存シートの取得・初期化ヘルパー（5列仕様）
 * 列構成: A列=保存名, B列=作業者名, C列=最終更新日時, D列=作成枚数, E列=JSONデータ
 */
function getOrCreateDraftSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let draftSheet = ss.getSheetByName('一時保存データ');
  if (!draftSheet) {
    draftSheet = ss.insertSheet('一時保存データ');
    draftSheet.appendRow(['保存名（案件・顧客名）', '作業者名', '最終更新日時', '作成枚数', 'スロットJSONデータ（編集厳禁）']);
    draftSheet.setColumnWidth(1, 220);
    draftSheet.setColumnWidth(2, 120);
    draftSheet.setColumnWidth(3, 150);
    draftSheet.setColumnWidth(4, 90);
    draftSheet.setColumnWidth(5, 300);
  } else {
    // 既存の4列仕様シートが存在する場合は、安全にB列へ「作業者名」列を挿入して5列仕様へ自動移行する
    const headerB = String(draftSheet.getRange(1, 2).getValue() || '').trim();
    if (headerB === '最終更新日時') {
      draftSheet.insertColumnBefore(2);
      draftSheet.getRange(1, 2).setValue('作業者名');
      draftSheet.setColumnWidth(2, 120);
    }
  }
  return draftSheet;
}

/**
 * 一時保存データの一覧を取得（作業者名を含む）
 */
function getDraftListFromSpreadsheet() {
  const sheet = getOrCreateDraftSheet();
  const vals = sheet.getDataRange().getValues();
  const list = [];

  for (let r = 1; r < vals.length; r++) {
    const title = String(vals[r][0] || '').trim();
    const worker = String(vals[r][1] || '未設定').trim();
    const updatedAt = vals[r][2] instanceof Date
      ? Utilities.formatDate(vals[r][2], 'Asia/Tokyo', 'MM/dd HH:mm')
      : String(vals[r][2] || '');
    const count = Number(vals[r][3] || 0);
    const jsonStr = String(vals[r][4] || '');

    if (!title || !jsonStr) continue;
    list.push({
      title: title,
      worker: worker,
      updatedAt: updatedAt,
      count: count,
      jsonStr: jsonStr
    });
  }
  // 新しいものが上に来るよう逆順にする
  return list.reverse();
}

/**
 * 作業途中の36枠データをスプレッドシートへ保存
 * 挙動B: 保存名が既存と一致すれば上書き、保存名が変更されていれば別名として新規行に追加
 */
function saveDraftToSpreadsheet(title, worker, jsonStr, usedCount) {
  const cleanTitle = String(title || '').trim();
  const cleanWorker = String(worker || '').trim() || '未設定';
  if (!cleanTitle) {
    throw new Error('保存名が空です。');
  }

  const sheet = getOrCreateDraftSheet();
  const vals = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  let targetRow = -1;
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][0] || '').trim() === cleanTitle) {
      targetRow = r + 1; // 1始まりの行番号
      break;
    }
  }

  if (targetRow > 0) {
    // 保存名が同じ場合は既存の行を上書き更新（B列:作業者名〜E列:JSON）
    sheet.getRange(targetRow, 2, 1, 4).setValues([[cleanWorker, nowStr, usedCount, jsonStr]]);
  } else {
    // 保存名が新しい場合は元のデータを残して新規行として追加（挙動B）
    sheet.appendRow([cleanTitle, cleanWorker, nowStr, usedCount, jsonStr]);
  }

  return getDraftListFromSpreadsheet();
}

/**
 * 指定した保存名のデータを削除
 */
function deleteDraftFromSpreadsheet(title) {
  const cleanTitle = String(title || '').trim();
  const sheet = getOrCreateDraftSheet();
  const vals = sheet.getDataRange().getValues();

  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][0] || '').trim() === cleanTitle) {
      sheet.deleteRow(r + 1);
      break;
    }
  }
  return getDraftListFromSpreadsheet();
}

/**
 * Webアプリから単品マスタへ新規商品を追加する関数
 * @param {Object} itemData - { recordNo, internalName, nameJa, nameEn, method }
 * @return {Object} 更新後のマスタデータ一式
 */
function addNewSingleMasterItem(itemData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('単品マスタ');
  if (!sheet) {
    throw new Error('「単品マスタ」シートが見つかりません。');
  }

  const recordNo = String(itemData.recordNo || '').trim();
  const internalName = String(itemData.internalName || '').trim();
  const nameJa = String(itemData.nameJa || '').trim();
  const nameEn = String(itemData.nameEn || '').trim();
  const method = String(itemData.method || 'レンジ 600W で 2分温め').trim();

  // シール商品名または内部商品名のいずれかは必須
  if (!nameJa && !internalName) {
    throw new Error('シール商品名または商品名を入力してください。');
  }

  // A列:レコード番号, B列:商品名, C列:シール商品名, D列:英語名, E列:温め方
  sheet.appendRow([recordNo, internalName, nameJa, nameEn, method]);

  // 追加後の最新マスタデータを取得して返却（画面側の即時同期用）
  return getMasterDataFromSpreadsheet();
}

