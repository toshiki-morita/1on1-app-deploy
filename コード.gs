/**
 * @file 1on1ヒアリング支援アプリ バックエンド処理
 */

// 管理用スプレッドシートのIDを定数として定義します。
const SPREADSHEET_ID = '1oeNGXFXgxqKCrP86Tps8mFTg2K21hglpN8_DLOJqGfU';

// gidが無い場合に優先表示するシート名（先頭ほど優先）
// 例: リンク先が常に「List」タブである場合にここへ 'List' を設定
const PREFERRED_SHEET_NAMES = ['List'];

/**
 * Webアプリにアクセスした際に呼び出されるメイン関数です。
 * index.htmlを画面に表示します。
 * @param {object} e - イベントオブジェクト（GASから自動的に渡されます）
 * @returns {HtmlOutput} 画面に表示するHTMLオブジェクト
 */
function doGet(e) {
  // 1行ごとにコメントを追加
  return HtmlService.createTemplateFromFile('index').evaluate(); // index.htmlをテンプレートとして評価し、表示します。
}

/**
 * 指定されたスプレッドシートに含まれるすべてのシート名を取得して返します。
 * @returns {string[] | {error: string}} シート名の配列、またはエラーオブジェクト
 */
function getSheetNames() {
  try {
    // 指定されたIDを使ってスプレッドシートを開きます。
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    // すべてのシートを取得します。
    const sheets = spreadsheet.getSheets();
    // 各シートの名前をmap関数で取り出し、新しい配列を作成して返します。
    return sheets.map(sheet => sheet.getName());
  } catch (e) {
    // エラーが発生した場合、内容をログに出力します。
    console.error('シート名の取得に失敗しました: ' + e.toString());
    // フロントエンド側でエラーハンドリングができるように、エラー情報を返します。
    return { error: 'シート名の取得に失敗しました。スプレッドシートIDや権限を確認してください。' };
  }
}

/**
 * Listシート(A列=会社名, B列=支店名[リンク付])からドロップダウン用データを作成します。
 * ラベルは「会社名 支店名」、valueはB列セルのリンクURL。
 * B列リンクは RichText のリンク > HYPERLINK関数 > 素のURL の順に抽出します。
 * @returns {{label:string, url:string}[] | {error:string}}
 * @example
 * // クライアント:
 * google.script.run.withSuccessHandler(console.log).getBranchLinks();
 */
function getBranchLinks() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('List');
    if (!sheet) return { error: 'Listシートが見つかりません。' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; // データ無し

    const n = lastRow - 1; // データ行数
    const rangeA = sheet.getRange(2, 1, n, 1); // A2:A 会社名
    const rangeB = sheet.getRange(2, 2, n, 1); // B2:B 支店名(リンク)

    const companyVals = rangeA.getDisplayValues();
    const bDisplays = rangeB.getDisplayValues();
    const bRich = rangeB.getRichTextValues();
    const bFormulas = rangeB.getFormulas();

    const items = [];
    for (let i = 0; i < n; i++) {
      const company = (companyVals[i][0] || '').trim();
      const branchDisplay = (bDisplays[i][0] || '').trim();
      if (!company && !branchDisplay) continue; // 空行スキップ

      let url = '';
      // 1) RichText のリンク（セル全体 or 部分リンク）
      const r = bRich[i][0];
      if (r) {
        const linkUrl = r.getLinkUrl && r.getLinkUrl();
        if (linkUrl) url = linkUrl;
        if (!url && r.getRuns) {
          const runs = r.getRuns();
          for (const run of runs) {
            const u = run.getLinkUrl && run.getLinkUrl();
            if (u) { url = u; break; }
          }
        }
      }
      // 2) HYPERLINK 関数: =HYPERLINK("url", "text")
      if (!url) {
        const f = bFormulas[i][0] || '';
        const m = f.match(/HYPERLINK\("([^"]+)"/i);
        if (m && m[1]) url = m[1];
      }
      // 3) 表示値がそのままURL
      if (!url && /^https?:\/\//i.test(branchDisplay)) url = branchDisplay;

      if (url) {
        const label = (company ? company + ' ' : '') + branchDisplay;
        items.push({ label, url });
      }
    }
    return items;
  } catch (e) {
    console.error('List!Bリンク取得失敗: ' + e.toString());
    return { error: 'List!Bのリンク取得に失敗しました。ID/権限/データをご確認ください。' };
  }
}

/**
 * 指定URLのスプレッドシートから表データを取得します。
 * URLから spreadsheetId と gid(任意) を抽出し、対象シートを特定して返却します。
 * @param {string} url - docs.google.com/spreadsheets のURL
 * @returns {{sheetName:string, headers:string[], rows:string[][]} | {error:string}}
 * @example
 * google.script.run.withSuccessHandler(console.log).getSheetDataFromUrl(url);
 */
function getSheetDataFromUrl(url) {
  try {
    const parsed = parseSpreadsheetUrl(url);
    if (!parsed || !parsed.spreadsheetId) {
      return { error: 'URLからスプレッドシートIDを抽出できませんでした。' };
    }

    const ss = SpreadsheetApp.openById(parsed.spreadsheetId);
    let sheet = null;
    if (parsed.gid != null) {
      sheet = findSheetByGid_(ss, parsed.gid);
    }
    if (!sheet) {
      // gidが無い/見つからない場合は候補名を優先
      const byName = ss.getSheets().find(sh => (PREFERRED_SHEET_NAMES || []).includes(sh.getName())) || null;
      sheet = byName || ss.getSheets()[0] || null;
    }
    if (!sheet) return { error: '対象シートを特定できませんでした。' };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      return { sheetName: sheet.getName(), headers: [], rows: [] };
    }
    const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = values[0] || [];
    const rows = values.slice(1);
    return { sheetName: sheet.getName(), headers, rows };
  } catch (e) {
    console.error('シートデータ取得失敗: ' + e.toString());
    return { error: 'シートデータの取得に失敗しました。URL/権限をご確認ください。' };
  }
}

/**
 * スプレッドシートURLからIDとgid(任意)を抽出します。
 * 例: https://docs.google.com/spreadsheets/d/{ID}/edit#gid=123456
 * @param {string} url
 * @returns {{spreadsheetId:string, gid:number|null} | null}
 */
function parseSpreadsheetUrl(url) {
  if (!url) return null;
  const idMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) return null;
  const spreadsheetId = idMatch[1];
  let gid = null;
  const gidMatch = url.match(/[?#&]gid=(\d+)/);
  if (gidMatch) gid = Number(gidMatch[1]);
  return { spreadsheetId, gid };
}

/**
 * gidで対象シートを検索します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {number} gid
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null}
 */
function findSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (sh.getSheetId && sh.getSheetId() === gid) return sh;
  }
  return null;
}

// 今後、指定されたシートの表を編集・保存する関数（セル更新/バリデーション再現など）を追加していきます。

/**
 * 指定URLから対象シートを解決して返します。
 * gidが無い場合、PREFERRED_SHEET_NAMESを優先し、それも無ければ先頭シートを返します。
 * @param {string} url - スプレッドシートURL
 * @returns {{ss: GoogleAppsScript.Spreadsheet.Spreadsheet, sheet: GoogleAppsScript.Spreadsheet.Sheet}}
 * @throws {Error} URL不正/権限不足/シート未特定時
 */
function resolveSheetFromUrl_(url) {
  const parsed = parseSpreadsheetUrl(url);
  if (!parsed || !parsed.spreadsheetId) {
    throw new Error('URLからスプレッドシートIDを抽出できませんでした。');
  }
  const ss = SpreadsheetApp.openById(parsed.spreadsheetId);
  let sheet = null;
  if (parsed.gid != null) {
    sheet = findSheetByGid_(ss, parsed.gid);
  }
  if (!sheet) {
    const byName = ss.getSheets().find(sh => (PREFERRED_SHEET_NAMES || []).includes(sh.getName())) || null;
    sheet = byName || ss.getSheets()[0] || null;
  }
  if (!sheet) throw new Error('対象シートを特定できませんでした。');
  return { ss, sheet };
}

/**
 * 指定のシートURL生成（対象シートのgid付き）。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {string}
 */
function buildSheetUrl_(ss, sheet) {
  return `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${sheet.getSheetId()}`;
}

/**
 * 指定URLのシートから、表示値の表データに加えて「プルダウン候補」や「日付カラムのISO値」を含むスキーマ情報を返します。
 * - プルダウン: DataValidation(リスト/範囲)を検出し、候補配列を返却
 * - 日付: DataValidation(各種DATE系) or 数値フォーマットの推定、さらに getValues() のDateを ISO(yyyy-MM-dd) で返却
 * @param {string} url
 * @returns {{
 *   sheetName:string,
 *   headers:string[],
 *   rows:string[][],
 *   dropdownByCol: {[colIndex:number]: string[]},
 *   dateIsoByCol: {[colIndex:number]: string[]}
 * } | {error:string}}
 */
function getSheetDataWithSchemaFromUrl(url) {
  try {
    const { ss, sheet } = resolveSheetFromUrl_(url);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      return { sheetName: sheet.getName(), headers: [], rows: [], dropdownByCol: {}, dateIsoByCol: {}, sheetUrl: buildSheetUrl_(ss, sheet), linksByCol: {}, checkboxByCol: {}, checkboxStateByCol: {} };
    }

    // 表示値/生値を取得
    const displayValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const rawValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = displayValues[0] || [];
    const rows = displayValues.slice(1);

    const dataRowCount = Math.max(0, lastRow - 1);
    const dataRange = dataRowCount > 0 ? sheet.getRange(2, 1, dataRowCount, lastCol) : null;
    const validations = dataRange ? dataRange.getDataValidations() : [];
    const numberFormats = dataRange ? dataRange.getNumberFormats() : [];

    /** @type {{[col:number]: string[]}} */
    const dropdownByCol = {};
    /** @type {{[col:number]: string[]}} */
    const dateIsoByCol = {};
    /** @type {{[col:number]: true}} */
    const checkboxByCol = {};
    /** @type {{[col:number]: boolean[]}} */
    const checkboxStateByCol = {};

    const DVC = SpreadsheetApp.DataValidationCriteria;
    const tz = Session.getScriptTimeZone();

    // 各列のバリデーション/書式を走査
    for (let col = 0; col < lastCol; col++) {
      let foundDropdown = false;
      let foundDate = false;
      let foundCheckbox = false;
      let dropdownOptions = null;

      for (let r = 0; r < dataRowCount; r++) {
        const v = validations && validations[r] ? validations[r][col] : null;
        if (!v) continue;
        const t = v.getCriteriaType && v.getCriteriaType();
        // 日付系
        if (!foundDate && t && String(t).indexOf('DATE') >= 0) {
          foundDate = true;
        }
        // チェックボックス
        if (!foundCheckbox && t && t === DVC.CHECKBOX) {
          foundCheckbox = true;
        }
        // プルダウン（リスト/範囲）
        if (!foundDropdown && t && (t === DVC.VALUE_IN_LIST || t === DVC.VALUE_IN_RANGE)) {
          const cv = v.getCriteriaValues ? v.getCriteriaValues() : [];
          if (t === DVC.VALUE_IN_LIST) {
            const list = (cv && cv[0]) ? cv[0] : [];
            if (Array.isArray(list)) {
              dropdownOptions = list.map(x => (x == null ? '' : String(x)));
              foundDropdown = true;
            }
          } else if (t === DVC.VALUE_IN_RANGE) {
            const refRange = cv && cv[0];
            if (refRange && refRange.getDisplayValues) {
              const opts = refRange.getDisplayValues().flat().map(x => (x == null ? '' : String(x))).filter(x => x !== '');
              dropdownOptions = Array.from(new Set(opts));
              foundDropdown = true;
            }
          }
        }
        if (foundDate && foundDropdown && foundCheckbox) break;
      }

      // バリデーションが無くても書式から日付を推定
      if (!foundDate && numberFormats && numberFormats[0]) {
        const fmt = (numberFormats[0][col] || '').toLowerCase();
        if (fmt.includes('yy') || fmt.includes('yyyy') || fmt.includes('m') && fmt.includes('d')) {
          foundDate = true;
        }
      }

      if (foundDropdown && dropdownOptions) {
        dropdownByCol[col] = dropdownOptions;
      }
      if (foundDate) {
        // 生値から ISO(yyyy-MM-dd) を準備
        const isoArr = [];
        for (let r = 1; r < rawValues.length; r++) {
          const val = rawValues[r][col];
          if (Object.prototype.toString.call(val) === '[object Date]') {
            isoArr.push(Utilities.formatDate(val, tz, 'yyyy-MM-dd'));
          } else {
            isoArr.push('');
          }
        }
        dateIsoByCol[col] = isoArr;
      }
      if (foundCheckbox) {
        checkboxByCol[col] = true;
        const states = [];
        for (let r = 1; r < rawValues.length; r++) {
          const val = rawValues[r][col];
          let b = false;
          if (typeof val === 'boolean') b = val;
          else if (val != null) b = String(val).toUpperCase() === 'TRUE';
          states.push(b);
        }
        checkboxStateByCol[col] = states;
      }
    }

    // マンション名列のハイパーリンクを抽出
    /** @type {{[col:number]: string[]}} */
    const linksByCol = {};
    const mansionIdx = headers.findIndex(h => (h || '').trim() === 'マンション名');
    if (mansionIdx >= 0 && dataRowCount > 0) {
      const rng = sheet.getRange(2, mansionIdx + 1, dataRowCount, 1);
      const rich = rng.getRichTextValues();
      const formulas = rng.getFormulas();
      const displays = rng.getDisplayValues();
      const arr = [];
      for (let i = 0; i < dataRowCount; i++) {
        let u = '';
        const r = rich[i][0];
        if (r) {
          const linkUrl = r.getLinkUrl && r.getLinkUrl();
          if (linkUrl) u = linkUrl;
          if (!u && r.getRuns) {
            const runs = r.getRuns();
            for (const run of runs) { const uu = run.getLinkUrl && run.getLinkUrl(); if (uu) { u = uu; break; } }
          }
        }
        if (!u) {
          const f = formulas[i][0] || '';
          const m = f.match(/HYPERLINK\("([^"]+)"/i);
          if (m && m[1]) u = m[1];
        }
        if (!u) {
          const d = displays[i][0] || '';
          if (/^https?:\/\//i.test(d)) u = d;
        }
        arr.push(u);
      }
      linksByCol[mansionIdx] = arr;
    }

    return { sheetName: sheet.getName(), headers, rows, dropdownByCol, dateIsoByCol, sheetUrl: buildSheetUrl_(ss, sheet), linksByCol, checkboxByCol, checkboxStateByCol };
  } catch (e) {
    console.error('スキーマ付きデータ取得失敗: ' + e.toString());
    return { error: 'スキーマ付きデータの取得に失敗しました。URL/権限/データをご確認ください。' };
  }
}

/**
 * 指定URLのスプレッドシートから、タブ名でシートを選んでスキーマ付きデータを取得します。
 * @param {string} url
 * @param {string} tabName
 * @returns {ReturnType<getSheetDataWithSchemaFromUrl>}
 */
function getTabDataWithSchemaFromUrl(url, tabName) {
  try {
    const parsed = parseSpreadsheetUrl(url);
    if (!parsed || !parsed.spreadsheetId) {
      return { error: 'URLからスプレッドシートIDを抽出できませんでした。' };
    }
    const ss = SpreadsheetApp.openById(parsed.spreadsheetId);
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      return { error: `タブ「${tabName}」が見つかりません。` };
    }
    // 既存ロジックを一部流用
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      return { sheetName: sheet.getName(), headers: [], rows: [], dropdownByCol: {}, dateIsoByCol: {}, sheetUrl: buildSheetUrl_(ss, sheet), linksByCol: {}, checkboxByCol: {}, checkboxStateByCol: {} };
    }

    const displayValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const rawValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = displayValues[0] || [];
    const rows = displayValues.slice(1);
    const dataRowCount = Math.max(0, lastRow - 1);
    const dataRange = dataRowCount > 0 ? sheet.getRange(2, 1, dataRowCount, lastCol) : null;
    const validations = dataRange ? dataRange.getDataValidations() : [];
    const numberFormats = dataRange ? dataRange.getNumberFormats() : [];

    /** @type {{[col:number]: string[]}} */
    const dropdownByCol = {};
    /** @type {{[col:number]: string[]}} */
    const dateIsoByCol = {};
    /** @type {{[col:number]: true}} */
    const checkboxByCol = {};
    /** @type {{[col:number]: boolean[]}} */
    const checkboxStateByCol = {};

    const DVC = SpreadsheetApp.DataValidationCriteria;
    const tz = Session.getScriptTimeZone();

    for (let col = 0; col < lastCol; col++) {
      let foundDropdown = false, foundDate = false, foundCheckbox = false; let dropdownOptions = null;
      for (let r = 0; r < dataRowCount; r++) {
        const v = validations && validations[r] ? validations[r][col] : null;
        if (!v) continue;
        const t = v.getCriteriaType && v.getCriteriaType();
        if (!foundDate && t && String(t).indexOf('DATE') >= 0) foundDate = true;
        if (!foundCheckbox && t && t === DVC.CHECKBOX) foundCheckbox = true;
        if (!foundDropdown && t && (t === DVC.VALUE_IN_LIST || t === DVC.VALUE_IN_RANGE)) {
          const cv = v.getCriteriaValues ? v.getCriteriaValues() : [];
          if (t === DVC.VALUE_IN_LIST) {
            const list = (cv && cv[0]) ? cv[0] : [];
            if (Array.isArray(list)) { dropdownOptions = list.map(x => (x == null ? '' : String(x))); foundDropdown = true; }
          } else if (t === DVC.VALUE_IN_RANGE) {
            const refRange = cv && cv[0];
            if (refRange && refRange.getDisplayValues) {
              const opts = refRange.getDisplayValues().flat().map(x => (x == null ? '' : String(x))).filter(x => x !== '');
              dropdownOptions = Array.from(new Set(opts)); foundDropdown = true;
            }
          }
        }
        if (foundDate && foundDropdown && foundCheckbox) break;
      }
      if (foundDropdown && dropdownOptions) dropdownByCol[col] = dropdownOptions;
      if (foundDate) {
        const isoArr = [];
        for (let r = 1; r < rawValues.length; r++) {
          const val = rawValues[r][col];
          if (Object.prototype.toString.call(val) === '[object Date]') isoArr.push(Utilities.formatDate(val, tz, 'yyyy-MM-dd')); else isoArr.push('');
        }
        dateIsoByCol[col] = isoArr;
      }
      if (foundCheckbox) {
        checkboxByCol[col] = true;
        const states = [];
        for (let r = 1; r < rawValues.length; r++) {
          const val = rawValues[r][col];
          let b = false; if (typeof val === 'boolean') b = val; else if (val != null) b = String(val).toUpperCase() === 'TRUE';
          states.push(b);
        }
        checkboxStateByCol[col] = states;
      }
    }

    // マンション名列のハイパーリンク
    /** @type {{[col:number]: string[]}} */
    const linksByCol = {};
    const mansionIdx = headers.findIndex(h => (h || '').trim() === 'マンション名');
    if (mansionIdx >= 0 && dataRowCount > 0) {
      const rng = sheet.getRange(2, mansionIdx + 1, dataRowCount, 1);
      const rich = rng.getRichTextValues();
      const formulas = rng.getFormulas();
      const displays = rng.getDisplayValues();
      const arr = [];
      for (let i = 0; i < dataRowCount; i++) {
        let u = '';
        const r = rich[i][0];
        if (r) {
          const linkUrl = r.getLinkUrl && r.getLinkUrl();
          if (linkUrl) u = linkUrl;
          if (!u && r.getRuns) { const runs = r.getRuns(); for (const run of runs) { const uu = run.getLinkUrl && run.getLinkUrl(); if (uu) { u = uu; break; } } }
        }
        if (!u) { const f = formulas[i][0] || ''; const m = f.match(/HYPERLINK\("([^"]+)"/i); if (m && m[1]) u = m[1]; }
        if (!u) { const d = displays[i][0] || ''; if (/^https?:\/\//i.test(d)) u = d; }
        arr.push(u);
      }
      linksByCol[mansionIdx] = arr;
    }

    return { sheetName: sheet.getName(), headers, rows, dropdownByCol, dateIsoByCol, sheetUrl: buildSheetUrl_(ss, sheet), linksByCol, checkboxByCol, checkboxStateByCol };
  } catch (e) {
    console.error('タブ別スキーマ付きデータ取得失敗: ' + e.toString());
    return { error: 'タブ別データの取得に失敗しました。URL/権限/データをご確認ください。' };
  }
}

/**
 * クライアントからのセル更新を対象シートへ反映します。
 * edits の各要素は 0-based index（ヘッダ除く行基準）。
 * 例: { rowIndex: 0, colIndex: 3, value: '新しい値' }
 * @param {string} url - 対象シートのURL
 * @param {{rowIndex:number, colIndex:number, value:any}[]} edits - 更新リスト
 * @returns {{updated:number}|{error:string}}
 */
function applyEdits(url, edits) {
  try {
    if (!Array.isArray(edits) || edits.length === 0) return { updated: 0 };
    const { sheet } = resolveSheetFromUrl_(url);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const dataRowCount = Math.max(0, lastRow - 1);
    let applied = 0;

    // 1件ずつ安全に反映（まずは可読性優先。パフォーマンス要件次第でバッチ最適化可）
    for (const e of edits) {
      if (!e) continue;
      const ri = Number(e.rowIndex);
      const ci = Number(e.colIndex);
      if (!isFinite(ri) || !isFinite(ci)) continue;
      if (ri < 0 || ri >= dataRowCount) continue;
      if (ci < 0 || ci >= lastCol) continue;

      const row = ri + 2; // データは2行目開始
      const col = ci + 1;
      let value = e.value;

      // ISO日付っぽい文字列はDateに変換（理事会日など）。"yyyy-MM-dd" のみ対応
      if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
        const parts = value.split('-');
        const dt = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        if (!isNaN(dt.getTime())) value = dt;
      }
      sheet.getRange(row, col).setValue(value);
      applied++;
    }

    return { updated: applied };
  } catch (e) {
    console.error('セル更新失敗: ' + e.toString());
    return { error: 'セル更新に失敗しました。URL/権限/編集対象をご確認ください。' };
  }
}
