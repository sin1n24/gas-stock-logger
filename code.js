// ==========================================
// 設定エリア
// ==========================================
const CONFIG = {
  TARGET_COLLECTION_URL: '[https://example.com/collections/all](https://example.com/collections/all)', // 監視したい販売品の一覧ページのURL
  DOMAIN: '[https://example.com](https://example.com)', // サイトのドメイン
  NOTIFY_EMAIL: 'your-email@example.com', // 通知先メールアドレス
  
  DEFAULT_CAUTION_LIMIT: 3,   
  DEFAULT_WARNING_LIMIT: 1    
};

// ==========================================
// メイン関数
// ==========================================

function updateInventoryAutomatic() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 今年の西暦シート (例: "2026")
  const now = new Date();
  const currentYear = now.getFullYear().toString();
  let sheet = ss.getSheetByName(currentYear);
  
  if (!sheet) {
    sheet = initializeYearSheet(ss, currentYear);
  }
  
  // --- 列インデックス定義 (0-based) ---
  // A:Name(0), B:Stock(1), C:Caution(2), D:Warning(3), E:Status(4)
  // F(5):1月 ... Q(16):12月
  const currentMonthColIndex = 5 + now.getMonth(); 
  
  // R(17):空白スペース列
  // S(18):1/1 ... (日別推移開始)
  const dayOfYear = getDayOfYear(now);
  const todayColIndex = 18 + (dayOfYear - 1); // 17(R列)を飛ばして18(S列)から開始

  // 1. 商品リスト取得
  console.log("商品リストを取得中...");
  const productUrls = getAllProductUrls(CONFIG.TARGET_COLLECTION_URL);
  if (productUrls.length === 0) return;

  // 2. 現在のデータを取得
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const dataMap = new Map();
  
  if (lastRow > 2) {
    const sheetValues = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    const formulas = sheet.getRange(3, 1, lastRow - 2, 1).getFormulas();

    sheetValues.forEach((row, i) => {
      let url = "";
      // 数式からURL抽出: =HYPERLINK("URL", "Name")
      const formula = formulas[i][0]; 
      const match = formula.match(/"(https:\/\/[^"]+)"/);
      if (match) {
        url = match[1];
      }
      
      if(url) {
        dataMap.set(url, {
          rowIndex: i,
          row: row
        });
      }
    });
  }

  // 3. 更新用データの構築
  const newSheetData = []; 
  const newFormulas = []; 
  const alertItems = [];   

  for (const url of productUrls) {
    try {
      console.log(`確認中: ${url}`);
      const html = fetchHtml(url);
      const name = extractProductName(html);
      let stock = extractStockCount(html);

      // --- データのマージ ---
      let rowData;
      let prevStock = null;
      let currentMonthSales = 0;
      
      if (dataMap.has(url)) {
        const existing = dataMap.get(url);
        rowData = [...existing.row];
        prevStock = rowData[1]; // B:在庫
        currentMonthSales = Number(rowData[currentMonthColIndex]) || 0;
      } else {
        // 新規 (列数を確保)
        const maxCols = Math.max(lastCol, todayColIndex + 1);
        rowData = new Array(maxCols).fill("");
        rowData[2] = CONFIG.DEFAULT_CAUTION_LIMIT; // C:注意
        rowData[3] = CONFIG.DEFAULT_WARNING_LIMIT; // D:警告
        rowData[4] = "Normal";                     // E:状態
        prevStock = null;
      }

      // --- 販売数カウント ---
      if (typeof stock === 'number' && typeof prevStock === 'number') {
        if (prevStock > stock) {
          const soldCount = prevStock - stock;
          currentMonthSales += soldCount;
        }
      }
      
      // --- データ更新 ---
      rowData[0] = name; 
      rowData[1] = stock; 
      
      // 月別販売数
      rowData[currentMonthColIndex] = currentMonthSales;

      // 空白列(R列 index 17)を確実に空にする
      if (rowData.length > 17) {
        rowData[17] = ""; 
      }

      // 日別ログ
      while (rowData.length <= todayColIndex) rowData.push("");
      rowData[todayColIndex] = stock;

      // --- 通知判定 ---
      let cautionLimit = rowData[2] === "" ? CONFIG.DEFAULT_CAUTION_LIMIT : rowData[2]; 
      let warningLimit = rowData[3] === "" ? CONFIG.DEFAULT_WARNING_LIMIT : rowData[3]; 
      let lastStatus = rowData[4] || "Normal"; 
      
      if (cautionLimit === 0) cautionLimit = -1; 
      if (warningLimit === 0) warningLimit = -1;

      let currentStatus = "Normal";
      let notifyType = null;

      if (typeof stock === 'number') {
        if (stock <= warningLimit) {
          currentStatus = "Warning";
        } else if (stock <= cautionLimit) {
          currentStatus = "Caution";
        }

        if (currentStatus === "Warning" && lastStatus !== "Warning") {
          notifyType = "【警告】在庫わずか";
        } else if (currentStatus === "Caution" && lastStatus !== "Caution" && lastStatus !== "Warning") {
          notifyType = "【注意】在庫減少";
        }
      }
      rowData[4] = currentStatus;

      if (notifyType) {
        alertItems.push({ type: notifyType, name: name, stock: stock, limit: (currentStatus === "Warning" ? warningLimit : cautionLimit), url: url });
      }

      newSheetData.push(rowData);
      newFormulas.push([`=HYPERLINK("${url}", "${name}")`]);
      
      Utilities.sleep(1000); 

    } catch (e) {
      console.error(`エラー (${url}): ${e.message}`);
      if (dataMap.has(url)) {
        const oldRow = dataMap.get(url).row;
        newSheetData.push(oldRow);
        newFormulas.push([`=HYPERLINK("${url}", "${oldRow[0]}")`]);
      }
    }
  }

  // 4. シートへの書き込みと装飾
  if (newSheetData.length > 0) {
    const numRows = newSheetData.length;
    const maxCols = newSheetData.reduce((max, row) => Math.max(max, row.length), 0);
    const formattedData = newSheetData.map(row => {
      while (row.length < maxCols) row.push("");
      return row;
    });

    // 既存クリア
    if (lastRow > 2) {
      sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clearContent();
      sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clearFormat(); 
    }
    
    // 書き込み
    sheet.getRange(3, 1, numRows, maxCols).setValues(formattedData);
    sheet.getRange(3, 1, numRows, 1).setFormulas(newFormulas);
    
    // 日付ヘッダー更新
    ensureDateHeader(sheet, now);

    // --- 装飾 ---
    sheet.getRange(3, 1, numRows, 2).setFontWeight("bold"); // A-B列太字
    
    // 背景色 (E列)
    const statusRange = sheet.getRange(3, 5, numRows, 1);
    const backgrounds = statusRange.getValues().map(row => {
      if (row[0] === "Warning") return ["#ea4335"]; 
      if (row[0] === "Caution") return ["#fbbc04"]; 
      return [null];
    });
    statusRange.setBackgrounds(backgrounds);

    // R列（空白列）の背景色をグレーにして区切りを分かりやすくする（任意）
    // sheet.getRange(3, 18, numRows, 1).setBackground("#f3f3f3");

    // --- 【重要】列幅の最適化 ---
    // 計算を確定させるためにFlushを実行
    SpreadsheetApp.flush(); 
    
    // 全列に対して自動調整を実行
    // ※ヘッダーの結合セルの影響を受けにくいよう、データ範囲に基づいて調整されます
    sheet.autoResizeColumns(1, maxCols);
  }

  // 5. メール送信
  if (alertItems.length > 0) {
    sendAlertEmail(alertItems);
  }
  
  console.log("処理完了");
}

// ==========================================
// シート初期化 & ヘッダー管理
// ==========================================

function initializeYearSheet(ss, yearStr) {
  const sheet = ss.insertSheet(yearStr);
  
  // A:Name, B:Stock, C:Caution, D:Warning, E:Status
  // F-Q: Sales(12cols)
  // R: Spacer(1col)
  // S+: Daily
  
  // --- 1行目 ---
  const header1 = [
    "",         
    "在庫数",   // B-D
    "",         
    "",         
    "",         
    "販売数",   // F-Q
    "", "", "", "", "", "", "", "", "", "", "", 
    "",         // R(空白列)
    "日別推移"  // S
  ];
  
  // --- 2行目 ---
  const header2 = [
    "商品名", 
    "現在", "注意", "警告", 
    "通知状態"
  ];
  // 1月〜12月
  for (let i = 1; i <= 12; i++) {
    header2.push(`${i}月`);
  }
  header2.push(""); // R列(スペーサー)用の空ヘッダー
  
  // 書き込み
  sheet.getRange(1, 1, 1, header1.length).setValues([header1]);
  sheet.getRange(2, 1, 1, header2.length).setValues([header2]);
  
  // 結合
  sheet.getRange("B1:D1").merge().setHorizontalAlignment("center"); 
  sheet.getRange("F1:Q1").merge().setHorizontalAlignment("center"); 
  
  // 固定
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1); 
  
  return sheet;
}

function ensureDateHeader(sheet, date) {
  // R列(index 17)は空白、S列(index 18)から日付 -> colIndex 19 (1-based)
  const dayOfYear = getDayOfYear(date);
  const colIndex = 19 + (dayOfYear - 1); 
  
  const headerRange = sheet.getRange(2, 1, 1, colIndex);
  const headers = headerRange.getValues()[0];
  
  if (headers.length < colIndex || headers[colIndex - 1] === "") {
    sheet.getRange(2, colIndex).setValue(`${date.getMonth() + 1}/${date.getDate()}`);
  }
}

function getDayOfYear(date) {
  const start = new Date(date.getFullYear(), 0, 0);
  const diff = date - start;
  const oneDay = 1000 * 60 * 60 * 24;
  return Math.floor(diff / oneDay);
}

// ==========================================
// 共通関数
// ==========================================
function sendAlertEmail(items) {
  const subject = `[在庫アラート] ${items.length}件の商品が基準値を下回りました`;
  let body = "以下の商品の在庫が少なくなっています。\n\n";
  items.forEach(item => {
    body += `--------------------------------------------------\n`;
    body += `${item.type}\n商品名: ${item.name}\n現在庫: ${item.stock} (基準: ${item.limit}以下)\nURL: ${item.url}\n`;
  });
  body += `\n--------------------------------------------------\nGoogle Apps Script Alert`;
  MailApp.sendEmail({ to: CONFIG.NOTIFY_EMAIL, subject: subject, body: body });
}

function getAllProductUrls(startUrl) {
  let urls = new Set();
  let nextUrl = startUrl;
  let pageCount = 0;
  while (nextUrl && pageCount < 5) {
    const html = fetchHtml(nextUrl);
    const regex = /<a\s+[^>]*href=["']([^"']*\/products\/[^"']+)["'][^>]*>/g;
    let match;
    while ((match = regex.exec(html)) !== null) {
      let path = match[1].split('?')[0];
      const productIndex = path.indexOf('/products/');
      if (productIndex !== -1) path = path.substring(productIndex);
      urls.add(CONFIG.DOMAIN + path);
    }
    const nextMatch = html.match(/<a\s+[^>]*href=["']([^"']+)["'][^>]*>\s*(?:Next|次へ|&rarr;|→)\s*<\/a>/i) || html.match(/<link\s+rel=["']next["']\s+href=["']([^"']+)["']/i);
    if (nextMatch) {
      let nextPath = nextMatch[1].replace(/&amp;/g, '&');
      nextUrl = nextPath.startsWith('http') ? nextPath : CONFIG.DOMAIN + nextPath;
      pageCount++;
      Utilities.sleep(1000); 
    } else { nextUrl = null; }
  }
  return Array.from(urls);
}

function extractProductName(html) {
  const match = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  return match ? match[1].replace(/<[^>]+>/g, '').trim() : "名称不明";
}

function extractStockCount(html) {
  const stockMatch = html.match(/在庫数\s*[:：]?\s*(\d+)/);
  if (stockMatch) return parseInt(stockMatch[1], 10);
  const stockMatch2 = html.match(/在庫\s*[:：]?\s*(\d+)/);
  if (stockMatch2) return parseInt(stockMatch2[1], 10);
  if (html.includes("在庫切れ") || html.includes("Sold out") || html.includes("sold-out")) return 0;
  return "不明";
}

function fetchHtml(url) {
  try {
    const options = { 'muteHttpExceptions': true };
    return UrlFetchApp.fetch(url, options).getContentText();
  } catch (e) { return ""; }
}
