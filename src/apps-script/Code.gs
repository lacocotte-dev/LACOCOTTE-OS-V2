function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('LACOCOTTE OS v2.2.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Productsシートが見つかりません。');
  return sheet;
}

function getImageFolder_() {
  return DriveApp.getFolderById(IMAGE_FOLDER_ID);
}

function getProducts() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);

  return rows.map((r, i) => ({
    rowNumber: i + 2,
    id: String(r[0] || ''),
    name: String(r[1] || ''),
    category: String(r[2] || ''),
    material: String(r[3] || ''),
    purchasePrice: Number(r[4] || 0),
price: Number(r[4] || 0), // 旧コード互換用
retailPrice: Number(r[5] || 0),
retail: Number(r[5] || 0), // 旧コード互換用
    box: String(r[6] || ''),
    purchaseDate: r[7] ? Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : '',
    country: String(r[8] || ''),
    status: String(r[9] || ''),
    memo: String(r[10] || ''),
    imageUrl: String(r[11] || ''),
    shippingEnabled: String(r[12] || '') === 'ON',
    invoicePrice: Number(r[13] || 0),
    customsName: String(r[14] || ''),
    shippingMethod: String(r[15] || '')
  })).filter(p => p.id);
}

function getDashboardData() {
  const products = getProducts();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const todayItems = products.filter(p =>
    p.purchaseDate && p.purchaseDate.indexOf(today) === 0
  ).length;

  const boxes = [...new Set(products.map(p => p.box).filter(Boolean))];

  return {
    totalItems: products.length,
    todayItems,
    openBoxes: boxes.length,
    latestProducts: products.slice(-5).reverse()
  };
}

function getBoxSummary() {
  const products = getProducts();
  const map = {};

  products.forEach(p => {
    const box = p.box || 'BOX未設定';

    if (!map[box]) {
      map[box] = {
        box,
        total: 0,
        shipping: 0,
        invoiceTotal: 0
      };
    }

    map[box].total += 1;

    if (p.shippingEnabled) {
      map[box].shipping += 1;
      map[box].invoiceTotal += Number(p.invoicePrice || p.price || 0);
    }
  });

  return Object.values(map).sort((a, b) => {
    return String(a.box).localeCompare(String(b.box));
  });
}

function saveProduct(product) {
  const sheet = getSheet_();
  const nextRow = sheet.getLastRow() + 1;
  const productId = createProductId(nextRow - 1);
  const imageUrl = saveImageIfExists_(product, productId);

  sheet.appendRow([
    productId,
    product.name || '',
    product.category || '',
    product.material || '',
    product.price || '',
    product.retail || '',
    product.box || '',
    new Date(),
    product.country || 'France',
    product.status || 'Draft',
    product.memo || '',
    imageUrl
  ]);

  return { success: true, id: productId, imageUrl };
}

function updateProduct(product) {
  const sheet = getSheet_();
  const row = Number(product.rowNumber);

  if (!row || row < 2) throw new Error('更新対象の行が不正です。');

  let imageUrl = sheet.getRange(row, 12).getValue();

  if (product.imageData && product.imageName) {
    imageUrl = saveImageIfExists_(product, product.id || 'NO-ID');
  }

  sheet.getRange(row, 2, 1, 15).setValues([[
    product.name || '',
    product.category || '',
    product.material || '',
    product.price || '',
    product.retail || '',
    product.box || '',
    new Date(),
    product.country || 'France',
    product.status || 'Draft',
    product.memo || '',
    imageUrl,
    product.shippingEnabled === true ? 'ON' : 'OFF',
    product.invoicePrice || product.price || '',
    product.customsName || '',
    product.shippingMethod || ''
  ]]);

  return { success: true, id: product.id, imageUrl };
}

function deleteProduct(rowNumber) {
  const sheet = getSheet_();
  const row = Number(rowNumber);

  if (!row || row < 2) throw new Error('削除対象の行が不正です。');

  sheet.deleteRow(row);
  return { success: true };
}

function saveImageIfExists_(product, productId) {
  if (!product.imageData || !product.imageName) return '';

  const folder = getImageFolder_();
  const base64 = product.imageData.split(',')[1];
  const mimeType = product.imageData.match(/^data:(.*?);base64,/)[1];

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, productId + '_' + product.imageName);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function analyzeProductImage(product) {
  if (!product.imageData) {
    throw new Error('画像がありません。先に写真を撮影してください。');
  }

  const base64 = product.imageData.split(',')[1];
  const mimeType = product.imageData.match(/^data:(.*?);base64,/)[1];

const prompt = `
あなたはLACOCOTTE専用のフランスアンティーク・ブロカント商品解析アシスタントです。

画像の商品を見て、買付登録用に推定してください。

必ずJSONのみで返してください。

{
  "name": "日本語の商品名",
  "category": "カテゴリ",
  "material": "材質",
  "era": "推定年代",
  "country": "推定国",
  "customsName": "英語の通関名",
  "description": "販売用の短い説明",
  "memo": "注意点・刻印・状態・補足",
  "authenticity": "アンティーク / ヴィンテージ / 現代品 / 不明",
  "confidence": "AI判定の信頼度 0〜100",
  "keywords": "検索用キーワード"
}

判断ルール：
- 現代品に見える場合は、無理にアンティーク扱いしない
- 缶、ペットボトル、家電、文具など明らかな現代品は「現代品」とする
- フランスアンティークらしい場合は年代を推定する
- 材質は Wood / Metal / Ceramic / Glass / Paper / Fabric なども参考にする
- 通関名はシンプルな英語にする
`;

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key='
    + GEMINI_API_KEY;

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [
          { text: prompt },
          {
            inlineData: {
              mimeType: mimeType,
              data: base64
            }
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.3,
      responseMimeType: 'application/json'
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code >= 400) {
    throw new Error('Gemini APIエラー：' + body);
  }

  const json = JSON.parse(body);
  const text = json.candidates?.[0]?.content?.parts?.[0]?.text;

  if (!text) {
    throw new Error('AI解析結果が空です。');
  }

  return parseGeminiJson_(text);
}

function parseGeminiJson_(text) {
  const cleaned = text
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();

  return JSON.parse(cleaned);
}

function authorizeDrive() {
  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  const file = folder.createFile('permission_test.txt', 'OK');
  file.setTrashed(true);
}

function authorizeGemini() {
  UrlFetchApp.fetch('https://www.google.com');
}

function createProductId(number) {
  return 'LC-2026FR-' + String(number).padStart(4, '0');
}

function getProductsByBox(boxName) {
  const products = getProducts();

  return products.filter(function(p) {
    const box = p.box || 'BOX未設定';
    return box === boxName;
  });
}
