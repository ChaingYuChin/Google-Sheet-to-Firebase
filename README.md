

### Google-Sheet-to-Firebase**  
---

### **這是一個把Google Sheet資料傳到Firebase的Firestore Database裡面**  

---

#### **使用到的工具：**  

- **1. Google Sheet**  

- **2. Google Drive**  

- **3. Firebase**  
  
- **4. Google Cloud**  

- **5. Apps Script**  

---

const CREDENTIALS_FILE_ID = "你的憑證文件的ID"; // 憑證文件的ID
const PROJECT_ID = "你的專案ID"; // 專案的ID


function getServiceAccountKey() {
  var file = DriveApp.getFileById(CREDENTIALS_FILE_ID); // 根據ID從Google雲端硬碟取得檔案
  return JSON.parse(file.getBlob().getDataAsString()); // 解析檔案內容並返回作為JSON物件
}

function createJWT(serviceAccount) {
  var header = {
    alg: "RS256", // 加密演算法
    typ: "JWT", // 類型
  };

  var now = Math.floor(Date.now() / 1000); // 取得當前時間戳記（秒）
  var claimSet = {
    iss: serviceAccount.client_email, // 服務帳戶的email
    scope: "https://www.googleapis.com/auth/datastore https://www.googleapis.com/auth/cloud-platform", // API的授權範圍
    aud: "https://www.googleapis.com/oauth2/v4/token", // Token請求的目的地
    exp: now + 3600, // Token的過期時間（一小時後）
    iat: now, // Token的發行時間
  };

  var signatureInput =
    Utilities.base64EncodeWebSafe(JSON.stringify(header)) + // 以Base64編碼標頭
    "." +
    Utilities.base64EncodeWebSafe(JSON.stringify(claimSet)); // 以Base64編碼Claim集
  var signature = Utilities.computeRsaSha256Signature(
    signatureInput,
    serviceAccount.private_key
  ); // 使用RSA SHA-256加密簽名
  var jwt = signatureInput + "." + Utilities.base64EncodeWebSafe(signature); // 合併生成完整的JWT

  return jwt; // 返回JWT
}

function getAccessToken() {
  var serviceAccount = getServiceAccountKey(); // 獲取服務帳戶金鑰
  var tokenUrl = "https://www.googleapis.com/oauth2/v4/token"; // Token請求的URL
  var payload = {
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", // 授權類型
    assertion: createJWT(serviceAccount), // JWT令牌
  };

  var options = {
    method: "post", // HTTP方法為POST
    payload: payload, // 請求的有效內容
  };

  var response = UrlFetchApp.fetch(tokenUrl, options); // 發送HTTP請求並獲取回應
  var token = JSON.parse(response.getContentText()).access_token; // 解析回應並提取Access Token

  return token; // 返回Access Token
}

function firestoreRequest(path, method, payload) {
  var token = getAccessToken(); // 獲取Access Token
  var url = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents${path}`; // Firestore API的請求URL
  var options = {
    method: method, // HTTP方法
    contentType: "application/json", // 請求的內容類型
    headers: {
      Authorization: "Bearer " + token, // 請求頭中的授權憑證
    },
    payload: payload ? JSON.stringify(payload) : null, // 如果有內容，將內容轉為JSON格式
  };

  var response = UrlFetchApp.fetch(url, options); // 發送請求並獲取回應
  return JSON.parse(response.getContentText()); // 解析回應並返回為JSON格式
}

function batchUpsertDocuments(documents, collectionName) {
  var collectionPath = "/" + collectionName; // 集合的路徑
  var writes = documents.map(function (doc) {
    return {
      update: {
        name: `projects/${PROJECT_ID}/databases/(default)/documents${collectionPath}/${doc.id}`, // 文件的完整路徑
        fields: doc.fields, // 文件的欄位內容
      },
      updateMask: {
        fieldPaths: Object.keys(doc.fields), // 指定需要更新的欄位
      },
    };
  });

  var payload = {
    writes: writes, // 整理批量寫入的資料
  };

  var response = firestoreBatchRequest(payload); // 發送批量寫入請求
  Logger.log(response); // 日誌記錄回應
}

function firestoreBatchRequest(payload) {
  var token = getAccessToken(); // 獲取Access Token
  var url = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents:batchWrite`; // 批量寫入的API端點

  var options = {
    method: "POST", // HTTP方法為POST
    contentType: "application/json", // 請求的內容類型為JSON
    headers: {
      Authorization: "Bearer " + token, // 授權憑證
    },
    payload: JSON.stringify(payload), // 將批量寫入的資料轉為JSON格式
  };

  var response = UrlFetchApp.fetch(url, options); // 發送請求並獲取回應
  return JSON.parse(response.getContentText()); // 解析回應並返回為JSON格式
}

function getJSONArray(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); // 取得當前活頁簿中的指定工作表
  const dataRange = sheet.getDataRange(); // 取得資料範圍
  const data = dataRange.getValues(); // 取得工作表中的所有資料
  const heads = data.shift(); // 取出表頭（第一行）
  return data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  ); // 將資料轉換成JSON格式，每一列對應一個物件
}

function createFieldValue(value) {
  if (typeof value === "string") {
    return { stringValue: value }; // 如果是字串，返回stringValue
  } else if (typeof value === "number") {
    return Number.isInteger(value)
      ? { integerValue: value } // 如果是整數，返回integerValue
      : { doubleValue: value }; // 如果是浮點數，返回doubleValue
  } else if (typeof value === "boolean") {
    return { booleanValue: value }; // 如果是布林值，返回booleanValue
  } else if (Array.isArray(value)) {
    return { arrayValue: { values: value.map(createFieldValue) } }; // 如果是陣列，遞迴處理每個元素
  } else if (value === null) {
    return { nullValue: null }; // 如果是null，返回nullValue
  }

  return { nullValue: null }; // 其他情況，返回nullValue
}

function convertToFirestoreObject(data) {
  const firestoreObject = {
    id: data.id || Math.random().toString(36).substring(7), // 如果沒有ID就生成一個隨機ID
    fields: {}, // 初始化fields物件
  };

  for (const key in data) {
    if (data[key] !== "") {
      firestoreObject.fields[key] = createFieldValue(data[key]); // 根據資料類型創建對應的欄位值
    }
  }
  return firestoreObject; // 返回Firestore格式的物件
}

function seedData() {
  const COLLECTION_NAME = SpreadsheetApp.getActiveSpreadsheet()
    .getActiveSheet()
    .getName(); // 取得當前工作表的名稱作為集合名稱
  const data = getJSONArray(COLLECTION_NAME); // 將資料轉換成JSON陣列
  const firestoreData = data.map((d) => convertToFirestoreObject(d)); // 轉換成Firestore物件
  batchUpsertDocuments(firestoreData, COLLECTION_NAME); // 批量寫入文件到Firestore
}

function onOpen() {
  var ui = SpreadsheetApp.getUi(); // 取得Google試算表的UI
  ui.createMenu("Firestore") // 建立一個名為"Firestore"的選單
    .addItem("Export To Firestore", "seedData") // 添加一個選項來導出資料到Firestore
    .addToUi(); // 將選單添加到UI中
}

