const SHEET_NAME = "Veritabani";

// 1. Veritabanı sekmesini otomatik oluşturan kurulum fonksiyonu
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Sütun Başlıkları
    sheet.appendRow(["Kurum/Birim ID", "Değerlendirme Tarihi", "Program Adı", "Toplam Puan", "Durum", "Genel Özet", "Güçlü Yönler", "Gelişmeye Açık Yönler", "Detay_JSON"]);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#f8fafc");
    sheet.setColumnWidth(9, 500); // JSON verisi uzun olacağı için sütunu genişletiyoruz
  }
}

// 2. OPTIONS metodu: Tarayıcının CORS (Preflight) kontrolünü geçmek için zorunludur.
// Özellikle büyük veri (uzun metin/görsel) gönderildiğinde tarayıcı önce bu fonksiyonu çağırır.
function doOptions(e) {
  return createCORSResponse("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function createCORSResponse(content) {
  return ContentService.createTextOutput(content);
}

function compressStringForStorage(text) {
  const value = String(text || '');
  if (!value) return value;
  const gzippedBlob = Utilities.gzip(Utilities.newBlob(value, 'application/json'));
  return '#GZ#' + Utilities.base64Encode(gzippedBlob.getBytes());
}

function decompressStringFromStorage(value) {
  if (typeof value !== 'string' || !value.startsWith('#GZ#')) return value;
  const decodedBytes = Utilities.base64Decode(value.slice(4));
  const blob = Utilities.newBlob(decodedBytes, 'application/octet-stream');
  const decompressed = Utilities.ungzip(blob);
  return decompressed.getDataAsString();
}

// 3. Frontend'den gelen verileri kaydetme (POST) işlemi
function doPost(e) {
  try {
    // PDF/dosya yükleme isteği (?action=uploadFile) -> Google Drive'a kaydet
    if (e && e.parameter && e.parameter.action === "uploadFile") return handleUpload(e);

    // Eğer postData boşsa güvenli çıkış yap
    if (!e || !e.postData || !e.postData.contents) {
      return createCORSResponse(JSON.stringify({status: "error", message: "Veri alınamadı veya boş gönderildi."}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

    const id = data.projectNumber;
    const timestamp = new Date();
    
    // Detaylar objesini string'e çevir
    const detailsJsonString = JSON.stringify(data.details);
    let storedDetails = detailsJsonString;

    // Google Sheets hücre limiti kontrolü (Maks: 50.000 karakter)
    // Detay JSON'u tek hücrede saklandığı için uzun veri doğrudan sınırı aşabilir.
    if (detailsJsonString.length > 50000) {
        const compressed = compressStringForStorage(detailsJsonString);
        if (compressed.length < 50000) {
            storedDetails = compressed;
        } else {
            return createCORSResponse(JSON.stringify({
                status: "error",
                message: `Veri boyutu çok büyük! Detay JSON uzunluğu ${detailsJsonString.length} karakter. Sıkıştırılmış hali ${compressed.length} karakter ve yine sınırı aşıyor. Lütfen yorumları kısaltın veya sayfayı daha az metinle kaydedin.`
            }))
            .setMimeType(ContentService.MimeType.JSON);
        }
    }

    // Satıra yazılacak veri dizisi
    const rowData = [
      id,
      timestamp,
      data.projectName,
      data.grandTotal,
      data.status,
      data.overall.summary,
      data.overall.strengths,
      data.overall.weaknesses,
      storedDetails
    ];

    // ID daha önce kaydedilmiş mi kontrol et (Güncelleme mantığı)
    const dataRange = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == id) {
        foundRow = i + 1; // Apps script satırları 1'den başlar
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]); // Var olanı güncelle
    } else {
      sheet.appendRow(rowData); // Yeni kayıt ekle
    }

    return createCORSResponse(JSON.stringify({status: "success", message: "Veritabanına başarıyla kaydedildi!"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return createCORSResponse(JSON.stringify({status: "error", message: "Sunucu Hatası: " + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- GitHub'a dosya/görsel yükleme ayarları ----
// Token KODA YAZILMAZ: Apps Script > Proje Ayarları > Komut dosyası özellikleri > GITHUB_TOKEN
const GITHUB_OWNER  = "mkemalaydin";
const GITHUB_REPO   = "avrasya";
const GITHUB_BRANCH = "main";
const UPLOAD_DIR    = "yuklenen_kanitlar";
const MAX_BASE64_LEN = 27000000; // yaklaşık 20 MB dosya

// Apps Script düzenleyicisinden bir kez çalıştırıp harici istek iznini onaylayın.
function authorizeExternalRequest() {
  UrlFetchApp.fetch("https://api.github.com/");
}

function safeFileName(name) {
  const map = {"ç":"c","Ç":"C","ğ":"g","Ğ":"G","ı":"i","İ":"I","ö":"o","Ö":"O","ş":"s","Ş":"S","ü":"u","Ü":"U"};
  const s = String(name || "kanit").replace(/[çÇğĞıİöÖşŞüÜ]/g, c => map[c]).replace(/[^A-Za-z0-9._-]+/g, "_");
  return Utilities.formatDate(new Date(), "UTC", "yyyyMMdd-HHmmss") + "_" + s.slice(-80);
}

function handleUpload(e) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
    if (!token) throw new Error("GITHUB_TOKEN, Komut dosyası özelliklerine eklenmemiş.");
    if (GITHUB_OWNER.indexOf("BURAYA") === 0 || GITHUB_REPO.indexOf("BURAYA") === 0) throw new Error("GITHUB_OWNER / GITHUB_REPO ayarlanmamış.");
    const f = JSON.parse(e.postData.contents);
    if (!f.data) throw new Error("Dosya verisi boş.");
    if (f.data.length > MAX_BASE64_LEN) throw new Error("Dosya çok büyük (en fazla ~20 MB).");
    const path = UPLOAD_DIR + "/" + safeFileName(f.fileName);
    const res = UrlFetchApp.fetch("https://api.github.com/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/contents/" + path, {
      method: "put",
      contentType: "application/json",
      headers: {Authorization: "Bearer " + token, Accept: "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"},
      payload: JSON.stringify({message: "Kanıt yüklendi: " + path, content: f.data, branch: GITHUB_BRANCH}),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 201 && code !== 200) {
      let msg = res.getContentText();
      try { msg = JSON.parse(msg).message || msg; } catch (x) {}
      throw new Error("GitHub " + code + ": " + msg);
    }
    const isImg = String(f.mimeType || "").indexOf("image/") === 0;
    const base = GITHUB_OWNER + "/" + GITHUB_REPO;
    const url = isImg ? "https://raw.githubusercontent.com/" + base + "/" + GITHUB_BRANCH + "/" + path
                      : "https://github.com/" + base + "/blob/" + GITHUB_BRANCH + "/" + path;
    return createCORSResponse(JSON.stringify({status: "success", url: url, path: path}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return createCORSResponse(JSON.stringify({status: "error", message: "Yükleme hatası: " + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 4. Frontend'e verileri ve ID listesini gönderme (GET) işlemi
function doGet(e) {
  // Preflight veya parametresiz durumlar için güvenlik
  if (!e || !e.parameter || !e.parameter.action) {
    return createCORSResponse(JSON.stringify({status: "error", message: "Eksik parametre."}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

  const data = sheet.getDataRange().getValues();
  let response;

  try {
    // Açılır liste (Select) için kayıtlı ID'leri getir
    if (action === "getList") {
      let list = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) list.push(data[i][0]);
      }
      response = JSON.stringify(list);
    }
    // Seçilen ID'nin tüm verilerini (JSON dahil) frontend'e yükle
    else if (action === "load") {
      const projNo = e.parameter.projNo;
      let result = null;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == projNo) {
          let detailsCell = data[i][8] || "{}";
          if (typeof detailsCell === 'string' && detailsCell.startsWith('#GZ#')) {
            detailsCell = decompressStringFromStorage(detailsCell);
          }
          result = {
            projectNumber: data[i][0],
            projectName: data[i][2],
            grandTotal: data[i][3],
            status: data[i][4],
            overall: {
              summary: data[i][5],
              strengths: data[i][6],
              weaknesses: data[i][7]
            },
            details: JSON.parse(detailsCell || "{}")
          };
          break;
        }
      }
      
      if (result) {
        response = JSON.stringify({status: "success", data: result});
      } else {
        response = JSON.stringify({status: "error", message: "Kayıt bulunamadı."});
      }
    } else {
      response = JSON.stringify({status: "error", message: "Geçersiz işlem."});
    }
  } catch(err) {
      response = JSON.stringify({status: "error", message: err.toString()});
  }

  return createCORSResponse(response)
    .setMimeType(ContentService.MimeType.JSON);
}