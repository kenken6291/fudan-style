/**
 * パーソナルスタイリストAI コーディネート提案Webアプリ
 * バックエンド（Google Apps Script）
 *
 * ■事前準備（スクリプトプロパティ）
 *   SPREADSHEET_ID  : このアプリで使うスプレッドシートのID
 *   GEMINI_API_KEY  : Google AI Studioで発行したGemini APIキー
 *   PASSWORD_PEPPER : パスワードハッシュ化用の固定文字列（自分で好きな長い文字列を設定）
 *   LIGHTX_API_KEY  : LightX（lightxeditor.com）のAPIキー（試着イメージ生成用）
 *
 * ■スプレッドシート構成
 *   シート「Users」   : UserID / Email / PasswordHash / Salt / RegisteredAt / AgeGroup / UsualStyle / NgElements / PhotoFront / PhotoBack / PhotoLeft / PhotoRight / Gender / HairStyle / PasswordResetRequired
 *   シート「History」 : HistoryID / UserID / CreatedAt / Brands / Scene / Mood / Temperature / OwnedItems / ResultJSON / TryOnImageUrl
 *                       / TopColors / BottomColors / TopStyles / BottomStyles / Mode / BudgetMin / BudgetMax / TotalPrice
 *   シート「Wardrobe」: ItemID / UserID / Category / Name / Color / Brand / Memo / CreatedAt / PhotoUrl （手持ちの服）
 *
 * ■Googleドライブ
 *   マイフォト保存先　: 「FudanStyle_UserPhotos」フォルダ（ユーザーごとにサブフォルダ）
 *   試着イメージ保存先: 「FudanStyle_TryOnResults」フォルダ（ユーザーごとにサブフォルダ）
 *   クローゼット写真　: 「FudanStyle_Wardrobe」フォルダ（ユーザーごとにサブフォルダ）
 *   ※LightX APIが画像を取得できるよう、保存した画像は「リンクを知っている全員が閲覧可」で共有されます。
 */

const GEMINI_MODEL = 'gemini-3.6-flash';
const SESSION_TTL_SEC = 21600; // 6時間（CacheServiceの上限）
const LIGHTX_OUTFIT_URL = 'https://api.lightxeditor.com/external/api/v1/outfit';
const LIGHTX_STATUS_URL = 'https://api.lightxeditor.com/external/api/v1/order-status';
const USER_PHOTOS_FOLDER = 'FudanStyle_UserPhotos';
const TRYON_RESULTS_FOLDER = 'FudanStyle_TryOnResults';
const WARDROBE_PHOTOS_FOLDER = 'FudanStyle_Wardrobe';

// ============================================================
// エントリーポイント
// ============================================================

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, message: 'Personal Stylist AI API is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {
      case 'register':
        result = handleRegister(body);
        break;
      case 'login':
        result = handleLogin(body);
        break;
      case 'forgotPassword':
        result = handleForgotPassword(body);
        break;
      case 'setNewPassword':
        result = handleSetNewPassword(body);
        break;
      case 'changePassword':
        result = handleChangePassword(body);
        break;
      case 'logout':
        result = handleLogout(body);
        break;
      case 'generateProposal':
        result = handleGenerateProposal(body);
        break;
      case 'getHistory':
        result = handleGetHistory(body);
        break;
      case 'getHistoryDetail':
        result = handleGetHistoryDetail(body);
        break;
      case 'uploadPhoto':
        result = handleUploadPhoto(body);
        break;
      case 'getProfile':
        result = handleGetProfile(body);
        break;
      case 'updateProfile':
        result = handleUpdateProfile(body);
        break;
      case 'generateTryOnImage':
        result = handleGenerateTryOnImage(body);
        break;
      case 'getWardrobe':
        result = handleGetWardrobe(body);
        break;
      case 'addWardrobeItem':
        result = handleAddWardrobeItem(body);
        break;
      case 'deleteWardrobeItem':
        result = handleDeleteWardrobeItem(body);
        break;
      case 'analyzeWardrobePhoto':
        result = handleAnalyzeWardrobePhoto(body);
        break;
      default:
        result = { success: false, error: '不明なアクションです: ' + action };
    }
  } catch (err) {
    result = { success: false, error: 'サーバーエラー: ' + err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// スプレッドシート／シートのユーティリティ
// ============================================================

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

function getUsersSheet_() {
  const ss = getSS_();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['UserID', 'Email', 'PasswordHash', 'Salt', 'RegisteredAt', 'AgeGroup', 'UsualStyle', 'NgElements', 'PhotoFront', 'PhotoBack', 'PhotoLeft', 'PhotoRight', 'Gender', 'HairStyle', 'PasswordResetRequired']);
  }
  return sheet;
}

function getHistorySheet_() {
  const ss = getSS_();
  let sheet = ss.getSheetByName('History');
  if (!sheet) {
    sheet = ss.insertSheet('History');
    sheet.appendRow(['HistoryID', 'UserID', 'CreatedAt', 'Brands', 'Scene', 'Mood', 'Temperature', 'OwnedItems', 'ResultJSON', 'TryOnImageUrl',
      'TopColors', 'BottomColors', 'TopStyles', 'BottomStyles', 'Mode', 'BudgetMin', 'BudgetMax', 'TotalPrice']);
  }
  return sheet;
}

function getWardrobeSheet_() {
  const ss = getSS_();
  let sheet = ss.getSheetByName('Wardrobe');
  if (!sheet) {
    sheet = ss.insertSheet('Wardrobe');
    sheet.appendRow(['ItemID', 'UserID', 'Category', 'Name', 'Color', 'Brand', 'Memo', 'CreatedAt', 'PhotoUrl']);
  }
  return sheet;
}

// ============================================================
// 認証まわり
// ============================================================

function generateSalt_() {
  return Utilities.getUuid();
}

function hashPassword_(password, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || '';
  const raw = password + ':' + salt + ':' + pepper;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function generateTempPassword_() {
  // 紛らわしい文字（0/O, 1/l/I など）を除いた文字セットから10文字生成
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let result = '';
  for (let i = 0; i < 10; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function sendTempPasswordEmail_(email, tempPassword, isNewRegistration) {
  const subject = isNewRegistration ? 'FUDAN, 会員登録のご案内（仮パスワード）' : 'FUDAN, 仮パスワードの発行';
  const intro = isNewRegistration
    ? 'FUDAN,にご登録いただきありがとうございます。'
    : 'パスワード再発行のリクエストを受け付けました。';

  const body = [
    intro,
    '',
    '下記の仮パスワードでログインし、ログイン後の画面で新しいパスワードを設定してください。',
    '',
    'メールアドレス: ' + email,
    '仮パスワード: ' + tempPassword,
    '',
    '※このメールに心当たりがない場合は、破棄してください。'
  ].join('\n');

  MailApp.sendEmail(email, subject, body);
}

function findUserRowByEmail_(sheet, email) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function findUserById_(sheet, userId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function handleRegister(body) {
  const email = (body.email || '').trim().toLowerCase();
  if (!email) {
    return { success: false, error: 'メールアドレスを入力してください。' };
  }

  const sheet = getUsersSheet_();
  if (findUserRowByEmail_(sheet, email)) {
    return { success: false, error: 'このメールアドレスは既に登録されています。' };
  }

  const userId = Utilities.getUuid();
  const tempPassword = generateTempPassword_();
  const salt = generateSalt_();
  const passwordHash = hashPassword_(tempPassword, salt);
  const registeredAt = new Date();

  sheet.appendRow([
    userId,
    email,
    passwordHash,
    salt,
    registeredAt,
    body.ageGroup || '',
    body.usualStyle || '',
    body.ngElements || '',
    '', '', '', '',
    body.gender || '',
    body.hairStyle || '',
    'TRUE'
  ]);

  try {
    sendTempPasswordEmail_(email, tempPassword, true);
  } catch (err) {
    return { success: false, error: '登録は完了しましたが、メール送信に失敗しました。時間をおいて「パスワードをお忘れの方」から再発行してください。' };
  }

  return {
    success: true,
    message: '仮パスワードをメールに送信しました。メールを確認してログインし、新しいパスワードを設定してください。'
  };
}

function handleLogin(body) {
  const email = (body.email || '').trim().toLowerCase();
  const password = body.password || '';
  const sheet = getUsersSheet_();
  const found = findUserRowByEmail_(sheet, email);

  if (!found) {
    return { success: false, error: 'メールアドレスまたはパスワードが違います。' };
  }

  const row = found.row;
  const storedHash = row[2];
  const salt = row[3];
  const computedHash = hashPassword_(password, salt);

  if (computedHash !== storedHash) {
    return { success: false, error: 'メールアドレスまたはパスワードが違います。' };
  }

  const userId = row[0];
  const token = createSession_(userId);
  return {
    success: true,
    token: token,
    passwordResetRequired: row[14] === 'TRUE',
    user: {
      userId: userId, email: email,
      ageGroup: row[5], usualStyle: row[6], ngElements: row[7],
      photoFront: row[8] || '', photoBack: row[9] || '', photoLeft: row[10] || '', photoRight: row[11] || '',
      gender: row[12] || '', hairStyle: row[13] || ''
    }
  };
}

function handleForgotPassword(body) {
  const email = (body.email || '').trim().toLowerCase();
  if (!email) {
    return { success: false, error: 'メールアドレスを入力してください。' };
  }

  const sheet = getUsersSheet_();
  const found = findUserRowByEmail_(sheet, email);
  if (!found) {
    return { success: false, error: 'このメールアドレスは登録されていません。' };
  }

  const tempPassword = generateTempPassword_();
  const salt = generateSalt_();
  const passwordHash = hashPassword_(tempPassword, salt);

  sheet.getRange(found.rowIndex, 3).setValue(passwordHash); // PasswordHash
  sheet.getRange(found.rowIndex, 4).setValue(salt);         // Salt
  sheet.getRange(found.rowIndex, 15).setValue('TRUE');      // PasswordResetRequired

  try {
    sendTempPasswordEmail_(email, tempPassword, false);
  } catch (err) {
    return { success: false, error: 'メール送信に失敗しました。時間をおいて再度お試しください。' };
  }

  return { success: true, message: '仮パスワードをメールに送信しました。' };
}

function handleSetNewPassword(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const newPassword = body.newPassword || '';
  if (newPassword.length < 6) {
    return { success: false, error: '新しいパスワードは6文字以上で設定してください。' };
  }

  const sheet = getUsersSheet_();
  const found = findUserById_(sheet, userId);
  if (!found) {
    return { success: false, error: 'ユーザーが見つかりません。' };
  }

  const salt = generateSalt_();
  const passwordHash = hashPassword_(newPassword, salt);

  sheet.getRange(found.rowIndex, 3).setValue(passwordHash); // PasswordHash
  sheet.getRange(found.rowIndex, 4).setValue(salt);         // Salt
  sheet.getRange(found.rowIndex, 15).setValue('FALSE');     // PasswordResetRequired

  return { success: true };
}

function handleChangePassword(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const currentPassword = body.currentPassword || '';
  const newPassword = body.newPassword || '';
  if (!currentPassword) {
    return { success: false, error: '現在のパスワードを入力してください。' };
  }
  if (newPassword.length < 6) {
    return { success: false, error: '新しいパスワードは6文字以上で設定してください。' };
  }

  const sheet = getUsersSheet_();
  const found = findUserById_(sheet, userId);
  if (!found) {
    return { success: false, error: 'ユーザーが見つかりません。' };
  }

  const row = found.row;
  const storedHash = row[2];
  const currentSalt = row[3];
  const computedHash = hashPassword_(currentPassword, currentSalt);

  if (computedHash !== storedHash) {
    return { success: false, error: '現在のパスワードが正しくありません。' };
  }

  const newSalt = generateSalt_();
  const newHash = hashPassword_(newPassword, newSalt);

  sheet.getRange(found.rowIndex, 3).setValue(newHash);  // PasswordHash
  sheet.getRange(found.rowIndex, 4).setValue(newSalt);  // Salt
  sheet.getRange(found.rowIndex, 15).setValue('FALSE'); // PasswordResetRequired

  return { success: true };
}

function handleLogout(body) {
  const token = body.token;
  if (token) {
    CacheService.getScriptCache().remove('session_' + token);
  }
  return { success: true };
}

function createSession_(userId) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, userId, SESSION_TTL_SEC);
  return token;
}

function getUserIdFromToken_(token) {
  if (!token) return null;
  return CacheService.getScriptCache().get('session_' + token);
}

// ============================================================
// Googleドライブ ユーティリティ
// ============================================================

function getOrCreateFolder_(name, parent) {
  const base = parent || DriveApp.getRootFolder();
  const iter = base.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return base.createFolder(name);
}

function getUserFolder_(rootFolderName, userId) {
  const root = getOrCreateFolder_(rootFolderName, DriveApp.getRootFolder());
  return getOrCreateFolder_(userId, root);
}

function saveBase64ImageToDrive_(rootFolderName, userId, fileName, base64Data, mimeType) {
  const folder = getUserFolder_(rootFolderName, userId);
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId: file.getId(), directUrl: buildDirectImageUrl_(file.getId()) };
}

function saveRemoteImageToDrive_(rootFolderName, userId, fileName, remoteUrl) {
  const response = UrlFetchApp.fetch(remoteUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('画像のダウンロードに失敗しました (' + response.getResponseCode() + ')');
  }
  const folder = getUserFolder_(rootFolderName, userId);
  const blob = response.getBlob().setName(fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId: file.getId(), directUrl: buildDirectImageUrl_(file.getId()) };
}

function buildDirectImageUrl_(fileId) {
  return 'https://lh3.googleusercontent.com/d/' + fileId + '=s1024';
}

// ============================================================
// マイフォト（前後左右）
// ============================================================

const PHOTO_ANGLE_COLUMN = { front: 8, back: 9, left: 10, right: 11 }; // 0-indexed列（配列インデックス）

function handleUploadPhoto(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const angle = body.angle;
  if (!PHOTO_ANGLE_COLUMN.hasOwnProperty(angle)) {
    return { success: false, error: '不正な撮影方向です: ' + angle };
  }
  if (!body.imageBase64) {
    return { success: false, error: '画像データがありません。' };
  }

  let saved;
  try {
    saved = saveBase64ImageToDrive_(USER_PHOTOS_FOLDER, userId, angle + '.jpg', body.imageBase64, body.mimeType || 'image/jpeg');
  } catch (err) {
    return { success: false, error: '写真の保存に失敗しました: ' + err.message };
  }

  const sheet = getUsersSheet_();
  const found = findUserById_(sheet, userId);
  if (!found) {
    return { success: false, error: 'ユーザーが見つかりません。' };
  }

  const colIndex = PHOTO_ANGLE_COLUMN[angle] + 1; // シートは1始まり
  sheet.getRange(found.rowIndex, colIndex).setValue(saved.directUrl);

  return { success: true, angle: angle, url: saved.directUrl };
}

function handleGetProfile(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const sheet = getUsersSheet_();
  const found = findUserById_(sheet, userId);
  if (!found) {
    return { success: false, error: 'ユーザーが見つかりません。' };
  }

  const row = found.row;
  return {
    success: true,
    profile: {
      email: row[1],
      ageGroup: row[5],
      usualStyle: row[6],
      ngElements: row[7],
      photoFront: row[8] || '',
      photoBack: row[9] || '',
      photoLeft: row[10] || '',
      photoRight: row[11] || '',
      gender: row[12] || '',
      hairStyle: row[13] || ''
    }
  };
}

function handleUpdateProfile(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const sheet = getUsersSheet_();
  const found = findUserById_(sheet, userId);
  if (!found) {
    return { success: false, error: 'ユーザーが見つかりません。' };
  }

  // 列: F=AgeGroup(6) G=UsualStyle(7) H=NgElements(8) M=Gender(13) N=HairStyle(14)
  const rowIndex = found.rowIndex;
  sheet.getRange(rowIndex, 6).setValue(body.ageGroup || '');
  sheet.getRange(rowIndex, 7).setValue(body.usualStyle || '');
  sheet.getRange(rowIndex, 8).setValue(body.ngElements || '');
  sheet.getRange(rowIndex, 13).setValue(body.gender || '');
  sheet.getRange(rowIndex, 14).setValue(body.hairStyle || '');

  return {
    success: true,
    profile: {
      ageGroup: body.ageGroup || '',
      usualStyle: body.usualStyle || '',
      ngElements: body.ngElements || '',
      gender: body.gender || '',
      hairStyle: body.hairStyle || ''
    }
  };
}

// ============================================================
// クローゼット（手持ちの服）
// ============================================================

const WARDROBE_CATEGORIES = ['トップス', 'ボトムス', 'アウター', 'ワンピース', '靴', 'バッグ', '小物'];

function rowToWardrobeItem_(row) {
  return {
    itemId: row[0],
    category: row[2] || '',
    name: row[3] || '',
    color: row[4] || '',
    brand: row[5] || '',
    memo: row[6] || '',
    createdAt: (row[7] instanceof Date) ? row[7].toISOString() : String(row[7] || ''),
    photoUrl: row[8] || ''
  };
}

function getWardrobeItems_(userId) {
  const sheet = getWardrobeSheet_();
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId) list.push(rowToWardrobeItem_(data[i]));
  }
  return list;
}

function handleGetWardrobe(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }
  const items = getWardrobeItems_(userId);
  items.sort(function (a, b) {
    const ca = WARDROBE_CATEGORIES.indexOf(a.category);
    const cb = WARDROBE_CATEGORIES.indexOf(b.category);
    if (ca !== cb) return (ca === -1 ? 99 : ca) - (cb === -1 ? 99 : cb);
    return new Date(b.createdAt) - new Date(a.createdAt);
  });
  return { success: true, items: items };
}

function handleAddWardrobeItem(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }
  const category = String(body.category || '').trim();
  const name = String(body.name || '').trim();
  if (!category || !name) {
    return { success: false, error: 'カテゴリとアイテム名を入力してください。' };
  }

  const sheet = getWardrobeSheet_();
  const itemId = Utilities.getUuid();
  const createdAt = new Date();
  const row = [
    itemId, userId, category, name.slice(0, 100),
    String(body.color || '').trim().slice(0, 50),
    String(body.brand || '').trim().slice(0, 50),
    String(body.memo || '').trim().slice(0, 200),
    createdAt,
    isOwnDriveImageUrl_(body.photoUrl) ? body.photoUrl : ''
  ];
  sheet.appendRow(row);
  return { success: true, item: rowToWardrobeItem_(row) };
}

function handleDeleteWardrobeItem(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }
  const sheet = getWardrobeSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.itemId && data[i][1] === userId) {
      const photoUrl = data[i][8] || '';
      sheet.deleteRow(i + 1);
      // 同じ写真を使っている他のアイテムが無ければドライブの画像もゴミ箱へ
      if (photoUrl) {
        const stillUsed = data.some(function (r, idx) { return idx !== i && idx > 0 && r[8] === photoUrl; });
        if (!stillUsed) trashDriveImageByUrl_(photoUrl);
      }
      return { success: true };
    }
  }
  return { success: false, error: 'アイテムが見つかりませんでした。' };
}

// ------------------------------------------------------------
// クローゼット：写真からAIで服を読み取る
// ------------------------------------------------------------

const WARDROBE_COLORS = ['白', 'ベージュ', 'グレー', 'ネイビー', 'ブラック', 'ブラウン', 'カーキ', 'デニムブルー', 'グリーン', 'ボルドー', 'ブルー', 'レッド', 'イエロー', 'ピンク', '柄物', 'その他'];

function isOwnDriveImageUrl_(url) {
  return typeof url === 'string' && url.indexOf('https://lh3.googleusercontent.com/d/') === 0;
}

function trashDriveImageByUrl_(url) {
  try {
    const m = String(url).match(/\/d\/([^=\/?]+)/);
    if (m) DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (e) {
    // 画像の削除失敗は無視（アイテム削除は成功扱い）
  }
}

function handleAnalyzeWardrobePhoto(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }
  if (!body.imageBase64) {
    return { success: false, error: '画像データがありません。' };
  }
  const mimeType = body.mimeType || 'image/jpeg';

  let items;
  try {
    items = callGeminiForWardrobePhoto_(body.imageBase64, mimeType);
  } catch (err) {
    return { success: false, error: '写真の読み取りに失敗しました: ' + err.message };
  }

  if (!items.length) {
    return { success: false, error: '写真から服を見つけられませんでした。服全体が写るように撮り直してみてください。' };
  }

  let saved;
  try {
    saved = saveBase64ImageToDrive_(WARDROBE_PHOTOS_FOLDER, userId, 'wardrobe_' + new Date().getTime() + '.jpg', body.imageBase64, mimeType);
  } catch (err) {
    return { success: false, error: '写真の保存に失敗しました: ' + err.message };
  }

  return { success: true, photoUrl: saved.directUrl, items: items };
}

function callGeminiForWardrobePhoto_(imageBase64, mimeType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEYが設定されていません。');
  }

  const prompt = [
    'あなたはアパレルショップの店員です。この写真に写っている「服・靴・バッグ・小物」を読み取り、クローゼット登録用のデータにしてください。',
    '',
    '【ルール】',
    '・写っているアイテムを1点ずつ、最大5点まで挙げてください（人物が着ている場合も、着用している服を1点ずつ）。',
    '・category は次のいずれかから選んでください: ' + WARDROBE_CATEGORIES.join('、'),
    '・color は次のいずれかから最も近いものを選んでください: ' + WARDROBE_COLORS.join('、'),
    '・name は「オックスフォードシャツ」「テーパードチノパン」「白スニーカー」のように、形や素材が分かる短い日本語名にしてください（色は含めない）。',
    '・brand はロゴやタグでブランドがはっきり読み取れる場合のみ入れ、分からなければ空文字にしてください。推測で入れないでください。',
    '・memo にはシルエットや素材感など、コーデ提案に役立つ特徴を20文字程度で入れてください（例: 「ゆったりめ・厚手コットン」）。',
    '・服が写っていない場合は items を空配列にしてください。',
    '',
    '【出力形式】JSONのみ',
    '{ "items": [ { "category": "トップス", "name": "アイテム名", "color": "白", "brand": "", "memo": "特徴" } ] }'
  ].join('\n');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  const payload = {
    contents: [{
      role: 'user',
      parts: [
        { inline_data: { mime_type: mimeType, data: imageBase64 } },
        { text: prompt }
      ]
    }],
    generationConfig: { temperature: 0.2, responseMimeType: 'application/json' }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini APIエラー (' + response.getResponseCode() + '): ' + response.getContentText());
  }

  const json = JSON.parse(response.getContentText());
  if (!json.candidates || !json.candidates.length) {
    throw new Error('Gemini APIから有効な応答が得られませんでした。');
  }
  const text = json.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('');
  const cleaned = text.replace(/```json/g, '').replace(/```/g, '').trim();

  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    throw new Error('AIの応答をJSONとして解析できませんでした。');
  }

  return (parsed.items || []).slice(0, 5).map(function (it) {
    return {
      category: WARDROBE_CATEGORIES.indexOf(it.category) !== -1 ? it.category : 'トップス',
      name: String(it.name || '').trim().slice(0, 100),
      color: WARDROBE_COLORS.indexOf(it.color) !== -1 ? it.color : '',
      brand: String(it.brand || '').trim().slice(0, 50),
      memo: String(it.memo || '').trim().slice(0, 200)
    };
  }).filter(function (it) { return it.name; });
}

// ============================================================
// スタイリスト診断・コーデ提案（Gemini API連携）
// ============================================================

const PROPOSAL_MODES = {
  new: '新しく購入してコーデ',
  owned: '手持ちの服だけでコーデ',
  mix: '手持ちの服＋買い足しでコーデ'
};

function toBudgetNumber_(v) {
  const n = parseInt(String(v || '').replace(/[^0-9]/g, ''), 10);
  return isNaN(n) || n <= 0 ? 0 : n;
}

function toPriceNumber_(v) {
  if (typeof v === 'number') return Math.max(0, Math.round(v));
  const n = parseInt(String(v || '').replace(/[^0-9]/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

// 提案に価格集計・予算判定を付与する
function annotateProposalPrices_(proposal, mode, budgetMin, budgetMax, wardrobeMap) {
  let total = 0;
  let purchaseCount = 0;
  let ownedCount = 0;

  (proposal.items || []).forEach(function (item) {
    const ownedById = item.wardrobe_id && wardrobeMap[item.wardrobe_id];
    const isOwned = mode === 'owned' || item.is_owned === true || item.is_owned === 'true' || !!ownedById;
    item.is_owned = isOwned;
    if (isOwned) {
      item.price = 0;
      ownedCount++;
      if (ownedById) {
        // 名前などは登録内容を正とする
        item.wardrobe_label = [ownedById.color, ownedById.brand, ownedById.name].filter(String).join(' ');
      }
    } else {
      item.price = toPriceNumber_(item.price);
      total += item.price;
      purchaseCount++;
    }
  });

  proposal.mode = mode;
  proposal.total_price = total;
  proposal.purchase_count = purchaseCount;
  proposal.owned_count = ownedCount;
  proposal.budget_min = budgetMin;
  proposal.budget_max = budgetMax;
  proposal.budget_ok = !((budgetMax && total > budgetMax) || (budgetMin && purchaseCount > 0 && total < budgetMin));
  return proposal;
}


function handleGenerateProposal(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const brands = Array.isArray(body.brands) && body.brands.length ? body.brands : ['指定なし/ミックス'];
  const topColors = Array.isArray(body.topColors) ? body.topColors : [];
  const bottomColors = Array.isArray(body.bottomColors) ? body.bottomColors : [];
  const topStyles = Array.isArray(body.topStyles) ? body.topStyles : [];
  const bottomStyles = Array.isArray(body.bottomStyles) ? body.bottomStyles : [];
  const scene = body.scene || '';
  const mood = body.mood || '';
  const temperature = body.temperature || '';
  const ownedItems = body.ownedItems || '';
  const mode = PROPOSAL_MODES.hasOwnProperty(body.mode) ? body.mode : 'new';
  let budgetMin = mode === 'owned' ? 0 : toBudgetNumber_(body.budgetMin);
  let budgetMax = mode === 'owned' ? 0 : toBudgetNumber_(body.budgetMax);
  if (budgetMin && budgetMax && budgetMin > budgetMax) {
    const tmp = budgetMin; budgetMin = budgetMax; budgetMax = tmp;
  }

  if (!scene || !mood) {
    return { success: false, error: 'シーン/TPOと気分・テイストを入力してください。' };
  }

  // 手持ちの服（クローゼット）
  let wardrobe = [];
  if (mode !== 'new') {
    const all = getWardrobeItems_(userId);
    const ids = Array.isArray(body.wardrobeIds) ? body.wardrobeIds : null;
    wardrobe = ids && ids.length ? all.filter(function (w) { return ids.indexOf(w.itemId) !== -1; }) : all;
    if (!wardrobe.length) {
      return { success: false, error: 'クローゼットに手持ちの服が登録されていません。「クローゼット」タブから登録してください。' };
    }
  }
  const wardrobeMap = {};
  wardrobe.forEach(function (w) { wardrobeMap[w.itemId] = w; });

  let proposal;
  try {
    const usersSheet = getUsersSheet_();
    const userRecord = findUserById_(usersSheet, userId);
    const profile = userRecord ? {
      ageGroup: userRecord.row[5],
      usualStyle: userRecord.row[6],
      ngElements: userRecord.row[7],
      gender: userRecord.row[12],
      hairStyle: userRecord.row[13]
    } : { ageGroup: '', usualStyle: '', ngElements: '', gender: '', hairStyle: '' };

    const options = {
      brands: brands, topColors: topColors, bottomColors: bottomColors, topStyles: topStyles, bottomStyles: bottomStyles,
      scene: scene, mood: mood, temperature: temperature, ownedItems: ownedItems, profile: profile,
      mode: mode, budgetMin: budgetMin, budgetMax: budgetMax, wardrobe: wardrobe, retryNote: ''
    };

    proposal = annotateProposalPrices_(callGeminiForProposal_(options), mode, budgetMin, budgetMax, wardrobeMap);

    // 予算外なら1回だけ作り直す
    if (!proposal.budget_ok && (budgetMin || budgetMax)) {
      options.retryNote = '前回の提案は購入合計が約' + proposal.total_price + '円で、予算の範囲外でした。必ず予算内に収まるようにアイテムを選び直してください。';
      const retry = annotateProposalPrices_(callGeminiForProposal_(options), mode, budgetMin, budgetMax, wardrobeMap);
      if (retry.budget_ok || distanceFromBudget_(retry, budgetMin, budgetMax) < distanceFromBudget_(proposal, budgetMin, budgetMax)) {
        proposal = retry;
      }
    }
  } catch (err) {
    return { success: false, error: 'AI提案の生成に失敗しました: ' + err.message };
  }

  // 履歴保存
  const historySheet = getHistorySheet_();
  const historyId = Utilities.getUuid();
  const createdAt = new Date();
  historySheet.appendRow([
    historyId,
    userId,
    createdAt,
    brands.join('、'),
    scene,
    mood,
    temperature,
    ownedItems,
    JSON.stringify(proposal),
    '',
    topColors.join('、'),
    bottomColors.join('、'),
    topStyles.join('、'),
    bottomStyles.join('、'),
    mode,
    budgetMin || '',
    budgetMax || '',
    proposal.total_price
  ]);

  return {
    success: true,
    historyId: historyId,
    createdAt: createdAt.toISOString(),
    proposal: proposal
  };
}

function distanceFromBudget_(p, min, max) {
  if (max && p.total_price > max) return p.total_price - max;
  if (min && p.total_price < min) return min - p.total_price;
  return 0;
}

function formatYen_(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '円';
}

function buildStylistPrompt_(o) {
  const brandList = (o.brands || []).join('、');
  const topColors = o.topColors || [];
  const bottomColors = o.bottomColors || [];
  const topStyles = o.topStyles || [];
  const bottomStyles = o.bottomStyles || [];
  const profile = o.profile || {};
  const mode = o.mode || 'new';
  const wardrobe = o.wardrobe || [];

  // 予算の説明
  let budgetText = '指定なし';
  if (o.budgetMin && o.budgetMax) budgetText = formatYen_(o.budgetMin) + '〜' + formatYen_(o.budgetMax);
  else if (o.budgetMax) budgetText = formatYen_(o.budgetMax) + '以内';
  else if (o.budgetMin) budgetText = formatYen_(o.budgetMin) + '以上';

  // クローゼット一覧
  const wardrobeLines = wardrobe.map(function (w) {
    return '・[ID:' + w.itemId + '] ' + w.category + '：' + [w.color, w.brand, w.name].filter(String).join(' ') + (w.memo ? '（' + w.memo + '）' : '');
  });

  const modeRules = [];
  if (mode === 'owned') {
    modeRules.push(
      '【今回の提案モード：手持ちの服だけでコーディネート】',
      '・下の「相談者のクローゼット」にある服だけを使ってコーディネートを組んでください。新しい服の購入は提案しないでください。',
      '・itemsの各アイテムは必ずクローゼットの中から選び、"is_owned": true、"wardrobe_id"に該当のID、"price": 0 としてください。',
      '・"name"にはクローゼットに登録された名前をそのまま使ってください。"brand"は登録がなければ「手持ち」としてください。',
      '・クローゼットに靴などが無いカテゴリは無理に含めなくて構いません。その場合は stylist_comment で「あると便利な買い足し候補」を一言添えてください。'
    );
  } else if (mode === 'mix') {
    modeRules.push(
      '【今回の提案モード：手持ちの服＋買い足しでコーディネート】',
      '・下の「相談者のクローゼット」の服を1点以上必ず活かし、足りないアイテムや印象を変えるアイテムを新しく買い足す提案をしてください。',
      '・手持ちの服を使うアイテムは "is_owned": true、"wardrobe_id"に該当のID、"price": 0 としてください。',
      '・買い足すアイテムは "is_owned": false、"wardrobe_id": "" とし、"price" に価格の目安を入れてください。',
      '・買い足しは「手持ちの服との相性」を description で具体的に説明してください。',
      '・買い足すアイテムの価格合計が予算内に収まるようにしてください。'
    );
  } else {
    modeRules.push(
      '【今回の提案モード：新しく購入してコーディネート】',
      '・すべてのアイテムを "is_owned": false、"wardrobe_id": "" とし、"price" に価格の目安を入れてください。',
      '・ただし「使いたい手持ちアイテム」の記載があるものをコーデに含める場合は、そのアイテムだけ "is_owned": true、"price": 0 としてください。',
      '・購入するアイテムの価格合計が予算内に収まるようにしてください。'
    );
  }

  const lines = [
    'あなたはプロのファッションスタイリストです。',
    '相談者に対して、身近な大衆服ブランド（ユニクロ・無印良品・GU・ワークマン・ZARA・H&M・GLOBAL WORK・Mac-House・coen・L.L.Bean・Belluna・Right-on・COMME CA ISM・DoCLASSE・ORIHICA・SUIT SELECT・UNITED ARROWS・鎌倉シャツ・Eddie Bauer・BEAMS F・SHIPS・Brooks Brothers・EDIFICEなど）を中心に、',
    '「誰でも真似しやすく、お洒落に見える」コーディネートを提案してください。',
    '',
    '【今回、相談者が希望するブランド】',
    brandList,
    '※「指定なし/ミックス」の場合は、上記の大衆ブランドから自由に組み合わせてください。',
    '',
    '【相談者の登録プロフィール】',
    '・性別: ' + (profile.gender || '未回答'),
    '・年代: ' + (profile.ageGroup || '不明'),
    '・髪型: ' + (profile.hairStyle || '未登録'),
    '・普段の系統: ' + (profile.usualStyle || '指定なし'),
    '・避けたい要素（必ず提案から除外すること）: ' + (profile.ngElements || 'なし'),
    '',
    '【相談者の情報】',
    '・シーン/TPO: ' + o.scene,
    '・気分・テイスト: ' + o.mood,
    '・気温・季節感: ' + (o.temperature || '指定なし'),
    '・使いたい手持ちアイテム: ' + (o.ownedItems || '特になし'),
    '・トップスの色の希望: ' + (topColors.length ? topColors.join('、') : '指定なし'),
    '・ボトムスの色の希望: ' + (bottomColors.length ? bottomColors.join('、') : '指定なし'),
    '・トップスのスタイルの希望: ' + (topStyles.length ? topStyles.join('、') : '指定なし'),
    '・ボトムスのスタイルの希望: ' + (bottomStyles.length ? bottomStyles.join('、') : '指定なし'),
    '・購入予算（新しく買うアイテムの税込合計）: ' + (mode === 'owned' ? '購入なし' : budgetText),
    ''
  ];

  if (wardrobe.length) {
    lines.push('【相談者のクローゼット（手持ちの服）】');
    wardrobeLines.forEach(function (l) { lines.push(l); });
    lines.push('');
  }

  modeRules.forEach(function (l) { lines.push(l); });
  lines.push('');

  lines.push(
    '【提案のルール】',
    '・購入するアイテムは、希望ブランドの定番アイテムを、実在しやすい具体的なアイテム名（例: 「ユニクロのスマートアンクルパンツ」「無印良品の洗いざらしオックスボタンダウンシャツ」）で挙げてください。',
    '・購入するアイテムの "price" は、そのブランドの通常販売価格（日本国内・税込・円）の目安を整数で入れてください。セール価格ではなく定価の目安にしてください。',
    '・トップス、ボトムス、靴、必要に応じて羽織りものや小物を含めて、全体のコーディネートを構成してください。',
    '・トップスの色の希望、ボトムスの色の希望、トップスのスタイルの希望、ボトムスのスタイルの希望が指定されている場合は、できる限りそれに沿ったアイテムを提案してください（手持ちの服を使う場合はクローゼットの内容を優先）。',
    '・登録プロフィールの「避けたい要素」に該当するアイテムやテイストは絶対に提案しないでください。',
    '・予算が指定されている場合、購入アイテムの price の合計は必ず予算内に収めてください。予算の下限がある場合は下限を大きく下回らないようにしてください。',
    '・「高見え」させるための着こなしのコツ（サイズ感の選び方、タックインの有無、ロールアップ、色合わせのルールなど）を、初心者にも分かるように論理的に説明してください。',
    '・最後にスタイリストからのワンポイントアドバイスを添えてください。'
  );

  if (o.retryNote) {
    lines.push('', '【重要な修正指示】', o.retryNote);
  }

  lines.push(
    '',
    '【出力形式】',
    '必ず以下のJSON形式のみで出力してください。前置きや説明文、Markdownのコードブロック記号は一切不要です。',
    '{',
    '  "items": [',
    '    { "category": "トップス", "brand": "ブランド名", "name": "アイテム名", "description": "選び方・色・サイズ感の説明", "styling_tip": "着こなしのコツ", "price": 2990, "is_owned": false, "wardrobe_id": "" }',
    '  ],',
    '  "coordination_summary": "全体のコーディネートの狙いや世界観の要約",',
    '  "styling_points": ["高見えテクニック1", "高見えテクニック2"],',
    '  "stylist_comment": "スタイリストからのワンポイントアドバイス"',
    '}',
    '',
    'itemsは3〜5点程度（トップス・ボトムス・靴を原則含める）にしてください。'
  );

  return lines.join('\n');
}

function callGeminiForProposal_(options) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEYが設定されていません。');
  }

  const prompt = buildStylistPrompt_(options);
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [
      { role: 'user', parts: [{ text: prompt }] }
    ],
    generationConfig: {
      temperature: options.retryNote ? 0.6 : 0.9,
      responseMimeType: 'application/json'
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode !== 200) {
    throw new Error('Gemini APIエラー (' + statusCode + '): ' + responseText);
  }

  const json = JSON.parse(responseText);
  const candidates = json.candidates;
  if (!candidates || !candidates.length) {
    throw new Error('Gemini APIから有効な応答が得られませんでした。');
  }

  const text = candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('');
  const cleaned = text.replace(/```json/g, '').replace(/```/g, '').trim();

  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    throw new Error('AIの応答をJSONとして解析できませんでした。');
  }

  return parsed;
}

// ============================================================
// 試着イメージ生成（LightX API + Googleドライブ保存）
// ============================================================

function buildTryOnPrompt_(proposal) {
  const items = (proposal.items || []).map(function (item) {
    return (item.category || '') + ': ' + (item.brand || '') + ' ' + (item.name || '');
  }).join('、');
  const summary = proposal.coordination_summary || '';
  return 'この人物に次のコーディネートを着せてください。' + summary + ' 着用アイテム: ' + items + '。自然な質感で、顔・体型・背景はそのままに、服だけを置き換えてください。';
}

function callLightXOutfit_(imageUrl, textPrompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('LIGHTX_API_KEY');
  if (!apiKey) {
    throw new Error('LIGHTX_API_KEYが設定されていません。');
  }

  const submitResponse = UrlFetchApp.fetch(LIGHTX_OUTFIT_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey },
    payload: JSON.stringify({ imageUrl: imageUrl, textPrompt: textPrompt }),
    muteHttpExceptions: true
  });

  if (submitResponse.getResponseCode() !== 200) {
    throw new Error('LightX APIエラー (' + submitResponse.getResponseCode() + '): ' + submitResponse.getContentText());
  }

  const submitJson = JSON.parse(submitResponse.getContentText());

  if (submitJson.status === 'FAIL' || submitJson.statusCode >= 5000) {
    if (submitJson.statusCode === 5040 || (submitJson.message || '').indexOf('API_CREDITS_CONSUMED') !== -1) {
      throw new Error('LightXのクレジット残高が不足しています。https://www.lightxeditor.com/pricing/api からクレジットを追加してください。');
    }
    throw new Error('LightX APIエラー: ' + (submitJson.message || submitResponse.getContentText()));
  }

  const orderId =
    (submitJson.body && submitJson.body.orderId) ||
    submitJson.orderId ||
    (submitJson.data && submitJson.data.orderId) ||
    (submitJson.body && submitJson.body.requestId) ||
    submitJson.requestId;

  if (!orderId) {
    throw new Error('LightX APIからorderIdが返されませんでした。レスポンス: ' + submitResponse.getContentText());
  }

  // 3秒間隔で最大10回ポーリング
  for (let i = 0; i < 10; i++) {
    Utilities.sleep(3000);

    const statusResponse = UrlFetchApp.fetch(LIGHTX_STATUS_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey },
      payload: JSON.stringify({ orderId: orderId }),
      muteHttpExceptions: true
    });

    if (statusResponse.getResponseCode() !== 200) continue;

    const statusJson = JSON.parse(statusResponse.getContentText());
    const body = statusJson.body || statusJson;
    const status = body.status;

    if (status === 'active') {
      return body.output;
    }
    if (status === 'failed') {
      throw new Error('試着イメージの生成に失敗しました（LightX側でエラー）。');
    }
    // 'init' の場合は継続してポーリング
  }

  throw new Error('試着イメージの生成がタイムアウトしました。しばらくしてから再度お試しください。');
}

function handleGenerateTryOnImage(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const historyId = body.historyId;
  if (!historyId) {
    return { success: false, error: '対象の提案が指定されていません。' };
  }

  const usersSheet = getUsersSheet_();
  const userRecord = findUserById_(usersSheet, userId);
  if (!userRecord || !userRecord.row[8]) {
    return { success: false, error: '正面の写真が登録されていません。「マイフォト」から正面写真を登録してください。' };
  }
  const frontPhotoUrl = userRecord.row[8];

  const historySheet = getHistorySheet_();
  const data = historySheet.getDataRange().getValues();
  let targetRowIndex = -1;
  let proposal = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === historyId && data[i][1] === userId) {
      targetRowIndex = i + 1;
      try {
        proposal = JSON.parse(data[i][8]);
      } catch (e) {
        proposal = null;
      }
      break;
    }
  }

  if (targetRowIndex === -1 || !proposal) {
    return { success: false, error: '対象の提案履歴が見つかりませんでした。' };
  }

  let outputUrl;
  try {
    const prompt = buildTryOnPrompt_(proposal);
    outputUrl = callLightXOutfit_(frontPhotoUrl, prompt);
  } catch (err) {
    return { success: false, error: err.message };
  }

  let saved;
  try {
    saved = saveRemoteImageToDrive_(TRYON_RESULTS_FOLDER, userId, historyId + '.jpg', outputUrl);
  } catch (err) {
    return { success: false, error: '試着イメージのドライブ保存に失敗しました: ' + err.message };
  }

  historySheet.getRange(targetRowIndex, 10).setValue(saved.directUrl);

  return { success: true, tryOnImageUrl: saved.directUrl };
}

// ============================================================
// 履歴
// ============================================================

function handleGetHistory(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const sheet = getHistorySheet_();
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === userId) {
      let summary = '';
      try {
        const parsed = JSON.parse(row[8]);
        summary = parsed.coordination_summary || '';
      } catch (e) {
        summary = '';
      }
      list.push({
        historyId: row[0],
        createdAt: (row[2] instanceof Date) ? row[2].toISOString() : String(row[2]),
        brands: row[3],
        scene: row[4],
        mood: row[5],
        temperature: row[6],
        summary: summary,
        tryOnImageUrl: row[9] || '',
        mode: row[14] || 'new',
        totalPrice: (row[17] === '' || row[17] === undefined) ? null : Number(row[17])
      });
    }
  }

  list.sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });

  return { success: true, history: list };
}

function handleGetHistoryDetail(body) {
  const userId = getUserIdFromToken_(body.token);
  if (!userId) {
    return { success: false, error: 'ログインが必要です。再度ログインしてください。' };
  }

  const sheet = getHistorySheet_();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === userId && row[0] === body.historyId) {
      let proposal;
      try {
        proposal = JSON.parse(row[8]);
      } catch (e) {
        proposal = null;
      }
      return {
        success: true,
        detail: {
          historyId: row[0],
          createdAt: (row[2] instanceof Date) ? row[2].toISOString() : String(row[2]),
          brands: row[3],
          scene: row[4],
          mood: row[5],
          temperature: row[6],
          ownedItems: row[7],
          proposal: proposal,
          tryOnImageUrl: row[9] || ''
        }
      };
    }
  }

  return { success: false, error: '履歴が見つかりませんでした。' };
}
