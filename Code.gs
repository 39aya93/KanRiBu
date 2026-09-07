/*************************************************
 * 備品管理アプリ
 * Google Apps Script
 *************************************************/


// ================================================
// 基本設定
// ================================================

const SHEET_NAME = '備品';
const SETTING_SHEET_NAME = '設定';
const IMAGE_FOLDER_NAME = '備品管理_画像';


// ================================================
// Webアプリ起動
// ================================================

function doGet() {

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('備品管理');

}


// ================================================
// 初期設定
// ================================================

function setup() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  // --------------------------------
  // 備品シート
  // --------------------------------

  let sheet =
    ss.getSheetByName(SHEET_NAME);

  if (!sheet) {

    sheet =
      ss.insertSheet(SHEET_NAME);

  }


  const headers = [
    'ID',
    '名称',
    '仕様（サイズ）',
    '画像URL',
    '画像ファイルID',
    '保管場所',
    '入手時期',
    '保管期限',
    '購入先',
    '価格',
    '備考',
    '登録日時'
  ];


  // 1行目に不足しているヘッダーを設定
  const currentHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .getValues()[0];


  headers.forEach(function(header, index) {

    if (
      currentHeaders[index] !== header
    ) {

      sheet
        .getRange(
          1,
          index + 1
        )
        .setValue(header);

    }

  });


  sheet.setFrozenRows(1);


  // --------------------------------
  // 設定シート
  // --------------------------------

  let settingSheet =
    ss.getSheetByName(
      SETTING_SHEET_NAME
    );


  if (!settingSheet) {

    settingSheet =
      ss.insertSheet(
        SETTING_SHEET_NAME
      );

  }


  // ヘッダー
  settingSheet
    .getRange('A1')
    .setValue('保管場所');

  settingSheet
    .getRange('B1')
    .setValue('購入先');


  // --------------------------------
  // 保管場所の初期値
  // --------------------------------

  const defaultLocations = [
    '工具棚A',
    '工具棚B',
    'ガレージ',
    '物置',
    'パントリー',
    '倉庫',
    '車庫'
  ];


  const existingLocations =
    getLocations();


  defaultLocations.forEach(
    function(location) {

      if (
        existingLocations.indexOf(
          location
        ) === -1
      ) {

        const row =
          Math.max(
            settingSheet.getLastRow() + 1,
            2
          );

        settingSheet
          .getRange(
            row,
            1
          )
          .setValue(location);

      }

    }
  );


  // --------------------------------
  // 購入先の初期値
  // --------------------------------

  const defaultStores = [
    'カインズ',
    'コーナン',
    'Amazon',
    '楽天市場',
    'モノタロウ',
    'その他'
  ];


  const existingStores =
    getStores();


  defaultStores.forEach(
    function(store) {

      if (
        existingStores.indexOf(
          store
        ) === -1
      ) {

        const row =
          Math.max(
            settingSheet.getLastRow() + 1,
            2
          );

        settingSheet
          .getRange(
            row,
            2
          )
          .setValue(store);

      }

    }
  );


  // --------------------------------
  // Google Drive画像フォルダ
  // --------------------------------

  getImageFolder();


  return '初期設定が完了しました。';

}


// ================================================
// 画像保存フォルダ取得
// ================================================

function getImageFolder() {

  const folders =
    DriveApp.getFoldersByName(
      IMAGE_FOLDER_NAME
    );


  if (folders.hasNext()) {

    return folders.next();

  }


  return DriveApp.createFolder(
    IMAGE_FOLDER_NAME
  );

}


// ================================================
// 保管場所一覧
// ================================================

function getLocations() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      SETTING_SHEET_NAME
    );


  if (!sheet) {
    return [];
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  return values
    .map(function(row) {

      return String(
        row[0]
      ).trim();

    })
    .filter(function(value) {

      return value !== '';

    });

}


// ================================================
// 保管場所追加
// ================================================

function addLocation(location) {

  location =
    String(
      location || ''
    ).trim();


  if (!location) {

    throw new Error(
      '保管場所を入力してください。'
    );

  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  let sheet =
    ss.getSheetByName(
      SETTING_SHEET_NAME
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        SETTING_SHEET_NAME
      );

    sheet
      .getRange('A1')
      .setValue('保管場所');

    sheet
      .getRange('B1')
      .setValue('購入先');

  }


  const locations =
    getLocations();


  if (
    locations.indexOf(
      location
    ) === -1
  ) {

    const row =
      Math.max(
        sheet.getLastRow() + 1,
        2
      );

    sheet
      .getRange(
        row,
        1
      )
      .setValue(location);

  }


  return getLocations();

}


// ================================================
// 購入先一覧
// ================================================

function getStores() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      SETTING_SHEET_NAME
    );


  if (!sheet) {
    return [];
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        2,
        lastRow - 1,
        1
      )
      .getValues();


  return values
    .map(function(row) {

      return String(
        row[0]
      ).trim();

    })
    .filter(function(value) {

      return value !== '';

    });

}


// ================================================
// 購入先追加
// ================================================

function addStore(store) {

  store =
    String(
      store || ''
    ).trim();


  if (!store) {

    throw new Error(
      '購入先を入力してください。'
    );

  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  let sheet =
    ss.getSheetByName(
      SETTING_SHEET_NAME
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        SETTING_SHEET_NAME
      );

    sheet
      .getRange('A1')
      .setValue('保管場所');

    sheet
      .getRange('B1')
      .setValue('購入先');

  }


  const stores =
    getStores();


  if (
    stores.indexOf(
      store
    ) === -1
  ) {

    const row =
      Math.max(
        sheet.getLastRow() + 1,
        2
      );

    sheet
      .getRange(
        row,
        2
      )
      .setValue(store);

  }


  return getStores();

}


// ================================================
// ID生成
// ================================================

function createId() {

  return (
    Utilities.getUuid()
  );

}


// ================================================
// 備品登録
// ================================================

function saveItem(item) {

  if (!item) {

    throw new Error(
      '登録データがありません。'
    );

  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      SHEET_NAME
    );


  if (!sheet) {

    throw new Error(
      '「備品」シートがありません。setup()を実行してください。'
    );

  }


  // --------------------------------
  // 保管場所チェック
  // --------------------------------

  if (
    item.location === '__ADD__'
  ) {

    throw new Error(
      '保管場所を選択してください。'
    );

  }


  // --------------------------------
  // 購入先チェック
  // --------------------------------

  if (
    item.store === '__ADD_STORE__'
  ) {

    throw new Error(
      '購入先を選択してください。'
    );

  }


  // --------------------------------
  // ID
  // --------------------------------

  const id =
    createId();


  // --------------------------------
  // 画像処理
  // --------------------------------

  let imageUrl = '';
  let imageFileId = '';


  if (
    item.imageData
  ) {

    const folder =
      getImageFolder();


    const base64 =
      item.imageData
        .split(',')[1];


    const bytes =
      Utilities.base64Decode(
        base64
      );


    const contentType =
      item.imageType ||
      'image/jpeg';


    let extension =
      'jpg';


    if (
      contentType ===
      'image/png'
    ) {

      extension = 'png';

    } else if (
      contentType ===
      'image/webp'
    ) {

      extension = 'webp';

    }


    const blob =
      Utilities.newBlob(
        bytes,
        contentType,
        item.name +
          '_' +
          id +
          '.' +
          extension
      );


    const file =
      folder.createFile(
        blob
      );


    imageFileId =
      file.getId();


    // 画像表示用
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );


    imageUrl =
      'https://drive.google.com/thumbnail?id=' +
      imageFileId +
      '&sz=w800';

  }


  // --------------------------------
  // 登録日時
  // --------------------------------

  const now =
    new Date();


  // --------------------------------
  // データ登録
  // --------------------------------

  sheet.appendRow([
    id,
    item.name || '',
    item.spec || '',
    imageUrl,
    imageFileId,
    item.location || '',
    item.acquiredDate || '',
    item.expiryDate || '',
    item.store || '',
    item.price || '',
    item.note || '',
    now
  ]);


  return {
    success: true,
    message: '登録しました。'
  };

}


// ================================================
// 全データ取得
// ================================================

function getItems() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      SHEET_NAME
    );


  if (!sheet) {
    return [];
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        12
      )
      .getValues();


  return values.map(
    function(row) {

      return rowToObject(row);

    }
  );

}


// ================================================
// 名称検索
// ================================================

function searchItems(keyword) {

  const items =
    getItems();


  keyword =
    String(
      keyword || ''
    )
      .trim()
      .toLowerCase();


  if (!keyword) {

    return items;

  }


  return items.filter(
    function(item) {

      return String(
        item.name || ''
      )
        .toLowerCase()
        .indexOf(keyword) !== -1;

    }
  );

}


// ================================================
// ID検索
// ================================================

function getItemById(id) {

  const items =
    getItems();


  return items.find(
    function(item) {

      return item.id === id;

    }
  ) || null;

}


// ================================================
// 備品削除
// ================================================

function deleteItem(id) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      SHEET_NAME
    );


  if (!sheet) {

    throw new Error(
      '備品シートがありません。'
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      '削除するデータがありません。'
    );

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        12
      )
      .getValues();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(id)
    ) {


      // 画像ファイルも削除
      const fileId =
        values[i][4];


      if (fileId) {

        try {

          DriveApp
            .getFileById(
              fileId
            )
            .setTrashed(true);

        } catch (e) {

          // 画像削除失敗でも
          // データ削除は続行

        }

      }


      sheet.deleteRow(
        i + 2
      );


      return {
        success: true,
        message: '削除しました。'
      };

    }

  }


  throw new Error(
    '指定されたデータが見つかりません。'
  );

}


// ================================================
// シート1行 → オブジェクト
// ================================================

function rowToObject(row) {

  return {

    id:
      String(row[0] || ''),

    name:
      String(row[1] || ''),

    spec:
      String(row[2] || ''),

    imageUrl:
      String(row[3] || ''),

    imageFileId:
      String(row[4] || ''),

    location:
      String(row[5] || ''),

    acquiredDate:
      formatDate(row[6]),

    expiryDate:
      formatDate(row[7]),

    store:
      String(row[8] || ''),

    price:
      String(row[9] || ''),

    note:
      String(row[10] || ''),

    registeredAt:
      formatDateTime(row[11])

  };

}


// ================================================
// 日付整形
// ================================================

function formatDate(value) {

  if (!value) {
    return '';
  }


  if (
    Object.prototype
      .toString
      .call(value) ===
    '[object Date]'
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  }


  return String(value);

}


// ================================================
// 日時整形
// ================================================

function formatDateTime(value) {

  if (!value) {
    return '';
  }


  if (
    Object.prototype
      .toString
      .call(value) ===
    '[object Date]'
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm'
    );

  }


  return String(value);

}