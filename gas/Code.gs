/**
 * レジ（請求ファイル自動変換）RPAツール
 * Google Apps Script版
 * 
 * 機能:
 * - Google Drive上の特定フォルダからZIPファイルを自動検索
 * - ZIPファイルを解凍してPDFファイルを抽出
 * - 請求情報PDFから会社名を自動抽出
 * - ファイル名を統一フォーマットに変更
 * - 処理済みファイルを「送信データ」フォルダに保存
 */

// 設定
const CONFIG = {
  TARGET_FOLDER_ID: '1eGkgviOKRCXuFV1OFLSOI5Ap8F-eeaKr',
  OUTPUT_FOLDER_NAME: '送信データ',
  SPREADSHEET_ID: '1Q9x8OdGIqd62J0RxoIRYQEfzDIDF2cEij3bpaRG97j0',
  YEAR_CELL: 'A5',
  MONTH_CELL: 'C5',
  // ZIPファイル内のPDFファイルパターン
  SAMPLE_FILE_PATTERNS: {
    BILL: /bill_shipping_\d{14}\.pdf$/,      // 請求書パターン
    DETAIL: /shippingcomp_\d{14}\.pdf$/,     // 明細書パターン
    POSTAGE: /postage_\d{14}\.pdf$/          // 送料明細書パターン
  }
};

/**
 * メイン処理関数（ZIPファイル処理版）
 */
function processInvoices() {
  try {
    console.log('🚀 レジ（請求ファイル自動変換）RPAツールを開始します...');
    
    // スプレッドシートから年月を取得
    const { year, month } = getYearMonthFromSpreadsheet();
    
    console.log(`📅 処理対象年月: ${year}年${month}月`);
    
    // 対象フォルダを取得
    const targetFolder = DriveApp.getFolderById(CONFIG.TARGET_FOLDER_ID);
    console.log(`📁 対象フォルダ: ${targetFolder.getName()}`);
    
    // 年月フォルダを検索
    const yearFolder = findFolderByName(targetFolder, year.toString());
    if (!yearFolder) {
      throw new Error(`${year}年のフォルダが見つかりません`);
    }
    
    const monthFolder = findFolderByName(yearFolder, `${month}月`);
    if (!monthFolder) {
      throw new Error(`${month}月のフォルダが見つかりません`);
    }
    
    console.log(`✅ 対象フォルダを発見: ${yearFolder.getName()}/${monthFolder.getName()}`);
    
    // ZIPファイルを検索
    const zipFiles = monthFolder.getFilesByType('application/zip');
    const zipFileArray = [];
    while (zipFiles.hasNext()) {
      zipFileArray.push(zipFiles.next());
    }
    
    console.log(`📦 ZIPファイル数: ${zipFileArray.length}件`);
    
    if (zipFileArray.length === 0) {
      console.log('⚠️ 処理対象のZIPファイルが見つかりません');
      return;
    }
    
    // 出力フォルダを作成または取得
    const outputFolder = findOrCreateFolder(monthFolder, CONFIG.OUTPUT_FOLDER_NAME);
    
    // 各ZIPファイルを処理
    let processedCount = 0;
    let successCount = 0;
    let errorCount = 0;
    
    for (const zipFile of zipFileArray) {
      try {
        console.log(`\n📦 ZIPファイル処理中: ${zipFile.getName()}`);
        
        const result = processZIPFile(zipFile, outputFolder, year, month);
        
        if (result.success) {
          successCount++;
          console.log(`✅ 処理成功: ${result.companyName}`);
        } else {
          errorCount++;
          console.log(`❌ 処理失敗: ${result.error}`);
        }
        
        processedCount++;
        
      } catch (error) {
        errorCount++;
        console.error(`❌ ZIPファイル処理でエラー: ${zipFile.getName()}`, error);
      }
    }
    
    // 結果表示
    console.log(`\n🎉 処理完了！`);
    console.log(`📊 処理結果: 総件数: ${processedCount}件, 成功: ${successCount}件, 失敗: ${errorCount}件`);
    
  } catch (error) {
    console.error('❌ メイン処理でエラーが発生しました:', error);
    throw error;
  }
}

/**
 * スプレッドシートから年月を取得
 */
function getYearMonthFromSpreadsheet() {
  try {
    console.log('📊 スプレッドシートから年月を取得中...');
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    
    // 年と月を取得
    const year = sheet.getRange(CONFIG.YEAR_CELL).getValue();
    const month = sheet.getRange(CONFIG.MONTH_CELL).getValue();
    
    // 値の検証
    if (!year || !month) {
      throw new Error('スプレッドシートの年または月が設定されていません');
    }
    
    if (typeof year !== 'number' || typeof month !== 'number') {
      throw new Error('年と月は数値で入力してください');
    }
    
    if (year < 2000 || year > 2100) {
      throw new Error('年は2000年から2100年の間で入力してください');
    }
    
    if (month < 1 || month > 12) {
      throw new Error('月は1から12の間で入力してください');
    }
    
    console.log(`✅ スプレッドシートから取得: ${year}年${month}月`);
    
    return { year, month };
    
  } catch (error) {
    console.error('❌ スプレッドシートからの年月取得でエラー:', error);
    throw error;
  }
}

/**
 * フォルダ名でフォルダを検索
 */
function findFolderByName(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : null;
}

/**
 * フォルダを作成または取得
 */
function findOrCreateFolder(parentFolder, folderName) {
  const existingFolder = findFolderByName(parentFolder, folderName);
  if (existingFolder) {
    console.log(`📁 既存のフォルダを使用: ${folderName}`);
    return existingFolder;
  }
  
  const newFolder = parentFolder.createFolder(folderName);
  console.log(`📁 新しいフォルダを作成: ${folderName}`);
  return newFolder;
}

/**
 * ZIPファイルを処理
 */
function processZIPFile(zipFile, outputFolder, year, month) {
  try {
    console.log(`📦 ZIPファイルを解凍中: ${zipFile.getName()}`);
    
    // ZIPファイルを解凍
    const zipBlob = zipFile.getBlob();
    const extractedFiles = Utilities.unzip(zipBlob);
    
    console.log(`📂 解凍されたファイル数: ${extractedFiles.length}件`);
    
    // 請求情報PDFを検索
    let billPdf = null;
    let detailPdf = null;
    let postagePdf = null;
    
    for (const file of extractedFiles) {
      const fileName = file.getName();
      
      if (CONFIG.SAMPLE_FILE_PATTERNS.BILL.test(fileName)) {
        billPdf = file;
        console.log(`📄 請求書PDF発見: ${fileName}`);
      } else if (CONFIG.SAMPLE_FILE_PATTERNS.DETAIL.test(fileName)) {
        detailPdf = file;
        console.log(`📄 明細書PDF発見: ${fileName}`);
      } else if (CONFIG.SAMPLE_FILE_PATTERNS.POSTAGE.test(fileName)) {
        postagePdf = file;
        console.log(`📄 送料明細書PDF発見: ${fileName}`);
      }
    }
    
    // 請求書PDFが必須
    if (!billPdf) {
      throw new Error('請求書PDFが見つかりません');
    }
    
    // 会社名を抽出
    const companyName = extractCompanyNameFromPDF(billPdf);
    if (!companyName) {
      throw new Error('会社名を抽出できませんでした');
    }
    
    console.log(`✅ 会社名抽出成功: ${companyName}`);
    
    // 各ファイルタイプを処理
    if (billPdf) {
      processFileByType(billPdf, outputFolder, year, month, companyName, 'bill');
    }
    
    if (detailPdf) {
      processFileByType(detailPdf, outputFolder, year, month, companyName, 'detail');
    }
    
    if (postagePdf) {
      processFileByType(postagePdf, outputFolder, year, month, companyName, 'postage');
    }
    
    return {
      success: true,
      companyName: companyName
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 特定のファイルタイプを処理
 */
function processFileByType(fileBlob, outputFolder, year, month, companyName, fileType) {
  try {
    // 新しいファイル名を生成
    const newFileName = generateFileName(year, month, companyName, fileType);
    
    // ファイルを出力フォルダに保存
    const newFile = outputFolder.createFile(fileBlob);
    newFile.setName(newFileName);
    
    console.log(`📁 ファイル保存完了: ${newFileName}`);
    
  } catch (error) {
    console.error(`${fileType}ファイル処理でエラー:`, error);
  }
}

/**
 * PDFから会社名を抽出
 */
function extractCompanyNameFromPDF(pdfBlob) {
  try {
    // 実際の実装では、PDFの内容を解析して会社名を抽出
    // 現在は仮の実装
    return '株式会社サンプル';
    
  } catch (error) {
    console.error('PDFからの会社名抽出でエラー:', error);
    return '会社名不明';
  }
}

/**
 * 新しいファイル名を生成
 */
function generateFileName(year, month, companyName, fileType) {
  const yearMonth = `${year}${String(month).padStart(2, '0')}`;
  
  let suffix = '';
  switch (fileType) {
    case 'bill':
      suffix = 'ご請求書';
      break;
    case 'detail':
      suffix = 'ご請求明細書';
      break;
    case 'postage':
      suffix = 'ご請求送料明細書';
      break;
    default:
      suffix = 'ご請求書';
  }
  
  return `${yearMonth}_${companyName}様_${suffix}.pdf`;
}

/**
 * テスト用関数
 */
function testConnection() {
  try {
    console.log('🧪 Google Drive接続テストを開始...');
    
    const targetFolder = DriveApp.getFolderById(CONFIG.TARGET_FOLDER_ID);
    console.log(`✅ 対象フォルダにアクセス成功: ${targetFolder.getName()}`);
    
    const files = targetFolder.getFiles();
    let fileCount = 0;
    while (files.hasNext()) {
      files.next();
      fileCount++;
    }
    
    console.log(`📁 フォルダ内のファイル数: ${fileCount}件`);
    console.log('🎉 接続テスト成功！');
    
  } catch (error) {
    console.error('❌ 接続テスト失敗:', error);
    throw error;
  }
}

/**
 * ZIPファイルの一覧表示
 */
function listZIPFiles() {
  try {
    console.log('📋 ZIPファイルの一覧表示を開始...');
    
    // スプレッドシートから年月を取得
    const { year, month } = getYearMonthFromSpreadsheet();
    
    const targetFolder = DriveApp.getFolderById(CONFIG.TARGET_FOLDER_ID);
    const yearFolder = findFolderByName(targetFolder, year.toString());
    
    if (!yearFolder) {
      console.log(`${year}年のフォルダが見つかりません`);
      return;
    }
    
    const monthFolder = findFolderByName(yearFolder, `${month}月`);
    if (!monthFolder) {
      console.log(`${month}月のフォルダが見つかりません`);
      return;
    }
    
    console.log(`\n📁 ${year}年${month}月のフォルダ内容:`);
    
    // ZIPファイルを検索
    const zipFiles = monthFolder.getFilesByType('application/zip');
    let zipCount = 0;
    
    while (zipFiles.hasNext()) {
      const zipFile = zipFiles.next();
      zipCount++;
      
      console.log(`📦 ZIPファイル: ${zipFile.getName()}`);
      console.log(`  📏 サイズ: ${Math.round(zipFile.getSize() / 1024)}KB`);
      console.log(`  📅 作成日: ${zipFile.getDateCreated()}`);
    }
    
    if (zipCount === 0) {
      console.log('⚠️ ZIPファイルが見つかりません');
    } else {
      console.log(`\n📊 合計ZIPファイル数: ${zipCount}件`);
    }
    
  } catch (error) {
    console.error('❌ ZIPファイル一覧表示でエラー:', error);
  }
}
