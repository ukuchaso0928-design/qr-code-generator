// DOM要素の取得
const idInput = document.getElementById('idInput');
const generateBtn = document.getElementById('generateBtn');
const resultSection = document.getElementById('resultSection');
const qrcodeDiv = document.getElementById('qrcode');
const generatedIdSpan = document.getElementById('generatedId');
const downloadBtn = document.getElementById('downloadBtn');

// バッチ処理用のDOM要素
const batchInput = document.getElementById('batchInput');
const generateBatchBtn = document.getElementById('generateBatchBtn');
const fileInput = document.getElementById('fileInput');
const fileUploadBtn = document.getElementById('fileUploadBtn');
const fileName = document.getElementById('fileName');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');

// バッチ処理結果表示用のDOM要素
const batchResultSection = document.getElementById('batchResultSection');
const batchCount = document.getElementById('batchCount');
const batchFileSize = document.getElementById('batchFileSize');
const batchPreview = document.getElementById('batchPreview');

// タブ要素
const tabBtns = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');

// 現在生成されたQRコードのデータを保存
let currentQRCodeData = null;
let batchQRCodeData = []; // バッチ処理用のQRコードデータ

// QRコードライブラリの読み込み状態（HTMLで定義済み）
// let qrCodeLibraryLoaded = false; // HTMLで定義されているため削除

// イベントリスナーの設定
generateBtn.addEventListener('click', generateQRCode);
idInput.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        generateQRCode();
    }
});
downloadBtn.addEventListener('click', downloadQRCode);

// バッチ処理用のイベントリスナー
generateBatchBtn.addEventListener('click', generateBatchQRCode);
fileUploadBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileUpload);
downloadExcelBtn.addEventListener('click', downloadExcelFile);

// タブ切り替えのイベントリスナー
tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        const tabName = btn.getAttribute('data-tab');
        switchTab(tabName);
    });
});

// QRコード生成関数
async function generateQRCode() {
    const id = idInput.value.trim();
    
    // 入力値の検証
    if (!id) {
        showError('IDを入力してください。');
        return;
    }
    
    if (id.length > 50) {
        showError('IDは50文字以内で入力してください。');
        return;
    }
    
    // 特殊文字の検証（基本的な文字のみ許可）
    if (!/^[a-zA-Z0-9\-_]+$/.test(id)) {
        showError('IDは英数字、ハイフン、アンダースコアのみ使用できます。');
        return;
    }
    
    try {
        // ローディング状態を表示
        setLoadingState(true);
        
        // 既存のQRコードをクリア
        qrcodeDiv.innerHTML = '';
        
        // QRコードライブラリの読み込み確認
        const qrApiType = checkQRCodeLibraryStatus();
        if (qrApiType) {
            let qrCodeDataURL;
            
            if (qrApiType === 'qrcode') {
                // 標準のQRCode APIを使用
                qrCodeDataURL = await QRCode.toDataURL(id, {
                    width: 256,
                    margin: 2,
                    color: {
                        dark: '#000000',
                        light: '#FFFFFF'
                    },
                    errorCorrectionLevel: 'M'
                });
            } else if (qrApiType === 'qrcode_alt') {
                // 代替のqrcode APIを使用
                const qr = qrcode(0, 'M');
                qr.addData(id);
                qr.make();
                qrCodeDataURL = qr.createDataURL(8, 0);
            }
            
            // QRコードにIDテキストを重ねて表示
            const qrWithIdDataURL = await addIdToQRCode(qrCodeDataURL, id);
            
            const img = document.createElement('img');
            img.src = qrWithIdDataURL;
            img.alt = `QR Code for ${id}`;
            img.style.maxWidth = '100%';
            img.style.height = 'auto';
            qrcodeDiv.appendChild(img);
            
            // 現在のQRコードデータを保存
            currentQRCodeData = qrWithIdDataURL;
            
            showSuccess('QRコードが正常に生成されました！');
        } else {
            // シンプルなQRコードを生成
            const simpleQR = generateSimpleQRCode(id);
            qrcodeDiv.appendChild(simpleQR);
            
            // ダウンロード機能は無効化
            currentQRCodeData = null;
            
            showSuccess('IDが表示されました（ライブラリ読み込み中）');
        }
        
        // 生成されたIDを表示
        generatedIdSpan.textContent = id;
        
        // 結果セクションを表示
        resultSection.style.display = 'block';
        
        // 入力フィールドをクリア
        idInput.value = '';
        
    } catch (error) {
        console.error('QRコード生成エラー:', error);
        showError('QRコードの生成中にエラーが発生しました。');
    } finally {
        setLoadingState(false);
    }
}

// ダウンロード機能
function downloadQRCode() {
    if (!currentQRCodeData) {
        showError('ダウンロードするQRコードがありません。');
        return;
    }
    
    try {
        // ダウンロード用のリンクを作成
        const link = document.createElement('a');
        link.download = `qrcode_${generatedIdSpan.textContent}_${getCurrentDateString()}.png`;
        link.href = currentQRCodeData;
        
        // ダウンロードを実行
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showSuccess('QRコードをダウンロードしました！');
    } catch (error) {
        console.error('ダウンロードエラー:', error);
        showError('ダウンロード中にエラーが発生しました。');
    }
}

// ローディング状態の設定
function setLoadingState(isLoading) {
    generateBtn.disabled = isLoading;
    if (isLoading) {
        generateBtn.textContent = '生成中...';
        generateBtn.classList.add('loading');
    } else {
        generateBtn.textContent = 'QRコードを生成';
        generateBtn.classList.remove('loading');
    }
}

// エラーメッセージの表示
function showError(message) {
    // 既存のメッセージを削除
    removeExistingMessages();
    
    const errorDiv = document.createElement('div');
    errorDiv.className = 'message error-message';
    errorDiv.textContent = message;
    errorDiv.style.cssText = `
        background: #f8d7da;
        color: #721c24;
        padding: 12px 20px;
        border-radius: 8px;
        margin: 15px 0;
        border: 1px solid #f5c6cb;
        animation: fadeIn 0.3s ease-out;
    `;
    
    idInput.parentNode.insertBefore(errorDiv, idInput.nextSibling);
    
    // 3秒後に自動削除
    setTimeout(() => {
        if (errorDiv.parentNode) {
            errorDiv.parentNode.removeChild(errorDiv);
        }
    }, 3000);
}

// 成功メッセージの表示
function showSuccess(message) {
    // 既存のメッセージを削除
    removeExistingMessages();
    
    const successDiv = document.createElement('div');
    successDiv.className = 'message success-message';
    successDiv.textContent = message;
    successDiv.style.cssText = `
        background: #d4edda;
        color: #155724;
        padding: 12px 20px;
        border-radius: 8px;
        margin: 15px 0;
        border: 1px solid #c3e6cb;
        animation: fadeIn 0.3s ease-out;
    `;
    
    idInput.parentNode.insertBefore(successDiv, idInput.nextSibling);
    
    // 3秒後に自動削除
    setTimeout(() => {
        if (successDiv.parentNode) {
            successDiv.parentNode.removeChild(successDiv);
        }
    }, 3000);
}

// 既存のメッセージを削除
function removeExistingMessages() {
    const existingMessages = document.querySelectorAll('.message');
    existingMessages.forEach(message => {
        if (message.parentNode) {
            message.parentNode.removeChild(message);
        }
    });
}

// 現在の日時を文字列で取得
function getCurrentDateString() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    
    return `${year}${month}${day}_${hours}${minutes}`;
}

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOMContentLoaded - 初期化開始');
    
    // DOM要素の存在確認
    console.log('DOM要素の確認:');
    console.log('- idInput:', idInput);
    console.log('- batchInput:', batchInput);
    console.log('- generateBatchBtn:', generateBatchBtn);
    console.log('- progressContainer:', progressContainer);
    console.log('- batchResultSection:', batchResultSection);
    
    // 入力フィールドにフォーカス
    if (idInput) {
        idInput.focus();
    }
    
    // 入力フィールドのプレースホルダーを動的に変更
    const placeholders = [
        '例: 1001',
        '例: 1002',
        '例: 1003',
        '例: 1004'
    ];
    
    let placeholderIndex = 0;
    setInterval(() => {
        if (idInput) {
            idInput.placeholder = placeholders[placeholderIndex];
            placeholderIndex = (placeholderIndex + 1) % placeholders.length;
        }
    }, 3000);
    
    console.log('初期化完了');
});

// QRコードライブラリの読み込み状態を確認する関数
function checkQRCodeLibraryStatus() {
    if (!qrCodeLibraryLoaded) {
        // ライブラリが読み込まれていない場合、シンプルなQRコードを生成
        console.log('外部ライブラリが利用できないため、シンプルなQRコードを生成します');
        return false;
    }
    
    // 利用可能なQRコードAPIを確認
    if (typeof QRCode !== 'undefined') {
        console.log('QRCode APIが利用可能です');
        return 'qrcode';
    } else if (typeof qrcode !== 'undefined') {
        console.log('qrcode APIが利用可能です');
        return 'qrcode_alt';
    } else {
        console.log('QRコードAPIが見つかりません');
        return false;
    }
}

// シンプルなQRコード生成（ライブラリが利用できない場合の代替）
function generateSimpleQRCode(id) {
    // シンプルなテキストベースのQRコード風表示
    const qrContainer = document.createElement('div');
    qrContainer.style.cssText = `
        width: 256px;
        height: 256px;
        border: 2px solid #333;
        background: white;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        font-family: monospace;
        text-align: center;
        padding: 20px;
        box-sizing: border-box;
    `;
    
    const title = document.createElement('div');
    title.textContent = 'QR CODE';
    title.style.cssText = `
        font-size: 16px;
        font-weight: bold;
        margin-bottom: 10px;
        color: #333;
    `;
    
    const idDisplay = document.createElement('div');
    idDisplay.textContent = id;
    idDisplay.style.cssText = `
        font-size: 14px;
        word-break: break-all;
        color: #666;
        margin-bottom: 10px;
    `;
    
    const note = document.createElement('div');
    note.textContent = '※ ライブラリ読み込み中';
    note.style.cssText = `
        font-size: 10px;
        color: #999;
    `;
    
    qrContainer.appendChild(title);
    qrContainer.appendChild(idDisplay);
    qrContainer.appendChild(note);
    
    return qrContainer;
}

// タブ切り替え関数
function switchTab(tabName) {
    // すべてのタブボタンとコンテンツからactiveクラスを削除
    tabBtns.forEach(btn => btn.classList.remove('active'));
    tabContents.forEach(content => content.classList.remove('active'));
    
    // 選択されたタブをアクティブにする
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    document.getElementById(`${tabName}-tab`).classList.add('active');
}

// ファイルアップロード処理
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    fileName.textContent = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        batchInput.value = e.target.result;
    };
    reader.readAsText(file);
}

// CSVデータから患者IDと患者名を抽出する関数
function extractPatientDataFromCSV(csvData) {
    console.log('CSVデータの解析開始');
    
    const lines = csvData.trim().split('\n');
    console.log(`CSV行数: ${lines.length}`);
    
    // ヘッダー行をスキップ（最初の行）
    const dataLines = lines.slice(1);
    console.log(`データ行数: ${dataLines.length}`);
    
    const patientData = [];
    
    for (let i = 0; i < dataLines.length; i++) {
        const line = dataLines[i].trim();
        if (!line) continue;
        
        // CSVの各行をカンマで分割
        const columns = line.split(',');
        console.log(`行 ${i + 1}:`, columns);
        
        // 最初の列が患者ID（数値のみ）、2番目の列が患者名
        const firstColumn = columns[0].trim();
        const patientName = columns[1] ? columns[1].trim() : ''; // 2番目の列（インデックス1）
        
        if (firstColumn && /^\d+$/.test(firstColumn)) {
            patientData.push({
                id: firstColumn,
                name: patientName || '名前不明'
            });
            console.log(`患者データ抽出: ID=${firstColumn}, 名前=${patientName}`);
        }
    }
    
    console.log(`抽出された患者データ数: ${patientData.length}`);
    console.log('最初の5個の患者データ:', patientData.slice(0, 5));
    
    return patientData;
}

// バッチQRコード生成関数
async function generateBatchQRCode() {
    console.log('バッチQRコード生成開始');
    console.log('batchInput要素:', batchInput);
    console.log('batchInput.value:', batchInput.value);
    
    const inputData = batchInput.value.trim();
    let patientData = [];
    
    // CSV形式かどうかを判定（カンマが含まれているか）
    if (inputData.includes(',')) {
        console.log('CSV形式のデータを検出');
        patientData = extractPatientDataFromCSV(inputData);
    } else {
        console.log('通常のテキスト形式のデータを検出');
        const ids = inputData.split('\n').filter(id => id.trim());
        patientData = ids.map(id => ({ id: id, name: '名前不明' }));
    }
    
    console.log(`最終的な患者データ数: ${patientData.length}`);
    console.log('最初の5個の患者データ:', patientData.slice(0, 5));
    
    if (patientData.length === 0) {
        console.log('患者データが入力されていません');
        showError('患者データを入力してください。');
        return;
    }
    
    if (patientData.length > 1000) {
        console.log('患者データ数が上限を超えています:', patientData.length);
        showError('一度に処理できる患者データは1000個までです。');
        return;
    }
    
    // 入力値の検証（患者IDは数字のみ）
    const invalidData = patientData.filter(data => !/^\d+$/.test(data.id.trim()));
    if (invalidData.length > 0) {
        console.log('無効なIDが見つかりました:', invalidData.slice(0, 5));
        showError(`無効なIDが見つかりました: ${invalidData.slice(0, 5).map(d => d.id).join(', ')}${invalidData.length > 5 ? '...' : ''}`);
        return;
    }
    
    // JSZipライブラリの確認
    console.log('JSZip読み込み状況:', jszipLoaded);
    if (!jszipLoaded) {
        console.log('JSZipライブラリが読み込まれていません');
        showError('JSZipライブラリが読み込まれていません。ページを再読み込みしてください。');
        return;
    }
    
    try {
        console.log('プログレス表示を開始');
        // プログレス表示を開始
        showProgress(true);
        batchQRCodeData = [];
        
        console.log('QRコードライブラリの状態を確認中...');
        const qrApiType = checkQRCodeLibraryStatus();
        console.log('QRコードライブラリの状態:', qrApiType);
        
        if (!qrApiType) {
            console.log('QRコードライブラリが利用できません');
            showError('QRコードライブラリが読み込まれていません。');
            showProgress(false);
            return;
        }
        
        console.log(`使用するQRコードAPI: ${qrApiType}`);
        
        // バッチ処理を実行
        for (let i = 0; i < patientData.length; i++) {
            const patient = patientData[i];
            const id = patient.id.trim();
            updateProgress(i + 1, patientData.length, `QRコード生成中: ${id} (${patient.name})`);
            
            try {
                let qrCodeDataURL;
                
                if (qrApiType === 'qrcode') {
                    qrCodeDataURL = await QRCode.toDataURL(id, {
                        width: 256,
                        margin: 2,
                        color: {
                            dark: '#000000',
                            light: '#FFFFFF'
                        },
                        errorCorrectionLevel: 'M'
                    });
                } else if (qrApiType === 'qrcode_alt') {
                    const qr = qrcode(0, 'M');
                    qr.addData(id);
                    qr.make();
                    qrCodeDataURL = qr.createDataURL(8, 0);
                }
                
                batchQRCodeData.push({
                    id: id,
                    name: patient.name,
                    dataURL: qrCodeDataURL
                });
                
                console.log(`QRコード生成完了: ${id} (${patient.name}) (${i + 1}/${patientData.length})`);
                
            } catch (error) {
                console.error(`QRコード生成エラー (${id}):`, error);
                showError(`QRコード生成エラー: ${id}`);
                showProgress(false);
                return;
            }
            
            // ブラウザの応答性を保つため少し待機
            if (i % 10 === 0) {
                await new Promise(resolve => setTimeout(resolve, 10));
            }
        }
        
        updateProgress(patientData.length, patientData.length, 'QRコード生成完了');
        
        showProgress(false);
        showSuccess(`${patientData.length}個のQRコードが正常に生成されました！`);
        downloadExcelBtn.style.display = 'inline-block';
        
        // バッチ処理結果を表示
        displayBatchResults(patientData.length);
        
    } catch (error) {
        console.error('バッチQRコード生成エラー:', error);
        showError('QRコードの生成中にエラーが発生しました。');
        showProgress(false);
    }
}

// プログレス表示の制御
function showProgress(show) {
    console.log('プログレス表示制御:', show);
    console.log('progressContainer要素:', progressContainer);
    
    if (progressContainer) {
        progressContainer.style.display = show ? 'block' : 'none';
        console.log('プログレスコンテナの表示状態:', progressContainer.style.display);
    } else {
        console.error('progressContainer要素が見つかりません');
    }
    
    if (generateBatchBtn) {
        generateBatchBtn.disabled = show;
        console.log('バッチボタンの無効化状態:', generateBatchBtn.disabled);
    } else {
        console.error('generateBatchBtn要素が見つかりません');
    }
    
    if (!show) {
        if (progressFill) {
            progressFill.style.width = '0%';
        }
        if (progressText) {
            progressText.textContent = '処理中...';
        }
    }
}

// プログレス更新
function updateProgress(current, total, message) {
    const percentage = (current / total) * 100;
    progressFill.style.width = `${percentage}%`;
    progressText.textContent = `${message} (${current}/${total})`;
}

// ZIPファイル作成
async function createZipFile() {
    console.log('ZIPファイル作成開始');
    
    if (typeof JSZip === 'undefined') {
        throw new Error('JSZipライブラリが読み込まれていません');
    }
    
    if (batchQRCodeData.length === 0) {
        throw new Error('QRコードデータがありません');
    }
    
    const zip = new JSZip();
    
    for (const qrData of batchQRCodeData) {
        try {
            // データURLからBase64データを抽出
            const base64Data = qrData.dataURL.split(',')[1];
            if (!base64Data) {
                throw new Error(`無効なデータURL: ${qrData.id}`);
            }
            zip.file(`${qrData.id}.png`, base64Data, { base64: true });
        } catch (error) {
            console.error(`ZIPファイル追加エラー (${qrData.id}):`, error);
            throw new Error(`ZIPファイル追加エラー: ${qrData.id}`);
        }
    }
    
    console.log(`${batchQRCodeData.length}個のファイルをZIPに追加`);
    
    // ZIPファイルを生成
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    console.log(`ZIPファイルサイズ: ${zipBlob.size} bytes`);
    
    // ダウンロード用のURLを作成
    const url = URL.createObjectURL(zipBlob);
    downloadZipBtn.onclick = () => {
        const link = document.createElement('a');
        link.download = `qrcodes_${getCurrentDateString()}.zip`;
        link.href = url;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    };
    
    console.log('ZIPファイル作成完了');
}

// バッチ処理結果の表示
function displayBatchResults(count) {
    console.log('バッチ処理結果を表示中...');
    
    // 結果セクションを表示
    batchResultSection.style.display = 'block';
    
    // 生成数を表示
    batchCount.textContent = count;
    
    // ファイルサイズを計算（概算）
    const estimatedSize = (count * 0.5).toFixed(1); // 1つのQRコード約0.5KB
    batchFileSize.textContent = estimatedSize;
    
    // プレビューを表示（最初の12個まで）
    batchPreview.innerHTML = '';
    const previewCount = Math.min(12, batchQRCodeData.length);
    
    for (let i = 0; i < previewCount; i++) {
        const qrData = batchQRCodeData[i];
        const previewItem = document.createElement('div');
        previewItem.className = 'batch-preview-item';
        
        const img = document.createElement('img');
        img.src = qrData.dataURL;
        img.alt = `QR Code for ${qrData.id}`;
        
        const label = document.createElement('div');
        label.className = 'id-label';
        label.textContent = qrData.id;
        
        previewItem.appendChild(img);
        previewItem.appendChild(label);
        batchPreview.appendChild(previewItem);
    }
    
    // プレビューが全て表示されていない場合のメッセージ
    if (batchQRCodeData.length > previewCount) {
        const moreMessage = document.createElement('div');
        moreMessage.style.cssText = `
            text-align: center;
            padding: 10px;
            color: #666;
            font-style: italic;
            border-top: 1px solid #e1e5e9;
            margin-top: 10px;
        `;
        moreMessage.textContent = `他 ${batchQRCodeData.length - previewCount} 個のQRコードが生成されました`;
        batchPreview.appendChild(moreMessage);
    }
    
    console.log(`バッチ処理結果表示完了: ${count}個のQRコード`);
}

// エクセルファイル作成（QRコード画像埋め込み版）
async function createExcelFile() {
    console.log('エクセルファイル作成開始（QRコード画像埋め込み）');
    
    if (typeof ExcelJS === 'undefined') {
        throw new Error('ExcelJSライブラリが読み込まれていません');
    }
    
    if (batchQRCodeData.length === 0) {
        throw new Error('QRコードデータがありません');
    }
    
    // 新しいワークブックを作成
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('QRコード一覧');
    
    // ヘッダー行を設定
    worksheet.columns = [
        { header: '患者ID', key: 'patientId', width: 15 },
        { header: 'QRコード画像', key: 'qrCode', width: 30 },
        { header: '患者名', key: 'patientName', width: 25 }
    ];
    
    // ヘッダー行のスタイルを設定
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
    };
    
    // QRコードデータを追加
    for (let i = 0; i < batchQRCodeData.length; i++) {
        const qrData = batchQRCodeData[i];
        const row = worksheet.addRow({
            patientId: qrData.id,
            qrCode: '',
            patientName: qrData.name
        });
        
        // 行の高さを設定（QRコード画像とテキスト用に高く）
        row.height = 140;
        
            // QRコード画像を埋め込み
            try {
                // QRコード画像にIDテキストを重ねて作成
                const qrWithIdDataURL = await addIdToQRCode(qrData.dataURL, qrData.id);
                
                // データURLからBase64データを抽出
                const base64Data = qrWithIdDataURL.split(',')[1];
                
                // 画像をワークシートに追加
                const imageId = workbook.addImage({
                    base64: base64Data,
                    extension: 'png'
                });
                
                // 画像の位置を設定（B列、対応する行）
                worksheet.addImage(imageId, {
                    tl: { col: 1, row: i + 1 }, // 左上の位置（B列、データ行）
                    ext: { width: 100, height: 140 } // 画像サイズ（大きなIDテキスト用に高く）
                });
                
                console.log(`QRコード画像とID埋め込み完了: ${qrData.id}`);
                
            } catch (error) {
                console.error(`QRコード画像埋め込みエラー (${qrData.id}):`, error);
            }
    }
    
    // エクセルファイルを生成
    const buffer = await workbook.xlsx.writeBuffer();
    
    // ダウンロード用のBlobを作成
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // ダウンロード用のURLを作成
    const url = URL.createObjectURL(blob);
    
    // 直接ダウンロードを実行
    const link = document.createElement('a');
    link.download = `qrcodes_${getCurrentDateString()}.xlsx`;
    link.href = url;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    console.log('エクセルファイル作成完了（QRコード画像埋め込み）');
}

// QRコード画像にIDテキストを重ねる関数
async function addIdToQRCode(qrDataURL, id) {
    return new Promise((resolve, reject) => {
        try {
            // 新しいCanvasを作成
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            // QRコード画像を読み込み
            const img = new Image();
            img.onload = function() {
                // Canvasサイズを設定（QRコード + IDテキスト用の余白）
                canvas.width = 256;
                canvas.height = 320; // QRコード256px + IDテキスト用64px（大きなフォント用）
                
                // 背景を白に設定
                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                
                // QRコード画像を描画
                ctx.drawImage(img, 0, 0, 256, 256);
                
                // IDテキストを描画（3倍のサイズ）
                ctx.fillStyle = '#000000';
                ctx.font = 'bold 48px Arial'; // 16px → 48px（3倍）
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                
                // IDテキストをQRコードの下に描画
                ctx.fillText(`ID: ${id}`, canvas.width / 2, 256 + 24);
                
                // CanvasをDataURLに変換
                const dataURL = canvas.toDataURL('image/png');
                resolve(dataURL);
            };
            
            img.onerror = function() {
                reject(new Error('QRコード画像の読み込みに失敗しました'));
            };
            
            img.src = qrDataURL;
            
        } catch (error) {
            reject(error);
        }
    });
}

// エクセルファイルダウンロード
async function downloadExcelFile() {
    try {
        console.log('エクセルファイルダウンロード開始');
        
        if (!exceljsLoaded) {
            showError('ExcelJSライブラリが読み込まれていません。ページを再読み込みしてください。');
            return;
        }
        
        if (batchQRCodeData.length === 0) {
            showError('エクセルファイルに出力するQRコードデータがありません。');
            return;
        }
        
        // プログレス表示
        showProgress(true);
        updateProgress(0, 100, 'エクセルファイル作成中...');
        
        // エクセルファイルを作成（この中でダウンロードも実行される）
        await createExcelFile();
        
        showProgress(false);
        showSuccess('QRコード画像付きエクセルファイルが正常に作成されました！');
        
    } catch (error) {
        console.error('エクセルファイル作成エラー:', error);
        showError('エクセルファイルの作成中にエラーが発生しました。');
        showProgress(false);
    }
}

// キーボードショートカット
document.addEventListener('keydown', function(e) {
    // Ctrl+Enter または Cmd+Enter で生成
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        const activeTab = document.querySelector('.tab-content.active');
        if (activeTab.id === 'single-tab') {
            generateQRCode();
        } else if (activeTab.id === 'batch-tab') {
            generateBatchQRCode();
        }
    }
    
    // Escape キーで結果セクションを非表示
    if (e.key === 'Escape') {
        resultSection.style.display = 'none';
        const activeTab = document.querySelector('.tab-content.active');
        if (activeTab.id === 'single-tab') {
            idInput.focus();
        } else {
            batchInput.focus();
        }
    }
});
