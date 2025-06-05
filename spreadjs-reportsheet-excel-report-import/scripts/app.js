// 日本語カルチャー設定
GC.Spread.Common.CultureManager.culture("ja-jp");
//GC.Spread.Sheets.LicenseKey = "ここにSpreadJSのライセンスキーを設定します";

// SpreadJSの設定
document.addEventListener("DOMContentLoaded", () => {
    const spread = new GC.Spread.Sheets.Workbook("ss");
    const printButton = document.getElementById('print');
    const pdfButton = document.getElementById('pdf');
    const previousButton = document.getElementById('previous');
    const nextButton = document.getElementById('next');
    let reportSheet;

    //------------------------------------------
    // PDFエクスポートに必要なフォントを登録します
    //------------------------------------------
    registerFont("IPAexゴシック", "normal", "fonts/ipaexg.ttf");

    //----------------------------------------------------------------
    // sjs形式のテンプレートシートを読み込んでレポートシートを実行します
    //----------------------------------------------------------------
    console.log("読み込み開始");
    const res = fetch('reports/invoice.sjs').then((response) => response.blob())
        .then((myBlob) => {
            console.log(myBlob);

            spread.open(myBlob, () => {
                console.log(`読み込み成功`);
                reportSheet = spread.getSheetTab(0);

                // レポートシートのオプション設定
                reportSheet.options.renderMode = 'PaginatedPreview';
                //reportSheet.options.renderMode = 'Preview';
                reportSheet.options.printAllPages = true;

                // レポートシートの印刷設定
                var printInfo = reportSheet.printInfo();
                printInfo.showBorder(false);
                printInfo.zoomFactor(1);
                reportSheet.printInfo(printInfo);
                reportSheet.refresh();
                initPage()
            }, (e) => {
                console.log(`***ERR*** エラーコード（${e.errorCode}） : ${e.errorMessage}`);
            });
        });

    //------------------------------------------
    // 印刷ボタン押下時の処理
    //------------------------------------------
    printButton.onclick = function () {
        spread.print();
    }

    //------------------------------------------
    // PDF出力ボタン押下時の処理
    //------------------------------------------
    pdfButton.onclick = function () {
        spread.savePDF(function (blob) {
            //saveAs(blob, 'download.pdf');
            var url = URL.createObjectURL(blob);
            window.open(url);
        }, function (error) {
            console.log(error);
        }, {
            title: '請求書',
            author: 'テストAuthor',
            subject: 'テストSubject',
            keywords: 'テストKeywords',
            creator: 'テストCreator'
        })
    }

    //------------------------------------------
    // 前のページボタン押下時の処理
    //------------------------------------------    
    previousButton.onclick = function () {
        const page = reportSheet.currentPage();
        if (page != 0) {
            reportSheet.currentPage(page - 1);
            initPage()
        }
    }

    //------------------------------------------
    // 次のページボタン押下時の処理
    //------------------------------------------    
    nextButton.onclick = function () {
        const page = reportSheet.currentPage();
        if (page < reportSheet.getPagesCount() - 1) {
            reportSheet.currentPage(page + 1);
            initPage()
        }
    }

    //------------------------------------------
    // 現在のページと全ページ数の表示
    //------------------------------------------      
    function initPage() {
        document.getElementById('current').innerHTML = reportSheet.currentPage() + 1;
        document.getElementById('all').innerHTML = reportSheet.getPagesCount();
    }

    //------------------------------------------
    // フォントファイルの読み込み
    //------------------------------------------      
    function registerFont(name, type, fontPath) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', fontPath, true);
        xhr.responseType = 'arraybuffer';
        xhr.onload = function (e) {
            if (this.status == 200) {
                var fontArrayBuffer = this.response;
                var fonts = {};
                fonts[type] = fontArrayBuffer;
                GC.Spread.Sheets.PDF.PDFFontsManager.registerFont(name, fonts);
            }
        };
        xhr.send();
    }
});

