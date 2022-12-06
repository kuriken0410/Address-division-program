<?php
date_default_timezone_set('Asia/Tokyo');

/**
 * 2022年11月26日、Excel住所データ分割プログラム
 *
 *
 *【プログラムの仕様】
 * A列に各規則性の無い住所データが入ったExcelファイルを読み込み、アパート・マンション名から住所を2つに分けて
 * 再度Excelファイルとして出力するプログラム。
 *
 *【事前準備】
 * 以下TODO EXのログファイルを設定して下さい。（こちらは一度設定すれば、今後変更する必要はありません。）
 *
 *【使用方法】
 * ① A列のみに住所データが入ったExcelを用意します。（※必ずA列のみに住所データを入れて下さい。B、C、D…etc列に入れても住所分割されません。）
 * ② 以下TODO①〜④に必要事項を入力してターミナルで実行して下さい。
 *
 *【注意事項】
 * 建物名が『55ビレッジ』や『５５小野ハウス』等のものは、『ビレッジ』や『小野ハウス』等の全角・半角数字の後から分割されてしまう。理由として、
 * 数字が続くと住所の号と建物名の境界線が区別不可のためです。
 *  例：東京都新宿区新宿1-1-155ビレッジ101号室
 *    ⇒ 東京都新宿区新宿1-1-155 と ビレッジ101号室 に分割される。
 *    ⇒ 『1-1-1』、『1-1-15』、『1-1-155』のどれなのか区別不可のため。
 */


/**
 * TODO EX: 以下『LOG_FILE』の『''』の中に、エラーが起こった際のエラー内容を記載するログファイルの作成先パスとファイル名を入力かコピペする
 * ※こちらは一度設定すれば、今後変更する必要はありません
 */
const LOG_FILE = '/Users/kuriken/phpspreadsheet/log.txt';



/** TODO①: 以下『READ_FILE』の『''』の中に、住所分割したいExcelファイルのパスを入力かコピペする */
const READ_FILE = '/Users/kuriken/phpspreadsheet/住所分割前テストデータ.xlsx';

/**
 * TODO②: Excelに住所分割をしない行があれば、以下『DELETE_EXCEL_NUMBER』の『[]』の中に対象行を入力すると削除される
 * 例1：A1セルを削除したい場合             ⇒ const DELETE_EXCEL_NUMBER = [1];
 * 例2：A1、A5、A6、A10セルを削除したい場合 ⇒ const DELETE_EXCEL_NUMBER = [1,5,6,10]; （複数の場合、順番に半角カンマで区切って入力する）
 * 例3：何も削除しない場合                 ⇒ const DELETE_EXCEL_NUMBER = [];
 */
const DELETE_EXCEL_NUMBER = [1,5,6,10];

/** TODO③: 以下『CREATE_FILE』の『'〜.xlsx'』の中に、住所分割完了したデータを入れたExcelファイルの作成先パスとファイル名を入力する */
define('CREATE_FILE', '/Users/kuriken/phpspreadsheet/' . date('Ymd') . '住所分割後テストデータ.xlsx');

/**
 * TODO④: 上記①〜③を終えたらターミナルで実行する
 * 例：このプログラム（division.php）が『/Users/kuriken/phpspreadsheet/division.php』に入っていたら
 *   ⇒ ① ターミナルを開き、cdコマンドで『/Users/kuriken/phpspreadsheet/』まで移動
 *   ⇒ ② 移動したら『php division.php』と入力し実行
 */




// ***************************************** ここから下は触らないで下さい *************************************************

const SUCCESS_MESSAGE = 'プログラムは正常に完了しました';
const ERROR_MESSAGE   = '不明なエラーが発生しました';
define('NOW', date('Y-m-d H:i:s'));

// エラーハンドリングここから
set_error_handler(
/**
 * @throws ErrorException
 */
static function($error_no, $error_msg, $error_file, $error_line, $error_vars) {
    if(error_reporting() === 0) {
        return;
    }
    throw new ErrorException($error_msg, 0, $error_no, $error_file, $error_line);
});

set_exception_handler(static function($throwable) {
    echo $throwable;
    send_error_log($throwable);
});

register_shutdown_function(static function() {
    $error = error_get_last();
    if($error === null) {
        return;
    }
    // fatal error の場合は既に何らかの出力がされているはずなので何もしない
    send_error_log(new ErrorException($error['message'], 0, 0, $error['file'], $error['line']));
});

/**
 * ログファイル作成と書込
 */
function send_error_log($throwable) {
    echo(PHP_EOL . ERROR_MESSAGE);

    if($throwable) {
        file_put_contents(LOG_FILE, NOW . PHP_EOL . $throwable->__toString() . PHP_EOL . PHP_EOL, FILE_APPEND | LOCK_EX);
    } else {
        file_put_contents(
            LOG_FILE,
            NOW . PHP_EOL .
                 ERROR_MESSAGE . PHP_EOL .
                 '住所分割前Excelファイル：' . READ_FILE . PHP_EOL .
                 '住所分割に失敗し未作成のExcelファイル：' . CREATE_FILE . PHP_EOL . PHP_EOL,
            FILE_APPEND | LOCK_EX
        );
    }
}
// エラーハンドリングここまで


include('./vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$reader      = new PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$spreadSheet = $reader->load(READ_FILE);
$sheet       = $spreadSheet->getActiveSheet();

$address = [];
$row = 1;
foreach($sheet->getRowIterator() as $k => $row) {
    $address[$k] = $sheet->getCell('A' . $row->getRowIndex())->getValue();
}

$new_address = [];
if(!empty($address)) {
    // 住所分割しないデータがあれば分割前に削除
    if(!empty(DELETE_EXCEL_NUMBER)) {
        foreach(DELETE_EXCEL_NUMBER as $k) {
            unset($address[$k]);
        }
    }

    // 地名の後、数字とハイフン以外の文字の前で分割
    foreach($address as $k => $v) {
        $new_address[$k] = preg_split('/^([^\d０-９ー−-]*[\d０-９ー−-]*[^\d０-９ー−-]*[\d０-９ー−-]*)/u', $v, 2, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
    }
}

if(!empty($new_address)) {
    foreach($new_address as $k => $v) {
        if(!empty($v[1])) {
            // 住所 +「号 or 号室」のみなら、一つの住所に統合
            if($v[1] === '号' || $v[1] === '号室') {
                $new_address[$k][0] = $v[0] . $v[1];
                unset($new_address[$k][1]);
            }

            // 「丁目」のみ抽出し住所側住所へ追加、建物側住所からは削除
            if(preg_match('/丁目/u', $v[1], $matches)) {
                $new_address[$k][0] = $v[0] . $matches[0];
                $new_address[$k][1] = str_replace($matches[0], '', $v[1]);
            }
        }
    }

    // 建物側の住所から地番や号に当たる箇所を抽出し住所側住所へ追加、建物側住所からは削除
    foreach($new_address as $k => $v) {
        if(!empty($v[1]) && (preg_match('/(^[\d０-９]{0,5}番[\d０-９]{1,5}号)/u', $v[1], $matches) ||
                             preg_match('/(^[\d０-９]{1,5}番[\d０-９]{1,5})/u', $v[1], $matches) ||
                             preg_match('/(^[\d０-９]{1,5}番)/u', $v[1], $matches) ||
                             preg_match('/(^[\d０-９]{1,5}[―-][\d０-９]{1,5}号)/u', $v[1], $matches) ||
                             preg_match('/(^[\d０-９]{1,5}[―-][\d０-９]{1,5})/u', $v[1], $matches) ||
                             preg_match('/(^[\d０-９]{1,5})/u', $v[1], $matches)))
        {
            $new_address[$k][0] = $v[0] . $matches[0];
            $new_address[$k][1] = str_replace($matches[0], '', $v[1]);
        }
    }


    $newSpreadSheet = new Spreadsheet();
    $newSheet       = $newSpreadSheet->getActiveSheet();

    $row = 1;
    foreach($new_address as $k => $v) {
        $newSheet->setCellValue('A' . $row, $v[0]);

        if(!empty($v[1])) {
            $newSheet->setCellValue('B' . $row, $v[1]);
        }

        $row++;
    }

    $writer = new Xlsx($newSpreadSheet);
    $writer->save(CREATE_FILE);

    if(file_exists(CREATE_FILE)) {
        file_put_contents(
            LOG_FILE,
            NOW . PHP_EOL .
                 SUCCESS_MESSAGE . PHP_EOL .
                 '住所分割前Excelファイル：' . READ_FILE . PHP_EOL .
                 '住所分割後Excelファイル：' . CREATE_FILE . PHP_EOL .
                 'ファイル作成日時：' . date('Y-m-d H:i:s', filectime(CREATE_FILE)) . PHP_EOL
                 //. '分割データ一覧：' . PHP_EOL . print_r($new_address, true)
                 . PHP_EOL,
            FILE_APPEND | LOCK_EX
        );

        exit(SUCCESS_MESSAGE);
    }
}

send_error_log('');
exit;
