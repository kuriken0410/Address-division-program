# address division

A列に各規則性の無い住所データが入ったExcelファイルを読み込み、アパート・マンション名から住所を2つに分けて、再度Excelファイルとして出力するプログラム。<br>

- 友人から上記の依頼を受けたため作成
- 友人のPCにphp環境があったため、プログラム内に必要事項を記入して使用します
- 友人はphpを含めプログラミングについて知識は無く、今後もプログラミングを学習することはありません
- ターミナルは別途取扱説明書に本プログラムの使用フローを載せるので使用可能
- 毎日100〜1000件の住所が入ったExcelファイルを数ファイル処理予定
- 上記5点の状況下の方でも使用できるプログラムとして作成しました

改善点があればリクエスト欲しいです。

<br>

## 分割例（プログラムで以下形式に変更する）
※規則性の無い住所データの『規則性の無さ』に範囲はありませんが、ほとんどが友人より取得した以下使用方法欄の『住所分割前テストデータ.xlsx』の形式と想定<br>
※アパート・マンション名が無ければ分割は無し。住所+部屋号室だけでも可<br>
<br>
例1：〒323-0822栃木県小山市駅南町6丁目15-19K.VillageR棟102<br>
- 〒323-0822栃木県小山市駅南町6丁目15-19
- K.VillageR棟102
<br>
例2：〒259-1322神奈川県秦野市渋沢1875-16 ・ 〒305-0856茨城県つくは市観音台-1-37-28103号<br>
⇒ 分割無し

<br><br>

## 使用方法 
プログラムリンク：　https://github.com/kuriken0410/address-division/blob/main/division.php

1．プログラム内のTODO EXのログファイルを設定して下さい

2．A列のみに住所データが入ったExcelを用意します
> 例： [住所分割前テストデータ.xlsx](https://github.com/kuriken0410/address-division/files/10170226/default.xlsx)

3．プログラム内のTODO①〜④に必要事項を入力してターミナルで実行して下さい

4．住所分割されたExcelファイルが指定したディレクトリ配下に出力されて処理終了
> 例： [20221207住所分割後テストデータ.xlsx](https://github.com/kuriken0410/address-division/files/10170273/20221207.xlsx)

<br>

## 必要環境
- php環境のあるPC
- ComposerでライブラリのPhpSpreadsheetがインストール済で使用可能

<br>

## 注意事項
建物名が『55ビレッジ』や『５５小野ハウス』等のものは、『ビレッジ』や『小野ハウス』等の全角・半角数字の後から分割されてしまう。<br>
理由として、数字が続くと住所の号と建物名の境界線が区別不可のためです。（プログラムの問題ではないと思いますが、念の為記載）<br>

> 例： 東京都新宿区新宿1-1-155ビレッジ101号室<br>
> &emsp;&emsp;⇒ 東京都新宿区新宿1-1-155 と ビレッジ101号室 に分割される。<br>
> &emsp;&emsp;⇒ 『1-1-1』、『1-1-15』、『1-1-155』のどれなのか区別不可のため。<br>
