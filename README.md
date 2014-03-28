# CsvToExcel

単一のCSVファイルを単一シートのExcelファイル(.xlsx)に変換する
apache poiベースのJavaプログラムです。

## 入力データ

### パラメータCSVファイル

4行固定のCSVファイルを用意します。
文字コードはUTF-8のみに対応しています。

```csv
出力シート名,参照テンプレートシートインデックス番号,ヘッダー行数,フッター行数
string,string,string
numeric,string,string
formula,string,string
```

1行目には4つの値を設定します。
2行目から4行目にはヘッダー部、ボディ部、フッター部の各列の表示形式を設定します。

* string: 文字列
* numeric: 数値
* formula: 式

### データCSVファイル

文字コードはUTF-8のみに対応しています。

### テンプレートExcel(.xlsx)ファイル

テンプレートシートを用意します。

1行目からヘッダー行数分は、ヘッダー部の書式を設定します。

ヘッダー部の次の行に、ボディ部の書式を設定します。

ボディ部の行の次行からフッター行数分は、フッター部の書式を設定します。

行の高さと列の幅も反映されます。

セルの結合も反映されます。ただし、ボディ部は同一行での結合のみに対応しています。

## 実行方法

```
java -Djava.awt.headless=true -jar CsvToExcel.jar -p parameter.csv -d data.csv -t template.xlsx -o output.xlsx
```

