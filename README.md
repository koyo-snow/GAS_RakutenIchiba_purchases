# GAS_RakutenIchiba_purchases

楽天市場の購入情報をメールから取得して、表にできるようにしました。
最初に１回だけinitialization()関数を実行することによって、
スプレッドシートの画面の設定が行えます。

allExecution()関数はこのスクリプトのinitialization()関数以外のすべての関数を実行する関数です。
この関数は、情報量によってはGASの最大実行時間である6分を超える恐れがあるので、
その場合はシートごとの関数（sheet0Execution()、sheet1counting0()、sheet1counting1()、sheet2totallingAmount()、sheet3totallingPayment()）を実行してください。

sheet0は「楽天市場購入履歴」シート、
sheet1は「件数」シート、
sheet2は「月別_価格の合計額」シート、
sheet3は「月別_支払金額の合計額」シート
に対応しています。

↓GASにおける関数の実行の仕方を紹介してくれている方がいます。↓
https://hirachin.com/post-3268/

Excelのマクロ用のボタンのように実行したい方は、
各々でボタンを作成し、スクリプトを割り当ててください。
↓ボタンの作り方を紹介してくれている方がいます。↓
https://www.acrovision.jp/service/gas/?p=269

使っていただいたり、コメントやレビューをいただけると嬉しいですが、
このコードで問題が起きた場合でも、責任は負えませんので、
各々の責任において実行してください。

2022/07/07
initialization関数を実装しました。
新しいスプレッドシートで最初に実行すると、
このスクリプトを使うための準備ができます。
