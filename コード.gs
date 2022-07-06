/**
 * スクリプトで使うシート名を取得する。
 * 
 * コンテナバインドスクリプトなので、SpreadsheetApp.getActiveSpreadsheet()を使っている。
 * スタンドアロンスクリプトにする場合、各々シートのIDを指定し、利用する。
 */
let SPREAD_SHEET = SpreadsheetApp.getActiveSpreadsheet();
let RAKUTEN_ICHIBA_MANAGEMENT_SHEET = SPREAD_SHEET.getSheetByName('楽天市場購入履歴');
let TOTALLING_SHEET = SPREAD_SHEET.getSheetByName('件数');
let TOTAL_AMOUNT_SHEET = SPREAD_SHEET.getSheetByName('月別_価格の合計額');
let TOTAL_PAYMENT_SHEET = SPREAD_SHEET.getSheetByName('月別_支払金額の合計額');

/** セルで指定した日数。メールで遡る日数を指定する。 */
let DATE_BACK_TO = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(3, 24).getValue();  // => 33.0

/** 全部の関数を実行する関数 */
function allExecution() {
  sheet0Execution();
  sheet1counting0();
  sheet1counting1();
  sheet2totallingAmount();
  sheet3totallingPayment();
}

/** 「楽天市場購入履歴」シートで実行する関数 */
function sheet0Execution() {
  addRakutenIchibaDetail();
  Logger.log('addRakutenIchibaDetailが終了しました。');
  canceledRakutenIchibaDetail();
  Logger.log('canceledRakutenIchibaDetailが終了しました。');
  coloring();
  Logger.log('coloringが終了しました。');
}

/** 「件数」シートで実行する関数 */
function sheet1counting0() {
  countingWithCategoryColor();
  Logger.log('countingWithCategoryColorが終了しました。');
}
/** 「件数」シートで実行する関数 */
function sheet1counting1() {
  countingWithMonthAndYear();
  Logger.log('countingWithMonthAndYearが終了しました。');
}

/** 「月別_価格の合計額」シートで実行する関数 */
function sheet2totallingAmount() {
  amountTotaled();
  Logger.log('amountTotaledが終了しました。');
}

/** 「月別_支払金額の合計額」シートで実行する関数 */
function sheet3totallingPayment() {
  totalPayment();
  Logger.log('totalPaymentが終了しました。');
}

function addRakutenIchibaDetail() {
  let SUBJECT = '【楽天市場】注文内容ご確認（自動配信メール）'; // 楽天市場の注文お知らせメールの題名
  let ADDRESS = 'order@rakuten.co.jp';   // 楽天市場の注文お知らせメールのアドレス

  /** 検索期間の初めはセルを参照し、終わりは明日にする。 */
  let beginningDate = new Date();
  beginningDate.setDate(beginningDate.getDate() - DATE_BACK_TO);
  let endDate = new Date();
  endDate.setDate(endDate.getDate() + 1);
  let BEGINNING_DATE = Utilities.formatDate(beginningDate, 'JST', 'yyyy/M/d');
  let END_DATE = Utilities.formatDate(endDate, 'JST', 'yyyy/M/d');

  let QUERY = 'subject:' + SUBJECT + ' from:' + ADDRESS + ' after:' + BEGINNING_DATE + ' before:' + END_DATE;

  /** メールを検索 */
  let threads = GmailApp.search(QUERY);

  /** 各々のスレッドの中身を入れる配列 */
  let eachThreadsPlainBodyArrForCollate = [];
  let eachThreadsPlainBodyArrForWrite = [];
  let eachThreadsDateMailGettedArr = [];

  /** 該当メールがあった場合 */
  if (threads.length > 0) {
    /** 最終行番号取得 */
    let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
    /** 新規で追加する行番号 */
    let TABLE_NEW_ROW = TABLE_LAST_ROW + 1;
    /** 最終行から実際の入力された結果の行を得る */
    let LENGTH_OF_ORDER_NUMBER = TABLE_LAST_ROW - 5;
    /**
     * 取得したsearchColの行が1の場合と、1以上の場合で分ける。
     * 1の場合はテーブルの検索を省く。
     * searchColが1の場合は、テーブルに何の値もないとき
     */
    let searchCol = LENGTH_OF_ORDER_NUMBER + 1;

    for (let th in threads) {
      let msgs = threads[th].getMessages();
      for (let i = 0; i < msgs.length; i++) {
        /** 照合用の配列にplainBodyを格納する */
        /** plainBodyにはメールの内容が入っている。 */
        let plainBodyForCollate = msgs[i].getPlainBody();
        eachThreadsPlainBodyArrForCollate.push(plainBodyForCollate);
        /** 配列にメールを受け取った日付を格納する */
        /** 
         * メールの受信日時を配列に格納する。
         * これはplainBodyにはない情報である。
         */
        let dateMailGetted = msgs[i].getDate();
        eachThreadsDateMailGettedArr.push(dateMailGetted);
      }
      msgs.push(eachThreadsPlainBodyArrForCollate.length);
    }

    for (let i = 0; i < eachThreadsDateMailGettedArr.length; i++) {
      /** 照合用のデータから受注番号を取得(メールからのデータ) */
      let orderNumberForCollateFromMails = eachThreadsPlainBodyArrForCollate[i].match(/\[受注番号\]\s.*/g);
      if (orderNumberForCollateFromMails && orderNumberForCollateFromMails.length) {
        orderNumberForCollateFromMails.forEach((eachOrderNumberOfMail, index) => {
          orderNumberForCollateFromMails[index] = eachOrderNumberOfMail.replace(/\[受注番号\]\s/g, '');
        });
      }
      /** orderNumberForCollateが配列形式になっているので、文字列形式にする。 */
      let orderNumberForCollateFromMailToString = orderNumberForCollateFromMails.toString();

      /** 
       * eachThreadsPlainBodyArrForWriteは、実際にスプレッドシートに書き込む準備をする配列である。メールの全てが入っているわけではない。
       * eachThreadsPlainBodyArrForCollateは、照合用の配列である。メールの全てが入っている。
       */
      if (searchCol === 1) {
        Logger.log('新しいスプレッドシートです。');
        /** 本文を取得 */
        let plainBody = eachThreadsPlainBodyArrForCollate[i];
        eachThreadsPlainBodyArrForWrite.push(plainBody);
      } else if (searchCol > 1) {
        /** searchColが1より大きい場合 */
        /** createTextFinder()関数でセルから値を探して取得 */
        /** createTextFinder()の引数に入れる値はメールの受注番号 */
        let orderNumberFinderFromTable = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.createTextFinder(orderNumberForCollateFromMailToString);
        let resultsOrderNumberCellFromTable = orderNumberFinderFromTable.findAll();
        /**
         * メールの受注番号（単体）からシートの受注番号（全部）を探して、見つけたらtrue、無かったらfalseを返す関数
         * 関数内でfor文を実行し、メールの受注番号がシート内にあった場合はtrueを返し、
         * 無かった場合はcontinueを使ってシートの最後まで検索を続行する。
         * falseは返さない。
        */
        /** メールの受注番号の数を取得する。繰り返すために用意した。 */
        // Logger.log(i);
        /** テーブルの行数を取得する。 */
        /** メールの受注番号は文字列型、テーブルの受注番号は探したばかりで配列型。 */
        let searchedValuesFromTableWithMailFlag = searchValuesFromTableWithMail(orderNumberForCollateFromMailToString, resultsOrderNumberCellFromTable);
        if (searchedValuesFromTableWithMailFlag) {
          continue;
        } else {
          /** 本文を取得 */
          let plainBody = eachThreadsPlainBodyArrForCollate[i];
          eachThreadsPlainBodyArrForWrite.push(plainBody);
        }
      }
    }

    /** 記入した日付を取得 */
    let WRITE_DATE = new Date();

    // Logger.log(eachThreadsPlainBodyArrForWrite.length);

    if (eachThreadsPlainBodyArrForWrite.length !== 0) {
      for (let i = 0; i < eachThreadsPlainBodyArrForWrite.length; i++) {
        /** 受注番号を取得 */
        let order_number = eachThreadsPlainBodyArrForWrite[i].match(/\[受注番号\]\s.*/g);
        if (order_number && order_number.length) {
          order_number.forEach((val, index) => {
            order_number[index] = val.replace(/\[受注番号\]\s/g, '');
          });
        }
        // Logger.log(order_number); 2つ漏れてきてる
        /** 注文日時の配列を取得 OK */
        let bought_days = eachThreadsPlainBodyArrForWrite[i].match(/\[日時\]\s.*/g);
        if (bought_days && bought_days.length) {
          bought_days.forEach((bought_day, index) => {
            let [year, month, day] = bought_day.split('/');
            bought_days[index] = new Date(Number(year), Number(month - 1), Number(day));
            bought_days[index] = bought_day.replace(/\[日時\]\s*/g, '');
          });
        }
        /** メール受信日時 dateMailGetted */
        let dateMailGetted = eachThreadsDateMailGettedArr[i];
        /******* ショップ情報 */
        /** ショップ名 */
        let shop_names = eachThreadsPlainBodyArrForWrite[i].match(/■\sショップ名.*/g);
        if (shop_names && shop_names.length) {
          shop_names.forEach((shop_name, index) => {
            shop_names[index] = shop_name.replace(/.*：\s/g, '');
          })
        }
        /** ショップ電話 */
        let shop_tels = eachThreadsPlainBodyArrForWrite[i].match(/■\s電話\s*：.*/g);
        if (shop_tels && shop_tels.length) {
          shop_tels.forEach((shop_tel, index) => {
            shop_tels[index] = shop_tel.replace(/.*：\s/g, '');
          })
        }
        /** ショップURL */
        let shop_urls = eachThreadsPlainBodyArrForWrite[i].match(/■\sショップURL.*/g);
        if (shop_urls && shop_urls.length) {
          shop_urls.forEach((shop_url, index) => {
            shop_urls[index] = shop_url.replace(/.*：\s/g, '');
          })
        }
        /** ショップお問い合わせフォーム */
        let shop_forms = eachThreadsPlainBodyArrForWrite[i].match(/■\s問い合わせフォーム:.*/g);
        if (shop_forms && shop_forms.length) {
          shop_forms.forEach((shop_form, index) => {
            shop_forms[index] = shop_form.replace(/.*:\s/g, '');
          })
        }
        /******* ショップ情報 終わり */

        /** 
         * 購入した商品（35文字で区切る）
         * メールによって複数の商品が記載されていることもあるので、
         * if文に含めるために順番をズラす。
         * */
        let bought_items = eachThreadsPlainBodyArrForWrite[i].match(/\[商品\][^*]*/g);
        /** 複数の商品を購入した時用のの配列を用意し、pushする。------\r\nの後の35文字をpushして、後で配列を並べて表示する。 */

        /*************** 単価、数量、合計金額を取得 */
        /** 購入した商品の数 */
        let quantity = eachThreadsPlainBodyArrForWrite[i].match(/[0-9]*\(個\)/g);
        if (quantity && quantity.length) {
          quantity.forEach((qty, index) => {
            quantity[index] = qty.replace(/[^0-9].*/g, '');
          });
        }
        /**合計金額 */
        let sumPrice = eachThreadsPlainBodyArrForWrite[i].match(/価格.*/g);

        if (sumPrice && sumPrice.length) {
          sumPrice.forEach((sum, index) => {
            sumPrice[index] = sum.replace(/.*=\s/g, '').replace(/[^0-9]*/, '').replace(/\(.*/g, '').replace(/,/g, '');
            sumPrice[index] = sumPrice[index]
          })
        }

        let sumSumPrice = 0;
        let qtySum = 0;

        if (quantity.length > 0) {
          for (let i = 0; i < quantity.length; i++) {
            qtySum += Number(quantity[i]);
            sumPrice[i] = Number(sumPrice[i]);
            sumSumPrice += sumPrice[i];
          }
        }

        let divideString = '----------\r\n';

        if (quantity.length === 1) {
          if (bought_items && bought_items.length) {
            /** [商品]の文字列を外す */
            bought_items = bought_items[0].replace(/\[商品\]\r\n/g, '').substring(0, 35);
          }
        } else if (quantity.length > 1) {
          /**まずは[商品]の文字列を外す */
          bought_items = bought_items[0].replace(/\[商品\]\r\n/g, '')
          /**文字列を配列にする */
          let bought_items_divided_by_string = bought_items.split(divideString);
          /**配列の要素それぞれをsubsringする */
          if (bought_items_divided_by_string && bought_items_divided_by_string.length) {
            bought_items_divided_by_string.forEach((each, index) => {
              bought_items_divided_by_string[index] = each.substring(0, 35);
            });
          }
          bought_items = bought_items_divided_by_string;
          /**配列を再び文字列にする。その際、改行を挟むと見やすい。 */
          bought_items = bought_items.join(',\r\n');
        }

        /*************** 単価、数量、合計金額終わり */
        /******** ポイント・クーポン（価格からマイナスされる分） */
        /** ポイント利用分を取得  OK */
        let result_points = '';
        let no_points = eachThreadsPlainBodyArrForWrite[i].match(/\[ポイント利用\] なし/g);
        let use_points = eachThreadsPlainBodyArrForWrite[i].match(/\[ポイント利用\] あり/g);
        if (no_points && no_points.length) {
          no_points.forEach((n_p, index) => {
            no_points[index] = n_p.replace(/\[ポイント利用\] /g, '');
            result_points = no_points[index];
          });
        } else if (use_points && use_points.length) {
          let used_points = eachThreadsPlainBodyArrForWrite[i].match(/ポイント利用\s\-.*/g);
          if (used_points && used_points.length) {
            used_points.forEach((p, index) => {
              used_points[index] = p.replace(/ポイント利用\s\-/g, '').replace(/\(円\)/g, '').replace(/,/g, '');
              result_points = used_points[index];
            });

          }
        }
        /** クーポン利用 クーポン利用 -200(円) */
        let result_coupons = eachThreadsPlainBodyArrForWrite[i].match(/クーポン利用\s.*/g);

        if (result_coupons && result_coupons.length) {
          result_coupons.forEach((result_c, index) => {
            result_coupons[index] = result_c.replace(/[^0-9]*/, '').replace(/\(円\)/g, '').replace(/,/g, '');
          });
        } else if (result_coupons === null) {
          result_coupons = 'なし';
        }
        /******** ポイント・クーポン（価格からマイナスされる分） 終わり */
        /** 送料を取得 OK */
        let postages = eachThreadsPlainBodyArrForWrite[i].match(/送料計.*/g);
        if (postages && postages.length) {
          postages.forEach((postage, index) => {
            postages[index] = postage.replace(/[^0-9]/g, '').replace(/,/g, '');
          })
        }
        /** 支払い金額を取得 OK */
        let all_prices = eachThreadsPlainBodyArrForWrite[i].match(/支払い金額.*/g);
        if (all_prices && all_prices.length) {
          all_prices.forEach((all_price, index) => {
            all_prices[index] = all_price.replace(/[^0-9]/g, '').replace(/,/g, '');
          })
        }
        /** 配送方法を取得 OK */
        let delivery = eachThreadsPlainBodyArrForWrite[i].match(/\[配送方法\]\s.*/g);
        if (delivery && delivery.length) {
          delivery.forEach((del, index) => {
            delivery[index] = del.replace(/\[配送方法\] /g, '');
          });
        }
        /** 獲得ポイントを取得 OK */
        let get_points = eachThreadsPlainBodyArrForWrite[i].match(/今回のお買い物で[0-9]*ポイント獲得予定/g);
        if (get_points && get_points.length) {
          get_points.forEach((g_p, index) => {
            get_points[index] = g_p.replace(/今回のお買い物で/g, '').replace(/ポイント獲得予定/g, '').replace(/,/g, '');
            get_points[index] = Number(get_points[index])
          });
        }
        /******* 注文者情報 */
        /** お届け先氏名 OK */
        let orderers = eachThreadsPlainBodyArrForWrite[i].match(/\[注文者\].*/g);
        if (orderers && orderers.length) {
          orderers.forEach((orderer, index) => {
            orderers[index] = orderer.replace(/\[注文者\]\s*/g, '').replace(/\s\(.*/g, '');
          })
        }
        /** 注文者お届け先住所 OK */
        let orderer_addresses = eachThreadsPlainBodyArrForWrite[i].match(/〒.*/g);
        /** 注文者お届け先電話番号 OK */
        let orderer_tels = eachThreadsPlainBodyArrForWrite[i].match(/\(TEL\).*/g);
        if (orderer_tels && orderer_tels.length) {
          orderer_tels.forEach((orderer_tel, index) => {
            orderer_tels[index] = orderer_tel.replace(/\(TEL\)/g, '');
          })
        }

        /******* 設定シーケンス */
        /** 記入日時: 記入した日時を設定 */
        let WRITE_DATE_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`A${TABLE_NEW_ROW + i}`);
        WRITE_DATE_cell.setValue(WRITE_DATE ?? new Date());

        /** 受注番号: 受注番号を設定 */
        let order_number_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`B${TABLE_NEW_ROW + i}`);
        order_number_cell.setValue(order_number ?? new Date());

        /** 注文日時: 注文日時を設定 */
        let bought_day_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`C${TABLE_NEW_ROW + i}`);
        bought_day_cell.setValue(bought_days ?? '');

        /** メール受信日時: メール受信日時に設定 */
        let mail_date_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`D${TABLE_NEW_ROW + i}`);
        mail_date_cell.setValue(dateMailGetted ?? new Date);

        /** ショップ名: ショップ名を設定 */
        let shop_names_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`E${TABLE_NEW_ROW + i}`);
        shop_names_cell.setValue(shop_names ?? '');

        /** ショップ電話番号: ショップ電話番号を設定 */
        let shop_tels_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`F${TABLE_NEW_ROW + i}`);
        shop_tels_cell.setValue(shop_tels ?? '');

        /** ショップURL: ショップURLを設定 */
        let shop_urls_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`G${TABLE_NEW_ROW + i}`);
        shop_urls_cell.setValue(shop_urls ?? '');

        /** ショップ問い合わせURL: ショップ問い合わせURLを設定 */
        let shop_forms_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`H${TABLE_NEW_ROW + i}`);
        shop_forms_cell.setValue(shop_forms ?? '');

        /** 購入した商品: 購入した商品を設定 */
        let bought_items_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`I${TABLE_NEW_ROW + i}`);
        bought_items_cell.setValue(bought_items ?? '');

        /** 合計金額: 合計金額を設定 */
        let prices_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`J${TABLE_NEW_ROW + i}`);
        prices_cell.setValue(sumSumPrice ?? 0);

        /** 数量：　購入した商品の合計数を設定 */
        let qty_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`K${TABLE_NEW_ROW + i}`);
        qty_cell.setValue(qtySum ?? 0);

        /** ポイント利用分: ポイント利用分を設定 */
        let result_points_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`L${TABLE_NEW_ROW + i}`);
        result_points_cell.setValue(result_points ?? 0);

        /** クーポン利用分 */
        let result_coupons_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`M${TABLE_NEW_ROW + i}`);
        result_coupons_cell.setValue(result_coupons ?? new Date());

        /** 送料 */
        let postages_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`N${TABLE_NEW_ROW + i}`);
        postages_cell.setValue(postages ?? 0);

        /** 支払い金額 */
        let all_prices_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`O${TABLE_NEW_ROW + i}`);
        all_prices_cell.setValue(all_prices ?? 0);

        /** 配送方法 */
        let delivery_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`P${TABLE_NEW_ROW + i}`);
        delivery_cell.setValue(delivery ?? '');

        /** 獲得予定P */
        let pricecell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`Q${TABLE_NEW_ROW + i}`);
        pricecell.setValue(get_points ?? 0);

        /** お届け先氏名 */
        let orderers_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`R${TABLE_NEW_ROW + i}`);
        orderers_cell.setValue(orderers ?? '');

        /** お届け先住所 */
        let orderer_addresses_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`S${TABLE_NEW_ROW + i}`);
        orderer_addresses_cell.setValue(orderer_addresses ?? '');

        /** お届け先電話番号 */
        let orderer_tels_cell = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`T${TABLE_NEW_ROW + i}`);
        orderer_tels_cell.setValue(orderer_tels ?? '');
      }
    }

    /** テーブルの最後の行番号 */
    let TABLE_LAST_ROW_2 = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
    /** テーブルの左端 */
    let TABLE_LEFT_MOST = 1;
    /** テーブルの右端 */
    let TABLE_RIGHT_MOST = 21;
    /** 枠線を書き込む処理 */

    /** を降順に並べ替え */
    let orderNumberSorting = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(`A6:U${TABLE_LAST_ROW_2}`);
    orderNumberSorting.sort({ column: 3, acsending: true });

    let sourceRange = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(
      `${getColName(TABLE_LEFT_MOST)}6:${getColName(TABLE_RIGHT_MOST)}${TABLE_LAST_ROW_2}`
    );
    sourceRange.setBorder(true, true, true, true, true, true);

  }
}

/** 時間のフォーマット */
function formatDate(date) {
  let yyyy = date.getFullYear();
  let mm = toTwoDigits(date.getMonth() + 1);
  let dd = toTwoDigits(date.getDate());

  return yyyy + '/' + mm + '/' + dd;
}
function formatYearAndMonth(year, month, day) {
  let yyyy = year;
  let mm = toTwoDigits(month);
  let dd = toTwoDigits(day);

  return yyyy + '/' + mm + '/' + dd;
}

/** 日付の0埋め */
function toTwoDigits(num) {
  num += "";
  if (num.length === 1) {
    num = "0" + num;
  }
  return num;
}


function getColName(num) {
  let result = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(1, num);
  result = result.getA1Notation();
  result = result.replace(/\d/, '');

  return result;
}

/** 
 * キャンセルメールを取得して、キャンセル済みを入力し、
 * coloringでは、ここの入力を見て色を付ける。
 * */
function canceledRakutenIchibaDetail() {
  let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
  if (TABLE_LAST_ROW === 5) {
    return;
  }

  let ADDRESS = 'order-cancel@order-rp.rms.rakuten.co.jp ';   // 楽天市場のキャンセル用アドレス

  /** メールボックス内でメールを探すための遡る日数を指定する。 */
  let beginningDate = new Date();
  beginningDate.setDate(beginningDate.getDate() - DATE_BACK_TO);
  let endDate = new Date();
  endDate.setDate(endDate.getDate() + 1);
  let BEGINNING_DATE = Utilities.formatDate(beginningDate, 'JST', 'yyyy/M/d');
  let END_DATE = Utilities.formatDate(endDate, 'JST', 'yyyy/M/d');

  let QUERY = 'from:' + ADDRESS + ' after:' + BEGINNING_DATE + ' before:' + END_DATE;

  /** メールを検索 */
  let threads = GmailApp.search(QUERY);

  /** 該当メールがあった場合 */
  if (threads.length > 0) {
    for (let th in threads) {
      let msgs = threads[th].getMessages();
      for (let i = 0; i < msgs.length; i++) {
        /** （キャンセルメールの）本文を取得 */
        let plainBody = msgs[i].getPlainBody();

        let order_number = plainBody.match(/\[注文番号\]\s.*/g);
        if (order_number && order_number.length) {
          order_number.forEach((val, index) => {
            order_number[index] = val.replace(/\[注文番号\]\s/g, '');
          });
        }

        collateOrderNumber(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, order_number);
      }
    }
  }
}

/** 
 * 受注番号から表内の同じ番号を探して、
 * あったらtrueを返す。
 */
function collateOrderNumber(sheet, orderNumberInMail) {
  /** createTextFinderで該当のセルを取得し、そのセルの行を取得する。 */
  let textFinder = sheet.createTextFinder(orderNumberInMail);
  let cellsOfTextFinder = textFinder.findAll();
  for (let i = 0; i < cellsOfTextFinder.length; i++) {
    let rowOfCell = cellsOfTextFinder[i].getRow();
    markCancelRow(sheet, rowOfCell);
    // Logger.log(rowOfCell);
  }
}

/**
 * 引数に行番号を渡して、セルに「キャンセル済み」と書き込む
 */
function markCancelRow(sheet, rowNumber) {
  sheet.getRange(rowNumber, 21).setValue('キャンセル済み');
}

/**
 * メールの受注番号（単体）からシートの受注番号（全部）を探して、見つけたらtrue、無かったら何も返さない関数
 * 関数内でfor文を実行し、メールの受注番号がシート内にあった場合はtrueを返し、
 * 無かった場合はcontinueを使ってシートの最後まで検索を続行する。
 * falseは返さない。
*/
function searchValuesFromTableWithMail(orderNumberOfMail, orderNumberOfTableArr) {
  for (let i = 0; i < orderNumberOfTableArr.length; i++) {
    if (orderNumberOfMail === orderNumberOfTableArr[i].getValue()) {
      return true;
    }
  }
}















/** 
 * 「件数」シートの表の背景色を使って、
 * 金額に応じて「楽天市場購入履歴」シートに色を付ける。
 */
function coloring() {
  let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
  Logger.log(TABLE_LAST_ROW);
  if (TABLE_LAST_ROW === 5) {
    return;
  }

  let STANDARD_NUM = 0;

  /** 0〜1000を取得する。B4の数値を変えるとスクリプトで使う値も変わる。 */
  let UNDER_NUM_0 = TOTALLING_SHEET.getRange(4, 2).getValue();
  let highNum0 = UNDER_NUM_0.replace(/[^0-9]/g, '');
  let num0ColorCode = TOTALLING_SHEET.getRange(4, 2).getBackground();
  let lowNum0 = STANDARD_NUM;

  /** 1001~5000を取得する。 */
  let UNDER_NUM_1 = TOTALLING_SHEET.getRange(5, 2).getValue();
  let highNum1 = UNDER_NUM_1.replace(/[^0-9]/g, '');
  let num1ColorCode = TOTALLING_SHEET.getRange(5, 2).getBackground();
  let lowNum1 = highNum0 + 1;

  /** 5001~10000を取得する。 */
  let UNDER_NUM_2 = TOTALLING_SHEET.getRange(6, 2).getValue();
  let highNum2 = UNDER_NUM_2.replace(/[^0-9]/g, '');
  let num2ColorCode = TOTALLING_SHEET.getRange(6, 2).getBackground();
  let lowNum2 = highNum1 + 1;

  /** 10001~50000を取得する。 */
  let UNDER_NUM_3 = TOTALLING_SHEET.getRange(7, 2).getValue();
  let highNum3 = UNDER_NUM_3.replace(/[^0-9]/g, '');
  let num3ColorCode = TOTALLING_SHEET.getRange(7, 2).getBackground();
  let lowNum3 = highNum2 + 1;

  /** 50001~を取得する。 */
  // let UNDER_NUM_4 = TOTALLING_SHEET.getRange(8, 2).getValue();
  // let highNum4 = UNDER_NUM_4.replace(/[^0-9]/g, '');
  let num4ColorCode = TOTALLING_SHEET.getRange(8, 2).getBackground();
  let lowNum4 = highNum3 + 1;

  let CANCELED = TOTALLING_SHEET.getRange(9, 2).getValue();
  let canceledColorCode = TOTALLING_SHEET.getRange(9, 2).getBackground();


  let willColorRowNumber0 = [];
  let willColorRowNumber1 = [];
  let willColorRowNumber2 = [];
  let willColorRowNumber3 = [];
  let willColorRowNumber4 = [];
  let willColorRowNumberCanceled = [];

  if (TABLE_LAST_ROW === 5) {
    return;
  } else if (TABLE_LAST_ROW > 5) {
    comparePricesToColor(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, highNum0, lowNum0, willColorRowNumber0, num0ColorCode);
    comparePricesToColor(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, highNum1, lowNum1, willColorRowNumber1, num1ColorCode);
    comparePricesToColor(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, highNum2, lowNum2, willColorRowNumber2, num2ColorCode);
    comparePricesToColor(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, highNum3, lowNum3, willColorRowNumber3, num3ColorCode);
    comparePricesToColor2(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, lowNum4, willColorRowNumber4, num4ColorCode);
    searchCancelsToColor(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, CANCELED, willColorRowNumberCanceled, canceledColorCode);
  }
}

/** 
 * この関数では、テーブル内の価格を調べてそれ以下の値があれば色を付けるようにする。
 * 以下以上はtextFinderでは難しそう。
 */
function comparePricesToColor(sheet, highStandardPrice, lowStandardPrice, willColorRowNumber, color) {
  let comparePricesRange = sheet.getRange(6, 10, sheet.getLastRow() - 5);

  for (let i = 0; i < comparePricesRange.getNumRows(); i++) {
    let priceCellOfTable = comparePricesRange.getCell(1 + i, 1);
    let priceValueOfTable = comparePricesRange.getCell(1 + i, 1).getValue();
    if (highStandardPrice > priceValueOfTable && lowStandardPrice <= priceValueOfTable) {
      willColorRowNumber.push(priceCellOfTable.getRow());
    }
  }
  for (let i = 0; i < willColorRowNumber.length; i++) {
    sheet.getRange(willColorRowNumber[i], 1, 1, 21).setBackground(color);
  }
}

/** 
 * 50001~は以上だけなので、別の関数を用意する。
 * comparePricesToColorをif文で分けて処理するよりも別の関数を作った方がわかりやすいはず。
 */
function comparePricesToColor2(sheet, lowStandardPrice, willColorRowNumber, color) {
  let comparePricesRange = sheet.getRange(6, 10, sheet.getLastRow() - 5);

  for (let i = 0; i < comparePricesRange.getNumRows(); i++) {
    let priceCellOfTable = comparePricesRange.getCell(1 + i, 1);
    let priceValueOfTable = comparePricesRange.getCell(1 + i, 1).getValue();
    if (lowStandardPrice < priceValueOfTable) {
      willColorRowNumber.push(priceCellOfTable.getRow());
    }
  }
  for (let i = 0; i < willColorRowNumber.length; i++) {
    sheet.getRange(willColorRowNumber[i], 1, 1, 21).setBackground(color);
  }
}

/**
 * この関数では、textFinderで「キャンセル済み」の行を調べて、
 * その行を紫にする。
 */
function searchCancelsToColor(sheet, word, willColorRowNumber, color) {
  let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
  let canceledRange = sheet.getRange(6, 21, TABLE_LAST_ROW - 5);

  for (let i = 0; i < canceledRange.getNumRows(); i++) {
    let canceledCellOfTable = canceledRange.getCell(1 + i, 1);
    let canceledValueOfTable = canceledRange.getCell(1 + i, 1).getValue();
    // Logger.log(canceledValueOfTable);
    if (word === canceledValueOfTable) {
      // Logger.log('true');
      willColorRowNumber.push(canceledCellOfTable.getRow());
    }
  }
  for (let i = 0; i < willColorRowNumber.length; i++) {
    sheet.getRange(willColorRowNumber[i], 1, 1, 21).setBackground(color);
  }
}










function countingWithCategoryColor() {
  let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
  /** テーブルにメールからの情報がない場合は関数を終了する。 */
  if (TABLE_LAST_ROW === 5) {
    return;
  }

  let willTotalRowNum0 = [];
  let willTotalRowNum1 = [];
  let willTotalRowNum2 = [];
  let willTotalRowNum3 = [];
  let willTotalRowNum4 = [];
  let willTotalRowNumCanceled = [];

  let num0ColorCode = TOTALLING_SHEET.getRange(4, 2).getBackground();
  let num1ColorCode = TOTALLING_SHEET.getRange(5, 2).getBackground();
  let num2ColorCode = TOTALLING_SHEET.getRange(6, 2).getBackground();
  let num3ColorCode = TOTALLING_SHEET.getRange(7, 2).getBackground();
  let num4ColorCode = TOTALLING_SHEET.getRange(8, 2).getBackground();
  let num5ColorCode = TOTALLING_SHEET.getRange(9, 2).getBackground();

  if (TABLE_LAST_ROW === 5) {
    return;
  } else if (TABLE_LAST_ROW > 5) {
    /**
     * 価格カテゴリごとの件数を調べる。
     */
    countCategories(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, TOTALLING_SHEET, TABLE_LAST_ROW, 4, num0ColorCode, willTotalRowNum0);
    countCategories(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, TOTALLING_SHEET, TABLE_LAST_ROW, 5, num1ColorCode, willTotalRowNum1);
    countCategories(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, TOTALLING_SHEET, TABLE_LAST_ROW, 6, num2ColorCode, willTotalRowNum2);
    countCategories(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, TOTALLING_SHEET, TABLE_LAST_ROW, 7, num3ColorCode, willTotalRowNum3);
    countCategories(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, TOTALLING_SHEET, TABLE_LAST_ROW, 8, num4ColorCode, willTotalRowNum4);
    countCategories(RAKUTEN_ICHIBA_MANAGEMENT_SHEET, TOTALLING_SHEET, TABLE_LAST_ROW, 9, num5ColorCode, willTotalRowNumCanceled);

  }
}

/** 
 * sheet0の背景色でカテゴリ分けをする。
 * sheet1の上の表に埋めるための関数。
 */
function countCategories(sheet0, sheet1, tableLastRow, indicatedCellInTotallingSheet, color, willTotalRowNum) {
  let totalRange = sheet0.getRange(6, 10, tableLastRow - 5);

  if (totalRange.getNumRows() === 0) {
    return;
  } else {
    for (let i = 0; i < totalRange.getNumRows(); i++) {
      let sheet0Range = sheet0.getRange(6 + i, 10);
      if (sheet0Range.getBackground() === color) {
        willTotalRowNum.push(sheet0Range.getValue());
      }
    }
    let totalLength = willTotalRowNum.length;
    /** 計算した合計値を「集計」タブのセルに入れる */
    sheet1.getRange(indicatedCellInTotallingSheet, 3).setValue(totalLength);
  }
}






/**
 * 先にシート１の年と月から基準となる年と月を取得する。
 * 取得した年と月を使って、シート０の注文日時を参照し、セル情報を配列に入れる。
 * 配列のlengthを取って、数を数える。
 * キャンセル分は最後にマイナスする。キャンセル分の計算は背景色で行う。
 * 「合計件数 - キャンセル件数」を年と月の情報を使ってシート１に記入する
 */
function countingWithMonthAndYear() {
  let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
  /** テーブルにメールからの情報がない場合は関数を終了する。 */
  if (TABLE_LAST_ROW === 5) {
    return;
  }
  /**
   * シート１から必要な年と月の配列を取得する。
   */
  let sheet1Year = TOTALLING_SHEET.getRange(12, 3, 1, TOTALLING_SHEET.getLastColumn() - 2);
  let sheet1YearValues = sheet1Year.getValues();
  let yearArr = [];
  for (let i = 0; i < sheet1YearValues.length; i++) {
    yearArr.push(sheet1YearValues[i]);
  }
  yearArr = yearArr[0];
  let sheet1Month = TOTALLING_SHEET.getRange(13, 2, 12);
  let sheet1MonthValues = sheet1Month.getValues();
  let monthArr = [];
  for (let i = 0; i < sheet1MonthValues.length; i++) {
    monthArr.push(sheet1MonthValues[i].toString());
  }

  /**シート０のキャンセル分の数を数える */
  /**まずは項目全体の数を数える */
  let sheet0ItemCount = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6, 1, RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow() - 5);
  /**カウントを保持する変数を用意する */
  let toCancelCount = [];
  let toCancelCountFromRange = [];

  /**その中から背景色が紫の部分だけを数える */
  for (let i = 0; i < yearArr.length; i++) {
    for (let j = 0; j < monthArr.length; j++) {
      /** 年と月の情報を使って検索する。 */
      let textFinder = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.createTextFinder(yearArr[i] + '-' + toTwoDigits(monthArr[j]) + ".*").useRegularExpression(true);
      let ranges = textFinder.findAll();
      /**紫の中から年と月が一致したものだけ数える */
      /** 数えて下の表に記入した後、初期化しないといけないので、0を代入している。 */
      toCancelCount.length = 0;
      toCancelCountFromRange.length = 0;
      for (let h = 0; h < sheet0ItemCount.getNumRows(); h++) {
        if (sheet0ItemCount.getCell(1 + h, 1).getBackground() === TOTALLING_SHEET.getRange(9, 2).getBackground()) {
          toCancelCount.push(RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6 + h, 3));
        }
      }

      /** 年と月が一致した場合、入力用の配列に行数をpushする。 */
      if (ranges.length > 0) {
        toCancelCount.forEach((rng) => {
          let value = rng.getValue();
          if (value.getFullYear() === yearArr[i] && (1 + Math.floor(value.getMonth())) === Math.floor(monthArr[j])) {
            toCancelCountFromRange.push(rng.getRow());
          }
        });
      }

      /** 
       * 下の表に件数を入力する。
       * 件数が0の場合グレーにする。
       */
      if (ranges.length) {
        TOTALLING_SHEET.getRange(13 + j, 3 + i).setValue(ranges.length - toCancelCountFromRange.length);
        TOTALLING_SHEET.getRange(13 + j, 3 + i).setBackground("#ffffff");
      } else if (ranges.length === 0) {
        TOTALLING_SHEET.getRange(13 + j, 3 + i).setValue(ranges.length);
        TOTALLING_SHEET.getRange(13 + j, 3 + i).setBackground("#dddddd");
      }
    }
  }
}











/**
 * 先にシート２の年と月から基準となる年と月を取得する。
 * 取得した年と月を使って、シート０の注文日時を参照し、セル情報を配列に入れる。
 * セル情報の行情報を使って、価格を値を配列に入れる。
 * reduceを使って、価格の値段を合計する。
 * キャンセル分は最後にマイナスする。キャンセル分の計算は背景色で行う。
 * 「合計値 - キャンセル分」を年と月の情報を使ってシート１に記入する
 */
function amountTotaled() {
  let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
  Logger.log(TABLE_LAST_ROW);
  /** テーブルにメールからの情報がない場合は関数を終了する。 */
  if (TABLE_LAST_ROW === 5) {
    return;
  }

  /** 最後の方で使う。for文のi（一般に）の初期化の際に、値を保持しておくために、変数にしておく。 */
  let countToamountDataArr = [];

  /**
   * シート２から必要な年と月の配列を取得する。
   */
  let sheet2Year = TOTAL_AMOUNT_SHEET.getRange(3, 3, 1, TOTAL_AMOUNT_SHEET.getLastColumn() - 2);
  // Logger.log(sheet2Year.getValues()); // CLEAR!
  let sheet2YearValues = sheet2Year.getValues();
  let yearArr = [];
  for (let i = 0; i < sheet2YearValues.length; i++) {
    yearArr.push(sheet2YearValues[i]);
  }
  yearArr = yearArr[0];
  // Logger.log(yearArr); // 数値だけの配列になっている。
  let sheet2Month = TOTAL_AMOUNT_SHEET.getRange(4, 2, 12);
  let sheet2MonthValues = sheet2Month.getValues();
  let monthArr = [];
  for (let i = 0; i < sheet2MonthValues.length; i++) {
    monthArr.push(sheet2MonthValues[i].toString());
  }
  // Logger.log(monthArr); // 数値だけの配列になっている。

  /********シート０のキャンセル分の数を数える */
  /**まずは項目全体の数を数える */
  let sheet0ItemCount = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6, 1, RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow() - 5);
  /**カウントを保持する変数を用意する */
  let toCancelDateArr = [];
  let toCancelDateArrFromRange = [];
  /**その中から背景色が紫の部分だけを数える */

  /** 
   * テーブルのRangeオブジェクトを作成した。
   * C列（注文日時）とJ列（価格）のRangeオブジェクトを作成した。
   * C列とJ列は行の数が同じなので、同じ変数を取れ、取り出した行データにも一貫性がある。
   */
  let rangelistOfTable = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRangeList([`C6:C${TABLE_LAST_ROW}`, `J6:J${TABLE_LAST_ROW}`]);
  let rangeOfTable = rangelistOfTable.getRanges();

  for (let i = 0; i < yearArr.length; i++) {
    for (let j = 0; j < monthArr.length; j++) {
      /** シート２の年情報と月情報でシート０のセルを検索する。 */
      let textFinder = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.createTextFinder(yearArr[i] + '-' + toTwoDigits(monthArr[j]) + ".*").useRegularExpression(true);;
      let ranges = textFinder.findAll();
      // Logger.log(ranges);
      /********紫(キャンセル分)の中から年と月が一致したものだけ数える */
      /** キャンセル日時のデータを格納する配列。初期化することで、年と月情報が変わるたびに数え直している。 */
      toCancelDateArr = [];
      toCancelDateArrFromRange = [];
      /** キャンセル分の数値を格納する配列 */
      toCancelAmountArr = [];
      for (let h = 0; h < sheet0ItemCount.getNumRows(); h++) {
        if (sheet0ItemCount.getCell(1 + h, 1).getBackground() === TOTALLING_SHEET.getRange(9, 2).getBackground()) {
          toCancelDateArr.push(RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6 + h, 3));
          toCancelAmountArr.push(RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6 + h, 10));
        }
      }
      // Logger.log(toCancelCount); /**キャンセルのRangeが入っている */
      // Logger.log(toCancelAmountArr); /**キャンセル分の価格の数値情報のRangeが入っている */

      /**toCancelDateArrの処理 */
      /** 年と月の情報で検索した情報が入っていた場合 */
      if (ranges.length > 0) {
        toCancelDateArr.forEach((rng) => {
          let value = rng.getValue();
          if (value.getFullYear() === yearArr[i] && (1 + Math.floor(value.getMonth())) === Math.floor(monthArr[j])) {
            toCancelDateArrFromRange.push(rng.getRow());
          }
        });
      }

      /**特定の日付のキャンセルした価格の合計値を調べる。 */
      let sumOfCancel = 0;
      if (ranges.length > 0) {
        /**キャンセルするデータの配列に中身が入っているときに実行 */
        if (toCancelDateArrFromRange.length > 0) {
          /**上でキャンセルデータの行番号を取り出したので、それを使って価格だけ取り出す */
          for (let h = 0; h < toCancelDateArrFromRange.length; h++) {
            let rangeOfAmount = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(toCancelDateArrFromRange[h], 10);
            let rangeValueOfAmount = rangeOfAmount.getValue();
            sumOfCancel += rangeValueOfAmount;
          }
        }
      }

      /*******
       * まずは表の価格の値全部のRangeオブジェクトを取る。
       * それぞれの年と月を使って検索し、一致したRangeオブジェクトを配列に格納する。
       * 初期化してあった配列には、年と月の一致したRangeオブジェクトが入っている。
       * .getValue()とfor文を使って、合計値を計算する。
       * それぞれの年と月ごとの合計値を変数に格納する。
      */
      // /**このログはテーブルの日付一覧を取り出している。 */
      // Logger.log(rangeOfTable[0].getValues());
      // /**このログはテーブルの価格一覧を取り出している。 */
      let amountData = rangeOfTable[1].getValues()
      // Logger.log(amountData);

      /**
       * 年データと月データで検索して一致したテーブルのセルの行を取得し、
       * その行を使って、価格のRangeオブジェクトのgetValue()にアクセスする。
       */

      let sumOfAllWithYearAndMonth = 0;

      let incrementData = countToamountDataArr.length;

      if (ranges.length > 0) {
        // 繰り返すのは検索にヒットした数。
        for (let h = incrementData; h < ranges.length + incrementData; h++) {
          sumOfAllWithYearAndMonth += parseInt(amountData[h]);
          countToamountDataArr.push(h);
        }
      }

      /** 年、月ごとに値がある場合は白、値がない場合はグレーにしている。 */
      if (ranges.length > 0) {
        TOTAL_AMOUNT_SHEET.getRange(4 + j, 3 + i).setValue(sumOfAllWithYearAndMonth - sumOfCancel);
        TOTAL_AMOUNT_SHEET.getRange(4 + j, 3 + i).setBackground("#ffffff");
      } else if (ranges.length === 0) {
        TOTAL_AMOUNT_SHEET.getRange(4 + j, 3 + i).setValue(ranges.length);
        TOTAL_AMOUNT_SHEET.getRange(4 + j, 3 + i).setBackground("#dddddd");
      }
    }
  }
}














/** 
 * シート３の処理。
 * シート２の処理と似ている。
 */
function totalPayment() {
  let TABLE_LAST_ROW = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow();
  Logger.log(TABLE_LAST_ROW);
  /** テーブルにメールからの情報がない場合は関数を終了する。 */
  if (TABLE_LAST_ROW === 5) {
    return;
  }
  /** 最後の方で使う。for文のi（一般に）の初期化の際に、値を保持しておくために、宣言しておく。 */
  let countToamountDataArr = [];

  /**
   * シート２から必要な年と月の配列を取得する。
   */
  let sheet3Year = TOTAL_PAYMENT_SHEET.getRange(3, 3, 1, TOTAL_PAYMENT_SHEET.getLastColumn() - 2);
  // Logger.log(sheet3Year.getValues()); // CLEAR!
  let sheet3YearValues = sheet3Year.getValues();
  let yearArr = [];
  for (let i = 0; i < sheet3YearValues.length; i++) {
    yearArr.push(sheet3YearValues[i]);
  }
  yearArr = yearArr[0];
  // Logger.log(yearArr); // 数値だけの配列になっている。
  let sheet3Month = TOTAL_PAYMENT_SHEET.getRange(4, 2, 12);
  let sheet3MonthValues = sheet3Month.getValues();
  let monthArr = [];
  for (let i = 0; i < sheet3MonthValues.length; i++) {
    monthArr.push(sheet3MonthValues[i].toString());
  }
  // Logger.log(monthArr); // 数値だけの配列になっている。

  /********シート０のキャンセル分の数を数える */
  /**まずは項目全体の数を数える */
  let sheet0ItemCount = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6, 1, RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getLastRow() - 5);
  /**カウントを保持する変数を用意する */
  let toCancelDateArr = [];
  let toCancelDateArrFromRange = [];
  /**その中から背景色が紫の部分だけを数える */



  /** 
   * テーブルのRangeオブジェクトを作成した。
   * C列（注文日時）とJ列（価格）のRangeオブジェクトを作成した。
   * C列とJ列は行の数が同じなので、同じ変数を取れ、取り出した行データにも一貫性がある。
   */
  let rangelistOfTable = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRangeList([`C6:C${TABLE_LAST_ROW}`, `O6:O${TABLE_LAST_ROW}`]);
  let rangeOfTable = rangelistOfTable.getRanges();

  for (let i = 0; i < yearArr.length; i++) {
    for (let j = 0; j < monthArr.length; j++) {
      let textFinder = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.createTextFinder(yearArr[i] + '-' + toTwoDigits(monthArr[j]) + ".*").useRegularExpression(true);;
      let ranges = textFinder.findAll();
      // Logger.log(ranges);
      /********紫(キャンセル分)の中から年と月が一致したものだけ数える */
      toCancelDateArr = [];
      toCancelDateArrFromRange = [];
      toCancelAmountArr = [];
      toCancelAmountArrFromRange = [];
      for (let h = 0; h < sheet0ItemCount.getNumRows(); h++) {
        if (sheet0ItemCount.getCell(1 + h, 1).getBackground() === TOTALLING_SHEET.getRange(9, 2).getBackground()) {
          toCancelDateArr.push(RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6 + h, 3));
          toCancelAmountArr.push(RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(6 + h, 15));
        }
      }
      // Logger.log(toCancelCount); /**キャンセルのRangeが入っている */
      // Logger.log(toCancelAmountArr); /**キャンセルのRangeが入っている */

      /**toCancelDateArrの処理 */
      if (ranges.length > 0) {
        toCancelDateArr.forEach((rng) => {
          let value = rng.getValue();
          if (value.getFullYear() === yearArr[i] && (1 + Math.floor(value.getMonth())) === Math.floor(monthArr[j])) {
            toCancelDateArrFromRange.push(rng.getRow());
          }
        });
      }

      /**特定の日付のキャンセルした価格の合計値を調べる。 */
      let sumOfCancel = 0;
      if (ranges.length > 0) {
        /**キャンセルするデータの配列に中身が入っているときに実行 */
        if (toCancelDateArrFromRange.length > 0) {
          /**上でキャンセルデータの行番号を取り出したので、それを使って価格だけ取りだそう */
          for (let h = 0; h < toCancelDateArrFromRange.length; h++) {
            let rangeOfAmount = RAKUTEN_ICHIBA_MANAGEMENT_SHEET.getRange(toCancelDateArrFromRange[h], 15);
            let rangeValueOfAmount = rangeOfAmount.getValue();
            sumOfCancel += rangeValueOfAmount;
          }
        }
      }


      /*******
       * まずは表の価格の値全部のRangeオブジェクトを取る。
       * それぞれの年と月を使って検索し、一致したRangeオブジェクトを配列に格納する。
       * 初期化してあった配列には、年と月の一致したRangeオブジェクトが入っている。
       * .getValue()とfor文を使って、合計値を計算する。
       * それぞれの年と月ごとの合計値を変数に格納する。
      */
      // /**このログはテーブルの日付一覧を取り出している。 */
      // Logger.log(rangeOfTable[0].getValues());
      // /**このログはテーブルの価格一覧を取り出している。 */
      let amountData = rangeOfTable[1].getValues()
      // Logger.log(amountData);

      /**
       * 年データと月データで検索して一致したテーブルのセルの行を取得し、
       * その行を使って、価格のRangeオブジェクトのgetValue()にアクセスする。
       */

      let sumOfAllWithYearAndMonth = 0;

      let incrementData = countToamountDataArr.length;

      if (ranges.length > 0) {
        // 繰り返すのは検索にヒットした数。
        for (let h = incrementData; h < ranges.length + incrementData; h++) {
          sumOfAllWithYearAndMonth += parseInt(amountData[h]);
          // Logger.log('h: ' + h);
          // Logger.log('amountData[h]: ' + amountData[h]);
          countToamountDataArr.push(h);
        }
        // Logger.log(sumOfAllWithYearAndMonth);
      }

      if (ranges.length > 0) {
        TOTAL_PAYMENT_SHEET.getRange(4 + j, 3 + i).setValue(sumOfAllWithYearAndMonth - sumOfCancel);
        TOTAL_PAYMENT_SHEET.getRange(4 + j, 3 + i).setBackground("#ffffff");
      } else if (ranges.length === 0) {
        TOTAL_PAYMENT_SHEET.getRange(4 + j, 3 + i).setValue(ranges.length);
        TOTAL_PAYMENT_SHEET.getRange(4 + j, 3 + i).setBackground("#dddddd");
      }
    }
  }
}