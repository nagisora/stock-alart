TARGET_STOCK = 2521; // 通知を設定する銘柄
TARGET_DROP_RATE = -5; // 騰落率(%)
TAEGET_VALUE = 5; // どの値で比較？2=始値、5=終値

// memo
// ワイルドカードを使ってtdを取得したい
// forを回さなくてもrange指定でスプレッドシートに書き込めそう。処理が早くなる気がする。

function send_message(rate){
　　// パラメーターの指定(分かりやすいように変数に格納します)
    var recipient = Session.getActiveUser().getUserLoginId(); // 実行者のメールアドレスを取得
    var subject = "株の購入タイミングです！";
    var body = "騰落率は" + Math.round(rate*100*100) / 100 + "%です\n";

    // メール送信
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log("お知らせメール送信済");
}

function myFunction() {
  // 株価を取得するシートを開く
  const workbook = SpreadsheetApp.getActive();
  const sheet= workbook.getSheetByName(TARGET_STOCK);

  // 株価の時系列データの取得
  let URL = "https://minkabu.jp/stock/" + TARGET_STOCK + "/daily_bar";
  let html = UrlFetchApp.fetch(URL).getContentText("UTF-8");
  
  let stock_data_table = Parser.data(html)
  .from('<table id="fourvalue_timeline" class="md_table">')
  .to('</table>')
  .build();
  // ここまででstock_data_tableに株価の時系列データのhtmlが取得できた

  // trの中身を配列で取得する
  let stock_data_tr = Parser.data(stock_data_table)
  .from('<tr>')
  .to('</tr>')
  .iterate();
  // ここまででstock_data_allに株価の時系列データのhtmlが取得できた  
  //Logger.log(stock_data_tr[1]);

  // 処理行の初期値を設定
  let row_now = 2;

  // ループを回して、株価の時系列データをスプレッドシートに書き込む
  for(let i=1; i<stock_data_tr.length; i++){

    // スプレッドシートの最初の行の日付を取得
    let date_first_row = dayjs.dayjs(sheet.getRange(row_now, 1).getValue()).format('YYYY/MM/DD');

    // 時系列データの日付を取得
    let stock_data_td_day = Parser.data(stock_data_tr[i])
    .from('<td class="tal">')
    .to('</td>')
    .build();

    // 日付を比較して同じじゃないなら処理を続ける
    if(date_first_row !== stock_data_td_day) {

      // その日付の株価データを取得
      let stock_data_td = Parser.data(stock_data_tr[i])
      .from('<td class="num">')
      .to('</td>')
      .iterate();

      // スプレッドシートに書き込み
      sheet.insertRows(row_now, 1);
      sheet.getRange(row_now, 1).setValue(stock_data_td_day);
      for(let j=0; j<stock_data_td.length; j++){
        sheet.getRange(row_now, j+2).setValue(stock_data_td[j]);
      }
    }
    row_now++;
  }

  // 最新データと5つ前のデータの比率を計算して書き込み
  let price_now = sheet.getRange(2, TAEGET_VALUE).getValue();
  let price_5ago = sheet.getRange(7, TAEGET_VALUE).getValue();
  let value_rate =  (price_now - price_5ago) / price_5ago;

  sheet.getRange(2, 8).setValue(value_rate).setNumberFormat('0.00%');

  // 騰落率が設定値を下回っていたらメールを送る
  if(value_rate*100 < TARGET_DROP_RATE) {
    send_message(value_rate);
  }
}
