function firestoreData() {
  //jsonファイルのemail、key、projectIdを読み込み、returnで返す
  const prop = PropertiesService.getScriptProperties().getProperties();
  const dataArray = {
    "email":prop.FIRESTORE_EMAIL,
    "key": prop.FIRESTORE_KEY,
    "projectId": prop.FIRESTORE_ID
  }
  return dataArray;
}

function start() {
  //実行時にポップアップを表示する
  let result = Browser.msgBox('在庫を更新してよろしいですか？', Browser.Buttons.OK_CANCEL);
   if (result != 'cancel'){
    // myFunction();
    addStockFunction();
  }
}






//メールを送る処理↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
function sendEmailFunction() {

  const stockBorder = 10; //在庫がこの値以下になったらメールする
  

  // const myEmail = Session.getActiveUser().getEmail();
  // const options = { cc: myEmail };
  const options = {};



  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('在庫リスト');
  const dataSheet = ss.getSheetByName('GAS用データ');
   //GAS用データから商品数を取得
  const dataSheetRange = dataSheet.getRange('A2');
  const ITEM = dataSheetRange.getValues();//商品数

  const emailRange = dataSheet.getRange('A5');
  const email = emailRange.getValues();

  // const email = "ne231146@senshu-u.jp";//メールの宛先
  console.log(email);
  //在庫リストから在庫データを取得
  let stockData = stockSheet.getSheetValues(2, 2, ITEM, 5);

  let needEmail = false;
  let body = `以下の商品の在庫数が少なくなっています。
  
  `;//この改行はメールに反映される

  for(let i=0; i<stockData.length; i++){
    let stock = parseInt(stockData[i][3]);
    if(stock <= stockBorder) {
      if(needEmail === false) needEmail = true;
      let itemName = stockData[i][1];
      body += `${itemName}   在庫数：${stock}
  `;//この改行はメールに反映される

    }

  }

  if(needEmail === true){


  body += `在庫シートURL: https://docs.google.com/spreadsheets/d/1qXUGBUq7liaBOKEGvZW8XuLNzW14oFmmU2iu7Zzjx1c/edit?usp=sharing`;
  const title =`無人コンビニNEst. 在庫減少のお知らせ`;
  GmailApp.sendEmail(email, title, body, options);
  console.log('emailを送信');
  }

}














//firebaseの在庫数をもとにスプレッドシートを更新する関数
function getStockFunction() {
  const dataArray = firestoreData();//jsonファイルから読み込む
  const firestore = FirestoreApp.getFirestore(dataArray.email, dataArray.key, dataArray.projectId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('在庫リスト');
  const dataSheet = ss.getSheetByName('GAS用データ');
  
  const products = firestore.getDocuments("products");

  const productsFirebase = products.map((product) => {

    //商品に関する情報が増えたら、ここを編集する↓ （増やしたら、下のスプレッドシートに書き込む処理の変数colNumも増やす)
    return [
      //product.fields.[firestore databaseの変数名].[(Stringとかintegerとか)Value]
      product.fields.category.stringValue, //カテゴリ
      product.fields.name.stringValue,//商品名
      parseInt(product.fields.price.integerValue),   //値段
      parseInt(product.fields.stock.integerValue),    //在庫数
      product.fields.jan.stringValue, //JANコード
      product.fields.status.booleanValue,   // 販売状況（販売可否）
    ]
  })
  console.log(productsFirebase);
  
  //カテゴリー順でソート（見やすくするため）
  productsFirebase.sort((a, b) => {
    if(b[0] > a[0]){
      return -1;
    }else{
      return 1;
    }
  });
  
  console.log(productsFirebase);
  //現在の時刻を取得
  const today = new Date();
  const month = today.getMonth() + 1;
  const day = today.getDate();
  const hours = today.getHours(); //時
  const minutes = today.getMinutes(); //分
  const seconds = today.getSeconds(); //秒
  const lastUpdated = `${month}月${day}日${hours} 時${minutes}分${seconds}秒`;



  // スプレッドシートに書き込む処理↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

  // データの形の宣言
  const row = 2             //最初のデータを入れる行番号
  const col = 2             //最初のデータを入れる列番号
  const rowNum = productsFirebase.length //行の長さ
  const colNum = 6          //列の長さ  今後商品の情報が増えたらここの値も増やす

  // スプレッドシートに記入
  stockSheet.getRange(row, col, rowNum, colNum).setValues(productsFirebase);//在庫データを在庫シートに書き込む
  stockSheet.getRange('I14').setValue(lastUpdated);
  stockColorFunction(row, productsFirebase);
  
  return productsFirebase;//addStockFunctionの使用時にデータを渡す　getのみなら何も起こらず終了
}










//商品の在庫を増やす関数
function addStockFunction() {
  const dataArray = firestoreData();//jsonファイルから読み込む
  const firestore = FirestoreApp.getFirestore(dataArray.email, dataArray.key, dataArray.projectId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('在庫リスト');
  const dataSheet = ss.getSheetByName('GAS用データ');
  const logSheet = ss.getSheetByName('在庫追加ログ');
  // const products = firestore.getDocuments("products");

  //GAS用データから商品数を取得
  const dataSheetItemRange = dataSheet.getRange('A2');
  const ITEM = dataSheetItemRange.getValues();//商品数
  //GAS用データからログ数を取得

  const dataSheetLogRange = dataSheet.getRange('C2');
  const LOG = dataSheetLogRange.getValues();
  let logData = logSheet.getSheetValues(2, 1, LOG, 5);

  let addStock = stockSheet.getSheetValues(2, 1, ITEM, 1);
  // console.log(addStock);//デバッグ用


  let stockData = getStockFunction();//もととなるデータ(旧データ)
  // console.log(stockData);//デバッグ用


  let newStockData = new Array();//スプレッドシートに反映させるデータ(新データ)
  let needLog = false;
  // let stockLogData = new Array();

  //追加の在庫がなければ、旧データをそのまま新データにプッシュ
  //追加があれば、在庫を更新&ログ用のデータ&firebaseの書き換えデータも合わせてプッシュ
  for(let i=0; i<stockData.length; i++){
    if(parseInt(addStock[i]) === 0){
      newStockData.push(stockData[i]);
    }else {
    //現在の時刻を取得
      const today = new Date();
      const year = today.getFullYear();
      const month = today.getMonth() + 1;
      const day = today.getDate();
      const hours = today.getHours(); //時
      const minutes = today.getMinutes(); //分
      if(needLog === false) needLog = true;

      let newStock = [
        stockData[i][0],
        stockData[i][1],
        stockData[i][2],
        parseInt(stockData[i][3]) + parseInt(addStock[i]),
        stockData[i][4],
      ]
      newStockData.push(newStock);
      let logStock = [
        `${year}年${month}月${day}日${hours} 時${minutes}分`,
        stockData[i][1],
        parseInt(addStock[i]),
        parseInt(stockData[i][3]),
        parseInt(stockData[i][3]) + parseInt(addStock[i])
      ]
      logData.push(logStock);
      let newFirebaseStock = {
            category:stockData[i][0],
            jan:stockData[i][4],
            name:stockData[i][1],
            price:parseInt(stockData[i][2]),
            stock:parseInt(stockData[i][3]) + parseInt(addStock[i]),
      }
      firestore.updateDocument("products/"+stockData[i][4], newFirebaseStock);

    }
  }
  console.log(newStockData);
  


  // スプレッドシートに書き込む処理↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
  // データの形の宣言
  const row = 2             //最初のデータを入れる行番号
  const col = 2             //最初のデータを入れる列番号
  const rowNum = newStockData.length //行の長さ
  const colNum = 6          //列の長さ  今後商品の情報が増えたらここの値も増やす
  stockSheet.getRange(2, 1, rowNum, 1).setValue(0);//「在庫の追加量」列を0にリセット
  stockSheet.getRange(row, col, rowNum, colNum).setValues(newStockData);//在庫データを在庫シートに書き込む

  //在庫の追加があれば、ログのデータを日付が新しい順にソート
  if(needLog === true)  {

    logData.sort((a, b) => {
    if(b[0] > a[0]){
      return 1;
    }else{
      return -1;
    }
  });

    logSheet.getRange(row, 1, logData.length, 5).setValues(logData);
  }
  
  stockColorFunction(row, newStockData);//在庫が少ない商品のセルの色を変える
}



//在庫が少ない商品のセルの色を変える処理↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
function stockColorFunction (row, stockData) {
  let stockBorder = 10;//在庫が1~stockBorder個のとき、黄色にする

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('在庫リスト');
  //0:赤  1~stockBorder：黄色  それ以上：青
  for(let i = 0; i < stockData.length;i++){
    let stock = parseInt(stockData[i][3]);
    if(stock === 0) {
      stockSheet.getRange(row + i,5).setBackground('#ea9999');

    }else if(stock <= stockBorder){
      stockSheet.getRange(row + i,5).setBackground('#ffe599');

    }else {
      stockSheet.getRange(row + i,5).setBackground('#c9dbf9');
    }
  }


}




















//購入データを日付の新しい順に取得してコンソールに出力するプログラム（在庫更新には使っていない）
function getPurchaseData() {
  const dataArray = firestoreData();//jsonファイルから読み込む
  const firestore = FirestoreApp.getFirestore(dataArray.email, dataArray.key, dataArray.projectId);
  const purchases = firestore.getDocuments("purchases");
  //　purchaseを取得するが、これだと商品が複数購入された際に、配列の中に商品の配列データが複数生成されてしまう(3次元配列になってしまう)ので、この後の処理で適切に処理する
const prePurchasesRes = purchases.map((purchase) => {
    

    const names = purchase.fields.items.arrayValue.values.map((name) => {
      return [
        Utilities.formatDate(new Date(Date.parse(purchase.fields.purchaseDate.timestampValue)), 'JST', 'yyyy-MM-dd(E) HH:mm:ss'),
        name.mapValue.fields.name.stringValue,
        name.mapValue.fields.quantity.integerValue,
        name.mapValue.fields.totalPrice.integerValue,
        // purchase.fields.purchaseDate.timestampValue
      ]

    });
    // const quantities = purchase.fields.items.arrayValue.values.map((quantity) => {
    //   return quantity.mapValue.fields.quantity.integerValue
    // });
    
    return [
    // //  purchase.fields.purchaseDate.timestampValue,
      names//購入された商品名
    // //  quantities.join(" "),//購入された数
    // //  purchase.fields.items.arrayValue.values[0].mapValue.fields

    ]
  });



  //ここでprePurchaseResのデータを、適切な2次元配列の購入履歴データに置き換える
  let purchasesRes = new Array;
  prePurchasesRes.forEach(function(x) {
    x.map(y => {
      for(let i=0;i<y.length;i++){
        purchasesRes.push(y[i]);
        // console.log(y[i]);  //デバッグ用
      }
      
      });
    
  });


  //購入履歴を日付の新しい順に並び替える
  purchasesRes.sort((a, b) => {
    if(b[0] > a[0]){
      return 1;
    }else{
      return -1;
    }
  });
    
    console.log(purchasesRes);  //デバッグ用
    // console.log(purchasesRes[0][4]);


const purchasesDate = purchases.map((purchase) => {
  
    
    return [
      Utilities.formatDate(new Date(Date.parse(purchase.fields.purchaseDate.timestampValue)), 'JST', 'yyyy年MM月'),
      purchase.fields.totalPrice.integerValue,  
    ]


});
console.log(purchasesDate);

  const purchaseMap = new Map();
  const profitMap = new Map();
   for(let i = 0; i<purchasesDate.length; i++){
    let date = purchasesDate[i][0];
    if(purchaseMap.has(date)){
      let count = purchaseMap.get(date);
      count++;
      purchaseMap.set(date, count);
      let sum = profitMap.get(date);
      sum += parseInt(purchasesDate[i][1]);
      profitMap.set(date, sum);
    }else {
      purchaseMap.set(date, 1);
      profitMap.set(date, parseInt(purchasesDate[i][1]));
    }
    
      // console.log(stockMap.get(itemName));  //デバッグ用
   }



  let purchaseCount = new Array();
  purchaseMap.forEach((value,key)=> {
      let sum = profitMap.get(key);
      purchaseCount.push([key, value, sum]);
  });

  console.log(purchaseCount);

purchaseCount.sort((a, b) => {
    if(b[0] > a[0]){
      return 1;
    }else{
      return -1;
    }
  });



  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const purchaseSheet = ss.getSheetByName('購入データ テスト');



  // データの形の宣言
  const row = 2             //最初のデータを入れる行番号
  const col = 1             //最初のデータを入れる列番号
  const rowNum = purchasesRes.length //行の長さ
  const colNum = 4          //列の長さ  
  // スプレッドシートに記入

  purchaseSheet.getRange(row, col, rowNum, colNum).setValues(purchasesRes);
  purchaseSheet.getRange(3, 8, purchaseCount.length, 3).setValues(purchaseCount);



}



