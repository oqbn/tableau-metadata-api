// Tableau情報用
let TABLEAU_SERVER;
let TABLEAU_SITE;
let TOKEN_NAME;
let TOKEN_SECRET;

function fetchData() {
  Logger.log('fetchData Start');

  // Tableau情報取得
  var ssInit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');

  TABLEAU_SERVER = ssInit.getRange(1,2).getValue();
  TABLEAU_SITE = ssInit.getRange(2,2).getValue();　// TableauServerのデフォルトサイトの場合は空文字列
  TOKEN_NAME = ssInit.getRange(3,2).getValue();
  TOKEN_SECRET = ssInit.getRange(4,2).getValue()
  var workbookName = ssInit.getRange(5,2).getValue();  
  
  // Tableauサインイン
  var t_res = signinTableau_();
  var t_json=JSON.parse(t_res.getContentText());
  
  // アクセストークンの取得
  var token = t_json['credentials']['token'];
    
  // GraphQLのクエリを設定
  var url = 'https://' + TABLEAU_SERVER + '/api/metadata/graphql';
  var graphql = HtmlService.createHtmlOutputFromFile('graphql').getContent();

  var headers = {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'X-Tableau-Auth': token
  };
 
  var options = {
    'method' : 'post',
    'muteHttpExceptions':true,
    'headers' : headers,
    'payload' : JSON.stringify({ query : graphql , variables : { "workbookName" : workbookName } })
  };

  // Logger.log(options);

  // GraphQLの呼び出し
  var res = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(res.getContentText());

  // Logger.log(json);

  var workbooks = json['data']['workbooks'];
  var views = workbooks[0]['views'];
  var datasources = workbooks[0]['embeddedDatasources'];

  // ダッシュボードとシートの一覧作成
  Logger.log('sheetView');

  var fnameView = 'V:' + workbookName + '_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss') ;
  var ssView = SpreadsheetApp.getActiveSpreadsheet().insertSheet(fnameView);

  // ヘッダ作成
  ssView.getRange(1,1).setValue("index");
  ssView.getRange(1,2).setValue("view");
  ssView.getRange(1,3).setValue("type");
  ssView.getRange(1,4).setValue("path");
  ssView.getRange(1,5).setValue("containedInDashboards");
  ssView.getRange(1,6).setValue("sheet");
  var urlBase = "https://" + TABLEAU_SERVER + "/#/site/" +  TABLEAU_SITE + "/views/";
  var row = 2;  // 2行目から

  // ビューごとに出力
  for (var i=0 ; i < views.length ; i++ ){
    ssView.getRange(row,1).setValue(views[i]['index']);
    ssView.getRange(row,2).setValue(views[i]['name']);
    ssView.getRange(row,3).setValue(views[i]['__typename']);
    if(views[i]['path'].length > 0){
      ssView.getRange(row,4).setValue(urlBase + views[i]['path']);
    }
    if(views[i]['containedInDashboards'] != undefined ){
      ssView.getRange(row,5).setValue(views[i]['containedInDashboards'].map(value => value['name']).join('\n'));
    }
    if(views[i]['sheets'] != undefined ){
      ssView.getRange(row,6).setValue(views[i]['sheets'].map(value => value['name']).join('\n'));
    }
    ssView.getRange(row,1,1,6).setVerticalAlignment('top');　// レイアウト整える
    row += 1;
  }

  // データソースと計算式の一覧作成
  Logger.log('sheetDatasource');

  var fnameDS = 'DS:' + workbookName + '_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss') ;
  var ssDS = SpreadsheetApp.getActiveSpreadsheet().insertSheet(fnameDS);

  // ヘッダ作成
  ssDS.getRange(1,1).setValue("datasource");
  ssDS.getRange(1,2).setValue("field name");
  ssDS.getRange(1,3).setValue("description");
  ssDS.getRange(1,4).setValue("dataType");
  ssDS.getRange(1,5).setValue("defaultFormat");
  ssDS.getRange(1,6).setValue("formula");
  ssDS.getRange(1,7).setValue("referencedByCalculations");
  ssDS.getRange(1,8).setValue("downstreamSheets");
  var row = 2;  // 2行目から

  // データソースごとに出力
  for (var i=0 ; i < datasources.length ; i++ ){
    // 各データソースのフィールドの取得
    var fields = datasources[i]['fields'];
    for ( var j=0 ; j < fields.length ; j++){
      ssDS.getRange(row,1).setValue(datasources[i]['name']);
      ssDS.getRange(row,2).setValue(fields[j]['name']);
      ssDS.getRange(row,3).setValue(fields[j]['description']);
      ssDS.getRange(row,4).setValue(fields[j]['dataType']);
      ssDS.getRange(row,5).setValue(fields[j]['defaultFormat']);
      ssDS.getRange(row,6).setValue(fields[j]['formula']);
      ssDS.getRange(row,7).setValue(fields[j]['referencedByCalculations'].map(value => value['name']).join('\n'));
      ssDS.getRange(row,8).setValue(fields[j]['downstreamSheets'].map(value => value['name']).join('\n'));
      ssDS.getRange(row,1,1,8).setVerticalAlignment('top');　// レイアウト整える
      row += 1;
    }
  }

  // Tableuサインアウト
  signoutTableau_(token);

  Logger.log('fetchData End');

}


/**
 * Tableauサインイン
 * @return response
 */
// https://help.tableau.com/current/api/rest_api/en-us/REST/rest_api_ref_authentication.htm#sign_in

function signinTableau_(){
  Logger.log('signinTableau');
  
  var url = 'https://' + TABLEAU_SERVER + '/api/3.17/auth/signin';
  
  // personal access token (PAT)を使用
  var payload = { 'credentials': 
    { 'personalAccessTokenName': TOKEN_NAME, 
    'personalAccessTokenSecret': TOKEN_SECRET, 
    'site': {'contentUrl': TABLEAU_SITE}
    }};

  var headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json'
  };
  
  // リクエストを送信するオプションを定義
  var options = {
    'method': 'POST',
    'muteHttpExceptions':true,
    'headers':headers,
    'payload': JSON.stringify(payload)
  };
  
  // リクエスト送信
  var res = UrlFetchApp.fetch(url, options);
  // Logger.log(res.getContentText());
  
  return res;
};


/**
 * Tableauサインアウト
 * @param token
 */

function signoutTableau_(token){
  Logger.log('signoutTableau');
    
  // SignOut用のURL
  var url = 'https://' + TABLEAU_SERVER + '/api/3.17/auth/signout';
  
  var headers = {
    'Accept': 'application/json',
    'X-Tableau-Auth': token
  }
  
  // リクエストを送信するオプションを定義
  var options = {
    'method': 'POST',
    'muteHttpExceptions':true,
    'headers': headers
  };
    
  // リクエスト送信
  var res = UrlFetchApp.fetch(url, options);
  
};
