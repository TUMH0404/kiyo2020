// メニューの作成  //

function onOpen(){
  var ui = DocumentApp.getUi();
  
  var menu = ui.createMenu('スタート');
  menu.addItem("作成",'main');
  
  menu.addSeparator();

  menu.addSubMenu(         
      ui.createMenu("解答受付")  
      .addItem("受付開始","permission").addItem("受付終了","reject")  
  );
   menu.addSeparator();

  
   menu.addSubMenu(    
      ui.createMenu("共有")   
      .addItem("共有on","Vieweron").addItem("共有off","Vieweroff")    
  ); 
  menu.addSeparator();
  menu.addSubMenu(                          
      ui.createMenu("採点")             
      .addItem("結果取得","GetAllResult")   
  );
  menu.addToUi();
  
  
}


// 初期情報  //
const Formstud = [];
const studlist = {};

//  初期情報の取得  //
function Initialize() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var para = body.getParagraphs();
  var drivename = para[0].getText().slice(para[0].getText().indexOf("#")+1).trim();
  var studentsfilename = para[1].getText().slice(para[1].getText().indexOf("#")+1).trim();
  var quizsheet = para[2].getText().slice(para[2].getText().indexOf("#")+1).trim();
  var quiznumber = para[3].getText().slice(para[3].getText().indexOf("#")+1).trim();
//  Logger.log(drivename);
//  Logger.log(studentsfilename);
//  Logger.log(quizsheet);
//  Logger.log(Number(quiznumber));
  return [drivename,studentsfilename,quizsheet,Number(quiznumber)];
  
}


//  メイン関数  //
function main(){
  var lists = Initialize();
  //var dat = "Javascript_Python_Question";
  var formIDforAnswer= MakeQuiz(lists[0],lists[1],lists[2],lists[3]);
  var Coment = DocumentApp.getActiveDocument();
  var BODY = Coment.getBody();
  var mm = Moment.moment();
  BODY.appendParagraph("");
  BODY.appendParagraph("***********************************************");
  BODY.appendParagraph("");
  BODY.appendParagraph(mm.format('YYYY年M月D日 H時m分'));
  BODY.appendParagraph("");
  BODY.appendParagraph(formIDforAnswer[3]);  
  BODY.appendParagraph(formIDforAnswer[0]);
  BODY.appendParagraph(formIDforAnswer[1]);
  BODY.appendParagraph(formIDforAnswer[2]);
  
  var ui = DocumentApp.getUi();
  var response = ui.alert("問題作成完了！");
}

//  以下，サブ関数   //
//
function CreateNewFolder(name){
  var app = DriveApp.createFolder(name);
  return app.getId();
}


//
function GetSpreadID(FileNameString){
  var FileIterator = DriveApp.getFilesByName(FileNameString);
  while (FileIterator.hasNext())
  {
    var file = FileIterator.next();
    if (file.getName() == FileNameString)
    {
      var Sheet = SpreadsheetApp.open(file);
      var fileurl = file.getUrl();
      var fileid = file.getId();    
    }    
  }
  return fileid;
}


//
function intRandom(min, max){
  return Math.floor( Math.random() * (max - min + 1)) + min;
}


//
function intRand(max,QQlist){
  var randoms = [];
for(var i = 0; i <= QQlist-1; ++i){
  if(i==max){
    break;
  }
  while(true){
    var tmp = intRandom(0, QQlist-1);
    if(!randoms.includes(tmp)){
      randoms.push(Number(tmp));
      break;
    }
  }
  
}

  return randoms;
} 

//
function intRand5(){
  var randoms = [];

for(var i = 0; i < 5; i++){
  while(true){
    var tmp = intRandom(0, 4);
    if(!randoms.includes(tmp)){
      randoms.push(Number(tmp));
      break;
    }
  }

}
  return randoms;
} 


//
function GetSpreaddata(FileNameString){
  
  var FileIterator = DriveApp.getFilesByName(FileNameString);
  while (FileIterator.hasNext())
  {
    var file = FileIterator.next();
    if (file.getName() == FileNameString)
    {
      var Sheet = SpreadsheetApp.open(file);
      var fileurl = file.getUrl();
      var fileid = file.getId();
    }    
  }
  return fileid;
}


//
function createSpreadsheetInfolder(folder,studfilename) {
  var studfile = {};
  var newfolder = DriveApp.createFolder(folder);
  var newfolderid = newfolder.getId();
  const folderdd = DriveApp.getFolderById(newfolderid);
  
  var sss = GetSpreaddata(studfilename);
  var Studfileid = SpreadsheetApp.openById(String(sss));
  var lists = Studfileid.getRange('A2:C'+String(Studfileid.getLastRow())).getValues();
  
  for(var i=0;i<lists.length;++i){
      
      var row = lists[i];
      var newSS=SpreadsheetApp.create(String(row[2]));
      var originalFile=DriveApp.getFileById(newSS.getId());
      folderdd.addFile(originalFile);
      studlist[String(row[2])]={"name":row[0],"email":row[1],"url":String(originalFile.getUrl())};
      DriveApp.getRootFolder().removeFile(originalFile);
  }
  
  var sheet = SpreadsheetApp.create("Answer");
  var sheeturl = SpreadsheetApp.create("sheeturl");
  const file = DriveApp.getFileById(sheet.getId());
  const fff = DriveApp.getFileById(sheeturl.getId());
  
  folderdd.addFile(file);
  folderdd.addFile(fff);
  
  var form = FormApp.create('試験解答一覧');
  
  var tmp = form.getPublishedUrl();
  var formFile = DriveApp.getFileById(form.getId());
  folderdd.addFile(formFile);
  
  var lll = [studlist,Studfileid.getLastRow()-1,sheet.getUrl(),sheeturl.getUrl(),formFile.getUrl(),newfolderid,formFile.getId()];
  DriveApp.getRootFolder().removeFile(file);
  DriveApp.getRootFolder().removeFile(fff);
  DriveApp.getRootFolder().removeFile(formFile);
  
  return lll;
}



//
function GetQQ(dat,quizmax,num){  

  var QQQ = {};
  var rt = intRand(num,quizmax);
  for(var x=0;x<rt.length;++x){
    var g = dat[rt[x]];
    
    var ans = g[7].toFixed();
    var ansarr = [false,false,false,false,false];
    for(var i=0;i<ans.length;++i){
      ansarr.splice(ans[i]-1, 1,true);
    }
    var newans=[];
    for(var i=0;i<5;++i){
      newans.push([ansarr[i],g[i+2]]);
    }
    
    var rn = intRand5();
    var ns = [];
    for(var i=0;i<rn.length;++i){
      ns.push(newans[rn[i]]);
    }
    
    var newa = "";
    var news = [];
    for(var i=0;i<5;++i){
      
      if(ns[i][0]==true){
        newa=newa+String(i+1);
      }
      news.push(ns[i][1]);
    }
    QQQ[rt[x]]={"q":g[0],"f":g[1],"a":newa,"s":news};    
  }
    return QQQ;
}




//
function addFormItems(form,question_list,nn){
  var ll = [];
  for(var i=1;i<=nn;++i){
    ll.push("Q."+String(i));
  }
  form.setIsQuiz(true);
 
  form.addListItem()
    .setTitle("学籍番号")
    .setChoiceValues(question_list)
    .setRequired(true);
  
  var gridItem = form.addGridItem();
gridItem.setTitle('適切な選択肢の番号にチェックをいれなさい。')
  .setRows(ll)
  .setColumns(["1","2","3","4","5"]);
  var gridValidation = FormApp.createGridValidation()
  .setHelpText("Select one item per column.")
  .build();
gridItem.setValidation(gridValidation);
  return form.getId();
}



//
function MakeQuiz(foldername,studdata,quizdata,num){
  var sss = GetSpreaddata(quizdata);
  var file = SpreadsheetApp.openById(sss);
  var sheet = file.getSheets()[0];
  var dat = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
  ///
  var studfortestdata = createSpreadsheetInfolder(foldername,studdata);
  var studdata = studfortestdata[0];
  var studnum = studfortestdata[1];
  var studansId = studfortestdata[2];
  var studansUrl = studfortestdata[3];
  var formFileUrl = studfortestdata[4];
  var folderid = studfortestdata[5];
  var studlistforForm =[];

  var ddd = [];
  for(var i=0;i<studnum;++i){
      ddd.push(GetQQ(dat,sheet.getLastRow()-1,num));
  }
        var Studsans = [];
        var StudURLs = [];
        var i=0;
        for(var y in studdata){
            studlistforForm.push(y);
            StudURLs.push([y,studdata[y]["email"],studdata[y]["url"]]);
            var file = SpreadsheetApp.openByUrl(studdata[y]["url"]);
            var StudSheet = file.getSheets()[0];
        var quizsheet = [];
        
        var ddd1 = ddd[i];
              u=1;
              var ansdata = [y];
              for(var x in ddd1){
                ansdata.push(ddd1[x]["a"]);
                quizsheet.push([String(u)+". "+ddd1[x]["q"]]);
                
                quizsheet.push([ddd1[x]["f"]]);
                quizsheet.push([""]);
                
               
                      for(var a=0;a<ddd1[x]["s"].length;++a){
                        quizsheet.push([String(a+1)+"--->> "+ddd1[x]["s"][a]]);
                      }
                quizsheet.push([""]);
                quizsheet.push([""]);
                u=u+1;
              }
          Studsans.push(ansdata);
        StudSheet.getRange(2, 1, quizsheet.length).setValues(quizsheet);
        var lastrow = StudSheet.getLastRow();
        
         StudSheet.getRange(lastrow+3, 1, 1, 1)
         .setFormula('=HYPERLINK("'+formFileUrl+'","解答先[クリック]")')
         .setFontSize(18).setHorizontalAlignment("left");
        StudSheet.getRange(2, 1, quizsheet.length)
        .setHorizontalAlignment("left").setFontSize(12)
        .setVerticalAlignment("top").setWrap(true);
       
              for(var j=0;j<num*10;++j){ 
                if((j%10)==1+1+1){

                  var rtt = StudSheet.getRange("A"+String(j)).getValue();
                  
                  if(rtt.length>10){  
                  StudSheet.getRange("A"+String(j)).setValue("");
                  var inkan = String(rtt);
                  
                  var blob = DriveApp.getFileById(inkan).getBlob();
                  
                  StudSheet.insertImage(blob, 2,j);
                  StudSheet.setColumnWidth(1, 330);
                  StudSheet.setRowHeight(j, 230);
                  }
            
                }
              }
          SpreadsheetApp.flush();
          i=i+1;
        }

  var fileanssheet = SpreadsheetApp.openByUrl(studansId);
  var StudSheet = fileanssheet.getSheetByName("シート1");
  StudSheet.getRange(1,1,studnum,1+num).setValues(Studsans);
  SpreadsheetApp.flush();
  var fileanssheet1 = SpreadsheetApp.openByUrl(studansUrl);
  var StudSheet1 = fileanssheet1.getSheetByName("シート1");
  StudSheet1.getRange(1,1,studnum,3).setValues(StudURLs);
  SpreadsheetApp.flush();
  var form = FormApp.openByUrl(formFileUrl);
  var formid=addFormItems(form,studlistforForm,num);

  return [formFileUrl,studansUrl,studansId,folderid];
}


//
function permission(){
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var para = body.getParagraphs();
  var key = para[para.length-3].getText();
 
  var form = FormApp.openByUrl(key); 
  form.setAcceptingResponses(true);
  
  var ui = DocumentApp.getUi();
  var response = ui.alert("解答受付中です。");
  
}


//
function reject(){
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var para = body.getParagraphs();
  var key = para[para.length-3].getText();
 
  var form = FormApp.openByUrl(key); 
  form.setAcceptingResponses(false);
  
  var ui = DocumentApp.getUi();
  var response = ui.alert("解答終了しました。");
 
}


//
function Vieweron(){
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var para = body.getParagraphs();
    var key = para[para.length-2].getText();
  Logger.log(key);
    var sheet =SpreadsheetApp.openByUrl(key);
    var sheetid = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
      for(var i=0;i<sheetid.length;++i){
         var dc = SpreadsheetApp.openByUrl(sheetid[i][2]);
         dc.addViewer(sheetid[i][1]);  
      }  
  var ui = DocumentApp.getUi();
  var response = ui.alert("閲覧可能にしました。");
}



//
function Vieweroff(){
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var para = body.getParagraphs();
    var key = para[para.length-2].getText();
    var sheet =SpreadsheetApp.openByUrl(key);
    var sheetid = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
      for(var i=0;i<sheetid.length;++i){
         var dc = SpreadsheetApp.openByUrl(sheetid[i][2]);
         dc.removeViewer(sheetid[i][1]);
      }  
    var ui = DocumentApp.getUi();
    var response = ui.alert("閲覧の許可を解除しました。");
}



//
function getdata(keys){
  RRR = {};
 var form = FormApp.openByUrl(keys);
 var formResponses = form.getResponses();
 for (var i = 0; i < formResponses.length; i++) {
   var formResponse = formResponses[i];
   var itemResponses = formResponse.getItemResponses();
       RRR[itemResponses[0].getResponse()]=[];
     
      var newobj = Object.values(itemResponses[1].getResponse());
       for(var y=0;y<newobj.length;++y){
         //RRR[itemResponses[0].getResponse()].push(merg(newobj[y]));
         RRR[itemResponses[0].getResponse()].push(newobj[y]);
       }
   }
  return RRR;
}

//
function getAnswers(keya){
  var ansdat ={};
  var sheet = SpreadsheetApp.openByUrl(keya);
  var sheetdata = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  for(var k=0;k<sheetdata.length;++k){
  result =[];
  for(var i=1;i<sheetdata[k].length;++i){
    result.push(sheetdata[k][i].toFixed());
  }
    ansdat[sheetdata[k][0]]=result;
  }
  return ansdat;
}


//
function setresult(keya,keys,ccc,num){
  var resultdat = setAnswer(keya,keys);
  var sheetss = SpreadsheetApp.create("Result");
  var filename = DriveApp.getFileById(sheetss.getId());
  
  var folder = DriveApp.getFolderById(ccc);
  folder.addFile(filename);
  var sheetn = sheetss.getSheetByName("シート1");
  
  var r = 1;
  for(var x in resultdat){
      var sheetnn = sheetn.getRange(r, 1).setValue(x);
      var sheetvv = sheetn.getRange(r, 2, 1, num).setValues([resultdat[x]]);
      r=r+1;
  }
  DriveApp.getRootFolder().removeFile(filename);
  }


//
function GetAllResult(){
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var para = body.getParagraphs();
  var studansform = para[para.length-3].getText();
  var ans = para[para.length-1].getText();
  var drivekey = para[para.length-4].getText();
  var quiznumber = para[3].getText().slice(14).trim();
  setresult(ans,studansform,drivekey,Number(quiznumber));

}

//
function setAnswer(keya,keys){
  var datA= getAnswers(keya);
  var datS = getdata(keys);
  var keyA = getkeys(datA);
  var keyS = getkeys(datS);
  
  var Point ={};
  for(x in datS){
    Point[x]=matching(datA[x],datS[x]);
    Logger.log(x);
    keyS = keyS.filter(n => n !== x);
    if(keyS.length==0){
      break;
    }
  }
  Logger.log(Point);
  return Point;
  
}

//
function getkeys(dict){
  keys = [];
  for(var x in dict){
    keys.push(x);
  }
  return keys;
}

//
function matching(a,b){
  result = []
  Logger.log(a);
  Logger.log(b);
  if(a!=null & b!=null){
    
  for(var i=0;i<a.length;++i){
      result.push(Number(a[i]===b[i]));
    }
  }
  else{
    Logger.log("回答数が合いません")
  }
  return result;
}


////
//function merg(d){
//    var k="";
//    for(var i=0;i<d.length;++i){
//      k+=String(d[i]);
//    }
//  return k;
//}
