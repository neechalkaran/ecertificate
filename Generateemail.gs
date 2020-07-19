/*
****************Steps**************

Upload the certificate template image in google drive and get fileid
Identify the middle pixel position for the texts in the image.
goto Tools->Script Editor
Custmize the code as per your certificates and text. Maximum 10 columns can be used
place your details in sheet List. ensure equal text input parameters of column number
Place your email body content. you can customize the code too.

To get latest code and latest version checkout https://github.com/neechalkaran/ecertificate
Reachout neechalkaran@gmail.com for any feedback
If the plugin is useful, feel free to sponsor for server cost at http://www.neechalkaran.com/p/donate.html
*/

function Generateemail() {
  var subject="Congratulation";/* Your subject*/
  var ccid="yourid@gmail.com";/* Your mailid*/
  
  var spreadsheet = SpreadsheetApp.getActive();
  var list = spreadsheet.getSheetByName("List");
  
  for(var i=2;i<=list.getLastRow();i++)
  {
  if(list.getRange(i, 5).getValue()!="")continue;
  var content =spreadsheet.getSheetByName("Body").getRange(1,1).getValue();
  content=content.replace("{{ Name }}", list.getRange(i, 2).getValue());
  content =content.replace("{{ image }}", "<img src='cid:userLogo1' class='certificate' />");
  content =content.replace(/\n/gi, "<br/>");
var data= Utilities.base64Encode(DriveApp.getFileById("<------fileid----->").getBlob().getBytes());

var imageurl="http://apps.neechalkaran.com/ecertificate.php";
var options = {"method" : "GET","followRedirects" : false,
"payload":{"template":data,
"text1": list.getRange(i, 2).getValue()+",1000,490,30",//customize the pixels. max 10 text
"text2": list.getRange(i, 3).getValue()+",1000,560,30"}
}


//send image as attachment
uBlob = UrlFetchApp.fetch(imageurl,options).getBlob().setName(list.getRange(i, 2).getValue()+" Certificate");
var allimage=new Object();
allimage["userLogo1"]=uBlob;
MailApp.sendEmail({to:list.getRange(i, 4).getValue(), cc:ccid, subject:subject, htmlBody: content, inlineImages:allimage}); 
list.getRange(i, 5).setValue("done");
  
  }
  
};

function convertutf(s) {
function parse(a, c){return string.fromCharCode(parseInt(c, 16));}
return encodeURIComponent(s.replace(/%u([0-f]{4})/gi, parse));
}
