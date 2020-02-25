function wipe_() {
  var html = HtmlService.createHtmlOutputFromFile('wipe')
  .setTitle("Gmail Deep Clean")
  .setWidth(350).setHeight(120);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

function stop_(e) {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  if (!e) {
    Browser.msgBox("The Gmail Deep Clean has been disabled. You can restart it anytime later."); 
  }  
}

function onOpen() {
  var menu = [     
    {name: "Clean", functionName: "wipe_"},
    null,
    {name: "Uninstall",  functionName: "stop_"},
    null
  ];  
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Gmail Deep Clean", menu);
}


function log_(status, subject, view, from, link) {
  var ss = SpreadsheetApp.getActive();
  ss.getActiveSheet().appendRow([status, subject, view, from, link]);
}


function doGmail() {
  try {
    var threads = GmailApp.getInboxThreads();
    Logger.log(threads)
    var url, urls, message, raw, body, formula, status;
    var hyperlink = '=HYPERLINK("#LINK#", "View")';
    var hrefs = new RegExp(/<a[^>]*href=["'](https?:\/\/[^"']+)["'][^>]*>(.*?)<\/a>/gi);
    for (var t in threads)  {
      url = "";
      status = "Could not unsubscribe";
      message = threads[t].getMessages()[0];
      raw = message.getRawContent();
      urls = raw.match(/^list\-unsubscribe:(.|\r\n\s)+<(https?:\/\/[^>]+)>/im);
      Logger.log(urls)
      if (urls) {
        url = urls[2];
        status = "Unsubscribed via header";
      } else {
        body = message.getBody().replace(/\s/g, "");
        while ( (url === "") && (urls = hrefs.exec(body)) ) {
          if (urls[1].match(/unsubscribe|optout|opt\-out|remove/i) || urls[2].match(/unsubscribe|optout|opt\-out|remove/i)) {
            url = urls[1];
            status = "Unsubscribed via link";
          }
        }
      }
      
      if (url === "") {
        urls = raw.match(/^list\-unsubscribe:(.|\r\n\s)+<mailto:([^>]+)>/im);
        if (urls) {
          url = parseEmail_(urls[2]);
          var subject = "Unsubscribe";
          GmailApp.sendEmail(url, subject, subject);
          status = "Unsubscribed via email";
        }
      }
      
      if (status.match(/unsubscribed/i)) {
        UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      }
      
      formula = hyperlink.replace("#LINK", threads[t].getPermalink());
      
      log_( status, message.getSubject(), formula, message.getFrom(), url );
      
    }
  } catch (e) {Logger.log(e.toString())}
}


function parseEmail_(email) {
  var result = email.trim().split("?");
  return result[0];
}


function saveConfig() {
  try {    
    stop_(true);
    
    ScriptApp.newTrigger('doGmail')
    .timeBased().everyMinutes(5).create();
    
    return "The Gmail Subscription Wipe will begin in 5 minutes";    
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}
