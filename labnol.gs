/**
* ================================================================
* G M A I L   U N S U B S C R I B E
* ================================================================
*
* @udpated: 31.05.2017
* @support: amit@labnol.org
* @see: http://www.labnol.org/internet/amazon-price-tracker/28156/
*
* @status: Refactoring / Code syntax optimization IN PROCESS; 
*          stopped 31.05.2017 at line #108 (`stop_` function)
* @todo: continue with "Code Syntax Optimization / Refactoring"
**/

/*** Kosytiantyn Didenko */
/**
* Timer Object BEGIN
*/
var Timer = (function() {
  
  var __now = function() {
      return (new Date()).getTime();
  };
  
  var __created = __now(),
      __started = 0,
      __stopped = 0;
  
  var __since = function(moment) {
    if (__stopped && (__stopped > moment)) {
      // till stopped
      return __stopped - moment;
    } else {
      // till now
      var tN = __now() - moment;      
      if (tN < 0) {        // the `moment` is in future!
        Logger.log("calculated {Timer.elapsed} for the moment of "
                   + __format(tN) + " in future!", "Result will be a negative `-` {number}");
      }
      return tN;
    }
  };
  
  // stub for returning formatted string
  //TODO: add seconds, minutes and hours formatting there.  
  var __format = function(time) {
      var _msec = time;      
      return time + " msec.";
  };
  
  var __toDateStr = function(time) {    
      return (new Date(__started)).toUTCString();
  };
  
  /**
  * @public methods go here
  */
  return {
    
    get started() {
        var message = "Stopped at: ",
            time, passed;
        if (__started) {
            time = __toDateStr(__started);
            passed = __format(__since(__started));
            message += time + "\nE.g: " + passed + " ago.";
        } else {
            time = false;
            message = "{Timer.started} property was accessed before it was actually {Timer.start()}'ed";
        }
        Logger.log(message);
        return time;
    },
      
    get stopped() {
        var message = "Stopped at: ",
            time, passed;
        if (__stopped && (__stopped >= __started)) {
            time    = __toDateStr(__started);
            passed  = __format(__since(__started));
            message = "Stopped at: " + time + "\nE.g: " + passed + " ago.";
        } else if (__stopped) {            // smth. wrong = !(__stopped > __started)             
            message = "Wrong {Timer.stopped} / {Timer.started} values!\n" +
                      "stopped = " + __stopped + "; started = " + __started;
            time = false;
        } else {            
            message = "{Timer.stopped} property was accessed before it was actually {Timer.stop()}'ed";
            time = false;
        }
        Logger.log(message);
        return (time);
    },
    
    get now() {
        return __toDateStr(__now());
    },
    
    get elapsed(label) {
        var elapsed;
        if(__started) {
            elapsed = __format(__since(__started));            
            return elapsed;
        } else {
            Logger.log("{Timer.elapsed} was accessed before it was {Timer.start}'ed");
            elapsed = __format(__since(__created));
            return elapsed;
        }
    },
    
    logTime: function(label) {
        Logger.log("Time Log: [" + label + "...] = +" + this.elapsed);
    },
    
    start: function() {
        __started = __now();
        __stopped = 0;
        Logger.log("Timer.start()'ed at " + __toDateStr(__started));
        return __started;
    },
    
    stop: function() {
        __stopped = __now();
        Logger.log(
            "Timer.stop()'ped at " + __toDateStr(__stopped) + "\n" +
            "E.g: " + __format(__since(__started)) + " after it was Timer.start()'ed"
        );
        return __stopped;
    },
    
    reset: function() {
        var time = __now();
        __stopped = 0;
        __started = 0;
        Logger.log("Timer.reset()'ed at " + __toDateStr(time));
        return time;
    }      
  };
  
})();

/**
* Timer Object END
*/

// some simple banchmarking of the script execution
// simple timing calculation
Timer.start();

/**
* @OnlyCurrentDoc
*/
function getConfig() {
    Timer.logTime("getConfig()");
    var params = {
        label: doProperty_("LABEL") || "Unsubscribe"
    };
    return params;
}

function config_() {
    Timer.logTime("config_()");
    var html = HtmlService.createHtmlOutputFromFile('config')
        .setTitle("Gmail Unsubscriber")
        .setWidth(300)
        .setHeight(200)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getActive().show(html);
}

function help_() {
    Timer.logTime("help_()");
    var html = HtmlService.createHtmlOutputFromFile('help')
        .setTitle("Google Scripts Support")
        .setWidth(350)
        .setHeight(120);
    SpreadsheetApp.getActive().show(html);
}

function createLabel_(name) {
    Timer.logTime("createLabel_('" + name + "')");
    var label = GmailApp.getUserLabelByName(name);
    return (!label) ? GmailApp.createLabel(name) : label;
}

function log_(status, subject, view, from, link) {
    Timer.logTime("log_('" + [status, subject, view, from, link].join("', '") + "')");
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getActiveSheet().appendRow(
        [status, subject, view, from, link]
    );
}

function init_() {
    Timer.logTime("init_()");
    return Browser.msgBox("The Unsubscriber was initialized. Please select the Start option from the Gmail menu to activate.");
}

function onOpen() {
    Timer.logTime("onOpen()");  
    var menu = [     
        {name: "Configure", functionName: "config_"},
         null,
        {name: "☎ Help & Support",functionName: "help_"},
        {name: "✖ Stop (Uninstall)",  functionName: "stop_"},
         null
    ];  
    SpreadsheetApp.getActiveSpreadsheet().addMenu("➪ Gmail Unsubscriber", menu);  
}

function stop_(error) {
    Timer.logTime("stop_(error)");
    if (!error) {
        var triggers = ScriptApp.getProjectTriggers();
        //TODO: pay attention to the iteration below:
        //      is `triggers` array a clone or a reference?
        //      e.g.: does `deleteTrigger()` also removes an item from `triggers` array
        //      isn't is better to use: 
        //            triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
        //      (to avoid the re-indexing issues)
        for(var i in triggers) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
        Browser.msgBox("The Gmail Unsubscriber has been stopped. You can start it anytime later."); 
    } else {
        Logger.log("An error occured when stopping the Gmail Unsubscriber:", error);
    }
}


function doProperty_(key, value) {
  Timer.logTime("doProperty_('" + [key, value].join("', '") + "')");
  var properties = PropertiesService.getUserProperties();
  
  if (value) {
    properties.setProperty(key, value);
  } else {
    return properties.getProperty(key);
  }
  
}

function doGmail() {
  Timer.logTime("doGmail()");
  try {    
    var label = doProperty_("LABEL") || "Unsubscribe";    
    var threads = GmailApp.search("label:" + label);    
    var gmail = createLabel_(label);    
    var url, urls, message, raw, body, formula, status;    
    var hyperlink = '=HYPERLINK("#LINK#", "View")';    
    var hrefs = new RegExp(/<a[^>]*href=["'](http[s]?:\/\/[^"']+)["'][^>]*>(.*?)<\/a>/gi);
    
    for (var t in threads)  {      
      url = "";      
      status = "Could not unsubscribe";      
      message = threads[t].getMessages()[0];

      Timer.logTime("doGmail() | Parsing message: " + message.getSubject());
      threads[t].removeLabel(gmail);
      raw = message.getRawContent();
      urls = raw.match(/^list\-unsubscribe:(.|\r\n\s)+<(http[s]?:\/\/[^>]+)>/im);
      
      if (urls) {
        url = urls[2];
        status = "Unsubscribed via header";
        Timer.logTime("doGmail() | Found in header: " + url);
      } else {
        body = message.getBody().replace(/\s/g, "");
        while ( (url === "") && (urls = hrefs.exec(body)) ) {
          if (urls[1].match(/unsubscribe|optout|opt\-out|remove/i) || urls[2].match(/unsubscribe|optout|opt\-out|remove/i)) {
            url = urls[1];
            status = "Unsubscribed via link";
            Timer.logTime("doGmail() | Found in body: " + url);
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
          Timer.logTime("doGmail() | Status: " + status + " URL: " + url);
        }
      }
      
      if (status.match(/unsubscribed/i)) {
        Timer.logTime("doGmail() | Status: " + status + " URL: " + url)
        var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
        Timer.logTime("doGmail() | fetched from " + url + " | Result: " + response.getContentText());
      }      
      formula = hyperlink.replace("#LINK", threads[t].getPermalink());      
      log_( status, message.getSubject(), formula, message.getFrom(), url );      
    }
  } catch (e) {Logger.log(e.toString())}  
}


function saveConfig(params) {
  Timer.logTime("saveConfig(params)");
  try {    
    doProperty_("LABEL", params.label);    
    stop_(true);    
    ScriptApp.newTrigger('doGmail')
             .timeBased()
             .everyMinutes(15)
             .create();    
    return "The Gmail unsubscriber is now active. You can apply the Gmail label " + params.label + " to any email and you'll be unsubscribed in 15 minutes. Please close this window.";    
  } catch (e) {    
    return "ERROR: " + e.toString();    
  }  
}

function parseEmail_(email) {
  Timer.logTime("parseEmail_(email)");
  var result = email.trim().split("?");
  return result[0];
}
