/**
 * @OnlyCurrentDoc
 */
//
var rawSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Data");
var cleanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Clean Pool");
var quarantinedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quarantined");
var finalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Final");
var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails");
var dirtySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dirty");
var curatedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Curated");
var referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reference");
var deadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cemetery");
var apiId = 'pmathur8';
var apiKey = 'WuOr5KWvrfFjIFY8aUZqI58W7bko3Dcip1mkA7ojMHcs3yqE';

var curated = []; //array to store awesome thankagrams

//Faux global variables
//B3 -  markerVolunteer import - used in volunteer spreadsheetsheets

function onOpen() {
	// Display a sidebar with custom HtmlService content.
	var htmlOutput = HtmlService
		.createHtmlOutput('<p>Data about submissions,errors,mails sent etc here</p>')
		.setTitle('Dashboard');
	//SpreadsheetApp.getUi().showSidebar(htmlOutput);
	//Add menu 
	SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
		.createMenu('Thankagram')
		.addItem('Profanity Check', 'profanityCheck')
		.addSeparator()
		.addItem('Push the reviewed dirty entries', 'review')
        .addSeparator()
        .addItem('Push the reviewed quarantined entries', 'quarantine_reviewed')
        .addSeparator()
		.addItem('Mail Merge', 'email')
		.addToUi();
}

//Profanity check

function profanityCheck() {
    Logger.log("Start");
    var batchSize = 20;
    var num_entries_to_process;
    var markerProfanityFilter = referenceSheet.getRange("B2");
    var startRow = parseInt(markerProfanityFilter.getValues());
    var lastRowTemp = lastRow(rawSheet);
    var num_remaining_entries = lastRowTemp - startRow + 1;
	if (startRow != (lastRowTemp + 1)) {
        if (num_remaining_entries >= batchSize) {
          num_entries_to_process = batchSize;
        } else {
          num_entries_to_process = num_remaining_entries;
        }
		var headers = rawSheet.getDataRange().getValues().shift();
		var headerIndexes = indexifyHeaders(headers);
		var range = rawSheet.getRange(startRow, 1, num_entries_to_process, lastCol(rawSheet)).getValues();
		var data = objectifyData(headerIndexes, range);
		var apiErrorEntries = [];
		//Logger.log(data);
      
		data.forEach(function (row) {
			var message = row['message'];
			var url = 'https://neutrinoapi.com/bad-word-filter?' + "user-id=" + apiId + "&api-key=" + apiKey + "&content=" + encodeURIComponent(message); //api endpoint as a string 
			var response = UrlFetchApp.fetch(url, {
				'muteHttpExceptions': true
			}); // get api endpoint
			var metaData = response.getResponseCode(); //Get the HTTP status code (200 for OK, etc.)
			if (metaData == 403 || metaData == 500) {
                apiErrorEntries.push(row);
                row['profanityCheck'] = 'ERROR';
				row['errorCode'] = "Forbidden/Internal Server Error";
                //send an email alert to me
                GmailApp.sendEmail('rachitth@umd.edu', '403 Error',Utilities.formatDate(new Date(), "EST", "MM-dd-yy HH:mm:ss") );
				return;
			}
			var json = response.getContentText(); // get the response content as text
			var profanityResult = JSON.parse(json); //parse text into json
			//If error add to the error array - appending to last row
			if (profanityResult['api-error'] > 0) {
				apiErrorEntries.push(row);
				row['profanityCheck'] = 'ERROR';
				row['errorCode'] = profanityResult['api-error'] + " - " + profanityResult['api-error-msg'];
                //send an email alert to me
                GmailApp.sendEmail('rachitth@umd.edu', 'Error 1234',Utilities.formatDate(new Date(), "EST", "MM-dd-yy HH:mm:ss") );
				return;
			}
			row['profanityCheck'] = profanityResult['is-bad'];
			row['wordsDetected'] = profanityResult['bad-words-list'];
			row['errorCode'] = "";
		});
        //Logger.log(data);
        
        // Change Bg color of analysed rows
        rawSheet.getRange(startRow, 1, num_entries_to_process, headerIndexes['sender_email']+1).setBackground("Cornsilk ");
        rawSheet.getRange(startRow,1).setNote(Utilities.formatDate(new Date(), "EST", "MM-dd-yy HH:mm:ss"));//Added timestamp in note 
        
        //Update marker
        startRow += num_entries_to_process; 
		markerProfanityFilter.setValue(startRow);

		//Push the analysed entries to the cleanPool and dirty sheets:
        var cleanPool = objectsToRange(data.filter(function (row) {
			return row['profanityCheck'] == false;
		}),cleanSheet.getRange(lastRow(cleanSheet) + 1, 1));
        
		var lastRowDS = lastRow(dirtySheet); // get the last row of the dirty sheet       
        var dirty = objectsToRange(data.filter(function (row) {
			return row['profanityCheck'] == true;
		}),dirtySheet.getRange(lastRow(dirtySheet) + 1, 1));

        if (dirty){
          //Add validation
          dirtySheet.getRange(lastRowDS + 1, headerIndexes['flag']+1, dirty.length).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['GOOD', 'BAD', 'WOW'], true).build());
        }
        
        //Add entries with error to the end of the list in Raw sheet
        if (apiErrorEntries.length != 0){
           objectsToRange(apiErrorEntries,rawSheet.getRange(lastRowTemp+1, 1,apiErrorEntries.length,headerIndexes['sender_email']+1));
        }
        //Logger.log(apiErrorEntries);
    }
    Logger.log("End");
}
//Review and push
function review() {
	var headers = dirtySheet.getDataRange().getValues().shift();
	var headerIndexes = indexifyHeaders(headers);
	var range = dirtySheet.getRange(2, 1, (lastRow(dirtySheet) - 1), lastCol(dirtySheet)).getValues();
	var data = objectifyData(headerIndexes, range);
   
    var pushToFinal = objectsToRange (data.filter(function (row) {
      return ((row['flag'] == "GOOD") || (row['flag'] == "WOW"));
    }),finalSheet.getRange(lastRow(finalSheet) + 1, 1));
    
    var pushToQuarantine = objectsToRange (data.filter(function (row) {
		return row['flag'] == 'BAD';
	}),quarantinedSheet.getRange(lastRow(quarantinedSheet) + 1, 1));
    
    var pushToCurated = objectsToRange (data.filter(function (row) {
		return (row['flag'] == 'WOW');
    }),curatedSheet.getRange(lastRow(curatedSheet) + 1, 1));
    
	//clear the dirty sheet and paste entries that were not reviewd
    dirtySheet.getRange(2, 1, lastRow(dirtySheet) - 1,lastCol(dirtySheet)).clear();
    var retainInDirty = objectsToRange (data.filter(function (row) {
		return (row['flag'] != "BAD" && row['flag'] != "GOOD" && row['flag'] != "WOW");
	}),dirtySheet.getRange(lastRow(dirtySheet) + 1, 1));
    //Add data validation to 'flag' column
    if (retainInDirty){
      dirtySheet.getRange(2, headerIndexes['flag']+1, retainInDirty.length).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['GOOD', 'BAD', 'WOW'], true).build());
    };    
}

//Review and push
function quarantine_reviewed() {
	var headers = quarantinedSheet.getDataRange().getValues().shift();
	var headerIndexes = indexifyHeaders(headers);
	var range = quarantinedSheet.getRange(2, 1, (lastRow(quarantinedSheet) - 1), lastCol(quarantinedSheet)).getValues();
	var data = objectifyData(headerIndexes, range);
  
  
    var pushToFinal = objectsToRange (data.filter(function (row) {
		return ((row['flag'] == "GOOD") || (row['flag'] == "WOW"));
	}),finalSheet.getRange(lastRow(finalSheet) + 1, 1));
    
    //To avoid duplicate reviewing by Scott, these thank-a-grams will not be sent.
    var pushToCemetary = objectsToRange (data.filter(function (row) {
		return row['flag'] == 'BAD';
	}),deadsheet.getRange(lastRow(deadsheet) + 1, 1));
    
    var pushToCurated = objectsToRange (data.filter(function (row) {
		return row['flag'] == 'WOW';
    }),curatedSheet.getRange(lastRow(curatedSheet) + 1, 1));
   
	//clear the quarantined sheet and paste entries that were not reviewd
    quarantinedSheet.getRange(2, 1, lastRow(quarantinedSheet) - 1,lastCol(quarantinedSheet)).clear();
    var retainInQuarantine = objectsToRange (data.filter(function (row) {
		return (row['flag'] != "BAD" && row['flag'] != "GOOD" && row['flag'] != "WOW");
	}),quarantinedSheet.getRange(lastRow(quarantinedSheet) + 1, 1));
    //Add data validation to 'flag' column
    if (retainInQuarantine){
      quarantinedSheet.getRange(2, headerIndexes['flag']+1, retainInQuarantine.length).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['GOOD', 'BAD', 'WOW'], true).build());
    };    
}


function email(){
    Logger.log("Start");
    var markerEmail = referenceSheet.getRange("B4");
    var startRow = parseInt(markerEmail.getValues());
    var lastRowFinalTemp = lastRow(finalSheet);
    var emailBatchSize = 50;
    var num_entries_to_process;
    var num_remaining_entries = lastRowFinalTemp - startRow + 1;
    
    if (startRow != (lastRowFinalTemp + 1)){
        var headerIndexesEmail = indexifyHeaders(emailSheet.getDataRange().getValues().shift());
        var rangeEmail = emailSheet.getRange(2, 1, (lastRow(emailSheet) - 1), lastCol(emailSheet)).getValues();
        var dataEmail = objectifyData(headerIndexesEmail, rangeEmail);
        
        if (num_remaining_entries >= emailBatchSize) {
          num_entries_to_process = emailBatchSize;
        } else {
          num_entries_to_process = num_remaining_entries;
        }
        var headers = finalSheet.getDataRange().getValues().shift();
        var headerIndexes = indexifyHeaders(headers);
        var range = finalSheet.getRange(startRow, 1, num_entries_to_process, lastCol(finalSheet)).getValues();
        var data = objectifyData(headerIndexes, range);
        Logger.log(data.length);
     
        var unique = [];
        
        //Image URLs for the Email template
        var postCardHeaderUrl = "https://tltc.umd.edu/sites/default/files/2022-04/Left%20upper%20corner.png";
        var postCardLogoUrl = "https://tltc.umd.edu/sites/default/files/2022-04/right%20upper%20corner.png";
        var emailHeaderUrl = "https://tltc.umd.edu/sites/default/files/2022-04/Banner%20type%203.png";  
        
        // Fetch images as blobs, set names for attachments
        var postCardHeader = UrlFetchApp
                                .fetch(postCardHeaderUrl)
                                .getBlob()
                                .setName("postCardHeader");
        var postCardLogo = UrlFetchApp
                                .fetch(postCardLogoUrl)
                                .getBlob()
                                .setName("postCardLogo");
        var emailHeader = UrlFetchApp
                                .fetch(emailHeaderUrl)
                                .getBlob()
                                .setName("emailHeader");
        //consolidate mutiple thankagrams addressed to the same teacher
        //Get a list of all unique teacher ids 
        data.forEach(function(row){
          unique.push(row['teacher_id']);
        })
        Logger.log(unique);
        unique = ArrayLib.unique(unique);//ArrayLib Library
        Logger.log(unique);
        //For each teacher(id), find all thankagrams 
        unique.forEach(function(each){
          //find the teacher's email from the id
          var teacherEmailRow = dataEmail.filter(function(entry){
                 return (entry['id'] == each);
             }); 
          var teacherEmail = datifyObjects(teacherEmailRow)[0][1];
          Logger.log(teacherEmail);
          var HTML = "<!DOCTYPE html><html><head><title>Thank-a-gram</title><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\" /><style type=\"text/css\">/* CLIENT-SPECIFIC STYLES */body,table,td,a {-webkit-text-size-adjust: 100%;-ms-text-size-adjust: 100%;}/* Prevent WebKit and Windows mobile changing default text sizes */table,td {mso-table-lspace: 0pt;mso-table-rspace: 0pt;}/* Remove spacing between tables in Outlook 2007 and up */img {-ms-interpolation-mode: bicubic;}/* Allow smoother rendering of resized image in Internet Explorer *//* RESET STYLES */img {border: 0;height: auto;line-height: 100%;outline: none;text-decoration: none;}table {border-collapse: collapse !important;}body {height: 100% !important;margin: 0 !important;padding: 0 !important;width: 100% !important;}/* iOS BLUE LINKS */a[x-apple-data-detectors] {color: inherit !important;text-decoration: none !important;font-size: inherit !important;font-family: inherit !important;font-weight: inherit !important;line-height: inherit !important;}/* MOBILE STYLES */@media screen and (max-width: 525px) {/* ALLOWS FOR FLUID TABLES */.wrapper {width: 100% !important;max-width: 100% !important;}/* ADJUSTS LAYOUT OF LOGO IMAGE */.logo img {margin: 0 auto !important;}/* USE THESE CLASSES TO HIDE CONTENT ON MOBILE */.mobile-hide {display: none !important; max-height:0px !important; overflow:hidden !important;}.mobile-show {display: block !important;max-height: none !important;overflow: visible !important;}.img-max {max-width: 100% !important;width: 100% !important;height: auto !important;}/* FULL-WIDTH TABLES */.responsive-table {width: 100% !important;}/* UTILITY CLASSES FOR ADJUSTING PADDING ON MOBILE */.padding {padding: 10px 5% 15px 5% !important;}.padding-meta {padding: 30px 5% 0px 5% !important;text-align: center;}.padding-copy {padding: 10px 5% 10px 5% !important;text-align: center;}.no-padding {padding: 0 !important;}.section-padding {padding: 50px 15px 50px 15px !important;}/* ADJUST BUTTONS ON MOBILE */.mobile-button-container {margin: 0 auto;width: 100% !important;}.mobile-button {padding: 15px !important;border: 0 !important;font-size: 16px !important;display: block !important;}}/* ANDROID CENTER FIX */div[style*=\"margin: 16px 0;\"] {margin: 0 !important;}</style></head><body style=\"margin: 0 !important; padding: 0 !important;\"><!-- HIDDEN PREHEADER TEXT --><div style=\"display: none; font-size: 1px; color: #fefefe; line-height: 1px; font-family: Helvetica, Arial, sans-serif; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden;\">I would like to thank you for ... </div><!-- HEADER --><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\"><tr><td bgcolor=\"#C71A30\" align=\"center\" style=\"padding: 15px;\"></td></tr><tr><td bgcolor=\"#ffffff\" align=\"center\"><!--[if (gte mso 9)|(IE)]>            <table align=\"center\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"500\">            <tr>            <td align=\"center\" valign=\"top\" width=\"500\">            <![endif]--><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"max-width: 640px;\" class=\"wrapper\"><tr><td align=\"center\" valign=\"top\" style=\"padding: 15px 0;\" class=\"logo\"><img class=\"mobile-hide\" alt=\"Logo\" src=\"cid:header\" style=\"display: block; font-family: Helvetica, Arial, sans-serif; color: #ffffff; font-size: 16px;\" border=\"0\"></td></tr></table><!--[if (gte mso 9)|(IE)]>            </td>            </tr>            </table>            <![endif]--></td></tr><tr><td bgcolor=\"#ffffff\" align=\"center\"><!--[if (gte mso 9)|(IE)]>            <table align=\"center\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"500\">            <tr>            <td align=\"center\" valign=\"top\" width=\"500\">            <![endif]--><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"70%\" style=\"display:none; max-width:640px;\" class=\"mobile-show\"><tr><td align=\"center\" valign=\"top\" style=\"padding: 15px 0;\" class=\"logo\"><img alt=\"Logo\" src=\"cid:pcLogo\" style=\"display: block; font-family: Helvetica, Arial, sans-serif; color: #ffffff; font-size: 16px; width:100%\" border=\"0\"></td></tr></table><!--[if (gte mso 9)|(IE)]>            </td>            </tr>            </table>            <![endif]--></td></tr><tr><td bgcolor=\"#ffffff\" align=\"center\" style=\"padding: 15px;\"><!--[if (gte mso 9)|(IE)]>            <table align=\"center\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"500\">            <tr>            <td align=\"center\" valign=\"top\" width=\"500\">            <![endif]--><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"max-width: 640px;\" class=\"responsive-table\"><tr><td><!-- COPY --><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\"><tr><td align=\"center\" style=\"font-size: 32px; font-family: Helvetica, Arial, sans-serif; color: #333333; padding-top: 30px;\" class=\"padding-copy\">You’ve Received a Thank-a-Gram!</td></tr></table></td></tr></table><!--[if (gte mso 9)|(IE)]>            </td>            </tr>            </table>            <![endif]--></td></tr><tr><td bgcolor=\"#ffffff\" align=\"center\" style=\"padding: 15px;\" class=\"padding\"><!--[if (gte mso 9)|(IE)]>            <table align=\"center\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"500\">            <tr>            <td align=\"center\" valign=\"top\" width=\"500\">            <![endif]-->"
          var count = 0;
          var rowNumbers =[];
          // Get message,sender name and email for each row and consolidate
          data.forEach(function(row,i){
             if (row['teacher_id'] == each){
               count ++;
               var courseName = String(row['course']);
               var courseCode = courseName.slice(0, courseName.indexOf(" "));
               courseCode=courseName;
               //Logger.log(courseCode);
               HTML += "<div class=\"mobile-hide\"><table style=\"border-collapse:collapse;table-layout:fixed;width:640px;height:426px;margin:25px auto;\" border=\"0\" width=\"640\" height=\"426\"><tbody><tr><td style=\"vertical-align:top;padding:0px;width:640px;height:426px;\" colspan=\"1\" rowspan=\"1\"><div style=\"display:block;background-color:#FFFFFF;border:2px solid #d7d7d7;color:#ffffff;width:638px;height:424px;\"><table style=\"border-collapse:collapse;table-layout:fixed;width:640px;height:426px;margin:auto;\" border=\"0\" width=\"640\" height=\"426\"><colgroup><col style=\"width:13px;\"><col style=\"width:2px;\"><col style=\"width:35px;\"><col style=\"width:6px;\"><col style=\"width:227px;\"><col style=\"width:39px;\"><col style=\"width:33px;\"><col style=\"width:12px;\"><col style=\"width:2px;\"><col style=\"width:19px;\"><col style=\"width:18px;\"><col style=\"width:6px;\"><col style=\"width:7px;\"><col style=\"width:6px;\"><col style=\"width:4px;\"><col style=\"width:5px;\"><col style=\"width:7px;\"><col style=\"width:7px;\"><col style=\"width:6px;\"><col style=\"width:158px;\"><col style=\"width:10px;\"><col style=\"width:2px;\"><col style=\"width:16px;\"></colgroup><tbody><tr><td colspan=\"23\" style=\"width:640px;height:12px\">­</td></tr><tr><td colspan=\"1\" style=\"width:13px;height:1px\">­</td><td style=\"vertical-align:top;padding:0px;width:270px;height:85px;\" colspan=\"4\" rowspan=\"2\"><div style=\"display:block;width:270px;height:85px;\"><img src=\"cid:pcHeader\" style=\"display:block;\" width=\"270\" height=\"85\" alt=\"header\"></div></td><td colspan=\"18\" style=\"width:357px;height:1px\">­</td></tr><tr><td colspan=\"1\" style=\"width:13px;height:84px\">­</td><td colspan=\"10\" style=\"width:146px;height:84px\">­</td><td style=\"vertical-align:top;padding:0px;width:183px;height:92px;\" colspan=\"5\" rowspan=\"2\"><div style=\"display:block;width:183px;height:92px;\"><img src=\"cid:pcLogo\" style=\"display:block;\" width=\"183\" height=\"92\" alt=\"stamp\"></div></td><td colspan=\"3\" style=\"width:28px;height:84px\">­</td></tr><tr><td colspan=\"15\" style=\"width:429px;height:8px\">­</td><td colspan=\"3\" style=\"width:28px;height:8px\">­</td></tr><tr><td colspan=\"23\" style=\"width:640px;height:23px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:33px\">­</td><td style=\"vertical-align:top;padding:0px;width:340px;height:200px;\" colspan=\"5\" rowspan=\"16\"><div style=\"display:block;font-family:TimesNewRomanPS-ItalicMT;font-size:16px;color:#000000;letter-spacing:0.05px;line-height:20px;text-align:left;width:340px;height:200px;\"><span style=\"font-family:'Arial';font-size:16px;color:#000000;opacity:1;\">"+row['message']+"</span></div></td><td colspan=\"1\" style=\"width:12px;height:33px\">­</td><td style=\"vertical-align:top;padding:0px;width:2px;height:271px;\" colspan=\"1\" rowspan=\"21\"><div style=\"display:block;border:1px solid #D1D1D1;color:#f2f2f2;width:0px;height:269px;\"></div></td><td colspan=\"1\" style=\"width:19px;height:33px\">­</td><td style=\"vertical-align:top;padding:0px;width:236px;height:33px;\" colspan=\"12\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:10px;color:#767C7C;letter-spacing:0.03px;text-align:left;width:236px;height:33px;\"><span style=\"font-family:'Arial';font-size:10px;color:#757b7b;opacity:1;\">The TLTC invites everyone to recognize the excellence in teaching that exemplifies the University of Maryland</span></div></td><td colspan=\"1\" style=\"width:16px;height:33px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:3px\">­</td><td colspan=\"1\" style=\"width:12px;height:3px\">­</td><td colspan=\"14\" style=\"width:271px;height:3px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:15px\">­</td><td colspan=\"1\" style=\"width:12px;height:15px\">­</td><td colspan=\"1\" style=\"width:19px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:236px;height:15px;\" colspan=\"12\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#767C7C;letter-spacing:0.04px;text-align:left;width:236px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#757b7b;opacity:1;\">Visit </span><span style=\"font-family:'Arial';font-size:13px;color:#bc091d;opacity:1;\">go.umd.edu/thanks</span><span style=\"font-family:'Arial';font-size:13px;color:#757b7b;opacity:1;\"> to learn more.</span></div></td><td colspan=\"1\" style=\"width:16px;height:15px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:20px\">­</td><td colspan=\"1\" style=\"width:12px;height:20px\">­</td><td colspan=\"14\" style=\"width:271px;height:20px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:15px\">­</td><td colspan=\"1\" style=\"width:12px;height:15px\">­</td><td colspan=\"1\" style=\"width:19px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:18px;height:15px;\" colspan=\"1\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#767C7C;letter-spacing:0.04px;text-align:left;width:18px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#757b7b;opacity:1;\">To:</span></div></td><td colspan=\"1\" style=\"width:6px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:210px;height:15px;\" colspan=\"9\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#000000;letter-spacing:0.04px;text-align:left;width:210px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#000000;opacity:1;\">"+row['teacher']+"</span></div></td><td colspan=\"2\" style=\"width:18px;height:15px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:3px\">­</td><td colspan=\"1\" style=\"width:12px;height:3px\">­</td><td colspan=\"14\" style=\"width:271px;height:3px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:2px\">­</td><td colspan=\"1\" style=\"width:12px;height:2px\">­</td><td colspan=\"3\" style=\"width:43px;height:2px\">­</td><td style=\"vertical-align:top;padding:0px;width:212px;height:2px;\" colspan=\"10\" rowspan=\"1\"><div style=\"display:block;border:1px solid #D1D1D1;color:#f2f2f2;width:210px;height:0px;\"></div></td><td colspan=\"1\" style=\"width:16px;height:2px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:25px\">­</td><td colspan=\"1\" style=\"width:12px;height:25px\">­</td><td colspan=\"14\" style=\"width:271px;height:25px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:15px\">­</td><td colspan=\"1\" style=\"width:12px;height:15px\">­</td><td colspan=\"1\" style=\"width:19px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:46px;height:15px;\" colspan=\"6\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#767C7C;letter-spacing:0.04px;text-align:left;width:46px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#757b7b;opacity:1;\">Course:</span></div></td><td colspan=\"1\" style=\"width:7px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:183px;height:15px;\" colspan=\"5\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#000000;letter-spacing:0.04px;text-align:left;width:183px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#000000;opacity:1;\">"+courseCode+"</span></div></td><td colspan=\"1\" style=\"width:16px;height:15px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:2px\">­</td><td colspan=\"1\" style=\"width:12px;height:2px\">­</td><td colspan=\"14\" style=\"width:271px;height:2px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:1px\">­</td><td colspan=\"1\" style=\"width:12px;height:1px\">­</td><td colspan=\"8\" style=\"width:72px;height:1px\">­</td><td style=\"vertical-align:top;padding:0px;width:183px;height:1px;\" colspan=\"5\" rowspan=\"1\"><div style=\"display:block;border:1px solid #D1D1D1;color:#f2f2f2;width:181px;height:-1px;\"></div></td><td colspan=\"1\" style=\"width:16px;height:1px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:25px\">­</td><td colspan=\"1\" style=\"width:12px;height:25px\">­</td><td colspan=\"14\" style=\"width:271px;height:25px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:15px\">­</td><td colspan=\"1\" style=\"width:12px;height:15px\">­</td><td colspan=\"1\" style=\"width:19px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:60px;height:15px;\" colspan=\"8\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#767C7C;letter-spacing:0.04px;text-align:left;width:60px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#757b7b;opacity:1;\">Semester:</span></div></td><td colspan=\"1\" style=\"width:6px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:170px;height:15px;\" colspan=\"3\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#000000;letter-spacing:0.04px;text-align:left;width:170px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#000000;opacity:1;\">"+row['semester']+"</span></div></td><td colspan=\"1\" style=\"width:16px;height:15px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:2px\">­</td><td colspan=\"1\" style=\"width:12px;height:2px\">­</td><td colspan=\"14\" style=\"width:271px;height:2px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:3px\">­</td><td colspan=\"1\" style=\"width:12px;height:3px\">­</td><td colspan=\"10\" style=\"width:85px;height:3px\">­</td><td style=\"vertical-align:top;padding:0px;width:170px;height:3px;\" colspan=\"3\" rowspan=\"1\"><div style=\"display:block;border:1px solid #D1D1D1;color:#f2f2f2;width:168px;height:0px;\"></div></td><td colspan=\"1\" style=\"width:16px;height:3px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:2px\">­</td><td colspan=\"1\" style=\"width:12px;height:2px\">­</td><td colspan=\"14\" style=\"width:271px;height:2px\">­</td></tr><tr><td colspan=\"8\" style=\"width:367px;height:23px\">­</td><td colspan=\"14\" style=\"width:271px;height:23px\">­</td></tr><tr><td colspan=\"2\" style=\"width:15px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:35px;height:15px;\" colspan=\"1\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#767C7C;letter-spacing:0.04px;text-align:left;width:35px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#757b7b;opacity:1;\">From:</span></div></td><td colspan=\"1\" style=\"width:6px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:266px;height:15px;\" colspan=\"2\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#000000;letter-spacing:0.04px;text-align:left;width:266px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#000000;opacity:1;\"><a href=\"mailto:"+row['sender_email']+"\" target=\"_blank\">"+row['sender_name']+"</a></span></div></td><td colspan=\"2\" style=\"width:45px;height:15px\">­</td><td colspan=\"1\" style=\"width:19px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:31px;height:15px;\" colspan=\"3\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#767C7C;letter-spacing:0.04px;text-align:left;width:31px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#757b7b;opacity:1;\">Year:</span></div></td><td colspan=\"1\" style=\"width:6px;height:15px\">­</td><td style=\"vertical-align:top;padding:0px;width:199px;height:15px;\" colspan=\"8\" rowspan=\"1\"><div style=\"display:block;font-family:ArialMT;font-size:13px;color:#000000;letter-spacing:0.04px;text-align:left;width:199px;height:15px;\"><span style=\"font-family:'Arial';font-size:13px;color:#000000;opacity:1;\">"+row['year']+"</span></div></td><td colspan=\"1\" style=\"width:16px;height:15px\">­</td></tr><tr><td colspan=\"8\" style=\"width:367px;height:3px\">­</td><td colspan=\"14\" style=\"width:271px;height:3px\">­</td></tr><tr><td colspan=\"4\" style=\"width:56px;height:3px\">­</td><td style=\"vertical-align:top;padding:0px;width:266px;height:3px;\" colspan=\"2\" rowspan=\"1\"><div style=\"display:block;border:1px solid #D1D1D1;color:#f2f2f2;width:264px;height:0px;\"></div></td><td colspan=\"2\" style=\"width:45px;height:3px\">­</td><td colspan=\"5\" style=\"width:56px;height:3px\">­</td><td style=\"vertical-align:top;padding:0px;width:199px;height:3px;\" colspan=\"8\" rowspan=\"1\"><div style=\"display:block;border:1px solid #D1D1D1;color:#f2f2f2;width:197px;height:0px;\"></div></td><td colspan=\"1\" style=\"width:16px;height:3px\">­</td></tr><tr><td colspan=\"8\" style=\"width:367px;height:27px\">­</td><td colspan=\"14\" style=\"width:271px;height:27px\">­</td></tr></tbody></table></div></td></tr></tbody></table></div><div style=\"display:none;margin-bottom:30px;\" class=\"mobile-show\"><table cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"display:none; max-width: 500px;border:2px solid #d7d7d7\" class=\"responsive-table mobile-show\"><tbody><tr><td><!-- HERO IMAGE --><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\"><tbody><tr><td class=\"padding\" align=\"center\"><img src=\"cid:pcHeader\" border=\"0\" style=\"display: block; color: #666666;  font-family: Helvetica, arial, sans-serif; font-size: 16px;\" class=\"img-max\"></td></tr><tr><td><!-- COPY --><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\"><tbody><tr><td align=\"left\" style=\"padding: 20px 0 0 0; font-size: 16px; font-family: 'Arial'; line-height:25px; color: #666666;\" class=\"padding\">"+row['message']+"</td></tr></tbody></table></td></tr><tr><td align=\"center\"><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\"><tbody><tr><td align=\"right\" style=\"padding-top: 25px;\" class=\"padding\"><table border=\"0\" cellspacing=\"0\" cellpadding=\"0\" class=\"mobile-button-container\"><tbody><tr><td align=\"right\" style=\"font-size: 16px; font-family: Helvetica, Arial, sans-serif; padding-bottom:10px;\">- <a href=\"mailto:"+row['sender_email']+"\" target=\"_blank\">"+row['sender_name']+"</a></td></tr><tr><td align=\"right\" style=\"font-size: 14px; font-family: Helvetica, Arial, sans-serif;\">"+courseCode+"&nbsp;"+row['semester']+"&nbsp;"+row['year']+"</td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></div><!--[if (gte mso 9)|(IE)]>            </td>            </tr>            </table>            <![endif]-->";
               //"<div class=\" \">"+row['message']+"<br>"+row['sender_name']+"<br>"+row['sender_email']+"</div>";
               rowNumbers.push(i);
             }
          });
          HTML += "</td></tr><tr><td bgcolor=\"#ffffff\" align=\"center\" style=\"padding: 15px;\"><!--[if (gte mso 9)|(IE)]>            <table align=\"center\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"500\">            <tr>            <td align=\"center\" valign=\"top\" width=\"500\">            <![endif]--><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"max-width: 640px;\" class=\"responsive-table\"><tr><td><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\"><tr><td><!-- COPY --><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\"><tr><td align=\"left\" style=\"padding: 0 0 20px 0; text-align: left; font-size: 16px; line-height: 25px; font-family: Helvetica, Arial, sans-serif; color: #666666;\" class=\"padding-copy\">During Thank a Teacher Week (May 2-6, 2022), students, alumni, and the whole campus community are invited to reflect on their experience and collectively recognize the excellence in teaching that exemplifies the University of Maryland.</td></tr><tr><td align=\"left\" style=\"padding: 0 0 0 0; font-size: 14px; line-height: 18px; font-family: Helvetica, Arial, sans-serif; color: #aaaaaa; font-style: italic; text-align: left \">The contents of the Thank-a-Gram were submitted through an online form. TLTC is doing its best to screen submissions, but the submitter is solely responsible for its content. If you do not wish to receive any more Thank-a-Grams, please let us know via email at tltc@umd.edu.</td></tr></table></td></tr></table></td></tr></table><!--[if (gte mso 9)|(IE)]>            </td>            </tr>            </table>            <![endif]--></td></tr><tr><td bgcolor=\"#ffffff\" align=\"center\" style=\"padding: 20px 0px;\"><!--[if (gte mso 9)|(IE)]>            <table align=\"center\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"500\">            <tr>            <td align=\"center\" valign=\"top\" width=\"500\">            <![endif]--><!-- UNSUBSCRIBE COPY --><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" align=\"center\" style=\"max-width: 640px;\" class=\"responsive-table\"><tr><td align=\"center\" style=\"font-size: 12px; line-height: 18px; font-family: Helvetica, Arial, sans-serif; color:#666666;\">Teaching and Learning Transformation Center<br><a href=\"http://litmus.com\" target=\"_blank\" style=\"color: #666666; text-decoration: none;\">University of Maryland,College Park</a></td></tr></table><!--[if (gte mso 9)|(IE)]>            </td>            </tr>            </table>            <![endif]--></td></tr></table></body></html>";

          //Send emails
          GmailApp.sendEmail(teacherEmail,'You’ve Received a Thank-a-Gram!','Your students send you thanks',
          {
          from:(GmailApp.getAliases())[0],
          name:'TLTC Thank-a-Gram',
          htmlBody:HTML,
            inlineImages:
            {
              pcLogo: postCardLogo,
              pcHeader: postCardHeader,
              header: emailHeader
            }
          }); 
    
          rowNumbers.forEach(function(value){
            data[value]['teacher email'] = teacherEmail;
            data[value]['sent time'] = Utilities.formatDate(new Date(), "EST", "MM-dd-yy HH:mm:ss");
            data[value]['email status'] = "sent";
          });
          
        
        })
        //update Final sheet by adding meta data
        objectsToRange(data,finalSheet.getRange(startRow, 1));
        
        //Update marker
        Logger.log(startRow);
        startRow += num_entries_to_process;
        markerEmail.setValue(startRow);
        Logger.log("End");
        
    } else {
        SpreadsheetApp.getUi().alert('No more entries to mail, please try after some time'); 
    }
};

//Dashboard
function dashBoard(){
    var emailCountData = finalSheet.getRange("T:T").getValues();
    var emailCount = ArrayLib.filterByText(emailCountData, 0, "sent").length;
    var thankagramCount = lastRow(rawSheet);
    var quarCount = lastRow(quarantinedSheet);
    var reviewCountData = cleanSheet.getRange("O:O").getValues();
    var reviewCount = ArrayLib.filterByText(reviewCountData, 0, "Reviewed").length;
}

// Functions for general data manipulation
/**
 * Get the last row of a sheet
 */
function lastRow(sheet) {
	return sheet.getDataRange().getLastRow();
}

function lastCol(sheet) {
	return sheet.getDataRange().getLastColumn();
}
/**
 * write objectified data to a sheet starting at given range
 * @param {[object]} dataObjects an array of objectified sheet data
 * @param {Range} range a starting range
 * @return {[[]]} a data array in sheet format
 */
function objectsToRange(dataObjects, range) {
	var newData = datifyObjects(dataObjects);
	var sheet = range.getSheet();
    if (newData){
      var newRange = sheet.getRange(
			range.getRow(), range.getColumn(), newData.length, newData[0].length
		)
		.setValues(newData);

      return newData;
    }
    else {
      return false;
    };
}

/**
 * turn objectifed data into sheet writable data
 * @param {[object]} dataObjects an array of objectified sheet data
 * @return {[[]]} a data array in sheet format
 */
function datifyObjects(dataObjects) {
  if (typeof(dataObjects[0]) != "undefined") {
	// get the headers from a row of objects
	var headers = Object.keys(dataObjects[0]);

	// turn the data back to an array and concat to header
	return (dataObjects.map(function (row) {
		return headers.map(function (cell) {
			return row[cell];
		});
	}));
  }
  else {
    return false;
  };

}

/**
 * given a range and property (eg values, backgrounds)
 * get the data and make it into an object
 * @param {Range} range the range
 * @return {[object]} the array of objects
 */
function objectifyRange(range) {
	// get the data from the range

	var data = range.getValues();

	// extract out the headers and data and objectify
	return objectifyData(
		indexifyHeaders(data.slice(0, 1)[0]), data.slice(1)
	);
}
/**
 * create an array of objects from data
 * @param {object} headerIndexes the map of header names to column indexes
 * @param {[[]]} data the data from the sheet
 * @return {[object]} the objectified data
 */
function objectifyData(headerIndexes, data) {
	return data.map(function (row) {
		return Object.keys(headerIndexes).reduce(function (p, c) {
			p[c] = row[headerIndexes[c]];
			return p;
		}, {});
	});
}
/**
 * create a map of indexes to properties
 * @param {[*]} headers an array of header names
 * @return {object} an object where the props are names & values are indexes
 */
function indexifyHeaders(headers) {

	var index = 0;
	return headers.reduce(function (p, c) {

		// skip columns with blank headers   
		if (c) {
			// for this to work, cant have duplicate column names
			if (p.hasOwnProperty(c)) {
				throw new Error('duplicate column name ' + c);
			}
			p[c] = index;
		}
		index++;
		return p;
	}, {});
}