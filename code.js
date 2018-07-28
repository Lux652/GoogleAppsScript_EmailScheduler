//globals
var app = SpreadsheetApp;
var ss = app.getActiveSpreadsheet().getActiveSheet();
var lr = ss.getLastRow();
var lastColumn = ss.getLastColumn();

var EMAIL_SENT = "EMAIL_SENT";
var POSTAVKE = "Postavke";


function onOpen() {
    app.getUi().createMenu('Email sidebar')
        .addItem('Show sidebar', 'showSidebar')
        .addToUi();
    triggered();

}

function onInstall() {
    onOpen();
    app.getUi().createMenu('Email sidebar')
        .addItem('Show sidebar', 'showSidebar')
        .addToUi();
}

function showSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('Index').setWidth(300).setTitle('Parametri');
    app.getUi().showSidebar(html);
    triggered();
}

function doGet() {
    return HtmlService.createHtmlOutputFromFile('Index').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

//convert column name to column number
function letterToColumn(letter) {
    var column = 0,
        length = letter.length;
    for (var i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
}

//check remaining emails(daily limit - 100)
function checkRemaining() {
    var c = MailApp.getRemainingDailyQuota();
    Logger.log(c);
}




//write user input from sidebar to sheet
function writeForm(form) {

    for (var i = 1; i <= lastColumn; i++) {

        var s = ss.getRange(1, lastColumn).getValues();

        //if column w/ header name "Postavke" exists
        if (s == POSTAVKE) {
            try {

                var email = form.email;

                var date1 = form.date1;
                var date2 = form.date2;

                var day1 = form.day1;
                var day2 = form.day2;

                var subject1 = form.subject1;
                var subject2 = form.subject2;

                var body1 = form.body1;
                var body2 = form.body2;
                var body3 = form.body3;
                var body4 = form.body4;
                var body5 = form.body5;
                var body6 = form.body6;
                var body7 = form.body7;

                var lastcolumn = ss.getLastColumn();

                var range = ss.getRange(2, lastColumn);
                range.setValue(email);

                range = ss.getRange(3, lastColumn);
                range.setValue(date1);
                range = ss.getRange(4, lastColumn);
                range.setValue(date2);

                range = ss.getRange(5, lastColumn);
                range.setValue(day1);
                range = ss.getRange(6, lastColumn);
                range.setValue(day2);

                range = ss.getRange(7, lastColumn);
                range.setValue(subject1);
                range = ss.getRange(8, lastColumn);
                range.setValue(subject2);

                range = ss.getRange(9, lastColumn);
                range.setValue(body1);
                range = ss.getRange(10, lastColumn);
                range.setValue(body2);
                range = ss.getRange(11, lastColumn);
                range.setValue(body3);
                range = ss.getRange(12, lastColumn);
                range.setValue(body4);
                range = ss.getRange(13, lastColumn);
                range.setValue(body5);
                range = ss.getRange(14, lastColumn);
                range.setValue(body6);
                range = ss.getRange(15, lastColumn);
                range.setValue(body7);

                var successMsg = "Parametri uspješno postavljeni!";
                return successMsg;
            } catch (error) {
                return error.toString();
            }

            //if column w/ header name "Postavke" does not exist
        } else {
            try {
                var email = form.email;

                var day1 = form.day1;
                var day2 = form.day2;

                var date1 = form.date1;
                var date2 = form.date2;

                var subject1 = form.subject1;
                var subject2 = form.subject2;

                var body1 = form.body1;
                var body2 = form.body2;
                var body3 = form.body3;
                var body4 = form.body4;
                var body5 = form.body5;
                var body6 = form.body6;
                var body7 = form.body7;

                var lc = ss.getLastColumn() + 2;

                var range = ss.getRange(1, lc);
                range.setValue(POSTAVKE);
                range = ss.getRange(2, lc);
                range.setValue(email);

                range = ss.getRange(3, lc);
                range.setValue(date1);
                range = ss.getRange(4, lc);
                range.setValue(date2);

                range = ss.getRange(5, lc);
                range.setValue(day1);
                range = ss.getRange(6, lc);
                range.setValue(day2);

                range = ss.getRange(7, lc);
                range.setValue(subject1);
                range = ss.getRange(8, lc);
                range.setValue(subject2);

                range = ss.getRange(9, lc);
                range.setValue(body1);
                range = ss.getRange(10, lc);
                range.setValue(body2);
                range = ss.getRange(11, lc);
                range.setValue(body3);
                range = ss.getRange(12, lc);
                range.setValue(body4);
                range = ss.getRange(13, lc);
                range.setValue(body5);
                range = ss.getRange(14, lc);
                range.setValue(body6);
                range = ss.getRange(15, lc);
                range.setValue(body7);

                var successMsg = "Parametri uspješno postavljeni!";
                return successMsg;
            } catch (error) {
                return error.toString();
            }
            break;
        }
    }
}



function sendEmail() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    var emailArr = [];
    var bodyArr = [];
    var bodySize = [];

    var dueDateFirst;
    var dueDateSecond;
    var firstDay;
    var secondDay;
    var Subject1;
    var Subject2;


    for (var x = 0; x <= sheets.length; x++) {

        var s = sheets[x];

        var row = s.getLastRow();

        var column = s.getLastColumn();
        var emailColumn = column - 1;

        for (var i = 2; i <= row; i++) {

            var pMail = s.getRange(2, column).getValues().join().toUpperCase();
            var mail = letterToColumn(pMail);
            if (pMail != "") {
                var mailContainer = s.getRange(i, mail).getValues().join();
                emailArr.push(mailContainer);
            }

        }


        for (var i = 2; i < row; i++) {

            var bodySlice = bodySize.length;

            var emailSent = s.getRange(i, emailColumn).getValue();

            var pDate1 = s.getRange(3, column).getValues().join().toUpperCase();
            var pDate2 = s.getRange(4, column).getValues().join().toUpperCase();

            firstDay = s.getRange(5, column).getValue();
            secondDay = s.getRange(6, column).getValue();

            var Subject1 = s.getRange(7, column).getValue();
            var Subject2 = s.getRange(8, column).getValue();

            for (var j = 9; j < 17; j++) {

                var pBody = s.getRange(j, column).getValues().join().toUpperCase();
                var body = letterToColumn(pBody);

                if (pBody != "") {

                    bodySize.push(body);

                    var bodyContainer = s.getRange(i, body).getValues().join();
                    bodyArr.push(bodyContainer)

                    if (pDate2 != "" && secondDay != "" && Subject2 != "") {

                        var firstDate = letterToColumn(pDate1);
                        var secondDate = letterToColumn(pDate2);

                        var firstDateContainer = s.getRange(i, firstDate).getValues().join();
                        var firstDateFormatted = Utilities.formatDate(new Date(firstDateContainer), Session.getScriptTimeZone(), "dd.MM.yyyy");
                        var secondDateContainer = s.getRange(i, secondDate).getValues().join();
                        var secondDateFormatted = Utilities.formatDate(new Date(secondDateContainer), Session.getScriptTimeZone(), "dd.MM.yyyy");
                        var mailDayFirst = new Date();
                        mailDayFirst = mailDayFirst.setDate(mailDayFirst.getDate() + firstDay);
                        dueDateFirst = Utilities.formatDate(new Date(mailDayFirst), Session.getScriptTimeZone(), "dd.MM.yyyy");



                        var mailDaySecond = new Date();
                        mailDaySecond = mailDaySecond.setDate(mailDaySecond.getDate() + secondDay);
                        dueDateSecond = Utilities.formatDate(new Date(mailDaySecond), Session.getScriptTimeZone(), "dd.MM.yyyy");

                    } else if (pDate2 == "" && secondDay == "" && Subject2 == "") {

                        var firstDate = letterToColumn(pDate1);
                        var firstDateContainer = s.getRange(i, firstDate).getValues().join();
                        var firstDateFormatted = Utilities.formatDate(new Date(firstDateContainer), Session.getScriptTimeZone(), "dd.MM.yyyy");

                        var mailDayFirst = new Date();
                        mailDayFirst = mailDayFirst.setDate(mailDayFirst.getDate() + firstDay);
                        dueDateFirst = Utilities.formatDate(new Date(mailDayFirst), Session.getScriptTimeZone(), "dd.MM.yyyy");
                    }
                }
            }


            if (dueDateFirst == firstDateFormatted && emailSent != EMAIL_SENT && secondDateFormatted == undefined) {


                var emailBody = bodyArr.slice(bodySlice).join(" | ");
                var email = emailArr.splice(0).join();
                Logger.log("prva petlja" + " " + emailBody);
                               MailApp.sendEmail({
                                   to: email,
                                   subject: Subject1,
                                   htmlBody: emailBody
                               });
                s.getRange(i, emailColumn).setValue(EMAIL_SENT).setBackground("green");

            }

            if (dueDateSecond == secondDateContainer && emailSent != EMAIL_SENT && secondDateFormatted != undefined) {

                var emailBody = bodyArr.slice(bodySlice).join(" | ");
                var email = emailArr.splice(0).join();
                Logger.log("druga petlja" + " " + emailBody);
                               MailApp.sendEmail({
                                   to: email,
                                   subject: Subject2,
                                   htmlBody: emailBody
                               });
                s.getRange(i, emailColumn).setValue(EMAIL_SENT).setBackground("green");

            }

            if (dueDateFirst == firstDateFormatted && emailSent != EMAIL_SENT && secondDateFormatted != undefined) {

                var emailBody = bodyArr.slice(bodySlice).join(" | ");
                var email = emailArr.splice(0).join();
                Logger.log("treca petlja" + " " + emailBody);
                               MailApp.sendEmail({
                                   to: email,
                                   subject: Subject1,
                                   htmlBody: emailBody
                               });
                s.getRange(i, emailColumn).setValue(EMAIL_SENT).setBackground("green");
            }
        }
    }
}

function trigger() {

    ScriptApp.newTrigger('sendEmail')
        .timeBased()
        .everyHours(1)
        .create();
}

function triggered() {

    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var triggers = ScriptApp.getUserTriggers(ss);

    for (var i = 0; i <= triggers.length; i++) {

        Logger.log(triggers[i]);
        if (triggers[i] != undefined) {
            ScriptApp.deleteTrigger(triggers[i]);
        } else {
            trigger();
        }
    }
}