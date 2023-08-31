# Mail-Automation
  function sendYy() {
  var sheetName = "combine email"; // Sheet name ADD here
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  var filterCell = sheet.getRange("J2"); // Filter cell
  var filterValue = filterCell.getValue(); // Filter value

  console.log("Filter Value: " + filterValue); // Debug logging

  if (filterValue !== "YES") {
    // value is not "YES", do not send email
    return;
  }

  var rangeValues = sheet.getRange("P2:P20").getValues(); // Get the range values from column P

  var ranges = [
    ];

  for (var i = 0; i < rangeValues.length; i++) {
    var rangeCell = rangeValues[i][0];

    if (rangeCell) {
      var rangeData = {
        range: rangeCell,
        paddingTop: "5px",
        paddingRight: "19px",
        paddingBottom: "1px",
        paddingLeft: "6px",
        fontSize: "105%",
        alignCenter: "true",
        margin: "2px",
        padding:"5px"
        
      };

      ranges.push(rangeData);
    }
  }
// Pick subject line from cell I6
  var subjectRange = sheet.getRange("K10");
  var emailSubject = subjectRange.getValue(); // Email subjec

  //var emailSubject = "Pre Quality Report | MP | GoShorts Annotations-Edit | RCA & Observations | JUNE'23"; // Email subject

  var recipientEmailRange = sheet.getRange("M3:M4"); // Range for recipient email addresses
  var recipientEmailValues = recipientEmailRange.getValues(); // Get the values in the range

  var checkboxRange = sheet.getRange("N3:N4"); // Range for checkbox values
  var checkboxValues = checkboxRange.getValues(); // Get the values in the range

  var message = ""; // Email message

  for (var i = 0; i < ranges.length; i++) {
    var rangeData = ranges[i];
    var dataRange = sheet.getRange(rangeData.range);
    var data = dataRange.getDisplayValues();
    var backgrounds = dataRange.getBackgrounds(); // Cell background colors
    var textColors = dataRange.getFontColors(); // Text colors
    var formulas = dataRange.getFormulas(); // Formulas in the range

    // Data ka HTML table mein convert karna
    var htmlTable = "<table style='border-collapse:collapse;'>";

// Fetch specific words from column O
var specificWordsRange = sheet.getRange("O2:O32").getValues();
var specificWords = specificWordsRange.map(function(row) {
  return row[0].toString().trim();
});


    for (var row = 0; row < data.length; row++) {
      htmlTable += "<tr>";

      for (var col = 0; col < data[row].length; col++) {
        var cellValue = data[row][col];
        var cellFormula = formulas[row][col];
        var cell = dataRange.getCell(row + 1, col + 1);
        var background = backgrounds[row][col]; // Cell background color
        var textColor = textColors[row][col]; // Text color

        // Cell ke style ko HTML inline style ke roop mein convert karna
        var fontWeight = cell.getFontWeight();
        var textDecoration = cell.getFontLine();
        var paddingTop = rangeData.paddingTop; // Top padding
        var paddingRight = rangeData.paddingRight; // Right padding
        var paddingBottom = rangeData.paddingBottom; // Bottom padding
        var paddingLeft = rangeData.paddingLeft; // Left padding
        var fontSize = rangeData.fontSize; // Font size

        var cellHTML = "";

        // Agar cell mein value hai
        if (cellValue !== "") {
          var words = cellValue.split(" ");
          var borderStyle = "1px solid black";

          // Agar sentence mein 5 se jyada words hai,
          if (words.length > 5) {
            borderStyle = "none";
          }

          // If sentence has more than 5 words, remove border-line
          if (rangeData.range === "A3:A8" && words.length > 2) {
            borderStyle = "none";
          }

         // If sentence has more than 5 words, remove border-line
          if (rangeData.range === "A28:C34" && words.length >3) {
            borderStyle = "1px solid black";
          }

          // Agar cell mein special words hai,
          var specialWords = [
            "Hello Everyone,",
            "Task Benchmark: 95.00%",
            "Quality Achieved:",
            "RCA Observations and Examples: ",
            "Top Performers Top 3 Date wise :",
            "Bottom Performers Date wise :",
            "MTD",
            "QA Observations:",
            "Observations:",
            "Task Benchmark: ",
            "Quality Achieved",
            "Rater Stats:",
            "Overall Rater Wise Stats:",
            "New Batch:",
            "MTD:",
            "RCA Observations and Examples:",
            "MTD :",
            "Entity observation:",
            "Rater wise stats:",
            "RCA Observations and Examples ",
            "Entity Analysis :",
            "Top Performers:",
           "Bottom Performers:",
           "**No error found",
           "Hello Aditya,",
           "RCA Observations and Examples",
           "Observations:",
           "Observation:",
           "sloppy mistake:",
           "Sloppy Mistake:"

          ];
          if (specialWords.includes(cellValue)) {
            borderStyle = "none";
          }

          // Check if the cell falls within the special range
          var specialRange = sheet.getRange("A9:B10");
          var cellRow = cell.getRow();
          var cellColumn = cell.getColumn();

          if (cellRow >= specialRange.getRow() && cellRow <= specialRange.getLastRow() &&
              cellColumn >= specialRange.getColumn() && cellColumn <= specialRange.getLastColumn()) {
            borderStyle = "none";
            background = specialRange.getBackground();
            fontSize = "18px"; // Increase font size for the special range
          }

 // Check if the cell contains a number or percentage, if so, center align it
    var isNumber = !isNaN(parseFloat(cellValue)) && isFinite(cellValue);
    var isPercentage = false;

    if (typeof cellValue === 'string') {
      var numericValue = parseFloat(cellValue.replace(/,/g, '').replace('%', ''));
      if (!isNaN(numericValue)) {
        if (cellValue.trim().endsWith('%')) {
          isPercentage = true;
        }
      }
    }

    var isNumericValue = isNumber || isPercentage;

    // Check if the cell contains any specific words to be center aligned
    var isCenterAlignWord = specificWords.includes(cellValue);

    var cellHTML = "<td style='border:" + borderStyle + "; font-weight:" + fontWeight + "; text-decoration:" + textDecoration + "; background-color:" + background + "; color:" + textColor + "; padding-top:" + paddingTop + "; padding-right:" + paddingRight + "; padding-bottom:" + paddingBottom + "; padding-left:" + paddingLeft + "; font-size:" + fontSize + "; text-align: " + ((isNumericValue || isCenterAlignWord) && rangeData.alignCenter ? "center" : "left") + ";'>";

         
          // Agar cell mein hyperlink hai, toh HTML anchor tag ka istemaal
          if (cellFormula.startsWith('=HYPERLINK')) {
            var hyperlinkURL = cellFormula.match(/"(.*?)"/)[1];
            cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";
          } else {
            // Convert specific words to hyperlinks
            if (rangeData.range === "A3:A8" && row === 2) {
              cellHTML += replaceSpecificWordsWithHyperlinks(cellValue);
            } else if (cell.getA1Notation() === "D18") {
              var hyperlinkURL = sheet.getRange("W2").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            } else if (cell.getA1Notation() === "D19") {
              var hyperlinkURL = sheet.getRange("W3").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            } else if (cell.getA1Notation() === "D20") {
              var hyperlinkURL = sheet.getRange("W4").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            } else if (cell.getA1Notation() === "D21") {
              var hyperlinkURL = sheet.getRange("W5").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            } else if (cell.getA1Notation() === "D22") {
              var hyperlinkURL = sheet.getRange("W6").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";
              
            } else if (cell.getA1Notation() === "D23") {
              var hyperlinkURL = sheet.getRange("W7").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            } else if (cell.getA1Notation() === "D24") {
              var hyperlinkURL = sheet.getRange("W8").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            }else if (cell.getA1Notation() === "C30") {
              var hyperlinkURL = sheet.getRange("I17").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            } else if (cell.getA1Notation() === "C31") {
              var hyperlinkURL = sheet.getRange("I18").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            } else if (cell.getA1Notation() === "C32") {
              var hyperlinkURL = sheet.getRange("I19").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";

            }else if (cell.getA1Notation() === "C33") {
              var hyperlinkURL = sheet.getRange("").getValue(); // Replace with the desired URL from W1 cell
              cellHTML += "<a href='" + hyperlinkURL + "'>" + cellValue + "</a>";
            } else {
              cellHTML += cellValue;
            }
          }

          cellHTML += "</td>";
        } else {
          // Agar cell khali hai, toh border-line hata dena
          cellHTML += "<td></td>";
        }

        htmlTable += cellHTML;
      }

      htmlTable += "</tr>";
    }

    htmlTable += "</table>";
    
    // Add line break with minimal spacing
   if (i < ranges.length - 1) {
    htmlTable += "<div style='height: 0px; line-height:1px; font-size: 1px; margin-bottom:0%; margin-top:0.3%; background-color: transparent; border: none;'>&nbsp;</div>";
  }


    message += htmlTable;
  }

  // Retrieve the hyperlink URLs from cells Y1 and Y2
var hyperlinkURL1 = sheet.getRange("Y2").getValue();
var hyperlinkURL2 = sheet.getRange("Z2").getValue();

// Add the note and hyperlink to the error log
//message += '<br><p style="color: black;"><span style="font-weight: bold; font-size: 15px;"></span> <span style="font-size:14px; font-weight: bold;"> * Kindly Review Your Errors on </span> <a href="' + hyperlinkURL1 + '" target="_blank" style="color:blue; font-weight: bold; font-size: 14px;">QASA</a> <span style="font-size:13px;"><a href="' + hyperlinkURL2 + '" target="_blank" style="color:blue; font-weight: bold; font-size: 13px;"></a> </span></p>';
message += '<br><br><strong style="color: #A52A2A; font-size: 14px;">Note: This is an autogenerated E-mail by Script</strong>';

 // Retrieve the values from the range "X2:X3"
var footerRange = sheet.getRange("X2:X3");
var footerValues = footerRange.getValues();
var footerTextX2 = footerValues[0][0]; // Value from X2 cell
var footerTextX3 = footerValues[1][0]; // Value from X3 cell

// Append the footerTextX2 to the message variable
message += "<br><br><strong style='color: black;'> " + footerTextX2 + "</strong>";

// Append the footerTextX3 to the message variable
message += "<br><strong style='color: black;'> " + footerTextX3 + "</strong>";





    // Add "Thanks and Regards" and the name
    // message += "<br><br><br><strong style='color: black;'>Thanks and Regards<br>Naina Sharma | Analyst- Quality </strong>";

  // Send email to each recipient
  for (var j = 0; j < recipientEmailValues.length; j++) {
    var recipientEmail = recipientEmailValues[j][0]; // Extract the email address from the range

    var checkboxValue = checkboxValues[j][0]; // Get the checkbox value

    if (checkboxValue === true) {
      // Checkbox is checked, send email
      MailApp.sendEmail({
        to: recipientEmail,
        subject: emailSubject,
        htmlBody: message,
      });

      console.log("Email sent to: " + recipientEmail); // Debug logging
    }
  }

  // Reset the filter value to NO after sending email
  //filterCell.setValue("NO");
}
// Function to replace specific words with hyperlinks
function replaceSpecificWordsWithHyperlinks(text) {
  var wordsToReplace = {
    "MLDO-Dashboard'23": "https://lookerstudio.google.com/u/0/reporting/1027659b-9874-476c-8d97-c27b143e572c/page/p_68hsz8nv5c",
    "go/mldoqualitydash23": "http://go/mldoqualitydash23"
  };

  var replacedText = text;

  for (var word in wordsToReplace) {
    var regex = new RegExp("\\b" + word + "\\b", "g");
    var replacement = "<a href='" + wordsToReplace[word] + "'>" + word + "</a>";
    replacedText = replacedText.replace(regex, replacement);
  }

  return replacedText;
}
