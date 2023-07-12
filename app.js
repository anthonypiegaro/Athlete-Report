function createReport() {
  /* This will be the function used to create the report */

  // Get all data
  const spreadsheet = SpreadsheetApp.getActive();
  const assesment_data_sheet = spreadsheet.getSheetByName("assessment_data");
  const documentation_sheet = spreadsheet.getSheetByName("Documentation");
  const drill_sheet = spreadsheet.getSheetByName("Drills");
  let dataRange = assesment_data_sheet.getDataRange();
  const data = dataRange.getValues();
  dataRange = documentation_sheet.getDataRange();
  const documentation_data = dataRange.getValues();
  dataRange = drill_sheet.getDataRange();
  const drillData = dataRange.getValues();


  // Get player data from the selected row
  const selectedRow = assesment_data_sheet.getCurrentCell();
  const row = selectedRow.getRow();
  const athleteData = data[row-1]

  const athleteName = athleteData[0];
  Logger.log(athleteName);

  // Create new Google Doc with "athleteName_report" as the name of the doc
  const doc = DocumentApp.create(athleteName + " Report");

  // Add the OA logo
  const body = doc.getBody();
  const imageUrl = "https://d1fdloi71mui9q.cloudfront.net/gp8N8MaITFGnHcwOXDui_DlZ15Av3Sb331UTJ";
  const response = UrlFetchApp.fetch(imageUrl);
  const image = response.getBlob();
  const insertedImage = body.insertImage(0, image);
  // Center the image
  let styles = {};
  styles[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  insertedImage.getParent().setAttributes(styles);

  // Add the Name of the Athlete
  let paragraph = body.appendParagraph(athleteName + " Report");
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  let textRange = paragraph.editAsText();
  textRange.setBold(true);
  textRange.setFontSize(15);
  let date = athleteData[1];
  if (date !== "") {
     date = Utilities.formatDate(date, "GMT+1", "MM/dd/yyyy");
     paragraph.appendText(" for " + date);
  }

  let tests = {
    "Horizontal Adduction Left": {
      idx: 5,
      description_row: 2,
    },
    "Horizontal Adduction Right": {
      idx: 6,
      description_row: 2,
    },
    "Total Arc Left": {
      idx: 11,
      description_row: 5,
    },
    "Total Arc Right": {
      idx: 12,
      description_row: 5,
    },
    "Total T/S Rotation": {
      idx: 15,
      description_row: 7,
    },
    "Hip IR Right": {
      idx: 16,
      description_row: 9,
    },
    "Hip IR Left": {
      idx: 17,
      description_row: 9,
    },
    "Hip ER Right": {
      idx: 18,
      description_row: 8,
    },
    "Hip ER Left": {
      idx: 19,
      description_row: 8,
    },
    "Ankle DF Right": {
      idx: 20,
      description_row: 10,
    },
    "Ankle DF Left": {
      idx: 21,
      description_row: 10,
    },
    "OH Squat": {
      idx: 22,
      description_row: 11,
    },
    "Hurdle Right": {
      idx: 23,
      description_row: 12,
    },
    "Hurdle Left": {
      idx: 24,
      description_row: 12,
    },
    "ASLR Right": {
      idx: 25,
      description_row: 13,
    },
    "ASLR Left": {
      idx: 26,
      description_row: 13,
    },
    "Pull-ups": {
      idx: 27,
      description_row: 14,
    },
    "Push-ups": {
      idx: 28,
      description_row: 15,
    },
    "Lateral Jump Right": {
      idx: 29,
      description_row: 16,
    },
    "Lateral Jump Left": {
      idx: 30,
      description_row: 16,
    },
    "Broad Jump": {
      idx: 31,
      description_row: 17,
    },
    "Vertical Jump": {
      idx: 32,
      description_row: 18,
    },
    "Medball Toss": {
      idx: 33,
      description_row: 19,
    },
    "Pro Agility": {
      idx: 34,
      description_row: 20,
    },
    "10 Yard Accel.": {
      idx: 35,
      description_row: 21,
    },
    "Plate Pinch Hold Right": {
      idx: 36,
      description_row: 22,
    },
    "Plate Pinch Hold Left": {
      idx: 37,
      description_row: 22,
    },
    "W Hold": {
      idx: 38,
      description_row: 23,
    }
  }
  // Loop through each metric and add to the report
  for (let test in tests) {
    pass = documentation_data[tests[test].description_row - 1][3];
    unit = documentation_data[tests[test].description_row - 1][2];
    let desc = documentation_data[tests[test].description_row - 1][1];
    let athleteValue = athleteData[tests[test].idx];
    let graphs = documentation_data[tests[test].description_row - 1][4].split(",");
    let idx = tests[test].idx;
    let drills = documentation_data[tests[test].description_row - 1][5].split(",");
    // Add title of metric
    paragraph = body.appendParagraph(test);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    let textRange = paragraph.editAsText();
    textRange.setBold(true);
    textRange.setFontSize(15);

  // Add description of metric and what constitutes a pass
    paragraph = body.appendParagraph(desc + " To pass, an athlete needs a score of " + pass + " or better.");
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    textRange = paragraph.editAsText();
    textRange.setBold(false);
    textRange.setFontSize(11);

    // If graph(s), add them
    for (let i = 0; i < graphs.length; i++) {
      let chart;
      let graph = graphs[i].trim();
      if (graph === "") {
        continue;
      } else if (graph == "bar") {
        chart = createBarChart(athleteName, test, pass, athleteValue, unit);
      } else if (graph == "perc") {
        // Get all OA data for the test
        let columnData = [];
        for (let j = 1; j < data.length; j++) {
          if (data[j][idx] !== 0 && data[j][idx] !== "") {
            columnData.push(data[j][idx]);
          }
        }
        chart = createPercChart(test, columnData, athleteValue, unit, athleteName);
      }
      // Create the chart in the document
      body.appendParagraph('').asParagraph().appendInlineImage(chart);
    }

    // Add athlete score
    paragraph = body.appendParagraph(athleteName + "'s score: " + athleteValue);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    textRange = paragraph.editAsText();
    textRange.setBold(false);
    textRange.setFontSize(11);

    // Add whether a pass or fail
    let result;
    if (["Horizontal Adduction Right", "Horizontal Adduction Left"].includes(test)) {
      if (athleteValue === "across") {
        Logger.log("Pass");
        result = "Pass"
      } else {
        result = "Fail";
      }
    } else {
      if (athleteValue >= pass) {
        result = "Pass";
      } else {
        result = "Fail";
      }
    }
    paragraph = body.appendParagraph("Result: " + result);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    textRange = paragraph.editAsText();
    textRange.setFontSize(11);

    // If Fail, add list of drills with links
    if (result === "Fail") {
      paragraph = body.appendParagraph("The following is a list of drills to improve your " + test + " test:");
      paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      textRange = paragraph.editAsText();
      textRange.setFontSize(11);

      // Create a numbered list
      for (let k = 0; k < drills.length; k++) {
        let drill = drills[k].trim();
        if (drill == "") {
          continue;
        }
        // Get drill from the Drill sheet
        let link;
        for (let l = 0; l < drillData.length; l++) {
          if (drill === drillData[l][0]) {
            link = drillData[l][1];
            break;
          }
        }
        // Add drill with link attatched
        let listItem = body.appendListItem(drill);
        listItem.setLinkUrl(link);
      }
    }


    Logger.log("Test: " + test);
    Logger.log("Description: " + desc);
    Logger.log("Result: " + athleteValue + " " + unit);
    Logger.log("Value needed needed to pass: " + pass + " " + unit);
    if (["Horizontal Adduction Right", "Horizontal Adduction Left"].includes(test)) {
      if (athleteValue === "across") {
        Logger.log("Pass");
      } else {
        Logger.log("Fail");
      }
    } else {
      if (athleteValue >= pass) {
      Logger.log("Pass");
      } else {
        Logger.log("Fail");
      }
    }

    for (let i = 0; i < graphs.length; i++) {
      let graph = graphs[i].trim();
      if (graph === "") {
        continue;
      }
      Logger.log(graph);
    }
    // If Fail, add list of drills with links
  }
  // Save and close the document
  doc.saveAndClose();
  const docUrl = doc.getUrl();
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Report ran and in Google Docs: ' + docUrl);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Report Menu')
      .addItem('Run Report', 'createReport')
      .addToUi();
}

function createBarChart(name, test, pass, value, unit) {
  const tableData = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Name")
    .addColumn(Charts.ColumnType.NUMBER, test)
    .addRow([name, value])
    .addRow(["Standard", pass])
    .build();
  
  const chart = Charts.newBarChart()
    .setTitle(name + " " + test + " compared to Standard")
    .setYAxisTitle(test + " " + "(" + unit + ")")
    .setDataTable(tableData)
    .setOption("width", 600)
    .setOption("height", 400)
    .build();

  // Think about hard coding the height and width of the graph
  
  return chart
}

function createPercChart(test, data, athleteValue, unit, athleteName) {
  const mean = calculateMean(data);
  const std = calculateStandardDeviation(data);

  // Generate x-values for the normal distribution curve
  const xValues = generateXValues(mean, std, -3, 3, 0.1);

  // Calculate the corresponding y-values using the normal distribution formula
  const yValues = xValues.map(function(x) {
    return normalDistribution(x, mean, std);
  });

  let dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.NUMBER, unit)
    .addColumn(Charts.ColumnType.NUMBER, "density")
    .addColumn(Charts.ColumnType.NUMBER, "Athlete's Score");

  for (let i = 0; i < xValues.length; i++) {
    dataTable.addRow([xValues[i], yValues[i]]);
  }

  const athleteValueHeight = Math.max(...yValues) * 1.05;
  dataTable.addRow([athleteValue,, 0]);
  dataTable.addRow([athleteValue,,athleteValueHeight]);

  dataTable = dataTable.build();

  const chartBuilder = Charts.newLineChart()
  .setTitle("Normal Distribution")
  .setXAxisTitle(test + " (" + unit + ")")
  .setYAxisTitle("Density")
  .setCurveStyle(Charts.CurveStyle.SMOOTH)
  .setDataTable(dataTable)
  .setOption("width", 600)
  .setOption("height", 400)
  .setOption("chartArea", {width: "50%", height: "75%"});

  let chart = chartBuilder.build();

  return chart;
}

function generateXValues(mean, stdDev, start, end, step) {
  let xValues = [];
  for (let x = start; x <= end; x += step) {
    xValues.push(x * stdDev + mean);
  }
  return xValues;
}

function normalDistribution(x, mean, stdDev) {
  const exponent = -Math.pow(x - mean, 2) / (2 * Math.pow(stdDev, 2));
  return 1 / (stdDev * Math.sqrt(2 * Math.PI)) * Math.exp(exponent);
}

function calculateMean(data) {
  let sum = data.reduce(function(a, b) {
    return a + b;
  }, 0);
  return sum / data.length;
}

function calculateStandardDeviation(data) {
  let mean = calculateMean(data);
  let squareDiffs = data.map(function(value) {
    let diff = value - mean;
    return diff * diff;
  });
  let avgSquareDiff = calculateMean(squareDiffs);
  return Math.sqrt(avgSquareDiff);
}

