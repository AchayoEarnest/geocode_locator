function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index");
}
function getLoc(value) {
  var destId = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById(destId);
  var respSheet = ss.getSheets()[0];
  var data = respSheet.getDataRange().getValues();
  var headers = data[0];
  var numColumns = headers.length;
  var numResponses = data.length;
  var c = value[0];
  var d = value[1];
  var e = c + "," + d;

  // the script will add device geocode data in last submit data row by clicking the link on confirmation page that shows up after hitting the submit button
  // geocode data here consists of 3 columns : time to click the link, longitude & latitude and address (using reverse geocode)
  // as long as clicking the link is done before next respondent submit button, data will be entered in the right row.
  // however particularly for the case of multiple devices that submit data at about the same time then whichever device clicks the link closer to the last submit time, their geodata will be entered in the last submit data row.
  // this may leave geocode column in the row above empty. In this case the geocode data will be in red font
  // If sometime later another device click the link then the geodata will be entered in closest missing geodata row. The data will also be in red font
  // Therefore for red font data you may want to check manually after completion for correct geodata entry. In the questionnaire add question about address to help checking.
  // script ini akan menambahkan data geocode pada row data terakhir yang disubmitdengan meng-klik link pada confirmation page yang muncul setelah kita mengklik submit
  // data geocode disini terdiri dari 3 kolom : waktu mengklik link, longitude dan latitude, dan alamat (dari geocode)
  // selama link di klik sebelum responden selanjutnya, data akan dimasukkan dalam row yang benar.
  // tetapi jika untuk beberapa device yang mengklik dalam waktu hampir bersamaan maka siapapun yang mengklik terdekat waktunya dengan waktu data submit terakhir, data geocode dari device tersebut dakan dimasukkan ke row terakhir.
  // Ini akan mengakibatkan geocode untuk device lain di row sebelumnya kosong. Data geocode akan dalam font merah.
  // Jika setelah itu device lain menekan link, data geocode akan dimasukkan dalam row kosong terdekat.
  // Karenanya untuk data yang berwarna merah perlu di check secara manual setelah pengisian kuesioner. Untuk membantu checking bisa ditambahkan pertanyaan mengenai alamat di kuesioner.

  if (respSheet.getRange(1, numColumns).getValue() == "GeoAddress") {
    //fill data for second respondents onwards no missing geo data
    // time here is Jakarta, you may need to change time to your local time (in GMT)

    if (
      respSheet.getRange(numResponses, numColumns - 2).getValue() == "" &&
      respSheet.getRange(numResponses - 1, numColumns - 2).getValue() != ""
    ) {
      respSheet
        .getRange(numResponses, numColumns - 2)
        .setValue(
          Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy HH:mm:ss")
        );
      respSheet.getRange(numResponses, numColumns - 1).setValue(e);
      var response = Maps.newGeocoder().reverseGeocode(value[0], value[1]);
      f = response.results[0].formatted_address;
      respSheet.getRange(numResponses, numColumns).setValue(f);
    }
    //fill data with previous geo data missing. red font
    else if (
      respSheet.getRange(numResponses, numColumns - 2).getValue() == "" &&
      respSheet.getRange(numResponses - 1, numColumns - 2).getValue() == ""
    ) {
      respSheet
        .getRange(numResponses, numColumns - 2)
        .setValue(
          Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy HH:mm:ss")
        )
        .setFontColor("red");
      respSheet
        .getRange(numResponses, numColumns - 1)
        .setValue(e)
        .setFontColor("red");
      var response = Maps.newGeocoder().reverseGeocode(value[0], value[1]);
      f = response.results[0].formatted_address;
      respSheet
        .getRange(numResponses, numColumns)
        .setValue(f)
        .setFontColor("red");
    }

    //to fill missing previous data. red font
    else if (
      respSheet.getRange(numResponses, numColumns - 2).getValue() != ""
    ) {
      for (i = 0; i < numResponses; i++) {
        if (
          respSheet.getRange(numResponses - i, numColumns - 2).getValue() == ""
        ) {
          respSheet
            .getRange(numResponses - i, numColumns - 2)
            .setValue(
              Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy HH:mm:ss")
            )
            .setFontColor("red");
          respSheet
            .getRange(numResponses - i, numColumns - 1)
            .setValue(e)
            .setFontColor("red");
          var response = Maps.newGeocoder().reverseGeocode(value[0], value[1]);
          f = response.results[0].formatted_address;
          respSheet
            .getRange(numResponses - i, numColumns)
            .setValue(f)
            .setFontColor("red");
          break;
        }
      }
    }
  } else if (respSheet.getRange(1, numColumns).getValue() != "GeoAddress") {
    //create labels in first row
    respSheet.getRange(1, numColumns + 1).setValue("GeoStamp");
    respSheet.getRange(1, numColumns + 2).setValue("GeoCode");
    respSheet.getRange(1, numColumns + 3).setValue("GeoAddress");
    //fill data for first respondent

    if (numResponses == 2) {
      respSheet
        .getRange(numResponses, numColumns + 1)
        .setValue(
          Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy HH:mm:ss")
        );
      respSheet.getRange(numResponses, numColumns + 2).setValue(e);
      var response = Maps.newGeocoder().reverseGeocode(value[0], value[1]);
      f = response.results[0].formatted_address;
      respSheet.getRange(numResponses, numColumns + 3).setValue(f);
    } else if (numResponses > 2) {
      respSheet
        .getRange(numResponses, numColumns + 1)
        .setValue(
          Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy HH:mm:ss")
        )
        .setFontColor("red");
      respSheet
        .getRange(numResponses, numColumns + 2)
        .setValue(e)
        .setFontColor("red");
      var response = Maps.newGeocoder().reverseGeocode(value[0], value[1]);
      f = response.results[0].formatted_address;
      respSheet
        .getRange(numResponses, numColumns + 3)
        .setValue(f)
        .setFontColor("red");
    }
  }
}
