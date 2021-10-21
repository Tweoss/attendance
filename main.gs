const time_column = 1,
  name_column = 3,
  fill_text = '1',
  output_sheet_name = "Output",
  spreadsheet_id = "1CrzoV6JpJo-Q8ZXQ_ilKbmsrfmM73MgCEinBT0r62cI";

function main() {
  let main_spreadsheet = SpreadsheetApp.openById(spreadsheet_id);
  let responses_sheet = main_spreadsheet.getSheetByName("Form Responses 1");
  let last_row = responses_sheet.getLastRow();

  // get the columns from the leftmost to the maximum necessary to see both time and name columns. start at row 2
  //getRange(starting Row, starting column, number of rows, number of columns)
  let data = responses_sheet.getRange(2, 1, last_row, Math.max(time_column, name_column)).getValues();


  const map_string_to_date = (el_string) => {
    let date = new Date(el_string);
    if (date instanceof Date && !isNaN(date)) {
      return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
    }
    return el_string
  };
  // turn the timestamps into month/day/year
  data = data.map(x => { x[0] = map_string_to_date(x[0]); return x });

  // create a list of lowercase, nonrepeating names
  let names = [...new Set(data.map(function (array) { return array[1].toLowerCase(); }).sort())].filter((v) => v.length > 0);
  // create a list of dates
  let dates = [...new Set(data.map(function (array) { return array[0]; }).sort(
    (string1, string2) => {
      let [month1, day1, year1] = string1.split('/');
      let [month2, day2, year2] = string2.split('/');
      return year1 == year2 ? month1 == month2 ? day1 > day2 : month1 > month2 : year1 > year2
    }
  ))].filter((v) => v.length > 0);
  // the matrix to be copied to the sheet
  let table = Array.from(Array(names.length), () => Array.from(Array(dates.length), () => false));

  // instantiate a map of maps to determine whether or not a date exists, and whether or not a name exists within that date
  let date_to_name_to_exist = new Map();
  for (i in data) {
    let [date, name] = [data[i][0], data[i][1].toLowerCase()];
    if (name.length > 0) {
      date_to_name_to_exist.has(date) ? true : date_to_name_to_exist.set(date, new Map());
      let name_to_exist = date_to_name_to_exist.get(date);
      name_to_exist.has(name) ? true : name_to_exist.set(name, true);
    }
  }

  for (row in table) {
    for (col in table) {
      let name_to_exist = date_to_name_to_exist.get(dates[col]);
      if (name_to_exist != null && name_to_exist.has(names[row])) {
        table[row][col] = true;
      }
    }
  }

  // map bools into an output table with text
  let output_table = table.map(x => x.map(c => c ? fill_text : ' '));
  // console.log(output_table)
  // set the range for output
  let output_spreadsheet = main_spreadsheet.getSheetByName(output_sheet_name);

  // clear formats and content
  output_spreadsheet.getRange(1, 1, output_spreadsheet.getMaxRows(), output_spreadsheet.getMaxColumns()).clearFormat().clearContent();

  let table_range = output_spreadsheet.getRange(2, 2, names.length, dates.length);
  table_range.setValues(output_table);
  output_spreadsheet.getRange(1, 2, 1, dates.length).setValues([dates]);
  output_spreadsheet.getRange(2, 1, names.length, 1).setValues(names.map(x => [x]));

  // Sets the formula for day and name sums
  output_spreadsheet.getRange(2 /* first header row and 1-indexing */ + names.length, 2 /* start at two b/c of left names */, 1, dates.length).setFormulaR1C1("=SUM(INDIRECT(\"R2C[0]:R[-1]C[0]\", false))");
  output_spreadsheet.getRange(2, 1 + dates.length + 1, names.length, 1).setFormulaR1C1("=SUM(INDIRECT(\"R[0]C2:R[0]C[-1]\", false))");

  let bold = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .build();
  // Sets the title for the sums
  output_spreadsheet.getRange(names.length + 2, 1).setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText("Daily Count")
    .setTextStyle(bold)
    .build());
  output_spreadsheet.getRange(1, dates.length + 2).setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText("Member Count")
    .setTextStyle(bold)
    .build());

  // reset formatting
  // Remove one of the existing conditional format rules.
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(fill_text)
    .setBackground("#00FF00")
    .setRanges([table_range])
    .build();
  output_spreadsheet.setConditionalFormatRules([rule]);
}


