var excelbuilder = require('msexcel-builder');

// Create a new workbook file in current working-path
var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx')

//input
obj = [{'col1': 1, 'col2': 2},{'col1': 3, 'col2': 4}] //input JSON

// Create a new worksheet with x columns and y rows
var sheet = workbook.createSheet('sheet1', Object.keys(obj[0]).length, obj.length+1);

//Map columns to values
numColumns = 0;
mapColumns = {}; //title
for (col in obj[0]) { //foreach column
        if (mapColumns[col] == undefined) { //this column wasn't mapped
            mapColumns[col] = ++numColumns;
            sheet.set(numColumns, 1, col);
        }
}

for (r = 0; r < obj.length; r++) { //foreach row
    row = obj[r];
    for (col in row) { //foreach column
        if (mapColumns[col] == undefined) { //this column wasn't mapped
            console.log("new column found")
            mapColumns[col] = ++numColumns;
        }
        rowNumber = r+2; //jump title row
        colNumber = mapColumns[col];
        console.log("New tuple:",colNumber, rowNumber, row[col]);
        sheet.set(colNumber, rowNumber, row[col]);
    }
}

// Save it
workbook.save(function(err){
  if (err) 
    workbook.cancel();
  else
    console.log('congratulations, your workbook created');
});