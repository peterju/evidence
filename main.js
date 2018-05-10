var app       = new ActiveXObject("Excel.Application");
var directory = WScript.CreateObject("WScript.Shell").CurrentDirectory + '\\'
// var wb = app.Workbooks.Open(directory + "excel_test.xlsx",  false)
// var lo = wb.Worksheets("Sheet1").ListObjects("table_test")
WScript.Echo(directory);

// var FS = new ActiveXObject("Scripting.FileSystemObject");
// eval(FS.OpenTextFile("Util_GlobalsMgmt.js", 1).ReadAll());

// // Set up Excel
// var app       = new ActiveXObject("Excel.Application");
// var directory = WScript.CreateObject("WScript.Shell").CurrentDirectory + '\\'
// var wb = app.Workbooks.Open(directory + "excel_test.xlsx",  false)
// var lo = wb.Worksheets("Sheet1").ListObjects("table_test")

// // Do something
// function change_data(lo) {
//    with (lo.DataBodyRange.Cells(2, 1)) {
//        Value2 = Math.random();
//        return Value2
//    }
// }

// var result = Utils_GlobalsMgmt.temp_disable_screenupdating(
//     app,
//     function() { return change_data(lo); }
// )

// // Tear down Excel
// Utils_GlobalsMgmt.temp_disable_displayalerts(app, function(){ wb.Save(); });
// app.Quit();

// WScript.Echo(result);

// ' if WScript.Arguments.Count < 2 Then
// '     WScript.Echo "Error! Please specify the source path and the destination. Usage: XlsToCsv SourcePath.xls Destination.csv"
// '     Wscript.Quit
// ' End If
// ' Dim oExcel
// ' Set oExcel = CreateObject("Excel.Application")
// ' Dim oBook
// ' Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))
// ' oBook.SaveAs WScript.Arguments.Item(1), 6
// ' oBook.Close False
// ' oExcel.Quit
// ' WScript.Echo "Done"