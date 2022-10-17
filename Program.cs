

using Microsoft.Office.Interop.Excel;


// list of persons
var persons = new List<Person>{
    new Person {
        ID = 1,
        Name = "John"
    },
    new Person {
        ID =2,
        Name="Susane"
    }
};


// Save Excel
SaveInExcel(persons);


// Garbage Collection will clean up process created by this program 
GC.Collect();
GC.WaitForPendingFinalizers();



static void SaveInExcel(IEnumerable<Person> persons)
{
    var excelApp = new Application();
    excelApp.DisplayAlerts = false;

    Workbook workbook = excelApp.Workbooks.Add();

    Worksheet workSheet = (Worksheet)workbook.Worksheets[1];

    workSheet.Cells[1, "A"] = "ID Number";
    workSheet.Cells[1, "B"] = "Name";

    var row = 1;

    foreach (var person in persons)
    {
        row++;
        workSheet.Cells[row, "A"] = person.ID;
        workSheet.Cells[row, "B"] = person.Name;
    }

    workbook.SaveAs(@"Test.xlsx");

    excelApp.Quit();
}
