

using System.Reflection;
using System.Runtime.InteropServices;
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

try
{
    //Save Excel
    SaveInExcel(persons);

    //Read excel
    var personsList = ReadFromExcel(@"Test.xlsx");

    foreach (var item in personsList)
    {
        Console.WriteLine($"Entry : {item.ID}  {item.Name}");
    }

}
catch (Exception e)
{
    Console.WriteLine(e.Message);
}
finally
{
    GC.Collect();
    GC.WaitForPendingFinalizers();
}



static void SaveInExcel(IEnumerable<Person> persons)
{
    var excelApp = new Application();
    excelApp.DisplayAlerts = false;

    Workbook workbook = excelApp.Workbooks.Add();

    Worksheet workSheet = (Worksheet)workbook.Worksheets[1];

    Type type = typeof(Person);
    var iterator = 'A';
    foreach (PropertyInfo prop in type.GetProperties())
    {
        workSheet.Cells[1, iterator.ToString()] = prop.Name;

        iterator++;
    }


    var row = 1;

    foreach (var person in persons)
    {
        row++;
        workSheet.Cells[row, "A"] = person.ID;
        workSheet.Cells[row, "B"] = person.Name;
    }

    workSheet.Columns.AutoFit();

    workbook.SaveAs(@"Test.xlsx");


    excelApp.Quit();
}


static List<Person> ReadFromExcel(string path)
{

    List<Person> persons = new();
    var excelApp = new Application();
    excelApp.DisplayAlerts = false;


    Workbooks workbooks = excelApp.Workbooks;
    Workbook workbook = workbooks.Open(path);

    Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

    int row = worksheet.UsedRange.Rows.Count;
    // int column = worksheet.UsedRange.Columns.Count;


    for (int i = 2; i <= row; i++)
    {

        Person person = new Person
        {
            ID = Convert.ToInt32((worksheet.Cells[i, "A"] as Microsoft.Office.Interop.Excel.Range)?.Value2),
            Name = Convert.ToString((worksheet.Cells[i, "B"] as Microsoft.Office.Interop.Excel.Range)?.Value2)
        };
        persons.Add(person);

    }


    excelApp.Quit();


    return persons;

}