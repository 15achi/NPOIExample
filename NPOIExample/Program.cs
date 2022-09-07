using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOIExample.Models;

List<Student> _oStudents = new List<Student>();


for (int i = 1; i <= 9; i++)
{
    _oStudents.Add(new Student()
    {
        Id = i,
        Name = "Student" + i,
        Roll = "100" + i,
        Address = "test" + i
    });
}

var memoryStream = new MemoryStream();
// --- Below code would create excel file with dummy data----  

    IWorkbook workbook = new XSSFWorkbook();
    ISheet excelSheet = workbook.CreateSheet("Student");
    int counter = 1;
    IRow row = excelSheet.CreateRow(0);
    row.CreateCell(0).SetCellValue("Id");
    row.CreateCell(1).SetCellValue("Name");
    row.CreateCell(2).SetCellValue("Roll");
    row.CreateCell(3).SetCellValue("Address");

    foreach (var student in _oStudents)
    {

        row = excelSheet.CreateRow(counter);
        row.CreateCell(0).SetCellValue(student.Id);
        row.CreateCell(1).SetCellValue(student.Name);
        row.CreateCell(2).SetCellValue(student.Roll);
        row.CreateCell(3).SetCellValue(student.Address);
       counter++;
    }

    using (var exportData = new MemoryStream())
    {

        workbook.Write(exportData);

      var conent = exportData.ToArray();
    //File.WriteAllBytes("C:\\Users\\User\\Desktop\\test\\excel_1.xlsx", StreamHelpers.GetBytes(exportData));
    File.WriteAllBytes("C:\\Users\\User\\Desktop\\test\\excel_1.xlsx", conent);

}

Console.WriteLine("Hello, World!");
