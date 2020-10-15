using System.IO;
using System.Linq;
using IronXL;

namespace DataController
{
    class DataControl
    {
        readonly string path = @"C:\Users\thepa\Desktop\Internship\WPF\PersonInfoExcel\test.xlsx";
        

        public void ToDatabase(string firstName, string lastName)
        {
            if (!File.Exists(path))
            {
                WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
                var sheet = workbook.CreateWorkSheet("PersonInfoXL");
                sheet["A1"].Value = "First Name";
                sheet["B1"].Value = "Last Name";
                int rowCount = sheet.Rows.Count() + 1;
                sheet["A" + rowCount].Value = firstName;
                sheet["B" + rowCount].Value = lastName;

                workbook.SaveAs(path);
            }

            else if (File.Exists(path))
            {
                WorkBook workbook = WorkBook.Load(path);
                WorkSheet sheet = workbook.GetWorkSheet("PersonInfoXL");
                int rowCount = sheet.Rows.Count() + 1;
                sheet["A" + rowCount].Value = firstName;
                sheet["B" + rowCount].Value = lastName;

                workbook.SaveAs(path);
            }
        }

        public void DeleteAllData()
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
}

