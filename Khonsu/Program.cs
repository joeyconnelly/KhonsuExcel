using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;


namespace Khonsu
{
    public class Program
    {
        public static DataTable CreateTable()
        {
            DataTable table = new DataTable("WhompData");

            string c0 = "DogName";
            string c1 = "DogAge";
            string c2 = "DogOwner";
            string c3 = "DogBirth";

            // Create columns.
            table.Columns.Add(c0, typeof (string));
            table.Columns.Add(c1, typeof (int));
            table.Columns.Add(c2, typeof (string));
            table.Columns.Add(c3, typeof (DateTime));

            // Create row data.
            DataRow row = table.NewRow();
            row[c0] = "Alvie";
            row[c1] = 12;
            row[c2] = "Joseph Bam";
            row[c3] = new DateTime(2018, 6, 14, 9, 0, 0); // date dnot working.              
            table.Rows.Add(row);

            return table;
        }

        public static void Main(string[] args)
        {
            string savePath = @"C:\JOEY_TomatoSoup\Khonsu\";
            string fileName = "coolest-ever";
            string tableName = "HappyThanksgiving";
            string fileExt = @".xlsx";

            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    using (DataTable table = CreateTable())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(tableName); 
                        worksheet.Cells["A1"].LoadFromDataTable(table, true);
                        worksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 12));
                        worksheet.Cells.AutoFitColumns();

                        // Remove existing file if exists.
                        string file = savePath + fileName + fileExt;
                        if (File.Exists(file))
                        {
                            File.Delete(file);
                        }
                
                        // Create file.
                        FileStream fstream = File.Create(file);
                        fstream.Close();

                        // Write contents of excel sheet to file.
                        File.WriteAllBytes(file, package.GetAsByteArray());
                    }
                }
            }
            catch (Exception)
            {

            }
        }
    }
}
