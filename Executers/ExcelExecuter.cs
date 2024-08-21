
using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Diagnostics;
using System.Drawing;
using System.IO.Packaging;
using System.Text;

namespace ConsoleApp1.Helpers;

internal static class ExcelExecuter
{
    public static string GetDigits(this string str) 
    {
        return new string(str.ToCharArray().Where(c => char.IsDigit(c)).ToArray());
    }
    public static string GetLetters(this string str)
    {
        return new string(str.ToCharArray().Where(c =>! char.IsDigit(c)).ToArray());
    }

    public static void EditDirSpecialities(string dir) 
    {
        var fileEntries = Directory.GetFiles(dir);
        foreach (var file in fileEntries) 
        {
            EditSpecialityDescriptions(file);
        }
    }

    public static void EditSpecialityDescriptions(string file_path) 
    {
        string file_name = Path.GetFileName(file_path);
        Console.WriteLine(file_name);

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string speciality_code = file_name.GetDigits();
            var table = ApplicationDbContext.SelectAll<Module>();

            using (var package = new ExcelPackage(file_path))
            {
                var sheet = package.Workbook.Worksheets[0];
                var cell = sheet.Cells.Where(x => x.Text
                .Contains("Описание дисциплины"))
                    .FirstOrDefault();

                var coll = cell.ToString().GetLetters();
                var row = cell.ToString().GetDigits();

                //Console.WriteLine(sheet.Cells["A"+coll].Value);
                int i = int.Parse(row) + 5;
                while (!string.IsNullOrEmpty(sheet.Cells["A" + i].Value?.ToString()))
                {
                    string excel_module_name = sheet.Cells["A" + i].Value.ToString().ToLower().Replace(" ", "");

                    Module? module_db = table.Where(
                        m => m.Name.ToLower().Replace(" ", "").Contains(excel_module_name)
                        && m.Speciality.GetDigits().Contains(speciality_code)).FirstOrDefault()
                        ?? 
                        table.Where(m => excel_module_name.Contains(m.Name.ToLower().Replace(" ", ""))
                        && m.Speciality.GetDigits().Contains(speciality_code)).FirstOrDefault();

                    sheet.Cells[coll + i].Value = module_db?.Description;

                    i += 1;
                }

                package.Save();
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось изменить файл {file_name} Excel. Ошибка: {ex.Message}");
        }
    }

    public static void WriteRow(List<Module> modules)
    {
        

        var filePath = "D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\Модули-Write\\InsertArrays.xlsx";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(filePath))
        {
            var sheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "My Sheet") ?? package.Workbook.Worksheets.Add("My Sheet");

            #region Styling

            
            var cf = sheet.ConditionalFormatting.AddBelowAverage("A1:A10");
            cf.Style.Fill.Style = eDxfFillStyle.PatternFill;

            //The most common type of fill "Solid Fill" is a pattern fill.
            //This property represents the Pattern Style drop down in excel and has enum options for all of them.
            cf.Style.Fill.PatternType = ExcelFillStyle.Solid;

            //This is how to pick "thin horizontal" equivalent in excel
            //Note that the name is as it needs to be written in the xml.
            cf.Style.Fill.PatternType = ExcelFillStyle.LightVertical;

            //Represents Pattern Color in excel .Gradient is the equivalent for gradient styles.
            cf.Style.Fill.BackgroundColor.Color = System.Drawing.Color.Gold;

            //.Border refers to the borders around a cell. You can set different options for each of the four or .BorderAround for all borders
            cf.Style.Border.Top.Style = ExcelBorderStyle.MediumDashed;
            cf.Style.Border.Top.Color.Color = System.Drawing.Color.RebeccaPurple;

            //This will overwrite the previous changes but also apply to all borders
            cf.Style.Border.BorderAround(ExcelBorderStyle.MediumDashDotDot, System.Drawing.Color.Red);

            //.Font has multiple standard properties like the below
            cf.Style.Font.Bold = true;
            cf.Style.Font.Italic = true;
            cf.Style.Font.Strike = true;
            cf.Style.Font.Underline = ExcelUnderLineType.Single;
            cf.Style.Font.Color.Color = System.Drawing.Color.ForestGreen;

            //NumberFormat represents the Number tab of the format UI in excel and is set via format string
            cf.Style.NumberFormat.Format = "0.00%";
            //You can also get the id of the numberformat but not set it
            var id = cf.Style.NumberFormat.NumFmtID;


            //sheet.Cells["B1"].Style.;


            #endregion

            //var table = sheet.Tables.Add(,"");
            //sheet.Range.AutoFitColumns();

            sheet.Cells["A1"].Value = "Факультет ответственный за дисциплину";
            sheet.Cells["B1"].Value = "Специальность";
            sheet.Cells["C1"].Value = "Наименование предмета";
            sheet.Cells["D1"].Value = "Описание предмета";

            for (int i = 2; i < modules.Count; i++)
            {

                sheet.Cells["A" +  i].Value = modules[i].DepartmentShortName;
                sheet.Cells["B" + i].Value = modules[i].Name;
                sheet.Cells["C" + i].Value = modules[i].Name;
                sheet.Cells["D" + i].Value = modules[i].Name;

            }


            // Save to file
            package.Save();


        }


           FileInfo fi = new FileInfo(filePath);
            if (fi.Exists)
            {
                var p = new Process();
                p.StartInfo = new ProcessStartInfo(filePath)
                {
                    UseShellExecute = true
                };
                p.Start();

            }
            else
            {
                //file doesn't exist
            }

            /*
            var sheet = package.Workbook.Worksheets.Add("Market Report");
            sheet.Cells["B2"].Value = "Company:";
            sheet.Cells[2, 3].Value = "Некое имя компании";

            sheet.Cells[8, 2, 8, 4].LoadFromArrays(new object[][] { new[] { "Capitalization", "SharePrice", "Date" } });
            var row = 9;
            var column = 2;
            var items = new List<Module>() { new Module() { Name = "1", Description = "11" }, new Module() { Name = "2", Description = "22" } };
            foreach (var item in items)
            {
                sheet.Cells[row, column].Value = item.Name;
                sheet.Cells[row, column + 1].Value = item.Description;
                row++;
            }
            */

        }

}
