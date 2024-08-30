
using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Diagnostics;
using System.Drawing;
using System.IO.Packaging;
using System.Linq;
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

    public static string? customNormalize(this string? str)
    {
        return str?.ToLower().Replace(" ", "").Replace("ё", "е").Replace("*", "").Replace("\n", "");
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
        bool chkDepartment = false;
        string file_name = Path.GetFileName(file_path);
        Console.Write(file_name);

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string speciality_code = file_name.GetDigits();
            var table = ApplicationDbContext.SelectAll<Module>();

            using (var package = new ExcelPackage(file_path))
            {
                var sheet = package.Workbook.Worksheets[0];
                var description_title_cell = sheet.Cells.Where(x => x.Text
                .Contains("Описание дисциплины"))
                    .FirstOrDefault();
                var description_title_coll = description_title_cell.ToString().GetLetters();
                var description_title_row = description_title_cell.ToString().GetDigits();

                var department_title_cell = sheet.Cells.Where(x => x.Text
                .Contains("Учебная дисциплина закреплена за кафедрой"))
                    .FirstOrDefault();
                var department_title_coll = department_title_cell.ToString().GetLetters();
                var department_title_row = department_title_cell.ToString().GetDigits();


                //Console.WriteLine(sheet.Cells["A"+coll].Value);
                int i = int.Parse(description_title_row) + 5;
                Module module_db;

                while (!sheet.Cells["A" + i].Value?.ToString()
                        .customNormalize()
                        .Contains("Учебные практики".customNormalize())
                        ?? true)
                {
                    if (string.IsNullOrEmpty(sheet.Cells["A" + i].Value?.ToString())) goto End;

                    var description_cell = sheet.Cells[description_title_coll + i];
                    description_cell.Value = null;
                    description_cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    description_cell.Style.Font.Color.SetColor(System.Drawing.Color.Black);

                    string excel_full_module_name = sheet.Cells["A" + i].Value.ToString();
                    string excel_module_name = sheet.Cells["A" + i].Value.ToString().customNormalize();
                    string excel_department_name = sheet.Cells[department_title_coll + i].Value?.ToString().customNormalize();
                    
                    if (excel_department_name is null) goto End;
                    if (excel_module_name.Contains("Курсовая работа".customNormalize())
                        || excel_module_name.Contains("Курсовой проект".customNormalize()))          
                        goto End;

                    var excel_modules = excel_full_module_name.Split("/");

                    if (excel_department_name.Contains("НГПиНХ".customNormalize())
                        || excel_department_name.Contains("ТДП".customNormalize())
                        || excel_department_name.Contains("ЭТИГ".customNormalize())
                        || excel_department_name.Contains("ПЭ".customNormalize())
                        || excel_department_name.Contains("ПКМ".customNormalize()))
                        chkDepartment = true;
                    else
                        chkDepartment = false;

                    foreach (var excel_module in excel_modules) 
                    {                        
                        IEnumerable<Module> selected_modules = table;
                        if (!excel_department_name.Contains("/") && !excel_department_name.Contains(",")) 
                        {
                            selected_modules = table.Where(m => m.DepartmentShortName.customNormalize().Equals(excel_department_name));
                            if (selected_modules.Count() == 0)
                            {
                                description_cell.Value = "Кафедра модуля отсутствует в базе.";
                                description_cell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
                                goto End;
                            }
                        }

                        selected_modules = selected_modules.Where(m => m.Name.customNormalize().Contains(excel_module.customNormalize()));
                        if (selected_modules.Count() == 0)
                        {
                            description_cell.Value += $"В базе нет дисциплины \"{excel_module}\".\n\n";
                            description_cell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
                            continue;
                        }

                        //Проверить, может ли быть больше 1 записи в базе
                        module_db = selected_modules.Where(m => m.Speciality.GetDigits().Contains(speciality_code)).FirstOrDefault();
                        if (module_db is null)
                        {
                            //description_cell.Value += $"Дисциплина \"{excel_module}\" без специальности.\n\n";
                            //description_cell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
                            //continue;

                            module_db = selected_modules.FirstOrDefault();
                        }
                        //module_db = table.Where(
                        //    m => m.Name.ToLower().Replace(" ", "").Contains(excel_module_name)
                        //    && m.Speciality.GetDigits().Contains(speciality_code)).FirstOrDefault()
                        //    ??
                        //    table.Where(m => excel_module_name.Contains(m.Name.ToLower().Replace(" ", ""))
                        //    && m.Speciality.GetDigits().Contains(speciality_code)).FirstOrDefault();

                        description_cell.Value += module_db?.Description + "\n\n";
                    }

                End:
                    i += 1;
                }

                package.Save();
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось изменить файл {file_name} Excel. Ошибка: {ex.Message}");
        }

        Console.WriteLine(chkDepartment ? " - Да" : " - Нет");
    }

    public static List<ModuleWrite> GetExcelDataIntoModel(string file_path)
    {
        string file_name = Path.GetFileName(file_path);
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(file_path))
            {
                var sheet = package.Workbook.Worksheets[0];
                List<ModuleWrite> modules = new List<ModuleWrite>();

                var module_title_cell = sheet.Cells.Where(x => x.Text
                .Contains("Название модуля, учебной дисциплины, курсового проекта (курсовой работы)"))
                    .FirstOrDefault();
                var module_title_coll = module_title_cell.ToString().GetLetters();
                var module_title_row = module_title_cell.ToString().GetDigits();

                int i = int.Parse(module_title_row) + 5;
                while (!sheet.Cells["A" + i].Value?.ToString()
                        .customNormalize()
                        .Contains("Учебные практики".customNormalize())
                        ?? true) 
                {
                    if (string.IsNullOrEmpty(sheet.Cells["A" + i].Value?.ToString())) goto End;
                    if (sheet.Cells["AG" + i].Value is null) goto End;
                    if (sheet.Cells["A" + i].Value.ToString().customNormalize().Contains("Курсовая работа".customNormalize())
                        || sheet.Cells["A" + i].Value.ToString().customNormalize().Contains("Курсовой проект".customNormalize()))
                        goto End;
                    //if (System.Drawing.ColorTranslator.FromHtml(sheet.Cells["AG" + i].Style.Font.Color.LookupColor())
                    //    == System.Drawing.Color.Red) 
                    //    goto End;
                    //Console.WriteLine(System.Drawing.ColorTranslator.FromHtml(sheet.Cells["AG" + i].Style.Font.Color.LookupColor()));

                    ModuleWrite module = new ModuleWrite();
                    module.Name = sheet.Cells["A" + i].Value?.ToString();
                    module.Exams = sheet.Cells["Q" + i].Value?.ToString();
                    module.Receives = sheet.Cells["S" + i].Value?.ToString() +
                                      " " + sheet.Cells["T" + i].Value?.ToString();
                    module.TotalHours = sheet.Cells["U" + i].Value?.ToString();
                    module.AuditoriumHours = sheet.Cells["W" + i].Value?.ToString();
                    module.LectureHours = sheet.Cells["Y" + i].Value?.ToString();
                    module.LabsHours = sheet.Cells["AA" + i].Value?.ToString();
                    module.PracticeHours = sheet.Cells["AC" + i].Value?.ToString();
                    module.SeminarHours = sheet.Cells["AE" + i].Value?.ToString();
                    module.DepartmentShortName = sheet.Cells["AG" + i].Value?.ToString();
                    module.ReceivedUnits = sheet.Cells["AH" + i].Value?.ToString();
                    module.Description = sheet.Cells["AI" + i].Value?.ToString();
                    module.FileName = file_name.Replace(".xlsx", "");
                    module.FullPath = file_path.Replace(".xlsx", "");

                    modules.Add(module);
                    
                    End:
                        i += 1;
                }

                return modules;
            }
        }
        catch (Exception ex) 
        {
            Console.WriteLine($"Не удалось считать файл {file_name} Excel. Ошибка: {ex.Message}");
            return null;
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
