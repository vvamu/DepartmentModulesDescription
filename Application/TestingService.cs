using ConsoleApp1.Helpers;
using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using MethodTimer;
using System.Linq;
namespace ConsoleApp1.Application;

internal static class TestingService
{
    private static IQueryable<Module> _allItems;
    private static IQueryable<Module> AllItems
    {
        get 
        { 
            if(_allItems == null) _allItems = ApplicationDbContext.SelectAll<Models.Module>().AsQueryable(); 
            return _allItems; 
        }
    }
    public static void AreEqualsCountFilesInDirectoryAndCountDepartmentsAndCountRowsInDatabase()
    {
        Console.WriteLine("----------------Error count in directory----------------");

        var targetDirectory = SettingsHelper.Path;

        DirectoryInfo dir = new DirectoryInfo(targetDirectory);
        foreach (var folder in dir.GetDirectories())
        {
            var folderName = targetDirectory + "\\" + folder.Name;
            List<string> filesToCheck = new List<string>();


            if (folderName.Contains("-Готово") || folderName.Contains("-Учебные планы")) { continue; }
            if (Directory.Exists(folderName))
            {
                string[] fileEntries = Directory.GetFiles(folderName);
                var departments = AllItems.GroupBy(x => new { x.DepartmentShortName, x.FileName }).Select(x => x.Key.DepartmentShortName).Where(x => x == folder.Name);
                var rowsInDatabaseByDepartment = AllItems.Where(x => x.FullFilePath != null && x.FullFilePath.Contains(folderName));
                
                var countFilesInDepartmentFolder = fileEntries.Length;
                var countDepartments = departments.Count();
                var countRowsInDatabaseByDepartment = rowsInDatabaseByDepartment.Count();

                var strNeedToCheck = "";
                if (countFilesInDepartmentFolder != countDepartments)
                {
                    strNeedToCheck = ". Need to check files: " + (countFilesInDepartmentFolder - countDepartments);
                    var filenamesNotInDatabase = fileEntries.Where(x => !AllItems.Any(c => c.FullFilePath == x)).ToList();
                    filesToCheck.AddRange(filenamesNotInDatabase);
                }
                if (fileEntries.Length != countRowsInDatabaseByDepartment)
                {
                    Console.WriteLine($"Count files in directory {folder.Name}/Count departments/Count rows in database: {countFilesInDepartmentFolder}/{countDepartments}/{countRowsInDatabaseByDepartment}{strNeedToCheck}" );
                }

                if (filesToCheck.Count() != 0)
                {
                    var oneOfItem = filesToCheck.FirstOrDefault() ?? "";
                    if (oneOfItem.Contains("МКиТП") || oneOfItem.Contains("НГПиНХ") || oneOfItem.Contains("ТДП") || oneOfItem.Contains("ЭТИГ")) continue;
                    Console.WriteLine("-----");
                    foreach (var item in filesToCheck)
                    {
                        
                        Console.WriteLine(item);
                    }
                    Console.WriteLine("-----");
                }
            }
        }
    }

    [Time] //Debug console output
    public static void PrintAllErrorDocsInDatabaseBySettingsHelper()
    {
        var items = AllItems
            .Where(x => x.Name == "Отсутствует в файле" || string.IsNullOrEmpty(x.Name)
            || x.Speciality == "Отсутствует в файле" || string.IsNullOrEmpty(x.Speciality)
            || x.Description == "Отсутствует в файле" || string.IsNullOrEmpty(x.Description));

        Console.WriteLine("----------------Error files----------------");
        foreach (var item in items)
        {
            Console.WriteLine(item.ToString());        
        }  
    }

    public static void CheckCountDepartments()
    {
        var items = AllItems.GroupBy(x=>x.DepartmentShortName);
        var listOfDepartmentShortNameInDatabase = items.Select(x => x.Key).ToList();
        var countDepartments = 48;
        
        DirectoryInfo dir = new DirectoryInfo(SettingsHelper.Path);    
        var countFolders = dir.GetDirectories().Count() - 2;
        Console.WriteLine($"Count departments in database/folder/need to be: {items.Count()}/{countFolders}/{countDepartments}; ");
        foreach (var folder in dir.GetDirectories())
        {
            if (!listOfDepartmentShortNameInDatabase.Contains(folder.Name) && folder.Name != "-Готово" && folder.Name != "-Учебные планы") Console.Write($" {folder.Name} ");
        }


    }

    public static void PrintAll()
    {

    }
}
