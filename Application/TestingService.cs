using ConsoleApp1.Helpers;
using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
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
    public static void AreEqualsCountFilesInDirectoryAndCountFilesInDatabase()
    {
        Console.WriteLine("----------------Error count in directory----------------");

        var targetDirectory = SettingsHelper.Path;

        DirectoryInfo dir = new DirectoryInfo(targetDirectory);
        foreach (var folder in dir.GetDirectories())
        {
            var folderName = targetDirectory + "\\" + folder.Name;

            if (folderName.Contains("-Готово") || folderName.Contains("-Учебные планы")) { continue; }
            if (Directory.Exists(folderName))
            {
                string[] fileEntries = Directory.GetFiles(folderName);
                var countFilesInDepartmentFolder = fileEntries.Length;
                var countRowsInDatabaseByDepartment = AllItems.Where(x => x.FullFilePath.Contains(folderName)).Count();
                if (fileEntries.Length != countRowsInDatabaseByDepartment)
                {
                    Console.WriteLine($"Count files in directory {folder.Name}/Count files in database: {countFilesInDepartmentFolder}/{countRowsInDatabaseByDepartment}");
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
        Console.WriteLine($"Count departments in folder: {items.Count()}/{countDepartments}; ");
        
        DirectoryInfo dir = new DirectoryInfo(SettingsHelper.Path);     
        foreach (var folder in dir.GetDirectories())
        {
            if (!listOfDepartmentShortNameInDatabase.Contains(folder.Name) && folder.Name != "-Готово" && folder.Name != "-Учебные планы") Console.Write($" {folder.Name} ");
        }


    }

}
