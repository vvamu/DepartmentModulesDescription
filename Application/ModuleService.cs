using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.InkML;
using MethodTimer;
using System.Reflection;
namespace ConsoleApp1.Application;

internal static class ModuleService
{
    
    public static void Create(Models.Module module)
    {
        var modules = ApplicationDbContext.SelectAll<Models.Module>();
        module.DateLastUpdateFileString = System.IO.File.GetLastWriteTime(module.FullFilePath).ToString();

       

        if (modules.Any(x => x.Equals(module))) return;
        if (module.FullFilePath.Contains(".docx"))
        {
            var newFullPath = module.FullFilePath.Replace(".docx", ".doc");
            //module.DateLastUpdateFileString = System.IO.File.GetLastWriteTime(newFullPathh).ToString();

            if (File.Exists(newFullPath)) return;
        }

        //if (modules.Any(x => x.Compare(module))) return
        //var any = modules.FirstOrDefault(x=>x.FullFilePath == module.FullFilePath);


        ApplicationDbContext.Insert(module);

    }
    //[Time] //Выво в консоль отладки
    [Time]
    public static void PrintAllErrorDocsInDatabase()
    {
        var items = ApplicationDbContext.SelectAll<Models.Module>()
            .Where(x => x.Name == "Отсутствует в файле" || string.IsNullOrEmpty(x.Name)
            || x.Speciality == "Отсутствует в файле" || string.IsNullOrEmpty(x.Speciality)
            || x.Description == "Отсутствует в файле" || string.IsNullOrEmpty(x.Description));
        foreach (var item in items)
        {
            Console.WriteLine(item.ToString());
        }

    }
}
