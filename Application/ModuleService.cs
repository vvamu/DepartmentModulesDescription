using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.InkML;
using MethodTimer;
namespace ConsoleApp1.Application;

internal static class ModuleService
{
    
    public static void Create(Models.Module module)
    {
        var modules = ApplicationDbContext.SelectAll<Models.Module>();
        module.DateLastUpdateFileString = System.IO.File.GetLastWriteTime(module.FullFilePath).ToString();


        var anyFoundModule = modules.FirstOrDefault(x => x.Equals(module));
        if (ConsoleApp1.Helpers.SettingsHelper.IsExistsEqualsDocFileForDocx(module)) return;
        if (anyFoundModule != null && anyFoundModule.IsDifferentDescriptionAndDateLastUpdate(module)) //If I found equals. Check if description and dateLastUpdate
        {
            ApplicationDbContext.Update(module);
            return;
        } 
        ApplicationDbContext.Insert(module);

    }

    

    [Time] //Debug console output
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
