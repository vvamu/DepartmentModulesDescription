using ConsoleApp1.Helpers;
using ConsoleApp1.Persistence;
namespace ConsoleApp1.Application;
internal static class ModuleService
{
    
    public static void Create(Models.Module module)
    {
        var modules = ApplicationDbContext.SelectAll<Models.Module>();
        module.DateLastUpdateFileString = System.IO.File.GetLastWriteTime(module.FullFilePath).ToString();


        var anyFoundModule = modules.FirstOrDefault(x => x.Equals(module));
        if (ConsoleApp1.Helpers.SettingsHelper.IsExistsEqualsDocFileForDocx(module)) return;
        
        if (anyFoundModule != null && !anyFoundModule.IsDifferentDescriptionAndDateLastUpdate(module)) //If I found equals. Check if description and dateLastUpdate
        {
            ApplicationDbContext.Update(module);
            SettingsHelper.countUpdatedRows++;
            return;
        } 
        ApplicationDbContext.Insert(module);

    }

}
