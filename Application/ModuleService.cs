using ConsoleApp1.Models;
using DocumentFormat.OpenXml.InkML;

namespace ConsoleApp1.Application;

internal static class ModuleService
{
    
    public static void Create(Module module)
    {
        var modules = ApplicationDbContext.SelectAll<Module>();
        module.DateLastUpdateFileString = System.IO.File.GetLastWriteTime(module.FullFilePath).ToString();

       

        if (modules.Any(x => x.Equals(module))) return;
        if (module.FullFilePath.Contains(".docx"))
        {
            var newFullPath = module.FullFilePath.Replace(".docx", ".doc");
            //module.DateLastUpdateFileString = System.IO.File.GetLastWriteTime(newFullPathh).ToString();

            if (File.Exists(newFullPath)) return;
        }


        #region StopWatch
        // var stopwatch = new System.Diagnostics.Stopwatch();
        // stopwatch.Start();
        //stopwatch.Stop();
        //Console.WriteLine($"Время выполнения Equals: {stopwatch.Elapsed}");

        //if (modules.Any(x => x.Compare(module))) return;
        #endregion
        //var any = modules.FirstOrDefault(x=>x.FullFilePath == module.FullFilePath);


        ApplicationDbContext.Insert(module);

    }

    public static void PrintAllErrorDocsInDatabase()
    {
        var items = ApplicationDbContext.SelectAll<Module>()
            .Where(x => x.Name == "Отсутствует в файле" || string.IsNullOrEmpty(x.Name)
            || x.Speciality == "Отсутствует в файле" || string.IsNullOrEmpty(x.Speciality)
            || x.Description == "Отсутствует в файле" || string.IsNullOrEmpty(x.Description));
        foreach (var item in items)
        {
            Console.WriteLine(item.ToString());
        }
    }



}
