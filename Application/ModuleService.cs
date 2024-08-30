using ConsoleApp1.Helpers;
using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
namespace ConsoleApp1.Application;
internal static class ModuleService
{
    private static IQueryable<Models.Module> _modules;
    private static IQueryable<Models.Module> Modules => _modules ?? ApplicationDbContext.SelectAll<Models.Module>().AsQueryable();
    public static void Create(Models.Module module)
    {
        if (string.IsNullOrEmpty(module.Speciality) && string.IsNullOrEmpty(module.Name)) return;
        if (string.IsNullOrEmpty(module.DepartmentShortName)) return;
        if (!string.IsNullOrEmpty(module.Speciality))
            module.LowerAndTrimSpeciality = module.Speciality.Trim().ToLower().Replace(" ", "");

        var modules = ApplicationDbContext.SelectAll<Models.Module>();
        if (module.FullFilePath != null && ConsoleApp1.Helpers.SettingsHelper.IsExistsEqualsDocFileForDocx(module)) return;
        if (!string.IsNullOrEmpty(module.FullFilePath)) module.DateLastUpdateFileString = System.IO.File.GetLastWriteTime(module.FullFilePath).ToString();
        var anyFoundModule = modules.FirstOrDefault(x => x.Equals(module));
        if (anyFoundModule != null)
        {
            if ((anyFoundModule?.DateLastUpdateFileString == module?.DateLastUpdateFileString
                && anyFoundModule?.Description == module?.Description)
                || string.IsNullOrEmpty(module?.Description)) return;
            else
            {
                if (IsFileNeedToNotUpdate(module)) return;
                module.Id = anyFoundModule.Id;
                ApplicationDbContext.Update(module);
                SettingsHelper.countUpdatedRows++;
                //Console.WriteLine("--------------------------");
                //Console.WriteLine($"Updated : {module.DepartmentShortName} {module.FileName} {module?.Name}");
                //Console.WriteLine($"From: {module?.Description} - {module?.DateLastUpdateFileString}");
                //Console.WriteLine($"To : {anyFoundModule.Description} {anyFoundModule.DateLastUpdateFileString}");
                //Console.WriteLine("--------------------------");

                return;
            }
        }
        if (string.IsNullOrEmpty(module.Speciality)) module.Speciality = "Отсутствует в файле";
        if (string.IsNullOrEmpty(module.Name)) module.Name = "Отсутствует в файле";
        if (string.IsNullOrEmpty(module.Description)) module.Description = "Отсутствует в файле";
        ApplicationDbContext.Insert(module);
    }
    public static void Create(string department, string speciality, string name, string description, string fullFulePath = "", string fileName = "", bool isDocxConvertedByDoc = false, string dateLastUpdateFileString = "")
    {
        var module = new Models.Module(department, speciality, name, description, fullFulePath, fileName, isDocxConvertedByDoc);
        Create(module);
    }

    public static bool IsFileAlreadyExists(string filePath)
    {

        var dateLastUpdateFileString = System.IO.File.GetLastWriteTime(filePath).ToString();
        var equals = Modules.FirstOrDefault(x => x.FullFilePath == filePath && x.DateLastUpdateFileString == dateLastUpdateFileString);
        if (equals != null) return true;
        return false;

    }

    public static void RemoveRepetitionsRows()
    {
        var modules = ApplicationDbContext.SelectAll<Models.Module>();
        /*
        var distinctModules = modules
        .GroupBy(m => m, new ModuleComparer())
        .Select(g => g.OrderByDescending(m => !string.IsNullOrEmpty(m.Description)).First())
        .ToList();*/



        var distinctModules = modules
            .GroupBy(m => new { m.Speciality, m.Name, m.DepartmentShortName })
            .SelectMany(g => g.OrderByDescending(m => !string.IsNullOrEmpty(m.Description)).Take(1))
            .ToList();

        // Update the database
        foreach (var module in modules)
        {
            if (!distinctModules.Contains(module) || string.IsNullOrEmpty(module.Description))
            {
                ApplicationDbContext.Remove(module);
            }
        }
    }

    public static bool IsFileNeedToNotUpdate(Module module)
    {
        if (!string.IsNullOrEmpty(module.FileName) &&
            (module.FileName.Contains("Загрузочно-формировочное оборудование лесопромышленных предприятий.docx") ||
             module.FileName.Contains("Загрузочно-формировочное оборудование лесопромышленных предриятий.docx") ||
             module.FileName.Contains("ЛПС_Декоративная дендрология(1).docx") ||
             module.FileName.Contains("ЛПС_Декоративная дендрология.docx")))
        {
            return true;
        }
        if (!string.IsNullOrEmpty(module.Name) && module.Name.Contains("Загрузочно-формировочное оборудование лесопромышленных предприятий")
            || module.Name.Contains("Загрузочно-формировочное оборудование лесопромышленных предриятий")
            || module.Name.Contains("ЛПС_Декоративная дендрология(1)")
            || module.Name.Contains("Декоративная дендрология")

            ) return true;
        return false;

    }
    public static void RemoveRepetitionsFiles()
    {

    }
}
