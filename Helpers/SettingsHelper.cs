using ConsoleApp1.Persistence;
using Microsoft.Extensions.Configuration;

namespace ConsoleApp1.Helpers;

public static class SettingsHelper
{
    public static string Path { 
        get {
            var builder = new ConfigurationBuilder();
            builder.SetBasePath(SettingsHelper.GetProjectPath()).AddJsonFile("appsetting.json", optional: false, reloadOnChange: true);
            IConfiguration config = builder.Build();
            return config?.GetSection("Path")?.Value; 
        } 
    }

    public static string PathToHandledExcelFiles
    {
        get
        {
            var builder = new ConfigurationBuilder();
            builder.SetBasePath(SettingsHelper.GetProjectPath()).AddJsonFile("appsetting.json", optional: false, reloadOnChange: true);
            IConfiguration config = builder.Build();
            return config?.GetSection("PathToHandledExcelFiles")?.Value;
        }
    }
    public static int CountRows => ApplicationDbContext.SelectAll<Models.Module>().Count();

    public static int countUpdatedRows = 0;

   
    public static string GetProjectPath()
    {
        string currentDirectory = Directory.GetCurrentDirectory();
        string projectPath = currentDirectory;

        while (!Directory.GetFiles(projectPath).Any(file => file.EndsWith(".csproj")))
        {
            projectPath = Directory.GetParent(projectPath).FullName;
        }

        return projectPath;
    }

    public static bool IsExistsEqualsDocFileForDocx(Models.Module module)  //if not exists equals. For new created files with .docx format just return because we will check 
    {
        if (module.FullFilePath.Contains(".docx"))
        {
            var newFullPath = module.FullFilePath.Replace(".docx", ".doc");
            if (File.Exists(newFullPath)) return true;
        }
        return false;
    }

   



}
