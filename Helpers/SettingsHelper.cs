using ConsoleApp1.Persistence;

namespace ConsoleApp1.Helpers;

public static class SettingsHelper
{
    private static string _path;
    public static string Path { get => _path ?? throw new Exception("Path is not set"); set { _path = value; } }
    public static int CountRows => ApplicationDbContext.SelectAll<Models.Module>().Count();

    public static int countUpdatedRows = 0;

    public static void SetGlobalPathToProject()
    {

    }

    
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
