using ConsoleApp1.Application;
using ConsoleApp1.Helpers;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp1;

internal class Program 
{
    public static string path => SettingsHelper.Path;
    static async Task Main(string[] args)
    {
        SettingsHelper.Path = "D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\Каталог учебных дисцилин";
        #region Word
        var wordExecuter = WordExecuter.getInstance(path);
        var countRowsBefore = SettingsHelper.CountRows;

        //await WordExecuter.ProcessRootDirectoryToFindOtherFoldersWithFiles();

        var countRowsAfter = SettingsHelper.CountRows;
        #endregion

        #region Testing
        Console.WriteLine($"Count rows before/after : {countRowsBefore}/{countRowsAfter} ");
        Console.WriteLine($"Count rows updated: {SettingsHelper.countUpdatedRows}");


        TestingService.PrintAllErrorDocsInDatabaseBySettingsHelper();
        //TestingService.AreEqualsCountFilesInDirectoryAndCountFilesInDatabase();
        //TestingService.CheckCountDepartments();

        #endregion



        #region Excel
        //ExcelExecuter.EditSpecialityDescriptions("6-05-0211-06 Example.xlsx");
        //ExcelExecuter.EditDirSpecialities("D:\\Ilya\\2024\\08\\Project\\Каталог учебных дисцилин\\-Готово_TEST");
        #endregion


        //GetFilesFromSubfolder("ТДиД");
        //GetFilesFromSubfolder("ФиП");



    }

    static void GetFilesFromSubfolder(string subfolder)
    {
        string rootFolderPath = Path.Combine(path, "ТДиД");
        List<string> allFilePaths = GetAllFilesInSubfolders(rootFolderPath , rootFolderPath);

        foreach (var filePath in allFilePaths)
        {
            Console.WriteLine(filePath);
        }
    }
    static List<string> GetAllFilesInSubfolders(string rootFolderPath , string directionToFiles)
    {
        List<string> filePaths = new List<string>();
        try
        {
            
            foreach (string filePath in Directory.GetFiles(rootFolderPath))
            {
                filePaths.Add(filePath);

                var resultPath = Path.Combine(directionToFiles, Path.GetFileName(filePath));
                File.Move(filePath, resultPath);
            }

            foreach (string subfolder in Directory.GetDirectories(rootFolderPath))
            {
                filePaths.AddRange(GetAllFilesInSubfolders(subfolder, directionToFiles));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }

        return filePaths;
    }


}
