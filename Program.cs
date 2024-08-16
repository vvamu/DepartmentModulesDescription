using ConsoleApp1.Application;
using ConsoleApp1.Helpers;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp1;

internal class Program 
{
    static async Task Main(string[] args)
    {
        WordExecuter._targetDirectory = "D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\Каталог учебных дисцилин";
        //WordExecuter.ProcessRootDirectoryToFindOtherFoldersWithFiles();
        ModuleService.PrintAllErrorDocsInDatabase();

        
    }

    
}
