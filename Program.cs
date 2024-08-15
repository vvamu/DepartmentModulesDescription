
using ConsoleApp1.Helpers;

namespace ConsoleApp1;

internal class Program
{
    static void Main(string[] args)
    {
        WordExecuter._targetDirectory = "D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\Каталог учебных дисцилин";
        WordExecuter.ProcessRootDirectoryToFindOtherFoldersWithFiles();
        //WordExecuter.PrintAllErrorDocsInDatabase();


    }
    
}
