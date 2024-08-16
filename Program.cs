using ConsoleApp1.Helpers;
using DocumentFormat.OpenXml.Bibliography;

namespace ConsoleApp1;

internal class Program
{
    static async Task Main(string[] args)
    {
        WordExecuter._targetDirectory = "D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\Каталог учебных дисцилин";
        //WordExecuter.ProcessRootDirectoryToFindOtherFoldersWithFiles();
        //ModuleService.PrintAllErrorDocsInDatabase();


        var translator = new GTranslatorAPI.GTranslatorAPIClient();
        var result = await translator.TranslateAsync(GTranslatorAPI.Languages.be, GTranslatorAPI.Languages.ru, "Практычная і функцыянальная стылістыка беларускай мовы");
        System.Console.WriteLine(result.TranslatedText);

    }
    
}
