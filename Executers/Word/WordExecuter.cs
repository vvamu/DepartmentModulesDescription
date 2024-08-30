using ConsoleApp1.Executers.Word.Write;
using ConsoleApp1.Helpers;
using ConsoleApp1.Models;
using ConsoleApp1.Persistence;
using DocumentFormat.OpenXml.Drawing;

namespace ConsoleApp1.Executers.Word;

public partial class WordExecuter
{

    private static string _targetDirectory = "";
    private static WordExecuter? instance; private WordExecuter(string targetDirectory) { _targetDirectory = targetDirectory; }
    public static WordExecuter getInstance(string targetDirectory) => instance ?? new WordExecuter(targetDirectory);
    //

    public static async Task ProcessRootDirectoryToFindOtherFoldersWithFilesToRead(string targetDirectory = "")
    {
        targetDirectory = string.IsNullOrEmpty(targetDirectory) ? string.IsNullOrEmpty(_targetDirectory) ? throw new Exception("Path for word file is not set") : _targetDirectory : targetDirectory;

        DirectoryInfo dir = new DirectoryInfo(targetDirectory);

        foreach (var folder in dir.GetDirectories())
        {
            var folderName = targetDirectory + "\\" + folder.Name;
            if (folderName.Contains("-Готово") || folderName.Contains("-Учебные планы")) { continue; }


            if (Directory.Exists(folderName))
            {
                //var folderd = new WordFolderExecuter();
                await ProcessFilesByDirectory(folderName);
            }

        }
    }
    public static async Task ProcessDirectoryToWrite(string path = "")
    {

        DirectoryInfo dir = string.IsNullOrEmpty(path) ? new DirectoryInfo(System.IO.Path.Combine(_targetDirectory, "-Готово")) : new DirectoryInfo(path);

        foreach (var file in dir.GetFiles())
        {
            
            var wordFileWriter = new WordFileWriter();

            //if (Directory.Exists(file.FullName))
            {
                try
                {
                    if (!file.Name.Contains(".docx")) continue; 
                    //await ProcessFilesByDirectory(folderName);
                    var itemsByExcelFile = ExcelExecuter.GetExcelDataIntoModel(file.FullName);
                    if (itemsByExcelFile == null || itemsByExcelFile.Count() < 1) continue;
                    await wordFileWriter.WriteIntoDocumentAsync(itemsByExcelFile);
                }
                catch (Exception ex) { 
                    Console.WriteLine($"Error with file {file.Name}"); }
            }

        }
    }

    public static async Task RemoveAllDocsFormattedByDoc()
    {
        var modules = ApplicationDbContext.SelectAll<Models.Module>().Where(x => x.IsDocxConvertedByDoc).Select(x => x.FullFilePath);
        foreach (var filePath in modules)
        {
            var filePathDocx = filePath + "x";
            File.Delete(filePathDocx);
            if (Directory.Exists(filePathDocx))
            {
                var fileDel = new FileInfo(filePathDocx);
                fileDel.Delete();
            }

        }

    }


    private static async Task ProcessFilesByDirectory(string targetDirectory)
    {
        targetDirectory = string.IsNullOrEmpty(targetDirectory) ? string.IsNullOrEmpty(_targetDirectory) ? throw new Exception("Path for word file is not set") : _targetDirectory : targetDirectory;
        string[] fileEntries = Directory.GetFiles(targetDirectory);

        foreach (string fileName in fileEntries)
        {
            var fileExecuter = new WordFileReader(fileName); // This should now find the WordFolderFileExecuter class
            await fileExecuter.HandleFile(fileName);
        }

        await RemoveAllDocsFormattedByDoc();
    }


}
