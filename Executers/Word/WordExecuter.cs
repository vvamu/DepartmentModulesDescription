using ConsoleApp1.Executers.Word.Write;
using ConsoleApp1.Helpers;
using ConsoleApp1.Persistence;

namespace ConsoleApp1.Executers.Word;

public partial class WordExecuter
{
    private static string _targetDirectory = "";
    public static async Task ProcessRootDirectoryToFindOtherFoldersWithFilesToRead(string targetDirectory = "")
    {
        _targetDirectory = string.IsNullOrEmpty(targetDirectory) ? string.IsNullOrEmpty(SettingsHelper.Path) ? throw new Exception("Path for word file is not set") : SettingsHelper.Path : targetDirectory;

        DirectoryInfo dir = new DirectoryInfo(_targetDirectory);

        foreach (var folder in dir.GetDirectories())
        {
            var folderName = _targetDirectory + "\\" + folder.Name;
            if (folderName.Contains("-Готово") || folderName.Contains("-Учебные планы")) { continue; }
            if (Directory.Exists(folderName))
            {
                await ProcessFilesByDirectory(folderName);
            }
        }
    }
    public static async Task ProcessDirectoryToWrite()
    {
        DirectoryInfo dir = new DirectoryInfo(SettingsHelper.PathToHandledExcelFiles);
        var wordFileWriter = new WordFileWriter();
        var files = dir.GetFiles();
        var countNotHandledFiles = 0;

        foreach (var file in files)
        {
            {
                try
                {
                    var itemsByExcelFile = ExcelExecuter.GetExcelDataIntoModel(file.FullName);
                    if (itemsByExcelFile == null || itemsByExcelFile.Count() < 1) continue;
                    await wordFileWriter.WriteIntoDocumentAsync(itemsByExcelFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error with file {file.Name}");
                    countNotHandledFiles++;
                }
                finally
                {
                    if (file.Name.Contains(".docx"))
                        Console.WriteLine($"WordExecuter : File {file.Name} write data by excel");
                }
            }
        }
        Console.WriteLine("");
        Console.WriteLine("Results of write into word:");
        Console.WriteLine($"Count excel files in folder/Count docx files was handled: {files.Count()}/{files.Count() - countNotHandledFiles}");
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
