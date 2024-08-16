using DocumentFormat.OpenXml.Packaging;

using System.Text;
using static System.Net.Mime.MediaTypeNames;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using System.Text.Json;
using ConsoleApp1.Models;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.InkML;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using _context = ConsoleApp1.Persistence.ApplicationDbContext;
using ConsoleApp1.Application;


using System;

namespace ConsoleApp1.Helpers;

internal static class WordExecuter
{


    public static string _currentFileName = "";
    public static string _currentFolderName = "";
    public static string _targetDirectory = "";
    public static string _currentFullFilePath = "";
    public static bool IsNeedToTranslate { get => _currentFolderName.Contains("БФ") ? true : false; }

    //private static WordExecuter? instance; private WordExecuter() { }
    //public static WordExecuter getInstance => instance ?? new WordExecuter();


    //
    public static async Task ProcessRootDirectoryToFindOtherFoldersWithFiles(string targetDirectory = "")
    {
        targetDirectory = string.IsNullOrEmpty(targetDirectory) ? (string.IsNullOrEmpty(_targetDirectory) ? throw new Exception("Path for word file is not set") : _targetDirectory) : targetDirectory;


        DirectoryInfo dir = new DirectoryInfo(targetDirectory);
        foreach (var folder in dir.GetDirectories())
        {
            var folderName = targetDirectory + "\\" + folder.Name;

            if (folderName.Contains("-Готово") || folderName.Contains("-Учебные планы")) { continue; }
            if (Directory.Exists(folderName))
            {
                await WordExecuter.ProcessFilesByDirectory(folderName);
            }

        }
    }
    private static async Task ProcessFilesByDirectory(string targetDirectory)
    {
        targetDirectory = string.IsNullOrEmpty(targetDirectory) ? (string.IsNullOrEmpty(_targetDirectory) ? throw new Exception("Path for word file is not set") : _targetDirectory) : targetDirectory;
        string[] fileEntries = Directory.GetFiles(targetDirectory);
        foreach (string fileName in fileEntries)
        {
            _currentFolderName = targetDirectory.Split("\\").Last();
            _currentFileName = fileName.Split("\\").Last();
            _currentFullFilePath = fileName;
            await HandleFile(fileName);
        }
    }

    private static async Task HandleFile(string filename)
    {
        if (filename.Contains("Биологические системы и методы в водоподготовке и очистке сточных вод"))
        {
            Console.Write("");
        }
        if (!filename.Contains(".docx") && filename.Contains(".doc"))
        {
            filename = ConvertDocToDocx(filename);
        }
        else if (!filename.Contains(".docx"))
        {
            Console.WriteLine($"File with incorrect format in method 'HandleFile' - {filename}");
            return;
        }
        await HandleTablesByWordFile(filename);

    }

    private static async Task HandleTablesByWordFile(string filename)
    {

        try
        {
            if (filename.Contains("Беларуская"))
            //if (filename.Contains("D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\Каталог учебных дисцилин\\ЭТИГ\\Каталог учебных дисциплин.docx"))
            {
                Console.Write("");
            }
            using (var wDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filename, false))
            {

                var parts = wDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
                if (parts != null)

                {
                    foreach (var node in parts.ChildElements)
                    {
                        /*
                        if (node is Paragraph)
                        {
                            ProcessParagraph((Paragraph)node, textBuilder);
                            textBuilder.AppendLine("");
                        }*/


                        if (node is Table)
                        {
                            ProcessTable((Table)node);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"File with incorrect format in method 'HandleTablesByWordFile' - {filename} with error {ex.Message}");
            return;
        }
    }

    //private static bool IsNeedConvertToDocs()
    //{
    //    var anyFoundModule = modules.FirstOrDefault(x => x.Equals(module));
    //    if (anyFoundModule != null && anyFoundModule.IsDifferentDescriptionAndDateLastUpdate(module)) //If I found equals. Check if description and dateLastUpdate
    //    {
    //        ApplicationDbContext.Update(module);
    //        return;
    //    }
    //    //if not exists equals. For new created files with .docx format just return because we will check 
    //    if (module.FullFilePath.Contains(".docx"))
    //    {
    //        var newFullPath = module.FullFilePath.Replace(".docx", ".doc");
    //        if (File.Exists(newFullPath)) return;
    //    }
    //    return true;
    //}

    public static string ConvertDocToDocx(string docPath, string docxPath = "")
    {
        var document = new Spire.Doc.Document();

        document.LoadFromFile(docPath);
        if (string.IsNullOrEmpty(docxPath)) docxPath += docPath + "x";

        document.SaveToFile(docxPath, Spire.Doc.FileFormat.Docx);
        return docxPath;
    }

    private static void RemoveAllDocsFormattedByDoc()
    {

    }


    private static async Task ProcessTable(Table node)
    {
        Module resultModule = new Module();
        foreach (var row in node.Descendants<TableRow>())
        {
            Module? module;
            ///TODO
            try
            {
                module = ProcessRowTopDown(row);
            }
            catch (Exception ex)
            {
                module = ProcessRowLeftRight(row);
            }

            if (module == null) continue;
            if (!string.IsNullOrEmpty(module.Description)) resultModule.Description = module.Description;
            if (!string.IsNullOrEmpty(module.Name)) resultModule.Name = module.Name;
            if (!string.IsNullOrEmpty(module.Speciality)) resultModule.Speciality = module.Speciality;

            if (IsNeedToTranslate) await TranslationHelper.TranslateToRu(module);

            /*

                foreach (var para in cell.Descendants<Paragraph>())
                {
                    ProcessParagraph(para, textBuilder);
                }
            */
        }
        if (string.IsNullOrEmpty(resultModule.Speciality))
            resultModule.Speciality = "Отсутствует в файле";
        if (string.IsNullOrEmpty(resultModule.Name))
            resultModule.Name = "Отсутствует в файле";
        if (string.IsNullOrEmpty(resultModule.Description))
            resultModule.Description = "Отсутствует в файле";


        resultModule.DepartmentShortName = _currentFolderName;
        resultModule.FileName = _currentFileName;
        resultModule.FullFilePath = _currentFullFilePath;

        ModuleService.Create(resultModule);
    }
    private static Module ProcessRowTopDown(TableRow row)
    {
        var module = new Module();
        var rowText = row.InnerText.ToLower().Replace(" ", "").Replace("-", "");
        if (ModuleWordHelper.IsDescriptionRow(rowText) || ModuleWordHelper.IsNameRow(rowText) || ModuleWordHelper.IsSpecialityRow(rowText))
        {
            var rowArray = row.Descendants<TableCell>().ToArray();

            for (int i = 0; i < row.Descendants<TableCell>().Count() - 1; i++)
            {
                var rowCellDescr = rowArray[i].InnerText.ToLower().Replace(" ", "").Replace("-", "");
                if (ModuleWordHelper.IsDescriptionRow(rowCellDescr))
                {
                    module.Description = rowArray[i + 1].InnerText;

                }
                else if (ModuleWordHelper.IsNameRow(rowCellDescr))
                {
                    module.Name = rowArray[i + 1].InnerText;
                }
                else if (ModuleWordHelper.IsSpecialityRow(rowCellDescr))
                {
                    module.Speciality = rowArray[i + 1].InnerText;
                }
                //else if (i == rowArray.Count() - 1){}

            }
            return module;
            //ExcelExecuter.WriteRow(new List<Models.Module>());
        }
        return null;

    }

    //TODO
    private static Module ProcessRowLeftRight(TableRow row)
    {
        return new Module();
    }
    private static void ProcessParagraph(Paragraph node, StringBuilder textBuilder)
    {
        foreach (var text in node.Descendants<Text>())
        {
            //textBuilder.Append('"' + text.InnerText + '"');
        }
    }


    
}
