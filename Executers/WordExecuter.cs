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
using _context = ConsoleApp1.ApplicationDbContext;
using ConsoleApp1.Application;



using System;

namespace ConsoleApp1.Helpers;

internal static class WordExecuter
{
    public static string _currentFileName = "";
    public static string _currentFolderName = "";
    public static string _targetDirectory = "";
    public static string _currentFullFilePath = "";

    

    //
    public static async Task ProcessRootDirectoryToFindOtherFoldersWithFiles(string targetDirectory = "")
    {
        targetDirectory = string.IsNullOrEmpty(targetDirectory) ? (string.IsNullOrEmpty(_targetDirectory) ? throw new Exception("Path for word file is not set") : _targetDirectory) : targetDirectory;


        DirectoryInfo dir = new DirectoryInfo(targetDirectory);
        foreach (var folder in dir.GetDirectories())
        {
            var folderName = targetDirectory + "\\" + folder.Name;
           
            if(folderName.Contains("-Готово") || folderName.Contains("-Учебные планы")) {  continue; }
            if (Directory.Exists(folderName))
            {
                await WordExecuter.ProcessFilesByDirectory(folderName);
            }

        }
    }
    public static void PrintAllErrorDocsInDatabase()
    {
        var items = _context.SelectAll<Module>()
            .Where(x => x.Name == "Отсутствует в файле" || string.IsNullOrEmpty(x.Name)
            || x.Speciality == "Отсутствует в файле" || string.IsNullOrEmpty(x.Speciality)
            || x.Description == "Отсутствует в файле" || string.IsNullOrEmpty(x.Description));
        foreach (var item in items)
        {
            Console.WriteLine(item.ToString());
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
        else if(!filename.Contains(".docx"))
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
            if (filename.Contains("Психология управления"))
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
            Console.WriteLine($"File with incorrect format in method 'HandleTablesByWordFile' - {filename}");
            return;
        }
    }

    public static string ConvertDocToDocx(string docPath, string docxPath = "")
    {
        var document = new Spire.Doc.Document();

        document.LoadFromFile(docPath);
        if (string.IsNullOrEmpty(docxPath)) docxPath += docPath + "x";
        
        document.SaveToFile(docxPath, Spire.Doc.FileFormat.Docx);
        return docxPath;
    }


    private static void ProcessTable(Table node)
    {
        Module resultModule = new Module();
        foreach (var row in node.Descendants<TableRow>())
        {
            Module? module;
            try
            {
                module = ProcessRowTopDown(row);
            }
            catch (Exception ex) 
            {
                module = ProcessRowTopDown(row);
            }

            if (module == null) continue;
            if (!string.IsNullOrEmpty(module.Description)) resultModule.Description = module.Description;
            if (!string.IsNullOrEmpty(module.Name)) resultModule.Name = module.Name;
            if (!string.IsNullOrEmpty(module.Speciality)) resultModule.Speciality = module.Speciality;



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
        if (((rowText.Contains("код") || rowText.Contains("название")) && rowText.Contains("специальности")) ||
            (rowText.Contains("название") && rowText.Contains("дисциплины")) ||
            (rowText.Contains("краткое") && rowText.Contains("содержание"))
        )
        {
            

            var rowArray = row.Descendants<TableCell>().ToArray();

            for (int i = 0; i < row.Descendants<TableCell>().Count(); i++)
            {
                var rowCellDescr = rowArray[i].InnerText.ToLower().Replace(" ", "").Replace("-", "");
                if (rowCellDescr.Contains("кратк") && rowCellDescr.Contains("содерж") && i < rowArray.Count())
                {
                    module.Description = rowArray[i + 1].InnerText;

                }
                else if (rowCellDescr.Contains("названи") && rowCellDescr.Contains("дисциплин") && i < rowArray.Count())
                {
                    module.Name = rowArray[i + 1].InnerText;
                }
                else if ((rowCellDescr.Contains("код") || rowCellDescr.Contains("назван")) && rowCellDescr.Contains("специальност") && i < rowArray.Count())
                {
                    module.Speciality = rowArray[i + 1].InnerText;

                }
                else if (i == rowArray.Count() - 1){}

            }
            return module;
            //ExcelExecuter.WriteRow(new List<Models.Module>());
        }
        return null;

    }

    private static void ProcessParagraph(Paragraph node, StringBuilder textBuilder)
    {
        foreach (var text in node.Descendants<Text>())
        {
            //textBuilder.Append('"' + text.InnerText + '"');
        }
    }
  



    //public static Module CreateModule()
    //{

    //}
}
