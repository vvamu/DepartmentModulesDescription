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

namespace ConsoleApp1.Helpers;

internal static class WordExecuter
{
    public static string _currentFileName = "";
    public static string _currentFolderName = "";
    public static string _targetDirectory = "";

    //
    public static void ProcessRootDirectoryToFindOtherFoldersWithFiles(string targetDirectory = "")
    {
        targetDirectory = string.IsNullOrEmpty(targetDirectory) ? (string.IsNullOrEmpty(_targetDirectory) ? throw new Exception("Path for word file is not set") : _targetDirectory) : targetDirectory;


        DirectoryInfo dir = new DirectoryInfo(targetDirectory);
        foreach (var folder in dir.GetDirectories())
        {
            var folderName = targetDirectory + "\\" + folder.Name;
            if (Directory.Exists(folderName))
            {
                WordExecuter.ProcessFilesByDirectory(folderName);
            }

        }
    }
    public static void ProcessFilesByDirectory(string targetDirectory)
    {
        targetDirectory = string.IsNullOrEmpty(targetDirectory) ? (string.IsNullOrEmpty(_targetDirectory) ? throw new Exception("Path for word file is not set") : _targetDirectory) :  targetDirectory;
        string[] fileEntries = Directory.GetFiles(targetDirectory);
        foreach (string fileName in fileEntries)
        {
            _currentFolderName = targetDirectory.Split("\\").Last();
            _currentFileName = fileName.Split("\\").Last().Replace(".docx", "");
            ProcessTablesByWordFile(fileName);
        }
    }
    public static async Task ProcessTablesByWordFile(string filename)
    {

        //TODO if exists not one table in document

        if (!filename.Contains(".docx"))
        {
            var module = new Module() { Description = "Файл содержит некорректный формат" };
            module.DepartmentShortName = _currentFolderName;
            module.Name = _currentFileName;

            _context.Insert(module);
            return;
        }
        
        try
        {

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filename, false))
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
        catch(Exception ex) { }

        /*
        var fullText = textBuilder.ToString();
        int index = fullText.IndexOf("3 , Краткое содержание учебной дисциплины  , ");
        string result = "";
        if (index >= 0)
        {
            index += "3 , Краткое содержание учебной дисциплины  , ".Length;
            result = fullText.Substring(index);
            result = result.Substring(0, result.Length - 7);
            Console.WriteLine(result);
        }
        else
        {
            Console.WriteLine("Substring not found in fullText string.");
            //Написать в excel 
        }
        return result;*/
    }

    private static void ProcessTable(Table node)
    {
        foreach (var row in node.Descendants<TableRow>())
        {
            CreateModuleByWordFile(row);
            /*

                foreach (var para in cell.Descendants<Paragraph>())
                {
                    ProcessParagraph(para, textBuilder);
                }
            */
        }
    }
    private static void CreateModuleByWordFile(TableRow row)
    {
        var module = new Module();
        var rowText = row.InnerText.ToLower().Replace(" ", "").Replace("-", "");
        if (//(rowText.Contains("код") || rowText.Contains("название") && rowText.Contains("специальности"))
            //(rowText.Contains("название") && rowText.Contains("дисциплины"))
            (rowText.Contains("краткое") && rowText.Contains("содержание"))
        )
        {
            module.DepartmentShortName = _currentFolderName;
            module.Name = _currentFileName;

            var rowArray = row.Descendants<TableCell>().ToArray();

            for (int i = 0; i < row.Descendants<TableCell>().Count(); i++)
            {
                var rowCellDescr = rowArray[i].InnerText.ToLower().Replace(" ", "").Replace("-", "");
                if (rowCellDescr.Contains("краткое") && rowCellDescr.Contains("содержание") && i < row.Descendants<TableCell>().Count())
                {
                    module.Description = rowArray[i+1].InnerText;
                }
                else
                {
                    module.Description = "Описание отсутствует либо находится не на своей позиции";
                }
            }
            _context.Insert(module);
            //ExcelExecuter.WriteRow(new List<Models.Module>());
        }

    }

    private static void ProcessParagraph(Paragraph node, StringBuilder textBuilder)
    {
        foreach (var text in node.Descendants<Text>())
        {
            //textBuilder.Append('"' + text.InnerText + '"');
        }
    }
    private static void PrintErrorDocType(string filePath)
    {
        //Error
    }

    //public static Module CreateModule()
    //{

    //}
}
