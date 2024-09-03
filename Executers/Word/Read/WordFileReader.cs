using ConsoleApp1.Application;
using ConsoleApp1.Helpers;
using ConsoleApp1.Persistence;
using System.Text;

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
using ConsoleApp1.Application;
using ConsoleApp1.Executers.Word;
using System.Xml.Linq;
using DocumentFormat.OpenXml;

namespace ConsoleApp1.Helpers;


public partial class WordFileReader
{
    public string FullFilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public string FolderName { get; set; } = "";
    public bool IsDocxConvertedByDoc { get; set; } = false;
    public bool IsNeedToTranslate { get => FolderName.Contains("БФ") ? true : false; }

    public WordFileReader(string fullPath)
    {
        var arrayOfFiles = fullPath.Split("\\");
        if (arrayOfFiles.Length < 2) { Console.WriteLine($"File by path {fullPath} can`t be handled"); return; }
        FullFilePath = fullPath;
        FolderName = arrayOfFiles[arrayOfFiles.Count() - 2];
        FileName = arrayOfFiles.Last();
    }



    public async Task HandleFile(string filePath)
    {
            
        if (!filePath.Contains(".docx") && filePath.Contains(".doc"))
        {
            filePath = ConvertDocToDocx(filePath);
        }
        else if (!filePath.Contains(".docx"))
        {
            Console.WriteLine($"File with incorrect format in method 'HandleFile' - {filePath}");
            return;
        }
        if(ModuleService.IsFileAlreadyExists(filePath)) return;
            

        await HandleTablesByWordFile(filePath);

    }

    private async Task HandleTablesByWordFile(string filename)
    {

        try
        {
            if (filename.Contains("ЛУ"))
            //if (filename.Contains("D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\Каталог учебных дисцилин\\ЭТИГ\\Каталог учебных дисциплин.docx"))
            {
                Console.Write("");
            }


            using (var wDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filename, false))
            {

                var parts = wDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
                    
                Module resultModule = new Module(FolderName, FileName, FullFilePath, IsDocxConvertedByDoc);
                List<Module> modulesInDifferentTables = new List<Module>();

                var IsTableHandled = false;
                var fullDescriptionForBrokenFilesStringBuilder = new StringBuilder();
                var fullDescriptionForBroken = "";
                if (parts != null)

                {
                    foreach (var node in parts.ChildElements)
                    {

                        if (node is Table)
                        {
                            var res = await ProcessTable((Table)node, resultModule);
                            modulesInDifferentTables.Add((await ProcessTable((Table)node, new Module(FolderName, FileName, FullFilePath, IsDocxConvertedByDoc))).Item1);
                            resultModule = res.Item1;
                            if(res.Item2 != null && res.Item2.Count() > 1)
                                foreach(var item in res.Item2)
                                {
                                    if (!string.IsNullOrEmpty(item.Name)) ModuleService.Create(item);
                                }
                            IsTableHandled = true;
                        }
                        if(((filename.Contains("ДОСИ") || filename.Contains("ТиО") || filename.Contains("ЭиУП"))) && IsTableHandled && node is Paragraph)
                        {
                            //fullDescriptionForBrokenFilesStringBuilder.Append(ProcessParagraph((Paragraph)node));
                            fullDescriptionForBroken += ProcessParagraph((Paragraph)node,true);

                        }  
                    }

                        
                    if(filename.Contains("ТиО") && string.IsNullOrEmpty(fullDescriptionForBroken) && string.IsNullOrEmpty(resultModule.Name))
                    {
                        resultModule = ProcessFormatInTIO(parts, resultModule);
                    }
                    if ((filename.Contains("ДОСИ") || filename.Contains("ТиО") || filename.Contains("ЭиУП")) &&  string.IsNullOrEmpty(resultModule.Description) && !string.IsNullOrEmpty(fullDescriptionForBroken)) resultModule.Description = fullDescriptionForBroken;
                    if (modulesInDifferentTables.Count > 0) { foreach (var item in modulesInDifferentTables) ModuleService.Create(item); return; }
                    ModuleService.Create(resultModule);
                }
            }
        }
        catch (Exception ex)
        {
            if (filename.Contains("ЛУ")) return;
            Console.WriteLine($"File with incorrect format in method 'HandleTablesByWordFile' - {filename} with error {ex.Message}");
            return;
        }
    }

    public  string ConvertDocToDocx(string docPath, string docxPath = "")
    {
        var document = new Spire.Doc.Document();

        document.LoadFromFile(docPath);
        if (string.IsNullOrEmpty(docxPath)) docxPath += docPath + "x";

        document.SaveToFile(docxPath, Spire.Doc.FileFormat.Docx);
        IsDocxConvertedByDoc = true;
        return docxPath;
    }
       
    private string ProcessParagraph(Paragraph node, bool isd = true)
    {
        var res = string.Empty;
        foreach (var text in node.Descendants<Text>())
        {
            res += text.InnerText;        
        }
        return res;
    }
    // :(
    //private IEnumerable<string> ProcessParagraph(Paragraph node)
    //{
            
    //    foreach (var text in node.Descendants<Text>())
    //    {
    //        var ress = text.InnerText;
    //       yield return ress ?? "";
    //    }
    //}
    private Module ProcessFormatInTIO(OpenXmlElement parts, Module module)
    {
        var fullText = string.Empty;
        foreach (var node in parts.ChildElements)
        {
            if (node is Paragraph)
            {
                fullText += ProcessParagraph((Paragraph)node, true);
            }
        }
        int index1 = fullText.IndexOf("Профилизация:", StringComparison.Ordinal);
        int index2 = fullText.IndexOf("Дисциплина ", StringComparison.Ordinal);
        if (index2 != -1) module.Speciality = ExtractTextBetween(fullText, "Специальность ", "Дисциплина") == "" ? ExtractTextBetween(fullText, "Специальности ", "Дисциплина") : ExtractTextBetween(fullText, "Специальность ", "Дисциплина"); 
        else module.Speciality = ExtractTextBetween(fullText, "Специальность ", "Профилизация") == "" ? ExtractTextBetween(fullText, "Специальности ", "Профилизация") : ExtractTextBetween(fullText, "Специальность ", "Профилизация");
           
        module.Name = ExtractTextBetween(fullText, "Дисциплина ", "Содержание");

        int indexDescription = fullText.IndexOf("Содержание:", StringComparison.Ordinal);
        if (indexDescription >= 0)  module.Description = fullText.Substring(indexDescription + "Содержание:".Length);

        return module;
    }

    private string ExtractTextBetween(string input, string start, string end)
    {
        int startIndex = input.IndexOf(start) + start.Length;
        int endIndex = input.IndexOf(end, startIndex);

        if (startIndex >= 0 && endIndex >= 0)
        {
            return input.Substring(startIndex, endIndex - startIndex).Trim();
        }

        return string.Empty;
    }
}



