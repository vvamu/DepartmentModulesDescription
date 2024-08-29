using ConsoleApp1.Models;
using GemBox.Document;


namespace ConsoleApp1.Executers.Word.Write;

public class WordFileWriter
{
    public async Task WriteIntoDocumentAsync(ModuleWrite moduleWrite)
    {
        var filePath = moduleWrite.FullPath.Replace("xlsx", "docx");
        //File.Create(filePath);
        
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();
        document.Content.Start.LoadText("Учебная дисциплина (модуль): ", new CharacterFormat() { FontName = "Times New Roman" ,Bold=true, Size = 14});
        document.Content.Start.LoadText(moduleWrite.Name, new CharacterFormat() { FontName = "Times New Roman", Size = 14 });

        document.Sections.Add(
            new Section(document,
                 new Paragraph(document,
                     new Run(document, "Экзамены: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } },
                     new Run(document, moduleWrite.Exams) { CharacterFormat = { FontName = "Times New Roman", Size = 14 } }),
                 new Paragraph(document,
                     new Run(document, "Зачеты: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } },
                     new Run(document, moduleWrite.Receives) { CharacterFormat = { FontName = "Times New Roman", Size = 14 } }),
                 new Paragraph(document,
                     new Run(document, "Всего: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } },
                     new Run(document, $"{moduleWrite.TotalHours} ({moduleWrite.AuditoriumHours} ауд. ч., {moduleWrite.LectureHours} лекционных ч., {moduleWrite.PracticeHours} практических ч., {moduleWrite.SeminarHours} семинарских ч.)") { CharacterFormat = { FontName = "Times New Roman", Size = 14 } }),
                 new Paragraph(document,
                     new Run(document, "Зачетных единиц: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } },
                     new Run(document, moduleWrite.ReceivedUnits) { CharacterFormat = { FontName = "Times New Roman", Size = 14 } }),
                 new Paragraph(document,
                     new Run(document, "Описание учебной дисциплины: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } }),
                 new Paragraph(document,
                     new Run(document, moduleWrite.Description) { CharacterFormat = { FontName = "Times New Roman", Size = 14 } })
            )
        );

        document.Sections.Add(new Section(document));

        //// Insert RTF formatted text at the beginning of the document.
        //var position = document.Content.Start.LoadText(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial Black;}}{\colortbl ;\red255\green128\blue64;}\f0\cf1 This is rich formatted text.}",
        //        LoadOptions.RtfDefault);

        //    // Insert HTML formatted text after the previous text.
        //    position.LoadText("<p style='font-family:Arial Narrow;color:royalblue;'>This is another rich formatted text.</p>",
        //        LoadOptions.HtmlDefault);

        //    // Save Word document to file's path.
         document.Save(filePath);



    }
}
