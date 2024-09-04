using ConsoleApp1.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;

namespace ConsoleApp1.Executers.Word.Write;

public class WordFileWriter
{
    public async Task WriteIntoDocumentAsync(List<ModuleWrite> moduleWrites)
    {
        var filePath = moduleWrites.FirstOrDefault().FullPath +  ".docx";
        if (File.Exists(filePath)) 
            File.Delete(filePath);
        using (File.Create(filePath)) { }
        System.Threading.Thread.Sleep(100);

        using (MemoryStream mem = new MemoryStream())
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                var docBody = new DocumentFormat.OpenXml.Wordprocessing.Body();
                moduleWrites = moduleWrites.DistinctBy(x=>x.Name).ToList();
                /*
                                Paragraph p = new Paragraph();
                                // Run 1
                                Run r1 = new Run();
                                Text t1 = new Text("Pellentesque ") { Space = SpaceProcessingModeValues.Preserve };
                                // The Space attribute preserve white space before and after your text
                                r1.Append(t1);
                                p.Append(r1);
                                // Run 2 - Bold
                                Run r2 = new Run();
                                RunProperties rp2 = new RunProperties();
                                rp2.Bold = new Bold();
                                // Always add properties first
                                r2.Append(rp2);
                                Text t2 = new Text("commodo ") { Space = SpaceProcessingModeValues.Preserve };
                                r2.Append(t2);
                                p.Append(r2);
                                // Run 3
                                Run r3 = new Run();
                                Text t3 = new Text("rhoncus ") { Space = SpaceProcessingModeValues.Preserve };
                                r3.Append(t3);
                                p.Append(r3);
                                // Run 4 – Italic
                                Run r4 = new Run();
                                RunProperties rp4 = new RunProperties();
                                rp4.Italic = new Italic();
                                // Always add properties first
                                r4.Append(rp4);
                                Text t4 = new Text("mauris") { Space = SpaceProcessingModeValues.Preserve };
                                r4.Append(t4);
                                p.Append(r4);
                                // Run 5
                                Run r5 = new Run();
                                Text t5 = new Text(", sit ") { Space = SpaceProcessingModeValues.Preserve };
                                r5.Append(t5);
                                p.Append(r5);
                                // Run 6 – Italic , bold and underlined
                                Run r6 = new Run();
                                RunProperties rp6 = new RunProperties();
                                rp6.Italic = new Italic();
                                rp6.Bold = new Bold();
                                rp6.Underline = new Underline();
                                // Always add properties first
                                r6.Append(rp6);
                                Text t6 = new Text("amet ") { Space = SpaceProcessingModeValues.Preserve };
                                r6.Append(t6);
                                p.Append(r6);
                                // Run 7
                                Run r7 = new Run();
                                Text t7 = new Text("faucibus arcu ") { Space = SpaceProcessingModeValues.Preserve };
                                r7.Append(t7);
                                p.Append(r7);
                                // Run 8 – Red color
                                Run r8 = new Run();
                                RunProperties rp8 = new RunProperties();
                                rp8.Color = new Color() { Val = "FF0000" };
                                // Always add properties first
                                r8.Append(rp8);
                                Text t8 = new Text("porttitor ") { Space = SpaceProcessingModeValues.Preserve };
                                r8.Append(t8);
                                p.Append(r8);
                                // Run 9
                                Run r9 = new Run();
                                Text t9 = new Text("pharetra. Maecenas quis erat quis eros iaculis placerat ut at mauris.") { Space = SpaceProcessingModeValues.Preserve };
                                r9.Append(t9);
                                p.Append(r9);
                                // Add your paragraph to docx body
                                docBody.Append(p);*/
                foreach (var moduleWrite in moduleWrites) 
                {
                    if (moduleWrite == null) continue;
                    docBody.Append(
                        new Paragraph(
                            new ParagraphProperties(
                                new Justification() { Val = JustificationValues.Both }
                            ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                new Text() { Text = "Учебная дисциплина (модуль): ", Space = SpaceProcessingModeValues.Preserve}
                            ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                new Text(moduleWrite.Name)
                            )
                        )
                    );

                    if (!string.IsNullOrEmpty(moduleWrite.Exams))
                    {
                        docBody.Append(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Both }
                                ),
                                new Run(
                                    new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                    new Text() { Text = "Экзамены, в каких семестрах: ", Space = SpaceProcessingModeValues.Preserve }
                                ),
                                new Run(
                                    new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                    new Text(moduleWrite.Exams)
                                )
                            )
                        );
                    }

                    if (!string.IsNullOrEmpty(moduleWrite.Receives.Replace(" ", "")))
                    {
                        docBody.Append(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Both }
                                ),
                                new Run(
                                    new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                    new Text() { Text = "Зачеты, в каких семестрах: ", Space = SpaceProcessingModeValues.Preserve }
                                ),
                                new Run(
                                    new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                    new Text(moduleWrite.GetReceives)
                                )
                            )
                        );
                    }

                    docBody.Append(
                        new Paragraph(
                            new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Both }
                                ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                new Text() { Text = "Всего: ", Space = SpaceProcessingModeValues.Preserve }
                            ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                new Text($"{moduleWrite?.TotalHours} ч. ({moduleWrite.GetAuditoriumHours} {moduleWrite.GetLectureHours} {moduleWrite.GetLabsHours} {moduleWrite.GetPracticeHours} {moduleWrite.GetSeminarHours})")
                            )
                        ),
                        new Paragraph(
                            new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Both }
                                ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                new Text("Описание учебной дисциплины: ")
                            )
                        )
                    );
                    if (moduleWrite.Description.Split("\n\n").Count() > 1)
                    {
                        foreach (string line in moduleWrite.Description.Split("\n\n"))
                        {
                            docBody.Append(
                               new Paragraph(
                                   new ParagraphProperties(
                                           new Justification() { Val = JustificationValues.Both }
                                       ),
                                   new Run(
                                       new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),

                                       new Text(line)
                                   )
                               )
                            );
                        }
                    }
                }

                mainPart.Document.Append(docBody);
                mainPart.Document.Save();
            }
        } 
    }
}
