using ConsoleApp1.Helpers;
using ConsoleApp1.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.Doc;
using Spire.Doc.Interface;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;


namespace ConsoleApp1.Executers.Word.Write;

public class WordFileWriter
{
    public async Task WriteIntoDocumentAsync(List<ModuleWrite> moduleWrites)
    {
        var filePath = moduleWrites.FirstOrDefault().FullPath +  ".docx";
        // Check if the file already exists and delete it if needed
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        // Create the file and close it immediately
        using (File.Create(filePath)) { }

        // Delay a bit to ensure the file handle is released
        System.Threading.Thread.Sleep(100);


        using (MemoryStream mem = new MemoryStream())
        {
            // Create Document
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                // Create the document structure and add some text.
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                var docBody = new DocumentFormat.OpenXml.Wordprocessing.Body();
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
                                new Text("Учебная дисциплина (модуль): "),
                                new Text(" ")
                            ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                new Text(" " + moduleWrite.Name)
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
                                    new Text("Экзамены: ")
                                ),
                                new Run(
                                    new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                    new Text(" " + moduleWrite.Exams)
                                )
                            )
                        );
                    }

                    if (!string.IsNullOrEmpty(moduleWrite.Receives.Replace(" ","")))
                    {
                        docBody.Append(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Both }
                                ),
                                new Run(
                                    new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                    new Text("Зачеты: ")
                                ),
                                new Run(
                                    new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                    new Text(" " + moduleWrite.Receives)
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
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                new Text("Всего: ")
                            ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                new Text(" " + $"{moduleWrite?.TotalHours} ч. ({moduleWrite.GetAuditoriumHours}, {moduleWrite.GetLectureHours}, {moduleWrite.GetLabsHours}, {moduleWrite.GetPracticeHours}, {moduleWrite.GetSeminarHours})")
                            )
                        ),

                        new Paragraph(
                            new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Both }
                                ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new Bold(), new FontSize() { Val = "28" }),
                                new Text("Описание учебной дисциплины: ")
                            )                    
                        ),
                        new Paragraph(
                            new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Both }
                                ),
                            new Run(
                                new RunProperties(new RunFonts() { ComplexScript = "Times New Roman", Ascii = "Times New Roman", HighAnsi = "Times New Roman" , EastAsia = "Times New Roman" }, new FontSize() { Val = "28" }),
                                new Text(moduleWrite.Description?.Replace("\n", "").Replace("\r", ""))
                            )
                        )/*,
                        new Paragraph(
                            new Run(
                                new Break() { Type = BreakValues.Page }
                            )
                        )*/

                    // Add more paragraphs following the same structure for other content
                    );
                }
                /*
                docBody.Append(new Paragraph(new Run
                    (new Text
                    (""
                    ))));*/
                mainPart.Document.Append(docBody);
                mainPart.Document.Save();
            }
            //Context.Response.AppendHeader("Content-Disposition", String.Format("attachment;filename=\"0}.docx\"", MyDocxTitle));
            //mem.Position = 0;
            //mem.CopyTo(Context.Response.OutputStream);
            //Context.Response.Flush();
            //Context.Response.End();
        }


        /*
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();
        foreach (var moduleWrite in moduleWrites);
                wordDocument.Save();
        {
            if (moduleWrite == null) continue;

            try
            {
                document.Sections.Add(
                    new Section(document,
                       
                        new Paragraph(document,
                            new Run(document, "Всего: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } },
                            new Run(document, $"{moduleWrite.TotalHours} ".Replace("\n", "").Replace("\r", "")) { CharacterFormat = { FontName = "Times New Roman", Size = 14 } } //({moduleWrite.AuditoriumHours} ауд. ч., {moduleWrite.LectureHours} лекционных ч., {moduleWrite.PracticeHours} практических ч., {moduleWrite.SeminarHours} семинарских ч.)
                        ),
                        new Paragraph(document,
                            new Run(document, "Зачетных единиц: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } },
                            new Run(document, moduleWrite.ReceivedUnits.Replace("\n", "").Replace("\r", "")) { CharacterFormat = { FontName = "Times New Roman", Size = 14 } }
                        ),
                        new Paragraph(document,
                            new Run(document, "Описание учебной дисциплины: ") { CharacterFormat = { FontName = "Times New Roman", Bold = true, Size = 14 } }
                        ),
                        new Paragraph(document,
                            new Run(document, moduleWrite.Description.Replace("\n", "").Replace("\r", "")) { CharacterFormat = { FontName = "Times New Roman", Size = 14 } }
                        )
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
            }
            catch (Exception ex) { 
                Console.WriteLine(ex.ToString()); }
            
        }
         document.Save(filePath);


        */
    }
}
