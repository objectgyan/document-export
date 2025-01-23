using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MasterFormatDocExportPOC.Models;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace MasterFormatDocExportPOC.Services
{
    public class DocumentGenerationService : IDocumentGenerationService
    {
        public void GenerateDocument(List<MasterFormatSection> sections, string outputPath)
        {
            using (WordprocessingDocument wordDocument = 
                WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Generate Table of Contents with image
                AddTableOfContents(body, sections);

                // Add page break after TOC
                AddPageBreak(body);

                // Add Detailed Sections header and add some spacing
                AddStyledParagraph(body, "DETAILED SECTIONS", 1);
                AddStyledParagraph(body, string.Empty, 3); // Add spacing after header

                // Process each section with full details
                bool isFirst = true;
                foreach (var section in sections)
                {
                    if (!isFirst)
                    {
                        AddPageBreak(body);
                    }
                    ProcessDetailedSection(body, section, 1);
                    isFirst = false;
                }

                mainPart.Document.Save();
            }
        }

        //private async Task AddImageFromUrl(Body body, string imageUrl)
        //{
        //    try
        //    {
        //        using (var httpClient = new HttpClient())
        //        {
        //            var imageBytes = await httpClient.GetByteArrayAsync(imageUrl);
        //            using (var memoryStream = new MemoryStream(imageBytes))
        //            {
        //                MainDocumentPart mainPart = body.Ancestors<Document>().First().MainDocumentPart;
        //                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                        
        //                // Feed the image bytes to the ImagePart
        //                memoryStream.Position = 0;
        //                imagePart.FeedData(memoryStream);

        //                // Load image for dimension calculation
        //                memoryStream.Position = 0;
        //                using (var image = System.Drawing.Image.FromStream(memoryStream))
        //                {
        //                    var maxWidthInches = 6.0;
        //                    var aspectRatio = (double)image.Height / image.Width;
        //                    var widthEmus = (long)(maxWidthInches * 914400);
        //                    var heightEmus = (long)(widthEmus * aspectRatio);

        //                    var element = new Drawing(
        //                        new DW.Inline(
        //                            new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
        //                            new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
        //                            new DW.DocProperties() { Id = 1U, Name = "Logo" },
        //                            new DW.NonVisualGraphicFrameDrawingProperties(
        //                                new A.GraphicFrameLocks() { NoChangeAspect = true }),
        //                            new A.Graphic(
        //                                new A.GraphicData(
        //                                    new PIC.Picture(
        //                                        new PIC.NonVisualPictureProperties(
        //                                            new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Logo.jpg" },
        //                                            new PIC.NonVisualPictureDrawingProperties()),
        //                                        new PIC.BlipFill(
        //                                            new A.Blip() { Embed = mainPart.GetIdOfPart(imagePart) },
        //                                            new A.Stretch(new A.FillRectangle())),
        //                                        new PIC.ShapeProperties(
        //                                            new A.Transform2D(
        //                                                new A.Offset() { X = 0L, Y = 0L },
        //                                                new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
        //                                            new A.PresetGeometry(new A.AdjustValueList()) 
        //                                            { Preset = A.ShapeTypeValues.Rectangle }))
        //                                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
        //                            )
        //                        ) {
        //                            DistanceFromTop = 0U,
        //                            DistanceFromBottom = 0U,
        //                            DistanceFromLeft = 0U,
        //                            DistanceFromRight = 0U
        //                        }
        //                    );

        //                    var paragraph = new Paragraph(new Run(element));
        //                    body.AppendChild(paragraph);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // Log error or handle gracefully
        //        Console.WriteLine($"Error loading image from URL: {ex.Message}");
        //    }
        //}

        private void AddTableOfContents(Body body, List<MasterFormatSection> sections)
        {
            // Add logo image from URL
            //string imageUrl = "https://res.cloudinary.com/acelab/image/upload/v1669815709/static/assets/project-placeholder_wlqhqb.png";
            //await AddImageFromUrl(body, imageUrl);
            
            // Add some spacing after the image
            AddStyledParagraph(body, string.Empty, 3);

            // Continue with existing TOC code
            AddStyledParagraph(body, "TABLE OF CONTENTS", 1);
            AddStyledParagraph(body, string.Empty, 3);

            foreach (var section in sections)
            {
                ProcessTocSection(body, section, 0);
            }
        }

        //private void AddImageToDocument(Body body, string imagePath)
        //{
        //    if (!File.Exists(imagePath)) return;

        //    MainDocumentPart mainPart = body.Ancestors<MainDocumentPart>().FirstOrDefault();
        //    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

        //    using (FileStream stream = new FileStream(imagePath, FileMode.Open))
        //    {
        //        imagePart.FeedData(stream);
        //    }

        //    // Calculate size while maintaining aspect ratio (max width 6 inches)
        //    var maxWidthInches = 6.0;
        //    using (var image = System.Drawing.Image.FromFile(imagePath))
        //    {
        //        var aspectRatio = (double)image.Height / image.Width;
        //        var widthEmus = (long)(maxWidthInches * 914400); // Convert inches to EMUs (914400 EMUs per inch)
        //        var heightEmus = (long)(widthEmus * aspectRatio);

        //        // Add the image to the document
        //        var element =
        //             new Drawing(
        //                 new DW.Inline(
        //                     new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
        //                     new DW.EffectExtent()
        //                     {
        //                         LeftEdge = 0L,
        //                         TopEdge = 0L,
        //                         RightEdge = 0L,
        //                         BottomEdge = 0L
        //                     },
        //                     new DW.DocProperties()
        //                     {
        //                         Id = 1U,
        //                         Name = "Logo"
        //                     },
        //                     new DW.NonVisualGraphicFrameDrawingProperties(
        //                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
        //                     new A.Graphic(
        //                         new A.GraphicData(
        //                             new PIC.Picture(
        //                                 new PIC.NonVisualPictureProperties(
        //                                     new PIC.NonVisualDrawingProperties()
        //                                     {
        //                                         Id = 0U,
        //                                         Name = "Logo.jpg"
        //                                     },
        //                                     new PIC.NonVisualPictureDrawingProperties()),
        //                                 new PIC.BlipFill(
        //                                     new A.Blip(
        //                                         new A.BlipExtensionList(
        //                                             new A.BlipExtension()
        //                                             {
        //                                                 Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
        //                                             })
        //                                     )
        //                                     {
        //                                         Embed = mainPart.GetIdOfPart(imagePart),
        //                                     },
        //                                     new A.Stretch(
        //                                         new A.FillRectangle())),
        //                                 new PIC.ShapeProperties(
        //                                     new A.Transform2D(
        //                                         new A.Offset() { X = 0L, Y = 0L },
        //                                         new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
        //                                     new A.PresetGeometry(
        //                                         new A.AdjustValueList()
        //                                     )
        //                                     { Preset = A.ShapeTypeValues.Rectangle }))
        //                         )
        //                         { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
        //                 )
        //                 {
        //                     DistanceFromTop = 0U,
        //                     DistanceFromBottom = 0U,
        //                     DistanceFromLeft = 0U,
        //                     DistanceFromRight = 0U,
        //                 });

        //        var paragraph = new Paragraph(new Run(element));
        //        body.AppendChild(paragraph);
        //    }
        //}

        private void ProcessTocSection(Body body, MasterFormatSection section, int level)
        {
            string indent = new string(' ', level * 2);
            AddStyledParagraph(body, $"{indent}{section.MasterFormatNumber} - {section.MasterFormatName}", 2);

            if (section.ChildSections?.Any() == true)
            {
                foreach (var childSection in section.ChildSections)
                {
                    ProcessTocSection(body, childSection, level + 1);
                }
            }
        }

        private void ProcessDetailedSection(Body body, MasterFormatSection section, int level)
        {
            // Add section header
            string headerText = level == 1
                ? $"{section.MasterFormatNumber} - {section.MasterFormatName}"
                : $"{section.MasterFormatNumber} - {section.MasterFormatName}";

            AddStyledParagraph(body, headerText, level);

            // Process products if any
            if (section.Products?.Any() == true)
            {
                AddStyledParagraph(body, "Products:", level + 1);

                foreach (var product in section.Products)
                {
                    ProcessProduct(body, product, level + 2);
                }

                // Add creation info once for all products
                var firstProduct = section.Products.First();
                AddStyledParagraph(body, 
                    $"Date Added - {firstProduct.CreatedDate:yyyy-MM-dd} {firstProduct.CreatedByUserName}", 
                    level + 2);
                AddStyledParagraph(body, string.Empty, 3); // Add spacing
            }

            // Process child sections
            if (section.ChildSections?.Any() == true)
            {
                foreach (var childSection in section.ChildSections)
                {
                    ProcessDetailedSection(body, childSection, level + 1);
                }
            }
        }

        private void ProcessProduct(Body body, Product product, int level)
        {
            // Product header
            string productText = product.ProductName;
            if (!string.IsNullOrEmpty(product.ProductSubName))
                productText += $" - {product.ProductSubName}";
            if (!string.IsNullOrEmpty(product.ManufacturerName))
                productText += $" ({product.ManufacturerName})";

            AddStyledParagraph(body, productText, level);

            // Process custom columns
            if (product.CustomColumns?.Any() == true)
            {
                int columnNumber = 1;
                foreach (var column in product.CustomColumns.OrderBy(c => c.DisplayOrder))
                {
                    string value = GetFormattedColumnValue(column);
                    AddStyledParagraph(body, $"{columnNumber}. {column.Title} - {value}", level + 1);
                    columnNumber++;
                }
                AddStyledParagraph(body, string.Empty, 3); // Add spacing between products
            }
        }

        private void AddStyledParagraph(Body body, string text, int level)
        {
            Paragraph para = new Paragraph();
            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            ParagraphProperties paraProperties = new ParagraphProperties();

            // Add Arial font
            runProperties.Append(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" });

            switch (level)
            {
                case 1: // Main headers
                    runProperties.Append(new Bold());
                    runProperties.Append(new FontSize() { Val = "28" });
                    paraProperties.Append(new SpacingBetweenLines() { Before = "240", After = "120" });
                    break;
                case 2: // Sub headers
                    runProperties.Append(new FontSize() { Val = "24" });
                    paraProperties.Append(new SpacingBetweenLines() { Before = "120", After = "60" });
                    break;
                case 3: // Spacing paragraph
                    paraProperties.Append(new SpacingBetweenLines() { Before = "60", After = "60" });
                    break;
                default: // Normal text
                    runProperties.Append(new FontSize() { Val = "22" });
                    paraProperties.Append(new SpacingBetweenLines() { Before = "40", After = "40" });
                    break;
            }

            run.Append(runProperties);
            run.Append(new Text(text));
            para.Append(paraProperties);
            para.Append(run);
            body.Append(para);
        }

        private void AddPageBreak(Body body)
        {
            Paragraph pageBreakPara = new Paragraph(
                new Run(
                    new Break() { Type = BreakValues.Page }
                )
            );
            body.Append(pageBreakPara);
        }

        private string GetFormattedColumnValue(CustomColumn column)
        {
            if (column.Data == null) return string.Empty;

            return column.Data.Type switch
            {
                "Bounded" => column.Data.BoundedData?.Any() == true
                    ? string.Join(", ", column.Data.BoundedData.Select(b => b.Name))
                    : string.Empty,

                "Metric" => column.Data.MetricData?.Any() == true
                    ? string.Join(", ", column.Data.MetricData.Select(m => 
                        m.Value.ToString($"F{column.Data.DecimalCount}")))
                    : string.Empty,

                "Text" => column.Data.Value ?? string.Empty,

                _ => string.Empty
            };
        }
    }
}
