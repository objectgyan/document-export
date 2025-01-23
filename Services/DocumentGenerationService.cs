using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MasterFormatDocExportPOC.Models;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace MasterFormatDocExportPOC.Services
{
    public class DocumentGenerationService : IDocumentGenerationService
    {
        public void GenerateDocument(List<MasterFormatSection> sections, string outputPath, Project project)
        {
            using (WordprocessingDocument wordDocument = 
                WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Add page numbers in footer
                AddPageNumbers(wordDocument);

                // Add project details
                AddProjectDetails(body, project);
                AddPageBreak(body);
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

        private void AddPageNumbers(WordprocessingDocument wordDocument)
        {
            var mainPart = wordDocument.MainDocumentPart;
            var footerPart = mainPart.AddNewPart<FooterPart>();
            
            // Create footer with page numbers and proper spacing
            var footer = new Footer(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                    ),
                    new Run(
                        new RunProperties(
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                            new FontSize() { Val = "20" }
                        ),
                        new Text("Page ") { Space = SpaceProcessingModeValues.Preserve },
                        new SimpleField() { Instruction = "PAGE" },
                        new Text(" of ") { Space = SpaceProcessingModeValues.Preserve },
                        new SimpleField() { Instruction = "NUMPAGES" }
                    )
                )
            );

            footer.Save(footerPart);

            var footerReference = new FooterReference() { 
                Type = HeaderFooterValues.Default, 
                Id = mainPart.GetIdOfPart(footerPart) 
            };

            var sectionProps = new SectionProperties(footerReference);
            mainPart.Document.Body.AppendChild(sectionProps);
        }

        private void AddProjectDetails(Body body, Project project)
        {
            // Add logo image
            string imageUrl = project.BannerImage.Url;

            AddImageFromUrl(body, imageUrl);

            // Add project name in bold and increased font size
            AddStyledParagraph(body, project.ProjectName, 100, true);

            // Create table for project details
            Table table = new Table();

            // Set table properties
            TableProperties tblProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 12 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 12 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 12 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 12 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 12 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 12 }
                ),
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct } // Set table width to 100%
            );
            table.AppendChild(tblProps);

            // Add rows to the table
            table.AppendChild(CreateTableRow("Location", project.LocationFullName));
            table.AppendChild(CreateTableRow("Type", project.Type));
            table.AppendChild(CreateTableRow("Budget", project.Budget));
            table.AppendChild(CreateTableRow("Phase", project.PhaseName));

            body.AppendChild(table);

            // Add "About Project" section
            AddStyledParagraph(body, "About Project:", 101, true);
            AddStyledParagraph(body, project.ProjectDescription, 2);
        }

        private TableRow CreateTableRow(string header, string value)
        {
            TableRow tr = new TableRow();

            TableCell tc1 = new TableCell(new Paragraph(new Run(new Text(header))));
            TableCell tc2 = new TableCell(new Paragraph(new Run(new Text(value))));

            TableCellProperties tcp1 = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1000" });
            TableCellProperties tcp2 = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "4000" });

            tc1.Append(tcp1);
            tc2.Append(tcp2);

            tr.Append(tc1, tc2);
            return tr;
        }

        private void AddImageFromUrl(Body body, string imageUrl)
        {
            try
            {
                using (var client = new System.Net.WebClient())
                using (var stream = new MemoryStream(client.DownloadData(imageUrl)))
                {
                    // Get MainDocumentPart through the document
                    var document = body.Parent as Document;
                    if (document == null)
                    {
                        throw new InvalidOperationException("Could not get Document");
                    }

                    var mainPart = document.MainDocumentPart;
                    if (mainPart == null)
                    {
                        throw new InvalidOperationException("Could not get MainDocumentPart");
                    }

                    var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    
                    stream.Position = 0;
                    imagePart.FeedData(stream);

                    // Fixed dimensions in EMUs (1 cm = 360000 EMUs)
                    const double widthCm = 16.8;
                    const double heightCm = 10.6;
                    const long emusPerCm = 360000;

                    var widthEmus = (long)(widthCm * emusPerCm);
                    var heightEmus = (long)(heightCm * emusPerCm);

                    var element = CreateImageElement(mainPart, imagePart, widthEmus, heightEmus);
                    var paragraph = new Paragraph(new Run(element));
                    
                    // Center align the image
                    paragraph.ParagraphProperties = new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                    );
                    
                    body.AppendChild(paragraph);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding image: {ex.Message}");
            }
        }

        private Drawing CreateImageElement(MainDocumentPart mainPart, ImagePart imagePart, long width, long height)
        {
            var relationshipId = mainPart.GetIdOfPart(imagePart);
            
            return new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = width, Cy = height },
                    new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties() { Id = 1U, Name = "Logo" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties() { Id = 1U, Name = "Logo.jpg" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip() { Embed = relationshipId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = width, Cy = height }),
                                    new A.PresetGeometry(new A.AdjustValueList()) 
                                    { Preset = A.ShapeTypeValues.Rectangle }))
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                }
            );
        }

        private void AddTableOfContents(Body body, List<MasterFormatSection> sections)
        {
            //// Add logo image
            //string imageUrl = "https://res.cloudinary.com/acelab/image/upload/v1669815709/static/assets/project-placeholder_wlqhqb.png";
            //AddImageFromUrl(body, imageUrl);
            
            //// Add spacing after image
            //AddStyledParagraph(body, string.Empty, 3);

            // Continue with TOC
            AddStyledParagraph(body, "TABLE OF CONTENTS", 1);
            AddStyledParagraph(body, string.Empty, 3);

            foreach (var section in sections)
            {
                ProcessTocSection(body, section, 0);
            }
        }

        private void ProcessTocSection(Body body, MasterFormatSection section, int level)
        {
            string indent = new string(' ', level * 2);
            string sectionText;

            if (level == 0) // Top-level section in TOC
            {
                sectionText = $"{indent}DIVISION {section.MasterFormatNumber} - {section.MasterFormatName}";
                AddStyledParagraph(body, sectionText, 2, true);
            }
            else
            {
                sectionText = $"{indent}{section.MasterFormatNumber} - {section.MasterFormatName}";
                AddStyledParagraph(body, sectionText, 2);
            }

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
            // Add section header with bold for parent and immediate child
            string headerText;
            if (level == 1) // Top-level section
            {
                headerText = $"DIVISION {section.MasterFormatNumber} - {section.MasterFormatName}";
            }
            else
            {
                headerText = $"{section.MasterFormatNumber} - {section.MasterFormatName}";
            }

            if (level == 1 || level == 2)
            {
                AddStyledParagraph(body, headerText, level, true);
            }
            else
            {
                AddStyledParagraph(body, headerText, level);
            }

            // Process products if any
            if (section.Products?.Any() == true)
            {
                char bulletPoint = 'A';
                foreach (var product in section.Products)
                {
                    ProcessProduct(body, product, level + 2, bulletPoint);
                    bulletPoint++;
                }
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

        private void ProcessProduct(Body body, Product product, int level, char bulletPoint)
        {
            // Create paragraph for product header
            Paragraph para = new Paragraph();
            ParagraphProperties paraProperties = new ParagraphProperties();
            paraProperties.Append(new Indentation() { Left = "360" });
            para.Append(paraProperties);
            
            // Add bullet point with bold and space after
            Run bulletRun = new Run();
            RunProperties bulletProps = new RunProperties();
            bulletProps.Append(new Bold());
            bulletProps.Append(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" });
            bulletProps.Append(new FontSize() { Val = "22" });
            bulletRun.Append(bulletProps);
            
            // Add bullet point with preserved spaces
            Text bulletText = new Text($"{bulletPoint}. ") { Space = SpaceProcessingModeValues.Preserve };
            bulletRun.Append(bulletText);

            // Add product name and details
            Run textRun = new Run();
            RunProperties textProps = new RunProperties();
            textProps.Append(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" });
            textProps.Append(new FontSize() { Val = "22" });
            textRun.Append(textProps);

            string productDetails = product.ProductName;
            if (!string.IsNullOrEmpty(product.ProductSubName))
                productDetails += $" - {product.ProductSubName}";
            if (!string.IsNullOrEmpty(product.ManufacturerName))
                productDetails += $" ({product.ManufacturerName})";

            Text productText = new Text(productDetails) { Space = SpaceProcessingModeValues.Preserve };
            textRun.Append(productText);

            // Combine bullet point and text
            para.Append(bulletRun);
            para.Append(textRun);
            body.Append(para);

            // Process custom columns with more indentation
            if (product.CustomColumns?.Any() == true)
            {
                int columnNumber = 1;
                foreach (var column in product.CustomColumns.OrderBy(c => c.DisplayOrder))
                {
                    string value = GetFormattedColumnValue(column);
                    
                    Paragraph subPara = new Paragraph();
                    ParagraphProperties subParaProps = new ParagraphProperties();
                    subParaProps.Append(new Indentation() { Left = "720" }); // Add more indentation (1 inch)
                    subPara.Append(subParaProps);

                    Run subRun = new Run();
                    RunProperties subRunProps = new RunProperties();
                    subRunProps.Append(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" });
                    subRunProps.Append(new FontSize() { Val = "22" });
                    subRun.Append(subRunProps);
                    subRun.Append(new Text($"{columnNumber}. {column.Title} - {value}"));
                    
                    subPara.Append(subRun);
                    body.Append(subPara);
                    
                    columnNumber++;
                }
            }

            // Add creation info with same indentation as custom columns
            Paragraph datePara = new Paragraph();
            ParagraphProperties dateParaProps = new ParagraphProperties();
            dateParaProps.Append(new Indentation() { Left = "720" });
            datePara.Append(dateParaProps);

            Run dateRun = new Run();
            RunProperties dateRunProps = new RunProperties();
            dateRunProps.Append(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" });
            dateRunProps.Append(new FontSize() { Val = "22" });
            dateRun.Append(dateRunProps);
            dateRun.Append(new Text($"4. Date Added - {product.CreatedDate:yyyy-MM-dd} {product.CreatedByUserName}"));
            
            datePara.Append(dateRun);
            body.Append(datePara);
            
            //AddStyledParagraph(body, string.Empty, 3); // Add spacing after product
        }

        private void AddStyledParagraph(Body body, string text, int level, bool isBold = false)
        {
            Paragraph para = new Paragraph();
            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            ParagraphProperties paraProperties = new ParagraphProperties();

            // Add Arial font
            runProperties.Append(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" });

            // Add bold if specified
            if (isBold)
            {
                runProperties.Append(new Bold());
            }

            switch (level)
            {
                case 1: // Main headers
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
                case 100: // Project Title
                    runProperties.Append(new FontSize() { Val = "48" });
                    paraProperties.Append(new SpacingBetweenLines() { Before = "240", After = "240" });
                    break;
                case 101: // Project Title
                    runProperties.Append(new FontSize() { Val = "32" });
                    paraProperties.Append(new SpacingBetweenLines() { Before = "240", After = "120" });
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
