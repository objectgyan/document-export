using MasterFormatDocExportPOC.Models;
using System.Text;

namespace MasterFormatDocExportPOC.Services
{
    public class TextDocumentGenerationService : IDocumentGenerationService
    {
        private readonly StringBuilder _stringBuilder;
        
        public TextDocumentGenerationService()
        {
            _stringBuilder = new StringBuilder();
        }

        public string GenerateDocument(IEnumerable<MasterFormatSection> sections)
        {
            _stringBuilder.Clear();
            
            // Add title
            _stringBuilder.AppendLine("TABLE OF CONTENTS");
            _stringBuilder.AppendLine();

            // Generate table of contents
            foreach (var section in sections)
            {
                AddToTableOfContents(section, 0);
            }

            _stringBuilder.AppendLine();
            _stringBuilder.AppendLine("DETAILED SECTIONS");
            _stringBuilder.AppendLine();

            // Generate detailed content
            foreach (var section in sections)
            {
                ProcessSection(section, 0);
            }

            return _stringBuilder.ToString();
        }

        private void AddToTableOfContents(MasterFormatSection section, int level)
        {
            _stringBuilder.AppendLine($"{section.MasterFormatNumber} - {section.MasterFormatName}");

            if (section.ChildSections?.Any() == true)
            {
                foreach (var childSection in section.ChildSections)
                {
                    AddToTableOfContents(childSection, level + 1);
                }
            }
        }

        private void ProcessSection(MasterFormatSection section, int level)
        {
            // Add section header
            _stringBuilder.AppendLine($"{section.MasterFormatNumber} - {section.MasterFormatName}");

            // Add products if any
            if (section.Products?.Any() == true)
            {
                _stringBuilder.AppendLine("  Products:");
                foreach (var product in section.Products)
                {
                    string productText = product.ProductName;
                    if (!string.IsNullOrEmpty(product.ProductSubName))
                        productText += $" - {product.ProductSubName}";
                    if (!string.IsNullOrEmpty(product.ManufacturerName))
                        productText += $" ({product.ManufacturerName})";

                    _stringBuilder.AppendLine($"    {productText}");

                    // Add custom columns with numbering
                    if (product.CustomColumns?.Any() == true)
                    {
                        int columnNumber = 1;
                        foreach (var column in product.CustomColumns)
                        {
                            string value = string.Empty;
                            
                            switch (column.Data.Type)
                            {
                                case "Bounded":
                                    if (column.Data.BoundedData?.Any() == true)
                                    {
                                        value = string.Join(", ", column.Data.BoundedData.Select(b => b.Name));
                                    }
                                    break;

                                case "Metric":
                                    if (column.Data.MetricData?.Any() == true)
                                    {
                                        var format = $"F{column.Data.DecimalCount}";
                                        value = string.Join(", ", column.Data.MetricData.Select(m => m.Value.ToString(format)));
                                    }
                                    break;

                                case "Text":
                                    value = column.Data.Value ?? string.Empty;
                                    break;
                            }
                            
                            _stringBuilder.AppendLine($"      {columnNumber}. {column.Title} - {value}");
                            columnNumber++;
                        }
                        _stringBuilder.AppendLine();
                    }
                    _stringBuilder.AppendLine();
                }

                // Add date added once after all products
                if (section.Products.Any())
                {
                    var firstProduct = section.Products.First();
                    _stringBuilder.AppendLine($"    Date Added - {firstProduct.CreatedDate:yyyy-MM-dd} {firstProduct.CreatedByUserName}");
                    _stringBuilder.AppendLine();
                }
            }

            // Process child sections
            if (section.ChildSections?.Any() == true)
            {
                foreach (var childSection in section.ChildSections)
                {
                    ProcessSection(childSection, level + 1);
                }
            }

            // Add extra line break after top-level sections (2-digit numbers)
            if (section.MasterFormatNumber.Length == 2)
            {
                _stringBuilder.AppendLine();
                _stringBuilder.AppendLine();
            }
        }

        public void GenerateDocument(List<MasterFormatSection> sections, string outputPath)
        {
            _stringBuilder.Clear();
            
            // Add title
            _stringBuilder.AppendLine("TABLE OF CONTENTS");
            _stringBuilder.AppendLine();

            // Generate table of contents
            foreach (var section in sections)
            {
                AddToTableOfContents(section, 0);
            }

            _stringBuilder.AppendLine();
            _stringBuilder.AppendLine("DETAILED SECTIONS");
            _stringBuilder.AppendLine();

            // Generate detailed content
            foreach (var section in sections)
            {
                ProcessSection(section, 0);
            }

            // Write to file
            File.WriteAllText(outputPath, _stringBuilder.ToString());
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
