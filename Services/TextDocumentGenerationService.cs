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

        public void GenerateDocument(List<MasterFormatSection> sections, string outputPath, Project project)
        {
            _stringBuilder.Clear();

            // Add Project Details
            AddProjectDetails(project);
            _stringBuilder.AppendLine();
            _stringBuilder.AppendLine();

            // Add Table of Contents
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

        private void AddProjectDetails(Project project)
        {
            // Project name
            _stringBuilder.AppendLine(project.ProjectName);
            _stringBuilder.AppendLine();

            // Project details table
            _stringBuilder.AppendLine($"Location: {project.LocationFullName}");
            _stringBuilder.AppendLine($"Type:     {project.Type}");
            _stringBuilder.AppendLine($"Budget:   {project.Budget}");
            _stringBuilder.AppendLine($"Phase:    {project.PhaseName}");
            _stringBuilder.AppendLine();

            // About Project section
            _stringBuilder.AppendLine("About Project:");
            _stringBuilder.AppendLine(project.ProjectDescription);
        }

        private void AddToTableOfContents(MasterFormatSection section, int level)
        {
            _stringBuilder.AppendLine($"DIVISION {section.MasterFormatNumber} - {section.MasterFormatName}");

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
            // Add section header with DIVISION prefix for top-level sections
            if (level == 0)
            {
                _stringBuilder.AppendLine($"DIVISION {section.MasterFormatNumber} - {section.MasterFormatName}");
            }
            else
            {
                _stringBuilder.AppendLine($"{section.MasterFormatNumber} - {section.MasterFormatName}");
            }

            // Process products if any
            if (section.Products?.Any() == true)
            {
                char bulletPoint = 'A';
                foreach (var product in section.Products)
                {
                    ProcessProduct(product, bulletPoint);
                    bulletPoint++;
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

            // Add extra line break after top-level sections
            if (section.MasterFormatNumber.Length == 2)
            {
                _stringBuilder.AppendLine();
            }
        }

        private void ProcessProduct(Product product, char bulletPoint)
        {
            // Product header with bullet point
            string productText = $"    {bulletPoint}. {product.ProductName}";
            if (!string.IsNullOrEmpty(product.ProductSubName))
                productText += $" - {product.ProductSubName}";
            if (!string.IsNullOrEmpty(product.ManufacturerName))
                productText += $" ({product.ManufacturerName})";

            _stringBuilder.AppendLine(productText);

            // Process custom columns
            if (product.CustomColumns?.Any() == true)
            {
                int columnNumber = 1;
                foreach (var column in product.CustomColumns.OrderBy(c => c.DisplayOrder))
                {
                    string value = GetFormattedColumnValue(column);
                    _stringBuilder.AppendLine($"        {columnNumber}. {column.Title} - {value}");
                    columnNumber++;
                }
                //_stringBuilder.AppendLine();
            }

            // Add creation info as point 4
            _stringBuilder.AppendLine($"        4. Date Added - {product.CreatedDate:yyyy-MM-dd} {product.CreatedByUserName}");
            //_stringBuilder.AppendLine();
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
