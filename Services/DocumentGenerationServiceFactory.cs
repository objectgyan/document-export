using MasterFormatDocExportPOC.Models;

namespace MasterFormatDocExportPOC.Services
{
    public class DocumentGenerationServiceFactory
    {
        public static IDocumentGenerationService CreateService(ExportType outputFormat)
        {
            return outputFormat switch
            {
                ExportType.Word => new DocumentGenerationService(),
                ExportType.TXT => new TextDocumentGenerationService(),
                _ => throw new ArgumentException("Unsupported document format", nameof(outputFormat))
            };
        }
    }
}
