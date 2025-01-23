using MasterFormatDocExportPOC.Models;

namespace MasterFormatDocExportPOC.Services
{
    public interface IDocumentGenerationService
    {
        void GenerateDocument(List<MasterFormatSection> sections, string outputPath);
    }
}
