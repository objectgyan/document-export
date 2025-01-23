using MasterFormatDocExportPOC.Models;

namespace MasterFormatDocExportPOC.Services
{
    public interface IDocumentExporter
    {
        Task ExportAsync(MasterFormatSection section, string outputPath);
    }
}
