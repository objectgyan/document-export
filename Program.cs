using MasterFormatDocExportPOC.Models;
using MasterFormatDocExportPOC.Services;
using Newtonsoft.Json;

namespace DocExportPOC
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Document Export POC Started");

            try
            {
                // Read and parse JSON data
                string jsonFilePath = Path.Combine(Directory.GetCurrentDirectory(), "masterformat-data.json");
                string jsonContent = File.ReadAllText(jsonFilePath);
                var sampleData = JsonConvert.DeserializeObject<MasterFormatResponse>(jsonContent);

                if (sampleData?.MasterFormatSections == null)
                {
                    throw new Exception("No data found in JSON file");
                }

                // Create Word document
                var docService = DocumentGenerationServiceFactory.CreateService(ExportType.Word);
                docService.GenerateDocument(sampleData.MasterFormatSections, "ExportedDocument.docx");

                // Create Text document
                var txtService = DocumentGenerationServiceFactory.CreateService(ExportType.TXT);
                txtService.GenerateDocument(sampleData.MasterFormatSections, "ExportedDocument.txt");

                Console.WriteLine("Documents generated successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating documents: {ex.Message}");
            }

            Console.WriteLine("Press any key to exit...");
            //Console.ReadKey();
        }
    }
}
