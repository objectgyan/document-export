namespace MasterFormatDocExportPOC.Models
{
    public class MasterFormatSection
    {
        public string SectionId { get; set; }
        public string MasterFormatNumber { get; set; }
        public string MasterFormatName { get; set; }
        public List<MasterFormatSection> ChildSections { get; set; }
        public List<Product> Products { get; set; }
        public MasterFormatSection? NextSection { get; set; }
    }

    public class Product
    {
        public string ProductId { get; set; }
        public string ProductName { get; set; }
        public string ProductSubName { get; set; }
        public string ManufacturerName { get; set; }
        public DateTime CreatedDate { get; set; }
        public string CreatedByUserName { get; set; }
        public List<CustomColumn> CustomColumns { get; set; }
    }

    public class CustomColumn
    {
        public string Title { get; set; }
        public CustomColumnData Data { get; set; }
        public int DisplayOrder { get; set; }
    }

    public class CustomColumnData
    {
        public string Type { get; set; }
        public bool AllowMultipleValues { get; set; }
        public List<BoundedDataItem> BoundedData { get; set; }
        public List<MetricDataItem> MetricData { get; set; }
        public int DecimalCount { get; set; }  // For Metric type
        public string Value { get; set; }      // For Text type
        public string ScheduleCustomColumnId { get; set; }
        public string ScheduleRowId { get; set; }
    }

    public class BoundedDataItem
    {
        public string BoundedDataOptionId { get; set; }
        public string Name { get; set; }
        public string Color { get; set; }
    }

    public class MetricDataItem
    {
        public string MetricValueId { get; set; }
        public decimal Value { get; set; }
    }

    public class MasterFormatResponse
    {
        public List<MasterFormatSection> MasterFormatSections { get; set; }
    }
}
