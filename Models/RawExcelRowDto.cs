namespace LISImportSPCApi.Models
{
    public class RawExcelRowDto
    {
        public int Task { get; set; }
        public string? TaskName { get; set; }
        public string TestPart { get; set; } = string.Empty;
        public string? TestPartDesc { get; set; }
        public double Value { get; set; }
        public string Unit { get; set; } = string.Empty;
        public DateTime TestDateTime { get; set; }

        // Optional fields for new table
        public string? Part { get; set; }
        public string? Serial { get; set; }
        public string? StoreLocation { get; set; } // New field for store_location column

        public string? Created_by { get; set; } // New field for store_location column
    }
}
