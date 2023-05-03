namespace QualityScoringBlazor.Models
{
    public class Order
    {
        public int Id { get; set; }
        public string? StringOrderNumber { get; set; }
        public string? OrderAddress { get; set; }
        public char[]? OrderNoCharArray { get; set; }
        public string? ShortCode { get; set; }
        public string? StringDate { get; set;}
    }
}
