namespace GoogleAnalytics4
{
    public class SageRecord : IGoogleRecord
    {
        public double TotalActive { get; set; }
        public double TotalNew { get; set; }
        public double TotalBounce { get; set; }
        public int CountryId { get; set; } 
    }
}
