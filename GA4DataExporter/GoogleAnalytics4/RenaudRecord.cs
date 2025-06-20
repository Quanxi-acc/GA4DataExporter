namespace GoogleAnalytics4
{
    public class RenaudRecord : IGoogleRecord
    {
        public double TotalScreenPageViews { get; set; }
        public Dictionary<string, double> SessionsParAppareil { get; set; } = new();
        public Dictionary<string, double> RevenueParAppareil { get; set; } = new();
        public Dictionary<string, double> LoadDurationParAppareil { get; set; } = new();

        public double TotalServerResponseDuration { get; set; }
        public double TotalLoadDuration { get; set; }
        public double TotalEventCountWebPerf { get; set; }

        public Dictionary<string, double> EventCountWebPerfParAppareil { get; set; } = new();
    }
}