
namespace GoogleAnalytics4
{
    internal class GoogleAnalytics4Settings
    {
        public string OutputExcelFilePath { get; set; } = @"C:\Users\stagiaireit\Desktop\OOG_GA4_Report.xlsx";
        public string OOGardenFranceId { get; set; } = "250897655";
        public string OOGardenAllemagneId { get; set; } = "401343913";
        public string OOGardenBelgiqueId { get; set; } = "280115024";
        public string GoogleSecretJSON { get; set; } = @"C:\Users\stagiaireit\Desktop\GoogleAnalytics\oogarden-ga4-edbc9bbe57a0.json";
        public string GoogleCredentialURL { get; set; } = "https://www.googleapis.com/auth/analytics.readonly";
        public string GrpcChannellURL { get; set; } = "https://analyticsdata.googleapis.com";

    }
}
