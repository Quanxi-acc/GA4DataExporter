
namespace GoogleAnalytics4
{
    internal class GoogleAnalytics4Settings
    {
        public string OutputExcelFilePath { get; set; } = @"C:\Users\stagiaireit\Desktop\OOG_GA4_Report.xlsx"; 
        public string OOGardenFranceId { get; set; } = "250897655";    
        public string OOGardenAllemagneId { get; set; } = "401343913"; 
        public string OOGardenBelgiqueId { get; set; } = "280115024";  
        public string GoogleSecretJSON { get; set; } = "C:\\Users\\stagiaireit\\Desktop\\GoogleAnalytics\\oogarden-ga4-edbc9bbe57a0.json"; /* Variable d'envrionnement GOOGLE_APPLICATION_CREDENTIALS à mettre en dure dans le code 
                                                                                      La variable est ici : C:\\Users\\stagiaireit\\Desktop\\GoogleAnalytics\\oogarden-ga4-edbc9bbe57a0.json */
        public string GoogleDriveJSON { get; set; } = "C:\\Users\\stagiaireit\\Desktop\\GoogleAnalytics\\GoogleAnalytics4\\bin\\Debug\\net6.0\\oog-stagiaire-sheet-report.json";

                                                          

    }
}
