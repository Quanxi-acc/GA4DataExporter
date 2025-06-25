using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleAnalytics4
{
    internal class GoogleAnalytics4Settings
    {
        public string OutputDirPath { get; set; } = @"C:\Users\stagiaireit\Desktop\OOG_GA4_Report.xlsx";
        public string oogardenFranceId { get; set; } = "250897655";
        public string oogardenAllemagneId { get; set; } = "401343913";
        public string oogardenBelgiqueId { get; set; } = "280115024";
        public string GoogleSecretJSON { get; set; } = "oog-stagiaire-sheet-report";  /* C:\\Users\\stagiaireit\\Desktop\\GoogleAnalytics\\GoogleAnalytics4\\bin\\Debug\\net6.0\\oog-stagiaire-sheet-report.json
                                                                                         C:\\Users\\stagiaireit\\Desktop\\GoogleAnalytics\\GoogleAnalytics4\\bin\\Debug\\net6.0\\token.json\... */
        public string GoogleDriveJSON { get; set; } = ""; /* Variable d'envrionnement GOOGLE_APPLICATION_CREDENTIALS à mettre en dure dans le code 
                                                           La variable est ici : C:\\Users\\stagiaireit\\Desktop\\GoogleAnalytics\\oogarden-ga4-edbc9bbe57a0.json */

    }
}
