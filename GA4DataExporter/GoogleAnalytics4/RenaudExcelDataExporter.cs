using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Google.Analytics.Data.V1Beta;
using System.Globalization;

namespace GoogleAnalytics4
{
    public class RenaudExcelDataExporter : IDataExporter
    {
        public void Export(IGoogleRecord data)
        {

        }
        
        public void Export(Dictionary<int, IGoogleRecord> dataFormatedByCountry)
        {
            var file = new XLWorkbook();
            var sheet = file.Worksheets.Add("Rapport");

            int currentRow = 1;

            foreach (var valueDataForm in dataFormatedByCountry)
            {
                int site = valueDataForm.Key;
                var results = (RenaudRecord)valueDataForm.Value;

                currentRow++;

                /* Entête de pays */
                sheet.Cell(currentRow, 1).Value = "Pays";
                sheet.Cell(currentRow, 2).Value = SwitchIdToCountryName(site);
                currentRow++;

                /* Nombre de pages vues */
                sheet.Cell(currentRow, 1).Value = "Nombre total de pages vues";
                sheet.Cell(currentRow, 2).Value = results.TotalScreenPageViews;
                currentRow++;

                currentRow++;

                /* Sessions avec engagement */
                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement mobile";
                sheet.Cell(currentRow, 2).Value = results.SessionsParAppareil.GetValueOrDefault("mobile", 0);
                currentRow++;

                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement desktop";
                sheet.Cell(currentRow, 2).Value = results.SessionsParAppareil.GetValueOrDefault("desktop", 0);
                currentRow++;

                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement tablet";
                sheet.Cell(currentRow, 2).Value = results.SessionsParAppareil.GetValueOrDefault("tablet", 0);
                currentRow++;

                double totalSessions = results.SessionsParAppareil.Values.Sum();
                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement Total";
                sheet.Cell(currentRow, 2).Value = totalSessions;
                currentRow++;

                currentRow++;

                /* Chiffre d'affaire */
                sheet.Cell(currentRow, 1).Value = "Revenue Mobile";
                sheet.Cell(currentRow, 2).Value = results.RevenueParAppareil.GetValueOrDefault("mobile", 0);
                currentRow++;

                sheet.Cell(currentRow, 1).Value = "Revenue Desktop";
                sheet.Cell(currentRow, 2).Value = results.RevenueParAppareil.GetValueOrDefault("desktop", 0);
                currentRow++;

                sheet.Cell(currentRow, 1).Value = "Revenue Tablet";
                sheet.Cell(currentRow, 2).Value = results.RevenueParAppareil.GetValueOrDefault("tablet", 0);
                currentRow++;

                double totalRevenue = results.RevenueParAppareil.Values.Sum();
                sheet.Cell(currentRow, 1).Value = "Revenue Total";
                sheet.Cell(currentRow, 2).Value = totalRevenue;
                currentRow++;

                currentRow++;

                /* Métriques spécifiques France */
                if (site == 1)
                {
                    sheet.Cell(currentRow, 1).Value = "Temps réponse serveur total (ms)";
                    sheet.Cell(currentRow, 2).Value = results.TotalServerResponseDuration;
                    currentRow++;

                    sheet.Cell(currentRow, 1).Value = "Temps réponse serveur / WebPerf Total (ms)";
                    double averageServerResponse = results.TotalServerResponseDuration / results.TotalEventCountWebPerf;
                    sheet.Cell(currentRow, 2).Value = (averageServerResponse / 1000);
                    currentRow++;

                    currentRow++;

                    /* LoadDuration moyens */
                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Mobile / WebPerf Mobile (s)";
                    double averageMobile = results.LoadDurationParAppareil.GetValueOrDefault("mobile", 0) / results.EventCountWebPerfParAppareil.GetValueOrDefault("mobile", 0);
                    sheet.Cell(currentRow, 2).Value = (averageMobile / 1000);
                    currentRow++;

                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Desktop / WebPerf Desktop (s)";
                    double averageDesktop = results.LoadDurationParAppareil.GetValueOrDefault("desktop", 0) / results.EventCountWebPerfParAppareil.GetValueOrDefault("desktop", 0);
                    sheet.Cell(currentRow, 2).Value = (averageDesktop / 1000);
                    currentRow++;

                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Tablet / WebPerf Tablet (s)";
                    double averageTablet = results.LoadDurationParAppareil.GetValueOrDefault("tablet", 0) / results.EventCountWebPerfParAppareil.GetValueOrDefault("tablet", 0);
                    sheet.Cell(currentRow, 2).Value = (averageTablet / 1000);
                    currentRow++;

                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Total / WebPerf Total (s)";
                    double averageLoadDuration = results.TotalLoadDuration / results.TotalEventCountWebPerf;
                    sheet.Cell(currentRow, 2).Value = (averageLoadDuration / 1000);
                    currentRow++;
                }
            }

            sheet.Range("A2:B21").Style.Fill.SetBackgroundColor(XLColor.FromArgb(174, 196, 210));    /* Couleur Bleu -- France */
            sheet.Range("A23:B34").Style.Fill.SetBackgroundColor(XLColor.FromArgb(231, 175, 184));   /* Couleur Rouge -- Allemagne */
            sheet.Range("A37:B48").Style.Fill.SetBackgroundColor(XLColor.FromArgb(238, 234, 156));   /* Couleur Jaune -- Belgique */

            
            sheet.Range("A2:B15").Style.NumberFormat.SetFormat("# ### ### 000.00");                  /* Format séparateur de milliers */
            sheet.Range("A16:B16").Style.NumberFormat.SetFormat("0.00");                             /* Format séparateur de milliers */
            sheet.Range("A17:B48").Style.NumberFormat.SetFormat("# ### ### ##0.00");                 /* Format séparateur de milliers */

            sheet.Columns().AdjustToContents();

            file.SaveAs(@"C:\Users\stagiaireit\Desktop\OOG_GA4_Report.xlsx");
        }

        private string SwitchIdToCountryName(int siteId)
        {
            return siteId switch
            {
                1 => "France",
                2 => "Allemagne",
                3 => "Belgique",
                _ => "Rien"
            };
        }
    }
}