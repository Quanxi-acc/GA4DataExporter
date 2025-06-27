using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Google.Analytics.Data.V1Beta;
using Google.Type;
using System.Globalization;

namespace GoogleAnalytics4
{
    public class RenaudExcelDataExporter : IDataExporter
    {

        private GoogleAnalytics4Settings settings = new GoogleAnalytics4Settings();

        private string sheetName;
        private string columnName;
        private static XLWorkbook file;
        private static int actualColumn = 2;
        public bool CreateNewSheet { get; set; } = true;
        private static IXLWorksheet currentSheet = null;

        public RenaudExcelDataExporter(string sheetName, string columnName)
        {
            if(file == null)
            {
                file = new XLWorkbook();
            }
            
            this.sheetName = sheetName;
            this.columnName = columnName;
        }

        public void Export(IGoogleRecord data)
        {
        }

        public void Export(Dictionary<int, IGoogleRecord> dataFormatedByCountry)
        {

            GenerateMainReport(file, dataFormatedByCountry);    

            file.SaveAs(settings.OutputExcelFilePath);
        }

        private void GenerateMainReport(XLWorkbook file, Dictionary<int, IGoogleRecord> dataFormatedByCountry)
        {
            if (CreateNewSheet)
            {
                actualColumn = 2;
                currentSheet = file.Worksheets.Add(sheetName);

                WriteHeader(dataFormatedByCountry, currentSheet);
            }


            WriteData(currentSheet, dataFormatedByCountry);
            WriteStyle(currentSheet);

        }

        private static void WriteStyle(IXLWorksheet sheet)
        {
            sheet.Range(sheet.Cell(2, 1), sheet.Cell(21, actualColumn - 1)).Style.Fill.SetBackgroundColor(XLColor.FromArgb(174, 196, 210));    /* Couleur Bleu -- France */
            sheet.Range(sheet.Cell(23, 1), sheet.Cell(34, actualColumn - 1)).Style.Fill.SetBackgroundColor(XLColor.FromArgb(231, 175, 184));   /* Couleur Rouge -- Allemagne */
            sheet.Range(sheet.Cell(37, 1), sheet.Cell(48, actualColumn - 1)).Style.Fill.SetBackgroundColor(XLColor.FromArgb(238, 234, 156));   /* Couleur Jaune -- Belgique */

            sheet.Range(sheet.Cell(2, 1), sheet.Cell(15, actualColumn - 1)).Style.NumberFormat.SetFormat("# ### ### 000.00");                  /* Format séparateur de milliers */
            sheet.Range(sheet.Cell(16, 1), sheet.Cell(16, actualColumn - 1)).Style.NumberFormat.SetFormat("0.00");                             /* Format séparateur de milliers */
            sheet.Range(sheet.Cell(17, 1), sheet.Cell(48, actualColumn - 1)).Style.NumberFormat.SetFormat("# ### ### ##0.00");                 /* Format séparateur de milliers */

            sheet.Columns().AdjustToContents();
        }

        private void WriteHeader(Dictionary<int, IGoogleRecord> dataFormatedByCountry, IXLWorksheet sheet)
        {
            int currentRow = 1;

            foreach (var valueDataForm in dataFormatedByCountry)
            {
                int site = valueDataForm.Key;
                var results = (RenaudRecord)valueDataForm.Value;


                currentRow++;
                sheet.Cell(currentRow, 1).Value = $"{SwitchIdToCountryName(site)} - {sheetName}";
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Nombre total de pages vues";
                currentRow++;
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement mobile";
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement desktop";
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement tablet";
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Sessions avec engagement Total";
                currentRow++;
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Revenue Mobile";
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Revenue Desktop";
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Revenue Tablet";
                currentRow++;
                sheet.Cell(currentRow, 1).Value = "Revenue Total";
                currentRow++;
                currentRow++;

                if (site == 1)
                {
                    sheet.Cell(currentRow, 1).Value = "Temps réponse serveur total (ms)";
                    currentRow++;
                    sheet.Cell(currentRow, 1).Value = "Temps réponse serveur / WebPerf Total (s)";
                    currentRow++;
                    currentRow++;
                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Mobile / WebPerf Mobile (s)";
                    currentRow++;
                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Desktop / WebPerf Desktop (s)";
                    currentRow++;
                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Tablet / WebPerf Tablet (s)";
                    currentRow++;
                    sheet.Cell(currentRow, 1).Value = "Vitesse du site Total / WebPerf Total (s)";
                    currentRow++;
                }
            }
        }

        private void WriteData(IXLWorksheet sheet, Dictionary<int, IGoogleRecord> dataFormatedByCountry)
        {
            int currentRow = 1;


            foreach (var valueDataForm in dataFormatedByCountry)
            {
                int site = valueDataForm.Key;
                var results = (RenaudRecord)valueDataForm.Value;


                currentRow++;
                /* Entête de pays */
                sheet.Cell(currentRow, actualColumn).Value = columnName;
                currentRow++;
                /* Nombre de pages vues */
                sheet.Cell(currentRow, actualColumn).Value = results.TotalScreenPageViews;
                currentRow++;
                currentRow++;

                /* Sessions avec engagement */
                sheet.Cell(currentRow, actualColumn).Value = results.SessionsParAppareil.GetValueOrDefault("mobile", 0);
                currentRow++;

                sheet.Cell(currentRow, actualColumn).Value = results.SessionsParAppareil.GetValueOrDefault("desktop", 0);
                currentRow++;

                sheet.Cell(currentRow, actualColumn).Value = results.SessionsParAppareil.GetValueOrDefault("tablet", 0);
                currentRow++;

                double totalSessions = results.SessionsParAppareil.Values.Sum();
                sheet.Cell(currentRow, actualColumn).Value = totalSessions;
                currentRow++;

                currentRow++;

                /* Chiffre d'affaire */
                sheet.Cell(currentRow, actualColumn).Value = results.RevenueParAppareil.GetValueOrDefault("mobile", 0);
                currentRow++;

                sheet.Cell(currentRow, actualColumn).Value = results.RevenueParAppareil.GetValueOrDefault("desktop", 0);
                currentRow++;

                sheet.Cell(currentRow, actualColumn).Value = results.RevenueParAppareil.GetValueOrDefault("tablet", 0);
                currentRow++;

                double totalRevenue = results.RevenueParAppareil.Values.Sum();
                sheet.Cell(currentRow, actualColumn).Value = totalRevenue;
                currentRow++;

                currentRow++;

                /* Métriques spécifiques France */
                if (site == 1)
                {
                    sheet.Cell(currentRow, actualColumn).Value = results.TotalServerResponseDuration;
                    currentRow++;

                    double averageServerResponse = results.TotalServerResponseDuration / results.TotalEventCountWebPerf;
                    sheet.Cell(currentRow, actualColumn).Value = (averageServerResponse / 1000);
                    currentRow++;

                    currentRow++;

                    /* LoadDuration moyens */
                    double averageMobile = results.LoadDurationParAppareil.GetValueOrDefault("mobile", 0) / results.EventCountWebPerfParAppareil.GetValueOrDefault("mobile", 0);
                    sheet.Cell(currentRow, actualColumn).Value = (averageMobile / 1000);
                    currentRow++;
                    double averageDesktop = results.LoadDurationParAppareil.GetValueOrDefault("desktop", 0) / results.EventCountWebPerfParAppareil.GetValueOrDefault("desktop", 0);
                    sheet.Cell(currentRow, actualColumn).Value = (averageDesktop / 1000);
                    currentRow++;

                    double averageTablet = results.LoadDurationParAppareil.GetValueOrDefault("tablet", 0) / results.EventCountWebPerfParAppareil.GetValueOrDefault("tablet", 0);
                    sheet.Cell(currentRow, actualColumn).Value = (averageTablet / 1000);
                    currentRow++;

                    double averageLoadDuration = results.TotalLoadDuration / results.TotalEventCountWebPerf;
                    sheet.Cell(currentRow, actualColumn).Value = (averageLoadDuration / 1000);
                    currentRow++;
                }
            }
            actualColumn++;
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

