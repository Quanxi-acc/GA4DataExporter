using ExcelToGoogle;
using Google.Analytics.Data.V1Beta;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;

namespace GoogleAnalytics4
{
    public enum GoogleDataOutputType
    {
        ExportWebServiceSage = 1,
        ExportExcelRenaud = 2,
        ExportGoogleSheetsRenaud = 3
    }

    public class GoogleDataCommand
    {
        public GoogleDataOutputType outputType;
        public void Execute()
        {
            if (outputType == GoogleDataOutputType.ExportWebServiceSage)
            {
                var googleDataForSageByCountry = FetchSageGoogleData();
                foreach (var countryData in googleDataForSageByCountry)
                {
                    ExportToWebServiceSage(countryData.Value, countryData.Key);
                }
            }
            else if (outputType == GoogleDataOutputType.ExportExcelRenaud)
            {
                var googleData = FetchGoogleData();
                ExportExcelRenaud(googleData);
            }
            else if (outputType == GoogleDataOutputType.ExportGoogleSheetsRenaud)
            {
                var googleData = FetchGoogleData();
                ExportGoogleSheetsRenaud(googleData);
            }
        }

        public static void ExportExcelRenaud(Dictionary<int, (RunReportResponse classicMetrics, RunReportResponse webperfMetrics)> googleData)
        {
            var dataExporter = new GoogleDataExporter
            {
                googleDataFormater = new RenaudGoogleDataFormater()
            };

            dataExporter.dataExporters.Add(new RenaudExcelDataExporter());
            dataExporter.Export(googleData);
        }

        public static void ExportGoogleSheetsRenaud(Dictionary<int, (RunReportResponse classicMetrics, RunReportResponse webperfMetrics)> googleData)
        {
            var dataExporter = new GoogleDataExporter
            {
                googleDataFormater = new RenaudGoogleDataFormater()
            };

            var excelExporter = new RenaudExcelDataExporter();
            dataExporter.dataExporters.Add(excelExporter);
            dataExporter.Export(googleData);

            var sheetService = ExportDriveRenaud.ConnectToGoogle();
            ExportDriveRenaud.ExportDataFromExcelToGoogleSheet(sheetService, @"C:\Users\stagiaireit\Desktop\OOG_GA4_Report.xlsx");
        }

        public static Dictionary<int, (RunReportResponse classicMetrics, RunReportResponse webperfMetrics)> FetchGoogleData()
        {
            var googleFetcher = new GoogleDataFetcher();
            googleFetcher.startDate = "2025-06-23";
            googleFetcher.endDate = "2025-06-24";

            return googleFetcher.Fetch();
        }

        public static void ExportToWebServiceSage(RunReportResponse countryData, int countryId)
        {
            var dataExporter = new GoogleDataExporter
            {
                googleDataFormater = new SageGoogleDataFormater()
            };

            dataExporter.dataExporters.Add(new SageWebServiceDataExporter());

            var sageRecord = dataExporter.googleDataFormater.Format(countryData, countryId);
            foreach (var exporter in dataExporter.dataExporters)
            {
                exporter.Export(sageRecord);
            }
        }

        public static Dictionary<int, RunReportResponse> FetchSageGoogleData()
        {
            var googleFetcher = new GoogleDataFetcher
            {
                startDate = "2025-06-23",
                endDate = "2025-06-24"
            };

            return googleFetcher.FetchDataForSage();
        }
    }
}