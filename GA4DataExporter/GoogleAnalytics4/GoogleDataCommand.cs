using ExcelToGoogle;
using Google.Analytics.Data.V1Beta;
using Google.Apis.Auth.OAuth2;
using Grpc.Net.Client;


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
        private GoogleAnalytics4Settings settings = new GoogleAnalytics4Settings();

        public GoogleDataOutputType outputType;

        public DateTime StartDate { get; internal set; }
        public DateTime EndDate { get; internal set; }

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
                ExportExcelForRenaud();
            }
            else if (outputType == GoogleDataOutputType.ExportGoogleSheetsRenaud)
            {
                ExportExcelForRenaud();
                ExportGoogleSheetsRenaud();
            }
        }

        private void ExportExcelForRenaud()
        {
            var googleData = FetchGoogleData();
            ExportExcelRenaud(googleData, $"{StartDate:yyyy-MM}", $"{StartDate:yyyy-MM}");

            for(int i = 0; i < (EndDate - StartDate).TotalDays; i++)
            {
                googleData = FetchGoogleData(StartDate.AddDays(i));
                ExportExcelRenaud(googleData, i == 0 ? "Journalier" : null, $"{StartDate.AddDays(i):MM-dd}");
            }
        }

        public static void ExportExcelRenaud(Dictionary<int, (RunReportResponse classicMetrics, RunReportResponse webperfMetrics)> googleData, string sheetName, string columnName)
        {
            var dataExporter = new GoogleDataExporter
            {
                googleDataFormater = new RenaudGoogleDataFormater()
            };

            dataExporter.dataExporters.Add(new RenaudExcelDataExporter(sheetName, columnName) { CreateNewSheet = sheetName != null});
            dataExporter.Export(googleData);
        }

        public void ExportGoogleSheetsRenaud()
        {
            var sheetService = ExportDriveRenaud.ConnectToGoogle();
            ExportDriveRenaud.ExportDataFromExcelToGoogleSheet(sheetService, settings.OutputExcelFilePath);
        }

        public Dictionary<int, (RunReportResponse classicMetrics, RunReportResponse webperfMetrics)> FetchGoogleData()
        {
            var googleFetcher = new GoogleDataFetcher();
            return googleFetcher.Fetch(StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"));
        }

        public Dictionary<int, (RunReportResponse classicMetrics, RunReportResponse webperfMetrics)> FetchGoogleData(DateTime dayToFetch)
        {
            var googleFetcher = new GoogleDataFetcher();
            return googleFetcher.Fetch(dayToFetch.ToString("yyyy-MM-dd"), dayToFetch.ToString("yyyy-MM-dd"));
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

        public Dictionary<int, RunReportResponse> FetchSageGoogleData()
        {
            var googleFetcher = new GoogleDataFetcher();
            return googleFetcher.FetchDataForSage(StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"));
        }
    }

    public static class DateTimeDayOfMonthExtensions
    {
        public static DateTime FirstDayOfMonth(this DateTime value)
        {
            return new DateTime(value.Year, value.Month, 1);
        }

        public static int DaysInMonth(this DateTime value)
        {
            return DateTime.DaysInMonth(value.Year, value.Month);
        }

        public static DateTime LastDayOfMonth(this DateTime value)
        {
            return new DateTime(value.Year, value.Month, value.DaysInMonth());
        }
    }
}
