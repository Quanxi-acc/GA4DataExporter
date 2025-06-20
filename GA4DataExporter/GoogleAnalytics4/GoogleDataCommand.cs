using Google.Analytics.Data.V1Beta;

namespace GoogleAnalytics4
{
    public enum GoogleDataOutputType
    {
        ExportWebServiceSage = 1,
        ExportExcelRenaud = 2,
        ExportSheetToDriveRenaud = 3
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


        public static Dictionary<int, (RunReportResponse general, RunReportResponse webperf)> FetchGoogleData()
        {
            var googleFetcher = new GoogleDataFetcher();
            googleFetcher.startDate = "2025-06-01";
            googleFetcher.endDate = "2025-06-18";
            return googleFetcher.Fetch();
        }

        public static void ExportExcelRenaud(Dictionary<int, (RunReportResponse general, RunReportResponse webperf)> googleData)
        {
            var dataExporter = new GoogleDataExporter
            {
                googleDataFormater = new RenaudGoogleDataFormater()
            };

            dataExporter.dataExporters.Add(new RenaudExcelDataExporter());

            dataExporter.Export(googleData);
        }



        public static Dictionary<int, RunReportResponse> FetchSageGoogleData()
        {
            var googleFetcher = new GoogleDataFetcher
            {
                startDate = "2025-06-01",
                endDate = "2025-06-18"
            };

            return googleFetcher.FetchDataForSage();
        }
    }
}