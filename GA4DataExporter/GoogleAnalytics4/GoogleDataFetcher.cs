using DocumentFormat.OpenXml.Wordprocessing;
using Google.Analytics.Data.V1Beta;

namespace GoogleAnalytics4
{
    class GoogleDataFetcher
    {
        private static GoogleAnalytics4Settings settings = new GoogleAnalytics4Settings();

        private  string oogardenFranceId = settings.OOGardenFranceId;
        private  string oogardenAllemagneId = settings.OOGardenAllemagneId;
        private  string oogardenBelgiqueId = settings.OOGardenBelgiqueId;

        public static string StartDate { get; set; }
        public static string EndDate { get; set; }


        public RunReportResponse FetchClassicExcelMetrics(string site)
        {
            var client = BetaAnalyticsDataClient.Create(); /* OAuth par variable d'environnement GOOGLE_APPLICATION_CREDENTIALS */
            var request = new RunReportRequest
            {
                Property = "properties/" + site,
                Dimensions = { new Dimension { Name = "deviceCategory" } },
                DateRanges = { new DateRange { StartDate = StartDate, EndDate = EndDate } },
                Metrics =
                {
                    new Metric { Name = "screenPageViews" },
                    new Metric { Name = "purchaseRevenue" },
                    new Metric { Name = "sessions" }
                }
            };
            return client.RunReport(request);
        }

        public RunReportResponse FetchExcelWebPerfMetrics(string site)
        {
            var client = BetaAnalyticsDataClient.Create(); /* OAuth par variable d'environnement GOOGLE_APPLICATION_CREDENTIALS */
            var request = new RunReportRequest
            {
                Property = "properties/" + site,
                Dimensions = { new Dimension { Name = "deviceCategory" } },
                DateRanges = { new DateRange { StartDate = StartDate, EndDate = EndDate } },
                DimensionFilter = new FilterExpression
                {
                    Filter = new Filter
                    {
                        FieldName = "eventName",
                        StringFilter = new Filter.Types.StringFilter
                        {
                            MatchType = Filter.Types.StringFilter.Types.MatchType.Exact,
                            Value = "webperf"
                        }
                    }
                },
                Metrics =
                {
                    new Metric { Name = "eventCount" },
                    new Metric { Name = "sessions" }
                }
            };

            if (site == oogardenFranceId)
            {
                request.Metrics.Add(new Metric { Name = "customEvent:ServerResponseDuration" });
                request.Metrics.Add(new Metric { Name = "customEvent:LoadDuration" });
            }
            return client.RunReport(request);
        }

        public RunReportResponse FetchSageMetrics(string site)
        {
            var client = BetaAnalyticsDataClient.Create(); /* OAuth par variable d'environnement GOOGLE_APPLICATION_CREDENTIALS */
            var request = new RunReportRequest
            {
                Property = "properties/" + site,
                Dimensions = { new Dimension { Name = "deviceCategory" } },
                DateRanges = { new DateRange { StartDate = StartDate, EndDate = EndDate } },
                Metrics =
                {
                    new Metric { Name = "bounceRate" },
                    new Metric { Name = "activeUsers" },
                    new Metric { Name = "newUsers" }
                }
            };
            return client.RunReport(request);
        }

        public Dictionary<int, (RunReportResponse fetchClassicExcelMetrics, RunReportResponse fetchWebPerfExcelMetrics)> Fetch(string startDate, string endDate)
        {
            StartDate = startDate;
            EndDate = endDate;

            return new Dictionary<int, (RunReportResponse, RunReportResponse)>
            {
                { 1, (FetchClassicExcelMetrics(oogardenFranceId), FetchExcelWebPerfMetrics(oogardenFranceId)) },
                { 2, (FetchClassicExcelMetrics(oogardenBelgiqueId), FetchExcelWebPerfMetrics(oogardenBelgiqueId)) },
                { 3, (FetchClassicExcelMetrics(oogardenAllemagneId), FetchExcelWebPerfMetrics(oogardenAllemagneId)) }
            };
        }

        public Dictionary<int, RunReportResponse> FetchDataForSage(string startDate, string endDate)
        {
            StartDate = startDate;
            EndDate = endDate;

            return new Dictionary<int, RunReportResponse>
            {
                { 1, FetchSageMetrics(oogardenFranceId) },
                { 2, FetchSageMetrics(oogardenBelgiqueId) },
                { 3, FetchSageMetrics(oogardenAllemagneId) }
            };
        }
    }
}