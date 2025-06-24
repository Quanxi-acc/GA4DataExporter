
using Google.Analytics.Data.V1Beta;

namespace GoogleAnalytics4
{
    class GoogleDataFetcher
    {
        private const string oogardenFranceId = "xxxxxxxxxxxxxx";
        private const string oogardenAllemagneId = "xxxxxxxxxxxxxx";
        private const string oogardenBelgiqueId = "xxxxxxxxxxxxxx";

        public string? startDate;
        public string? endDate;

        public RunReportResponse FetchClassicExcelMetrics(string site)
        {
            var client = BetaAnalyticsDataClient.Create(); //OAuth par variable d'environnement GOOGLE_APPLICATION_CREDENTIALS
            var request = new RunReportRequest
            {
                Property = "properties/" + site,
                Dimensions = { new Dimension { Name = "deviceCategory" } },
                DateRanges = { new DateRange { StartDate = startDate, EndDate = endDate } },
                Metrics =
                {
                    new Metric { Name = "screenPageViews" },
                    new Metric { Name = "purchaseRevenue" },
                    new Metric { Name = "sessions" }
                }
            };
            var result = client.RunReport(request);
            return result;
        }

        public RunReportResponse FetchExcelWebPerfMetrics(string site)
        {
            var client = BetaAnalyticsDataClient.Create();
            var request = new RunReportRequest
            {
                Property = "properties/" + site,
                Dimensions = { new Dimension { Name = "deviceCategory" } },
                DateRanges = { new DateRange { StartDate = startDate, EndDate = endDate } },
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
            var client = BetaAnalyticsDataClient.Create();
            var request = new RunReportRequest
            {   
                Property = "properties/" + site,
                Dimensions = { new Dimension { Name = "deviceCategory" } },
                DateRanges = { new DateRange { StartDate = startDate, EndDate = endDate } },
                Metrics =
                {
                    new Metric { Name = "bounceRate" },
                    new Metric { Name = "activeUsers" },
                    new Metric { Name = "newUsers" }
                }
            };
            return client.RunReport(request);
        }

        public Dictionary<int, (RunReportResponse fetchClassicExcelMetrics, RunReportResponse fetchWebPerfExcelMetrics)> Fetch()
        {
            return new Dictionary<int, (RunReportResponse, RunReportResponse)>
            {
                { 1, (FetchClassicExcelMetrics(oogardenFranceId), FetchExcelWebPerfMetrics(oogardenFranceId)) },
                { 2, (FetchClassicExcelMetrics(oogardenBelgiqueId), FetchExcelWebPerfMetrics(oogardenBelgiqueId)) },
                { 3, (FetchClassicExcelMetrics(oogardenAllemagneId), FetchExcelWebPerfMetrics(oogardenAllemagneId)) }
            };
        }

        public Dictionary<int, RunReportResponse> FetchDataForSage()
        {
            return new Dictionary<int, RunReportResponse>
            {
                { 1, FetchSageMetrics(oogardenFranceId) },
                { 2, FetchSageMetrics(oogardenBelgiqueId) },
                { 3, FetchSageMetrics(oogardenAllemagneId) }
            };
        }
    }
}
