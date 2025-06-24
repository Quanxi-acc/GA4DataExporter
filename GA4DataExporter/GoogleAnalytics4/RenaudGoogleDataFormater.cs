using Google.Analytics.Data.V1Beta;
using System.Globalization;

namespace GoogleAnalytics4
{
    public class RenaudGoogleDataFormater : IGoogleDataFormater
    {
        public IGoogleRecord Format(RunReportResponse generalMetrics, RunReportResponse webperfMetrics)
        {
            var results = new RenaudRecord();

            foreach (var row in generalMetrics.Rows)
            {
                string device = row.DimensionValues[0].Value;

                double screenPageViews = double.Parse(row.MetricValues[0].Value.Trim('"').Replace(".", ","));
                double revenue = double.Parse(row.MetricValues[1].Value.Trim('"').Replace(".", ","));
                double sessions = double.Parse(row.MetricValues[2].Value.Trim('"').Replace(".", ","));

                results.TotalScreenPageViews += screenPageViews;
                results.RevenueParAppareil[device] = revenue;
                results.SessionsParAppareil[device] = sessions;
            }

            foreach (var row in webperfMetrics.Rows)
            {
                string device = row.DimensionValues[0].Value;

                double eventCount = double.Parse(row.MetricValues[0].Value.Trim('"').Replace(".", ","));
                results.TotalEventCountWebPerf += eventCount;

                if (!results.EventCountWebPerfParAppareil.ContainsKey(device))
                    results.EventCountWebPerfParAppareil[device] = 0;
                    results.EventCountWebPerfParAppareil[device] += eventCount;

                if (row.MetricValues.Count > 3)
                {
                    double serverResponseDuration = double.Parse(row.MetricValues[2].Value.Trim('"').Replace(".", ","));
                    double loadDuration = double.Parse(row.MetricValues[3].Value.Trim('"').Replace(".", ","));

                    results.TotalServerResponseDuration += serverResponseDuration;
                    results.TotalLoadDuration += loadDuration;
                    results.LoadDurationParAppareil[device] = loadDuration;
                }
            }

            return results;
        }

        public IGoogleRecord Format(RunReportResponse site)
        {
            throw new NotImplementedException();
        }

        public IGoogleRecord Format(RunReportResponse site, int countryId)
        {
            throw new NotImplementedException();
        }
    }
}