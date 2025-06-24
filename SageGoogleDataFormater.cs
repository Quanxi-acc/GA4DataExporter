using Google.Analytics.Data.V1Beta;
using System.Globalization;

namespace GoogleAnalytics4
{
    public class SageGoogleDataFormater : IGoogleDataFormater
    {
        public IGoogleRecord Format(RunReportResponse reportResponse, int countryId)
        {
            var sageRecord = new SageRecord
            {
                TotalActive = 0,
                TotalNew = 0,
                TotalBounce = 0,
                CountryId = countryId 
            };

            foreach (var row in reportResponse.Rows)
            {
                if (row.DimensionValues[0].Value != null)
                {
                    sageRecord.TotalBounce += double.Parse(row.MetricValues[0].Value.Trim('"'), CultureInfo.InvariantCulture);
                    sageRecord.TotalActive += double.Parse(row.MetricValues[1].Value.Trim('"'), CultureInfo.InvariantCulture);
                    sageRecord.TotalNew += double.Parse(row.MetricValues[2].Value.Trim('"'), CultureInfo.InvariantCulture);
                }
            }

            return sageRecord;
        }

        public IGoogleRecord Format(RunReportResponse general, RunReportResponse webperf)
        {
            return Format(general, 0);
        }
    }
}