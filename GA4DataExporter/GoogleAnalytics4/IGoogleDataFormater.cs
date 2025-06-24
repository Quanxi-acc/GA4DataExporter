using Google.Analytics.Data.V1Beta;

namespace GoogleAnalytics4
{
    public interface IGoogleDataFormater
    {
        IGoogleRecord Format(RunReportResponse classicMetrics, RunReportResponse webperfMetrics);

        IGoogleRecord Format(RunReportResponse site, int countryId);
    }
}