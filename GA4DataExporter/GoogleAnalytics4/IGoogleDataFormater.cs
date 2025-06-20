using Google.Analytics.Data.V1Beta;

namespace GoogleAnalytics4
{
    public interface IGoogleDataFormater
    {
        IGoogleRecord Format(RunReportResponse genesiteral, RunReportResponse webperf);

        IGoogleRecord Format(RunReportResponse site, int countryId);
    }
}