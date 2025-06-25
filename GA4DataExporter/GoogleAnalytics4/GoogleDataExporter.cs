using Google.Analytics.Data.V1Beta;

namespace GoogleAnalytics4
{
    public class GoogleDataExporter
    {
        public IGoogleDataFormater googleDataFormater;

        public List<IDataExporter> dataExporters = new List<IDataExporter>();
        public void Export((RunReportResponse general, RunReportResponse webperf) reportResponses)
        {
            var data = googleDataFormater.Format(reportResponses.general, reportResponses.webperf);
            foreach (var exporter in dataExporters)
            {
                exporter.Export(data);
            }
        }

        internal void Export(Dictionary<int, (RunReportResponse general, RunReportResponse webperf)> googleData)
        {
            var dataFormatedByCountry = new Dictionary<int, IGoogleRecord>();

            foreach (var countryId in googleData.Keys)
            {
                var responses = googleData[countryId];
                dataFormatedByCountry.Add(countryId, googleDataFormater.Format(responses.general, responses.webperf));
            }

            foreach (var exporter in dataExporters)
            {
                exporter.Export(dataFormatedByCountry);
            }
        }
    }
}