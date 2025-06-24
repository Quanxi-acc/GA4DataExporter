using Google.Analytics.Data.V1Beta;

namespace GoogleAnalytics4
{
    public interface IDataExporter
    {
        void Export(IGoogleRecord data);
        void Export(Dictionary<int, IGoogleRecord> dataFormatedByCountry);
    }
}
