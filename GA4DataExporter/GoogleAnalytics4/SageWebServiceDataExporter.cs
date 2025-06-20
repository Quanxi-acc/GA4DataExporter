using System.Text;
using System.Globalization;

namespace GoogleAnalytics4
{
    public class SageWebServiceDataExporter : IDataExporter
    {
        public void Export(Dictionary<int, IGoogleRecord> dataFormatedByCountry)
        {
            int[] countryIds = new int[] { 1, 2, 3 }; 

            foreach (var countryId in countryIds)
            {
                if (dataFormatedByCountry.TryGetValue(countryId, out var record))
                {
                    Export(record);
                }
            }
        }
        public void Export(IGoogleRecord data)
        {
            var sageData = data as SageRecord;

            var stringBuilder = new StringBuilder();

            stringBuilder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            stringBuilder.Append("<soapenv:Envelope xmlns:xsi =\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:wss=\"http://www.adonix.com/WSS\">");
            stringBuilder.Append("<soapenv:Header />");
            stringBuilder.Append("<soapenv:Body>");
            stringBuilder.Append("<wss:run soapenv:encodingStyle =\"http://schemas.xmlsoap.org/soap/encoding/\">");
            stringBuilder.Append("<callContext xsi:type=\"wss:CAdxCallContext\">");
            stringBuilder.Append("<codeLang xsi:type=\"xsd:string\">FRA</codeLang>");
            stringBuilder.Append("<poolAlias xsi:type=\"xsd:string\">XXXXXXX</poolAlias>");
            stringBuilder.Append("<poolId xsi:type=\"xsd:string\"/>");
            stringBuilder.Append("<requestConfig xsi:type=\"xsd:string\"/>");
            stringBuilder.Append("</callContext>");
            stringBuilder.Append("<publicName xsi:type=\"xsd:string\">XXXXXX</publicName>");
            stringBuilder.Append("<inputXml xsi:type=\"xsd:string\"><![CDATA[");
            stringBuilder.Append("<PARAM>");
            stringBuilder.Append("<GRP ID=\"IN\">");
            stringBuilder.Append($"<FLD NAM=\"ZDAT\">{DateTime.Now:yyyyMMdd}</FLD>");
            stringBuilder.Append($"<FLD NAM=\"ZIDPAYS\">{sageData.CountryId.ToString()}</FLD>");
            stringBuilder.Append($"<FLD NAM=\"ZNBRVIS\">{sageData.TotalActive.ToString("F2", CultureInfo.InvariantCulture)}</FLD>");
            stringBuilder.Append($"<FLD NAM=\"ZNBRNEW\">{sageData.TotalNew.ToString("F2", CultureInfo.InvariantCulture)}</FLD>");
            stringBuilder.Append($"<FLD NAM=\"ZTXREBOND\">{(sageData.TotalBounce * 100).ToString("F2", CultureInfo.InvariantCulture)}</FLD>");
            stringBuilder.Append("</GRP>");
            stringBuilder.Append("</PARAM>");
            stringBuilder.Append("]]></inputXml>");
            stringBuilder.Append("</wss:run>");
            stringBuilder.Append("</soapenv:Body>");
            stringBuilder.Append("</soapenv:Envelope>");

            var myRequest = stringBuilder.ToString();

            var httpClientHandler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
            };

            var client = new HttpClient(httpClientHandler)
            {
                Timeout = TimeSpan.FromSeconds(90)
            };

            var request = new HttpRequestMessage(HttpMethod.Post, "https://XXXXXXXXXXX")
            {
                Content = new StringContent(myRequest, Encoding.UTF8, "text/xml")
            };

            request.Headers.Add("Authorization", $"Basic {Convert.ToBase64String(Encoding.Default.GetBytes("XXXXX"))}");
            request.Headers.Add("cache-control", "no-cache");
            request.Headers.Add("Connection", "Keep-Alive");
            request.Headers.Add("soapAction", "run");

            var response = client.Send(request);

            var content = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        }

        
    }
}