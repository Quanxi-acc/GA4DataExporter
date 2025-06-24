using Google.Apis.Sheets.v4;

namespace GoogleAnalytics4
{
    class ExportExcelDataToDriveForRenaud
    {
        public ExportExcelDataToDriveForRenaud()
        {
        }

        private static int Main()
        {
            var command = new GoogleDataCommand();

         // command.outputType = GoogleDataOutputType.ExportExcelRenaud;         // Export vers Excel uniquement
         // command.outputType = GoogleDataOutputType.ExportGoogleSheetsRenaud;  // Export vers Excel puis Google Sheets
         // command.outputType = GoogleDataOutputType.ExportWebServiceSage;      // Export vers Sage

            command.Execute();

            return 0;
        }
    }
}
