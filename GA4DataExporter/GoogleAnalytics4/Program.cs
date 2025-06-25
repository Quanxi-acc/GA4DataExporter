using Google.Apis.Sheets.v4;

namespace GoogleAnalytics4
{
    class ExportExcelDataToDriveForRenaud
    {
        public ExportExcelDataToDriveForRenaud()
        {
        }

        private static int Main(string[] args)
        {
            string driveStartDate = args[0];
            string driveEndDate   = args[1];
            string sageStartDate  = args[2];
            string sageEndDate    = args[3];
        //  string pathway        = args[4];
            
            var command = new GoogleDataCommand();

            GoogleDataFetcher.SageStartDate = sageStartDate;
            GoogleDataFetcher.SageEndDate = sageEndDate;

            GoogleDataFetcher.DriveStartDate = driveStartDate;
            GoogleDataFetcher.DriveEndDate = driveEndDate;

         // GoogleDataCommand.Pathway = pathway;


          // command.outputType = GoogleDataOutputType.ExportExcelRenaud;         // Export vers Excel uniquement
           command.outputType = GoogleDataOutputType.ExportGoogleSheetsRenaud;  // Export vers Excel puis Google Sheets
          // command.outputType = GoogleDataOutputType.ExportWebServiceSage;      // Export vers Sage

            command.Execute();

            return 0;
        }
    }
}