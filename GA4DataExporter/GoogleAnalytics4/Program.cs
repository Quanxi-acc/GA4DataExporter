
namespace GoogleAnalytics4
{
    class ExportExcelDataToDriveForRenaud
    {
        public ExportExcelDataToDriveForRenaud()
        {
        }

        private static int Main(string[] args)
        {
            
            var command = new GoogleDataCommand();

            switch (args[0])
            {
                case "SageExport":
                    string startDate = args[1];
                    string endDate   = args[2];
                    command.StartDate = DateTime.Parse(startDate); 
                    command.EndDate = DateTime.Parse(endDate);
                    command.outputType = GoogleDataOutputType.ExportWebServiceSage;         /* Export vers Sage */
                    break;
                case "Excel":
                    string month = args[1];
                    command.StartDate = DateTime.Parse(month);
                    command.EndDate = DateTime.Parse(month).LastDayOfMonth();
                    command.outputType = GoogleDataOutputType.ExportExcelRenaud;            /* Export vers Excel uniquement */
                    break;
                case "Googlesheets":
                    command.outputType = GoogleDataOutputType.ExportGoogleSheetsRenaud;     /* Export vers Excel puis Google Sheets */
                    break;
                default:
                    Console.WriteLine("Usage: GoogleAnalytics4 <startDate> <endDate> <outputType>");
                    Console.WriteLine("outputType: sage | excel | googlesheets");
                    return 1;
            }

            command.Execute();

            return 0;
        }
    }
}