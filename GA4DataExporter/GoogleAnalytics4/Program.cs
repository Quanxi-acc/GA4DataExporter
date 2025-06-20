namespace GoogleAnalytics4
{
    class Program
    {
        private static int Main()
        {
            var command = new GoogleDataCommand();

       //   command.outputType = GoogleDataOutputType.ExportExcelRenaud; 

            command.outputType = GoogleDataOutputType.ExportWebServiceSage;
 
            command.Execute();

            return 0;
        }
    }
}
