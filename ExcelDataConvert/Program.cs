using ExcelDataConvert.configurations;
using ExcelDataConvert.conversion;
using ExcelDataConvert.resources;

namespace ExcelDataReader 
{

    class Program
    {

        static async Task Main(string[] args)
        {

            configurations config = new configurations();

            if (config.getconfigurations())
            {
                Console.WriteLine();
                Console.WriteLine(stringResource.stringConvertProcess);
                Console.WriteLine();
                conversionProcess conver = new conversionProcess();
                await conver.convertCsv(config.pathExcel,config.pathCsv,config.sheetOption);

            }
            else
            {
                Console.WriteLine(stringResource.stringBye);
            }

            //config.getconfigurations();

            //conversion conver = new conversion();

        }



    }

}
