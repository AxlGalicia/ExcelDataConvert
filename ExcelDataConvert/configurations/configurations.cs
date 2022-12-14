using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataConvert.conversion;
using ExcelDataConvert.resources;


namespace ExcelDataConvert.configurations
{
    public  class configurations
    {
        public string pathExcel { get; set; } = String.Empty;
        public string pathCsv { get; set; } = String.Empty;
        public string sheetOption { get; set; } = String.Empty; 
        public configurations()
        {
            //Console.WriteLine("Welcome to configurations");
        }

        public Boolean getconfigurations()
        {
            Console.Write(stringResource.stringPathExcel);
            pathExcel = Console.ReadLine() ?? "";

            Console.Write(stringResource.stringNumberSheet);
            sheetOption = Console.ReadLine()?? "";

            Console.Write(stringResource.stringPathCsv);
            pathCsv = Console.ReadLine() ?? "";

            return confirmOperations();

        }

        public Boolean  confirmOperations()
        {
            string confirmation = "";

            Console.WriteLine(stringResource.stringSeparetor);
            Console.WriteLine(stringResource.stringSummaryHead);
            Console.WriteLine();
            Console.WriteLine(stringResource.stringSummaryBodyExcel + pathExcel);
            Console.WriteLine(stringResource.stringSummaryBodyCsv + pathCsv);
            Console.WriteLine(stringResource.stringSummaryBodySheet+ sheetOption);
            Console.WriteLine();
            Console.Write(stringResource.stringSummaryConfirmation);
            confirmation = Console.ReadLine()??"Cancel";

            if(confirmation == "Si" || confirmation == "si")
            {
                Console.WriteLine();
                //Console.WriteLine("Convirtiendo documento a Csv..................................................");
                return true;
                //conversionProcess conver = new conversionProcess();
                //conver.convertCsv(pathExcel,pathCsv,sheetOption);
            }
            else if(confirmation == "Cancel" || confirmation == "cancel")
            {
                return false;
            }else
            {
                Console.Clear();
                return getconfigurations();
            }


        }

    }
}
