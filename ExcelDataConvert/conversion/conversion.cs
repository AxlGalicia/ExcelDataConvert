using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataConvert.conversion
{
    public class conversionProcess
    {

        public conversionProcess()
        {

        }

        public async Task convertCsv(string pathExcel, string pathCvs, string numberCvs)
        {
            string pathExcelComplete = completePath(pathExcel);
            string pathCvsComplete = completePath(pathCvs);

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string resulCvs = "";

            using (var stream = File.Open(pathExcelComplete, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    // Ejemplos de acceso a datos
                    DataTable table = result.Tables[(int.Parse(numberCvs)-1)];

                    var totalRows = table.Rows.Count;
                    var totalColumns = table.Columns.Count;
                    //Console.WriteLine("Filas totales: " + totalRows.ToString());
                    //Console.WriteLine("Columnas totales: " + totalColumns.ToString());

                    string dato = "";

                    for (int x = 0; x < totalRows; x++)
                    {

                        for (int i = 0; i < totalColumns; i++)
                        {
                            dato = table.Rows[x][i].ToString();

                            if (i < (totalColumns - 1))
                            {
                                resulCvs = resulCvs + dato + ",";
                            }
                            else
                            {
                                resulCvs = resulCvs + dato;
                            }

                        }

                        if (x < (totalRows - 1))
                        {
                            resulCvs = resulCvs + "\n";
                        }

                    }

                   // Console.WriteLine(resulCvs);
                   Console.WriteLine("Process Success!");

                }
            }

            await save(resulCvs,pathCvsComplete);


        }

        public string completePath(string pathFile)
        {

            var pathParts = pathFile.Trim().Split((char)92);

            string pathResult = "";
            for (int i = 0; i < pathParts.Length; i++)
            {
                if (i < (pathParts.Length - 1))
                {
                    pathResult = pathResult + pathParts[i] + (char)92 + (char)92;
                }
                else
                {
                    pathResult = pathResult + pathParts[i];
                }
            }
            Console.WriteLine("Este es el path ingresado: " + pathResult);

            return pathResult;


        }

        public static async Task save(string resulCvs,string pathCvs)
        {

            await File.WriteAllTextAsync(pathCvs, resulCvs, Encoding.Latin1);

        }



    }
}
