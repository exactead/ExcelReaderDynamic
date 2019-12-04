using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderDynamic
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelControl = new ExcelControl(@"JinjiFile.xlsx");
            Console.WriteLine("----DynamicObject----");
            var resultDynamic = excelControl.ReadExcelDataDynamic();
            foreach (var item in resultDynamic)
            {
                Console.WriteLine($"{item.Id},{item.Name},{item.Affiliated},{item.Age},{item.Position}");
            }

            Console.WriteLine("----T Parameter----");
            var resultT = excelControl.ReadExcelData<Model>();
            foreach (var item in resultT)
            {
                Console.WriteLine($"{item.Id},{item.Name},{item.Affiliated},{item.Age},{item.Position}");
            }

            Console.ReadKey();
        }
    }
}
