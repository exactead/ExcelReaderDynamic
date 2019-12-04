using System;
using System.Linq;
using ExcelReaderDynamic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelControl = new ExcelControl(@"JinjiFile.xlsx");

            var result = excelControl.ReadExcelDataDynamic().ToArray();

            foreach (var item in result)
            {
                Console.WriteLine($"{item.Id},{item.Name},{item.Affiliated},{item.Age},{item.Position}");
            }

            result[0].Age = (double)50;
            result[0].Affiliated = "役員";
            result[0].Position = "執行役員";

            Console.WriteLine($"{result[0].Id},{result[0].Name},{result[0].Affiliated},{result[0].Age},{result[0].Position}");
        }  

        [TestMethod]
        public void TestMethod2()
        {
            var excelControl = new ExcelControl(@"JinjiFile.xlsx");

            var result = excelControl.ReadExcelData<Model>().ToArray();
            foreach (var item in result)
            {
                Console.WriteLine($"{item.Id},{item.Name},{item.Affiliated},{item.Age},{item.Position}");
            }

            result[0].Age = 50;
            result[0].Affiliated = "役員";
            result[0].Position = "執行役員";

            Console.WriteLine($"{result[0].Id},{result[0].Name},{result[0].Affiliated},{result[0].Age},{result[0].Position}");
        }
    }
}
