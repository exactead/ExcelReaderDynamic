using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ExcelReaderDynamic
{
    public class ExcelControl
    {
        public string Path { get; set; }

        public ExcelControl(string path)
        {
            this.Path = path;
        }

        /// <summary>
        /// DynamicObject使用による読み込み
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> ReadExcelDataDynamic()
        {
            using (IXLWorkbook workbook = new XLWorkbook(Path))
            {
                IXLWorksheet worksheet = workbook.Worksheet(1);

                // 項目名称の取得
                var tables = worksheet.RangeUsed().AsTable();
                var columnNames = tables.Fields.Select(field => field.Name);
                var values = tables.DataRange.Rows();

                // 生成開始
                var generator = new DataRecordGenerator(columnNames, values);
                return generator.Generate();
            }
        }

        /// <summary>
        /// 比較用　Tパラメータ使用による読み込み
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public IEnumerable<T> ReadExcelData<T>() where T : new()
        {
            using (IXLWorkbook workbook = new XLWorkbook(Path))
            {
                IXLWorksheet worksheet = workbook.Worksheet(1);

                var tables = worksheet.RangeUsed().AsTable();
                var fields = tables.Fields.Select(field => field.Name);
                foreach (var tableRow in tables.DataRange.Rows())
                {
                    var obj = new T();
                    foreach (var field in fields)
                    {
                        PropertyInfo propertyInfo = typeof(T).GetProperties().Where(p => p.CanRead && p.CanWrite).FirstOrDefault(x => x.Name == field);
                        propertyInfo?.SetValue(obj, Convert.ChangeType(tableRow.Field(field).Value, propertyInfo.PropertyType));
                    }
                    yield return obj;
                }
            }
        }
    }
}
