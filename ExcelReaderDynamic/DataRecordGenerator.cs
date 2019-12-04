using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReaderDynamic
{
    internal class DataRecordGenerator
    {
        IEnumerable<string> columnNames;
        IXLTableRows xLTableRows;

        public DataRecordGenerator(IEnumerable<string> columnNames, IXLTableRows rows)
        {
            this.columnNames = columnNames;
            xLTableRows = rows;
        }

        public IEnumerable<dynamic> Generate()
        {
            foreach (var row in xLTableRows)
            {
                var dic = columnNames.Select((name, index) => (name, index))
                    .Select(x => (x.name, row.Cell(x.index + 1).Value))
                    .ToDictionary(k => k.name, v => v.Value);
                yield return new DataRecord(dic);
            }
        }
    }
}
