# DynamicRecordTestProject

Test project for reading Excel files with DynamicObject

## Sample

Example use ClosedXML

```csharp
public IEnumerable<dynamic> ReadExcelData()
{
    using (IXLWorkbook workbook = new XLWorkbook(Path))
    {
        IXLWorksheet worksheet = workbook.Worksheet(1);

        // Get XLSX ColumnItemNames
        var tables = worksheet.RangeUsed().AsTable();
        var columnNames = tables.Fields.Select(field => field.Name);

        // Get XLSX Row Data(IXLRangeRows)
        var values = tables.DataRange.Rows();

        var generator = new DataRecordGenerator(columnNames, values);
        return generator.Generate();
    }
}

public IEnumerable<dynamic> Generate()
{
    foreach (var row in rows)
    {
        var dic = columnNames.Select((name, index) => (name, index))
            .Select(x => (x.name, row.Cell(x.index + 1).Value))
            .ToDictionary(k => k.name, v => v.Value);
        yield return new DataRecord(dic);
    }
}
```

## SampleData and TestCode

Sample File Extension is (.xlsx).

| Id  | Name | Affiliated      | Age | Position         |
| --- | ---- | --------------- | --- | ---------------- |
| 1   | A1   | General Affairs | 28  | General manager  |
| 2   | B2   | Human Resources | 19  | Chief            |
| 3   | C3   | Developments    | 34  | Manager          |
| 4   | D4   | Developments    | 23  | Chief            |
| 5   | E5   | Productions     | 18  | General manager  |
| 6   | F6   | Productions     | 69  | General employee |

```csharp
[TestMethod]
public void TestMethod1()
{
    var excelControl = new ExcelControl(path);

    var result = excelControl.ReadExcelData().ToArray();

    foreach (var item in result)
    {
        Console.WriteLine($"{item.Id},{item.Name},{item.Affiliated},{item.Age},{item.Position}");
    }

    result[0].Age = (double)50;
    result[0].Affiliated = "Officer";
    result[0].Position = "OperatingOfficer";

    Console.WriteLine($"{result[0].Id},{result[0].Name},{result[0].Affiliated},{result[0].Age},{result[0].Position}");
}
```

## SampleResult

```txt
1,A1,General Affairs,28,General manager
2,B2,Human Resources,19,Chief
3,C3,Developments,34,Manager
4,D4,Developments,23,Chief
5,E5,Productions,18,General manager
6,F6,Productions,69,General employee
1,A1,Officer,50,OperatingOfficer
```
