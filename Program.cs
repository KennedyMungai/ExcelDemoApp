using OfficeOpenXml;
using System.IO;

namespace ExcelDemoApp;

public class Program
{
    public static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var file = new FileInfo(@"C:\Excel\ExcelDemo.xlsx");
    }
}
