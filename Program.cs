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


    static List<PersonModel> GetSetupData()
    {
        List<PersonModel> output = new()
        {
            new() {Id =1, FirstName="Kill", LastName="Monger"},
            new() {Id =2, FirstName="Storm", LastName="Front"},
            new() {Id =3, FirstName="Bear", LastName="Hug"}
        };

        return output;
    }
}
