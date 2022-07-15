using OfficeOpenXml;
using System.IO;

namespace ExcelDemoApp;

public class Program
{
    public static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var file = new FileInfo(@"C:\Excel\ExcelDemo.xlsx");

        var people = GetSetupData();

        await SaveExcelFile(people, file);
    }


    private static List<PersonModel> GetSetupData()
    {
        List<PersonModel> output = new()
        {
            new() {Id =1, FirstName="Kill", LastName="Monger"},
            new() {Id =2, FirstName="Storm", LastName="Front"},
            new() {Id =3, FirstName="Bear", LastName="Hug"}
        };

        return output;
    }

    private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
    {
        DeleteIfExists(file);

        using var package = ExcelPackage(file);

        
    }

    private static void DeleteIfExists(FileInfo file)
    {
        if (file.Exists)
        {
            file.Delete();
        }
    }
}
