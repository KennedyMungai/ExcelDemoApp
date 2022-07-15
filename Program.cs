using OfficeOpenXml;
using System.IO;

namespace ExcelDemoApp;

public class Program
{
    public static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string userRoot = System.Environment.GetEnvironmentVariable("USERPROFILE");
        string downloadPath=Path.Combine(userRoot, "Downloads");

        var file = new FileInfo(downloadPath);

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

        using var package = new ExcelPackage(file);

        var ws = package.Workbook.Worksheets.Add("MainReport");

        var range = ws.Cells["A2"].LoadFromCollection(people, true);

        range.AutoFitColumns();

        ws.Cells["A1"].Value = "Our cool report";
        ws.Cells["A1:C1"].Merge = true;

        await package.SaveAsync();
    }

    private static void DeleteIfExists(FileInfo file)
    {
        if (file.Exists)
        {
            file.Delete();
        }
    }
}
