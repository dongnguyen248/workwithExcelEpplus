using excel.Services;
internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        var res = new ExcelWork();
        res.PivotData();
    }
}