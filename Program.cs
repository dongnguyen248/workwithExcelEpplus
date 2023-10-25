using excel.Services;
internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Bắt đầu tính công:");
        var res = new ExcelWork();
        res.PivotData();
    }
}