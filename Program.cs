using excel.Services;

namespace excel;
class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        var res = new ExcelWork();
        res.ReadExcelFile();
    }
}
