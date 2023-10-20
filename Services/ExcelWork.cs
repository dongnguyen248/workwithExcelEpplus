using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using excel.Model;
using OfficeOpenXml;

namespace excel.Services
{
    public class ExcelWork
    {
        public void PivotData(){
            var listEmp = ReadExcelFile();
            for(var i = 0 ; i<listEmp.Count ; i++){
                
            }
        }
        private static List<EmployeeModel> ReadExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var res = new List<EmployeeModel>();
          
            string FilePath = "C:\\Users\\HP\\Documents\\dongproject\\congt5.xlsx";
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                int colCount = ws.Dimension.End.Column;  //get Column Count
                int rowCount = ws.Dimension.End.Row;
                  //get row count
                for (int row = 2; row <= rowCount; row++)
                {
                      var emp = new EmployeeModel(); 
                    for (int col = 1; col <= colCount; col++)
                    {
                        if(col == 1){
                            emp.Id = row;
                            emp.Name = ws.Cells[row,col].Value.ToString();
                        }else{
                            emp.WorkDate =ws.Cells[row,col].Value.ToString();
                        }
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + ws.Cells[row, col].Value?.ToString().Trim());
                    }
                    res.Add(emp);
                }
            }
            return res;
        }


        
    }
}