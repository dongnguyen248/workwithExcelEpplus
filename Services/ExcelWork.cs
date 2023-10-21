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
        public void PivotData()
        {
            var listEmp = ReadExcelFile();
            var listEmpDateTime = ConvertToDateTime(listEmp);
            var listEachEmpWt = new List<EmpWorkingTime>();
            for (var i = 1; i <= listEmpDateTime.GroupBy(x=>x.Name).Count(); i++)
            {
                var lEmpInMonth = listEmpDateTime.Where(x => x.Name == listEmpDateTime[i].Name);
                for (var day = 1; day <= listEmpDateTime[i].NumberDayOfMonth; day++)
                {
                    DateTime dayOfMonth = new DateTime(listEmpDateTime[i].WorkDay.Value.Year, listEmpDateTime[i].WorkDay.Value.Month, day);
                    var workingDayemp = lEmpInMonth.Where(x => x.DateFormat == dayOfMonth.ToString("MM/dd/yyyy").Replace("/", ""));
                    var EachEmpWt = CalculateWokingTime(workingDayemp.ToList(), dayOfMonth);

                }
            }



        }

        private EmpWorkingTime CalculateWokingTime(List<EmployeeModel> workingDayemp, DateTime dayOfMonth)
        {
            var startTime1 = new System.TimeSpan(7, 40, 0);
            var startTime2 = new System.TimeSpan(8, 15, 0);
            var endTime1 = new System.TimeSpan(16, 55, 0);
            var endTime2 = new System.TimeSpan(17, 15, 0);
            var EachEmpWt = new EmpWorkingTime();
            if (workingDayemp == null)
            {
                EachEmpWt.WorkingDay = dayOfMonth.ToString("dd/MM/yyyy");
                EachEmpWt.EmpNm = workingDayemp.FirstOrDefault().Name;
                EachEmpWt.TimeIn = "";
                EachEmpWt.TimeOut = "";
                EachEmpWt.OffOrWork = false;
                EachEmpWt.Early = 0;
                EachEmpWt.Late = 0;
                EachEmpWt.OverTime = 0;
                return EachEmpWt;
            }
            else if (workingDayemp.Count() == 1)
            {
                var TimeInDay = workingDayemp.FirstOrDefault()?.WorkDay?.TimeOfDay;
                EachEmpWt.TimeOut = TimeInDay >= startTime1 && TimeInDay<= startTime2 ? TimeInDay?.ToString("HH:mm"):"Quen bam vao";
                EachEmpWt.TimeOut = TimeInDay >=endTime1 && TimeInDay <= endTime2 ? TimeInDay?.ToString("HH:mm"):"Quen Bam Ra";
                EachEmpWt.OffOrWork =true;
                EachEmpWt.WorkingDay =dayOfMonth.ToString("dd/MM/yyyy");
                return EachEmpWt;
            }else{
                // EachEmpWt.TimeIn =workingDayemp.Where(workingDayemp.Max(x=>x.WorkDay.Value.Hour) )

            return EachEmpWt;
            }

        }
        private List<EmployeeModel> ConvertToDateTime(List<EmployeeModel> listEmp)
        {

            for (var i = 0; i < listEmp.Count; i++)
            {
                DateTime workdate = DateTime.ParseExact(listEmp[i].WorkDate, "d/M/yyyy h:mm tt", System.Globalization.CultureInfo.InvariantCulture);
                listEmp[i].WorkDay = workdate;
                listEmp[i].NumberDayOfMonth = DateTime.DaysInMonth(workdate.Year, workdate.Month);
                listEmp[i].DateFormat = workdate.ToString("MM/dd/yyyy").Replace("/", "");
            }
            return listEmp;
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
                        if (col == 1)
                        {
                            emp.Id = row;
                            emp.Name = ws.Cells[row, col].Value.ToString();
                        }
                        else
                        {
                            emp.WorkDate = ws.Cells[row, col].Value.ToString();
                        }
                    }
                    res.Add(emp);
                }
            }
            return res;
        }



    }
}