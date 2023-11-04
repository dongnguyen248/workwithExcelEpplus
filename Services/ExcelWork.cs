using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using excel.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace excel.Services
{
    public class ExcelWork
    {
        public void PivotData()
        {
            var listEmp = ReadExcelFile();
            var listEmpDateTime = ConvertToDateTime(listEmp);
            var listEachEmpWt = new List<EmpWorkingTime>();
            var listGroupBy = listEmpDateTime.GroupBy(x => x.Name).Select(x => new EmployeeModel
            {
                Name = x.Key,
            }).ToList();
            for (var i = 0; i < listGroupBy.Count(); i++)
            {

                var lEmpInMonth = listEmpDateTime.Where(x => x.Name == listGroupBy[i].Name).ToList();
                for (var day = 1; day <= listEmpDateTime[i].NumberDayOfMonth; day++)
                {
                    DateTime dayOfMonth = new DateTime(listEmpDateTime[i].WorkDay.Value.Year, listEmpDateTime[i].WorkDay.Value.Month, day);

                    var workingDayemp = lEmpInMonth.Where(x => x.DateFormat == dayOfMonth.ToString("MM/dd/yyyy").Replace("/", ""));

                    var EachEmpWt = CalculateWokingTime(workingDayemp.ToList(), dayOfMonth, listGroupBy[i].Name);

                    listEachEmpWt.Add(EachEmpWt);
                }
                Console.WriteLine("Tinh cong xong cho  nhan vien: " + listGroupBy[i].Name);
            }
            WriteExcelFile(listEachEmpWt);


        }
        private void WriteExcelFile(List<EmpWorkingTime> listEmp)
        {
            ExcelPackage excel = new ExcelPackage();

            // name of the sheet 
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            // setting the properties 
            // of the work sheet  
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties 
            // of the first row 
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet 
            workSheet.Cells[1, 1].Value = "S.No";
            workSheet.Cells[1, 2].Value = "Tên Nhân Viên";
            workSheet.Cells[1, 3].Value = "Ngày";
            workSheet.Cells[1, 4].Value = "Giờ Vào";
            workSheet.Cells[1, 5].Value = "Giờ Ra";
            workSheet.Cells[1, 6].Value = "Tăng ca 150%";
            workSheet.Cells[1, 7].Value = "Tăng ca 185%";
            workSheet.Cells[1, 8].Value = "Nghỉ hay làm";
            workSheet.Cells[1, 9].Value = "Đi Trễ";
            workSheet.Cells[1, 10].Value = "Về Sớm";
            workSheet.Cells[1, 11].Value = "Số giờ làm việc";



            // Inserting the article data into excel 
            // sheet by using the for each loop 
            // As we have values to the first row  
            // we will start with second row 
            int recordIndex = 2;

            foreach (var emp in listEmp)
            {
                workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheet.Cells[recordIndex, 2].Value = emp.EmpNm;
                workSheet.Cells[recordIndex, 3].Value = emp.WorkingDay;
                workSheet.Cells[recordIndex, 4].Value = emp.TimeIn;
                workSheet.Cells[recordIndex, 5].Value = emp.TimeOut;
                workSheet.Cells[recordIndex, 6].Value = emp.OverTime150;
                workSheet.Cells[recordIndex, 7].Value = emp.OverTime185;
                workSheet.Cells[recordIndex, 8].Value = emp.OffOrWork ? "Làm" : "Nghỉ";
                workSheet.Cells[recordIndex, 9].Value = emp.Late;
                workSheet.Cells[recordIndex, 10].Value = emp.Early;
                var wtHour = emp.OffOrWork ?  8 : 0;
                workSheet.Cells[recordIndex, 11].Value = wtHour > 0 ? wtHour : 0;

                recordIndex++;
            }

            // By default, the column width is not  
            // set to auto fit for the content 
            // of the range, so we are using 
            // AutoFit() method here.  
            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();

            // file name with .xlsx extension  
            string p_strPath = $"C:\\report\\bangchamCong{DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss")}.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk  
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package 
            excel.Dispose();
        }

        private EmpWorkingTime CalculateWokingTime(List<EmployeeModel> workingDayemp, DateTime dayOfMonth, string EmpNm)
        {
            var shift1In = new System.TimeSpan();
            var startTime1 = new System.TimeSpan(8, 00, 0);
            var endTime1 = new System.TimeSpan(16, 00, 0);
            var overTime1 = new System.TimeSpan(16, 30, 00);
            var overTime2 = new System.TimeSpan(22, 00, 00);

            var EachEmpWt = new EmpWorkingTime();
            if (workingDayemp.Count() == 0)
            {
                EachEmpWt.WorkingDay = dayOfMonth.ToString("dd/MM/yyyy");
                EachEmpWt.EmpNm = EmpNm;
                EachEmpWt.TimeIn = "";
                EachEmpWt.TimeOut = "";
                EachEmpWt.OffOrWork = false;
                EachEmpWt.Early = 0;
                EachEmpWt.Late = 0;
                EachEmpWt.OverTime150 = 0;
                EachEmpWt.OverTime185 = 0;
                return EachEmpWt;
            }
            else if (workingDayemp.Count() == 1)
            {
                var TimeInDay = workingDayemp.FirstOrDefault()?.WorkDay?.TimeOfDay;
                EachEmpWt.EmpNm = EmpNm;
                EachEmpWt.TimeIn = TimeInDay <= startTime1 ? TimeInDay?.ToString(@"hh\:mm") : "Quen bam vao";
                EachEmpWt.TimeOut = TimeInDay >= endTime1 ? TimeInDay?.ToString(@"hh\:mm") : "Quen Bam Ra";
                EachEmpWt.OffOrWork = true;
                EachEmpWt.WorkingDay = dayOfMonth.ToString("dd/MM/yyyy");
                EachEmpWt.OverTime150 = 0;
                EachEmpWt.OverTime185 = 0;
                return EachEmpWt;
            }
            else
            {
                var timeIn = workingDayemp.Where(x => x.WorkDay.Value.Ticks == workingDayemp.Min(x => x.WorkDay.Value.Ticks)).FirstOrDefault().WorkDay.Value.TimeOfDay;
                var timeOut = workingDayemp.Where(x => x.WorkDay.Value.Ticks == workingDayemp.Max(x => x.WorkDay.Value.Ticks)).FirstOrDefault().WorkDay.Value.TimeOfDay;
                EachEmpWt.EmpNm = EmpNm;
                EachEmpWt.OffOrWork = true;
                EachEmpWt.TimeIn = timeIn.ToString(@"hh\:mm");
                EachEmpWt.TimeOut = timeOut.ToString(@"hh\:mm");
                EachEmpWt.WorkingDay = dayOfMonth.ToString("dd/MM/yyyy");
                EachEmpWt.OverTime185 = timeOut > overTime2 ? Math.Round(timeOut.Subtract(overTime2).TotalHours * 2, MidpointRounding.AwayFromZero) / 2 : 0;
                EachEmpWt.OverTime150 = timeOut > overTime1 ? Math.Round(timeOut.Subtract(timeIn > endTime1 ? timeIn : overTime1).TotalHours * 2, MidpointRounding.AwayFromZero) / 2 - EachEmpWt.OverTime185 : 0;
                EachEmpWt.Early = timeOut < endTime1 ? Math.Round(endTime1.Subtract(timeOut).TotalHours * 2, MidpointRounding.AwayFromZero) / 2 : 0;
                EachEmpWt.Late = timeIn > startTime1 ? Math.Round(timeIn.Subtract(startTime1).TotalHours * 2, MidpointRounding.AwayFromZero) / 2 : 0;
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

            string FilePath = "C:\\Users\\HP\\Documents\\dongproject\\T8. Cong Ca ngay.xlsx";
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet ws = package.Workbook.Worksheets["Dien ÐK"];
                int colCount = ws.Dimension.End.Column;  //get Column Count
                int rowCount = ws.Dimension.End.Row;
                //get row count
                for (int row = 2; row <= rowCount; row++)
                {
                    var emp = new EmployeeModel();
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (col == 3)
                        {
                            emp.Id = row;
                            emp.Name = ws.Cells[row, col].Value.ToString();
                        }
                        else if(col == 4)
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