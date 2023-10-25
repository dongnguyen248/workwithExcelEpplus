using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace excel.Model
{
    public class EmpWorkingTime
    {
        public string EmpNm { get; set; }
        public string? WorkingDay { get; set; }
        public string? TimeIn { get; set; }
        public string? TimeOut { get; set; }
        public double? OverTime150 { get; set; }
        public double? OverTime185 { get; set; }
        public double? Late { get; set; }
        public double? Early { get; set; }
        public bool OffOrWork { get; set; } = false;
    }
}