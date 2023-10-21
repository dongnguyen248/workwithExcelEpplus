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
        public int? OverTime { get; set; }
        public int? Late { get; set; }
        public int? Early { get; set; }
        public bool OffOrWork { get; set; } = false;
    }
}