using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace excel.Model
{
    public class EmployeeModel
    {
        [Key]
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? WorkDate {get;set;}
        public DateTime? WorkDay { get; set; }
        public string? DateFormat { get; set; }
        public string? TimeIn { get; set; }
        public string? TimeOut { get; set; }
        public int? OT { get; set; }
        public int? NumberDayOfMonth { get; set; }
    }
}