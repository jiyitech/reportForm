using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ServiceStack.DataAnnotations;

namespace reportform
{
    [Alias("ReportInfo"), Table("ReportInfo")]
    class ReportInfo
    {
        [Alias("id"), PrimaryKey, AutoIncrement, ReadOnly(true)]
        public long id { get; set; }
        public DateTime? date { get; set; }
        public string dutyName { get; set; }
        public double? autoTime { get; set; }
        public double? autoEnd { get; set; }
        public double? totalTime { get; set; }
        public double? totalEnd { get; set; }
        public double? percent { get; set; }
    }
}
