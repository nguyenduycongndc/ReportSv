using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.MigrationsModels
{
    [Table("REPORT_AUDIT_WORK_FILE")]
    public class ReportAuditWorkFile
    {
        [Key]
        [Column("id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonPropertyName("id")]
        public int id { get; set; }
        [JsonPropertyName("report_auditwork_id")]
        public int report_auditwork_id { get; set; }
        [ForeignKey("report_auditwork_id")]
        public virtual ReportAuditWork ReportAuditWork { get; set; }
        [Column("path")]
        [JsonPropertyName("path")]
        public string path { get; set; }
        [Column("file_type")]
        [JsonPropertyName("file_type")]
        public string file_type { get; set; }
        [Column("isdelete")]
        [JsonPropertyName("isdelete")]
        public bool? isdelete { get; set; }
    }
}
