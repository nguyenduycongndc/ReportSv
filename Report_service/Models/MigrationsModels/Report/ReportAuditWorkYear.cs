using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.MigrationsModels
{
    [Table("REPORT_AUDIT_WORK_YEAR")]
    public class ReportAuditWorkYear
    {
        [Key]
        [Column("id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonPropertyName("id")]
        public int Id { get; set; }

        //year : năm kiểm toán
        [Column("year")]
        [JsonPropertyName("year")]
        public int? year { get; set; }

        //name : Tên báo cáo
        [Column("name")]
        [JsonPropertyName("name")]
        public string name { get; set; }

        //evaluation : Đánh giá tổng quan của KTNB
        [Column("evaluation")]
        [JsonPropertyName("evaluation")]
        public string evaluation { get; set; }

        //concers : Quan ngại chính của KTNB
        [Column("concerns")]
        [JsonPropertyName("concerns")]
        public string concerns { get; set; }

        //reason : Lý do không hoàn thành kế hoạch
        [Column("reason")]
        [JsonPropertyName("reason")]
        public string reason { get; set; }

        //note : Các vấn đề cần lưu ý
        [Column("note")]
        [JsonPropertyName("note")]
        public string note { get; set; }

        //quality : Chất lượng của hoạt động KTNB
        [Column("quality")]
        [JsonPropertyName("quality")]
        public string quality { get; set; }

        [Column("is_deleted")]
        [JsonPropertyName("is_deleted")]
        public bool? IsDeleted { get; set; }

        [Column("created_at")]
        [JsonPropertyName("created_at")]
        public DateTime? CreatedAt { get; set; }

        [Column("created_by")]
        [JsonPropertyName("created_by")]
        public int? CreatedBy { get; set; }

        [Column("modified_at")]
        [JsonPropertyName("modified_at")]
        public DateTime? ModifiedAt { get; set; }

        [Column("modified_by")]
        [JsonPropertyName("modified_by")]
        public int? ModifiedBy { get; set; }

        [Column("deleted_at")]
        [JsonPropertyName("deleted_at")]
        public DateTime? DeletedAt { get; set; }

        [Column("deleted_by")]
        [JsonPropertyName("deleted_by")]
        public int? DeletedBy { get; set; }

        [Column("overcome")]
        [JsonPropertyName("overcome")]
        public string overcome { get; set; }

    }

    /// <summary>
    /// Tình hình kế hoạch kiểm toán nội bô
    /// </summary>
    public class ReportTable1
    {
        public int type { get; set; }
        public double? audit_plan_current { get; set; }
        public double? audit_completed_current { get; set; }
        public double? completed_current { get; set; }
        public double? audit_expected_current { get; set; }

        public double? audit_plan_previous { get; set; }
        public double? audit_completed_previous { get; set; }
        public double? completed_previous { get; set; }
        public double? audit_expected_previous { get; set; }

        public double? audit_plan_volatility { get; set; }
        public double? audit_completed_volatility { get; set; }
    }

    public class ReportTable2
    {
        public string audit_name { get; set; }
        public string audit_time { get; set; }
        public string report_date { get; set; }
        public string level { get; set; }
        public string risk_high { get; set; }
        public string risk_medium { get; set; }
    }

    public class ReportTable3
    {
        public string audit_name { get; set; }
        public string audit_summary { get; set; }
        public string audit_request_content { get; set; }
        public string audit_request_status { get; set; }
    }

    public class ReportAuditRequest
    {
        public string code { get; set; }
        public DateTime? extendat { get; set; }
        public int? rating { get; set; }
        public DateTime? completed { get; set; }
        public DateTime? actualcompleted { get; set; }
        public int? conclusion { get; set; }
        public int timestatus { get; set; }
        public int? processstatus { get; set; }
        public int? year { get; set; }
        public double day { get; set; }
    }
    public class ReportTable4
    {
        public int type { get; set; }
        public int beginning_high { get; set; }
        public int notclose_high { get; set; }
        public int close_high { get; set; }
        public int ending_high { get; set; }

        public int beginning_medium { get; set; }
        public int notclose_medium { get; set; }
        public int close_medium { get; set; }
        public int ending_medium { get; set; }
        public int total { get; set; }
    }
}
