using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels
{
    public class ReportAuditWorkYearSearchModel
    {
        [JsonPropertyName("year")]
        public string Year { get; set; }
        [JsonPropertyName("status")]
        public string Status { get; set; }
        [JsonPropertyName("start_number")]
        public int StartNumber { get; set; }
        [JsonPropertyName("page_size")]
        public int PageSize { get; set; }
    }

    public class ReportAuditWorkYearModel
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        //year : năm kiểm toán
        [JsonPropertyName("year")]
        public int? year { get; set; }

        //name : Tên báo cáo
        [JsonPropertyName("name")]
        public string name { get; set; }

        //evaluation : Đánh giá tổng quan của KTNB
        [JsonPropertyName("evaluation")]
        public string evaluation { get; set; }

        //concers : Quan ngại chính của KTNB
        [JsonPropertyName("concerns")]
        public string concerns { get; set; }

        //reason : Lý do không hoàn thành kế hoạch
        [JsonPropertyName("reason")]
        public string reason { get; set; }

        //note : Các vấn đề cần lưu ý
        [JsonPropertyName("note")]
        public string note { get; set; }

        //quality : Chất lượng của hoạt động KTNB
        [JsonPropertyName("quality")]
        public string quality { get; set; }

        [JsonPropertyName("is_deleted")]
        public bool? IsDeleted { get; set; }

        [JsonPropertyName("created_at")]
        public DateTime? CreatedAt { get; set; }

        [JsonPropertyName("created_by")]
        public int? CreatedBy { get; set; }

        [JsonPropertyName("modified_at")]
        public DateTime? ModifiedAt { get; set; }

        [JsonPropertyName("modified_by")]
        public int? ModifiedBy { get; set; }

        [JsonPropertyName("deleted_at")]
        public DateTime? DeletedAt { get; set; }

        [JsonPropertyName("deleted_by")]
        public int? DeletedBy { get; set; }

        [JsonPropertyName("overcome")]
        public string OverCome { get; set; }

        [JsonPropertyName("status")]
        public string Status { get; set; }
        [JsonPropertyName("statusname")]
        public string StatusName { get; set; }

        [JsonPropertyName("reviewerid")]
        public int? Reviewerid { get; set; }

        [JsonPropertyName("approval_user")]
        public int? approval_user { get; set; } // người duyệt
        [JsonPropertyName("approval_user_last")]
        public int? approval_user_last { get; set; } // người duyệt
    }

    public class StatusApprove
    {
        [JsonPropertyName("id")]
        public int id { get; set; }

        [JsonPropertyName("status_code")]
        public string status_code { get; set; }

        [JsonPropertyName("status_name")]
        public string status_name { get; set; }
    }
}
