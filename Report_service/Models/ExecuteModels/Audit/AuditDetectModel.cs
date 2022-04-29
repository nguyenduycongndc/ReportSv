using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels.Audit
{
    public class AuditDetectSearchModel
    {
        [JsonPropertyName("year")]
        public int? year { get; set; }
        [JsonPropertyName("auditwork_id")]
        public int? auditwork_id { get; set; }
        [JsonPropertyName("auditprocess_id")]
        public int? auditprocess_id { get; set; }
        [JsonPropertyName("auditfacilities_id")]
        public int? auditfacilities_id { get; set; }
        [JsonPropertyName("code")]
        public string code { get; set; }
        [JsonPropertyName("title")]
        public string title { get; set; }
        [JsonPropertyName("working_paper_code")]
        public string working_paper_code { get; set; }
        [JsonPropertyName("audit_report")]
        public int audit_report { get; set; }
        [JsonPropertyName("start_number")]
        public int StartNumber { get; set; }
        [JsonPropertyName("page_size")]
        public int PageSize { get; set; }
    }
    public class AuditDetectModel
    {
        [JsonPropertyName("id")]
        public int id { get; set; }
        [JsonPropertyName("year")]
        public int? year { get; set; }
        [JsonPropertyName("auditwork_id")]
        public int? auditwork_id { get; set; }
        [JsonPropertyName("auditwork_name")]
        public string auditwork_name { get; set; }
        [JsonPropertyName("auditprocess_id")]
        public int? auditprocess_id { get; set; }
        [JsonPropertyName("auditprocess_name")]
        public string auditprocess_name { get; set; }
        [JsonPropertyName("auditfacilities_id")]
        public int? auditfacilities_id { get; set; }
        [JsonPropertyName("auditfacilities_name")]
        public string auditfacilities_name { get; set; }
        [JsonPropertyName("code")]
        public string code { get; set; }
        [JsonPropertyName("title")]
        public string title { get; set; }
        [JsonPropertyName("working_paper_code")]
        public string working_paper_code { get; set; }
        [JsonPropertyName("audit_report")]
        public int? audit_report { get; set; }
        //Xếp hạng rủi ro: 1-Cao, 2-Trung bình, 3-Thấp
        [JsonPropertyName("rating_risk")]
        public int? rating_risk { get; set; }
    }
    public class AuditDetectDetail
    {
        [JsonPropertyName("id")]
        public int id { get; set; }

        [JsonPropertyName("code")]
        public string code { get; set; }

        [JsonPropertyName("name")]
        public string name { get; set; }

        [JsonPropertyName("status")]
        public int? status { get; set; }

        [JsonPropertyName("title")]
        public string title { get; set; }//Tiêu đề phát hiện kiểm toán

        [JsonPropertyName("short_title")]
        public string short_title { get; set; }//Tiêu đề rút gọn phát hiện kiểm toán

        [JsonPropertyName("description")]
        public string description { get; set; }//mô tả

        [JsonPropertyName("evidence")]
        public string evidence { get; set; }//Bằng chứng phất hiện KT

        [JsonPropertyName("path_audit_detect")]
        public string path_audit_detect { get; set; }//File

        [JsonPropertyName("affect")]
        public string affect { get; set; }//ảnh hưởng

        [JsonPropertyName("rating_risk")]
        public int? rating_risk { get; set; }//xếp hạng rủi ro

        [JsonPropertyName("cause")]
        public string cause { get; set; }//Nguyên nhân

        [JsonPropertyName("audit_report")]
        public bool audit_report { get; set; }//Đưa vào báo cáo kiểm toán

        [JsonPropertyName("classify_audit_detect")]
        public int? classify_audit_detect { get; set; }//Phân loại phát hiện
        [JsonPropertyName("str_classify_audit_detect")]
        public string str_classify_audit_detect { get; set; }//Phân loại phát hiện

        [JsonPropertyName("summary_audit_detect")]
        public string summary_audit_detect { get; set; }//Tóm tắt phát hiện

        [JsonPropertyName("followers")]
        public int? followers { get; set; }//Người theo dõi

        [JsonPropertyName("str_followers")]
        public string str_followers { get; set; } //Người theo dõi id+ name

        [JsonPropertyName("year")]
        public int? year { get; set; }
        [JsonPropertyName("str_year")]
        public string str_year { get; set; }//Năm id+ name

        [JsonPropertyName("opinion_audit")]
        public bool opinion_audit { get; set; }//Ý kiến của ĐVĐKT

        [JsonPropertyName("reason")]
        public string reason { get; set; }//lý do

        [JsonPropertyName("auditwork_id")]
        public int? auditwork_id { get; set; }//id của cuộc kiểm toán ở trạng thái "Đã duyệt" thuộc năm đã chọn

        [JsonPropertyName("auditwork_name")]
        public string auditwork_name { get; set; }
        [JsonPropertyName("str_auditwork_name")]
        public string str_auditwork_name { get; set; }

        [JsonPropertyName("auditprocess_id")]
        public int? auditprocess_id { get; set; }//id của quy trình

        [JsonPropertyName("auditprocess_name")]
        public string auditprocess_name { get; set; }
        [JsonPropertyName("str_auditprocess_name")]
        public string str_auditprocess_name { get; set; }

        [JsonPropertyName("auditfacilities_id")]
        public int? auditfacilities_id { get; set; }//id của đơn vị 

        [JsonPropertyName("auditfacilities_name")]
        public string auditfacilities_name { get; set; }
        [JsonPropertyName("str_auditfacilities_name")]
        public string str_auditfacilities_name { get; set; }
        [JsonPropertyName("working_paper_code")]
        public string working_paper_code { get; set; }

        [JsonPropertyName("is_active")]
        public bool? IsActive { get; set; }

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

        [JsonPropertyName("listauditobserve")]
        public List<ListAuditObserve> ListAuditObserve { get; set; }

        [JsonPropertyName("listauditrequestmonitor")]
        public List<ListAuditRequestMonitor> ListAuditRequestMonitor { get; set; }
    }
    public class ListAuditObserve
    {
        [JsonPropertyName("id")]
        public int? id { get; set; }
        [JsonPropertyName("code")]
        public string code { get; set; }
        [JsonPropertyName("name")]
        public string name { get; set; }
        [JsonPropertyName("description")]
        public string description { get; set; }
        [JsonPropertyName("working_paper_code")]
        public string working_paper_code { get; set; }
    }
    public class ListAuditRequestMonitor
    {
        [JsonPropertyName("id")]
        public int? id { get; set; }
        [JsonPropertyName("code")]
        public string code { get; set; }
        [JsonPropertyName("content")]
        public string content { get; set; }
        [JsonPropertyName("auditrequesttypeid")]
        public int? auditrequesttypeid { get; set; }
        [JsonPropertyName("auditrequesttype_name")]
        public string auditrequesttype_name { get; set; }//phân loại kiến nghị
        [JsonPropertyName("user_id")]
        public int? user_id { get; set; }
        [JsonPropertyName("user_name")]
        public string user_name { get; set; }//người chịu trách nghiệm
        [JsonPropertyName("unit_id")]
        public int? unit_id { get; set; }
        [JsonPropertyName("unit_name")]
        public string unit_name { get; set; }//Đơn vị đầu mối
        [JsonPropertyName("cooperateunit_id")]
        public int? cooperateunit_id { get; set; }
        [JsonPropertyName("cooperateunit_name")]
        public string cooperateunit_name { get; set; }//Đơn vị phối hợp
        [JsonPropertyName("completeat")]
        public string completeat { get; set; }//Thời hạn hoàn thành
    }

    public class UncheckedModel
    {
        [JsonPropertyName("id")]
        public int? id { get; set; }
        [JsonPropertyName("audit_detect_id")]
        public int? audit_detect_id { get; set; }
    }
    public class DropListCatDetectTypeModel
    {
        [JsonPropertyName("id")]
        public int id { get; set; }
        [JsonPropertyName("name")]
        public string name { get; set; }
    }
}
