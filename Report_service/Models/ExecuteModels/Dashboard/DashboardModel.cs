using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels.Dashboard
{
    /// <summary>
    /// Tình hình thực hiện kế hoạch kiểm toán năm hiện tại
    /// </summary>
    public class InternalAudit1Model
    {
        /// <summary>
        /// Cuộc kiểm toán đã hoàn thành
        /// </summary>
        [JsonPropertyName("total_audit_completed")]
        public double Total_Audit_Completed { get; set; }
        /// <summary>
        /// Cuộc kiểm toán chưa hoàn thành
        /// </summary>
        [JsonPropertyName("total_audit_pending")]
        public double Total_Audit_Pending { get; set; }
        /// <summary>
        /// Cuộc kiểm toán theo kế hoạch
        /// </summary>
        [JsonPropertyName("total_audit_plan")]
        public double Total_Audit_Plan { get; set; }
        /// <summary>
        /// Cuộc kiểm toán đột xuất
        /// </summary>
        [JsonPropertyName("total_audit_expect")]
        public double Total_Audit_Expect { get; set; }
    }

    /// <summary>
    /// Tình hình thực hiện của kiến nghị kiểm toán
    /// </summary>
    public class InternalAudit2Model
    {
        /// <summary>
        /// Trạng thái thời gian
        /// </summary>
        [JsonPropertyName("type")]
        public int Type { get; set; }
        /// <summary>
        /// Kiến nghị đã đóng
        /// </summary>
        [JsonPropertyName("close")]
        public int Close { get; set; }
        /// <summary>
        /// kiến nghị chưa đóng
        /// </summary>
        [JsonPropertyName("open")]
        public int Open { get; set; }
    }

    /// <summary>
    /// Số lượng phát hiện của kiểm toán trong năm theo mức rủi ro
    /// </summary>
    public class InternalAudit3Model
    {
        /// <summary>
        /// Cấp độ rủi ro
        /// </summary>
        [JsonPropertyName("ratingrisk")]
        public int Ratingrisk { get; set; }
        /// <summary>
        /// số lượng rủi ro
        /// </summary>
        [JsonPropertyName("total")]
        public int Total { get; set; }
    }


    /// <summary>
    /// Lịch sử đánh giá rủi ro
    /// </summary>

    public class AudittedUnit1Model
    {
        /// <summary>
        /// Năm
        /// </summary>

        [JsonPropertyName("year")]
        public int Year { get; set; }
        /// <summary>
        /// cấp độ rủi ro từng năm
        /// </summary>
        [JsonPropertyName("risk")]
        public string Risk { get; set; }
    }

    /// <summary>
    /// Thống kê số lượng phát hiện kiểm toán
    /// </summary>
    public class AudittedUnit2Model
    {
        /// <summary>
        /// quy trình
        /// </summary>

        [JsonPropertyName("processname")]
        public string Processname { get; set; }
        /// <summary>
        /// phát hiện rủi ro cao
        /// </summary>
        [JsonPropertyName("risk_high")]
        public int Risk_High { get; set; }
        /// <summary>
        /// phát hiện rủi ro trung bình
        /// </summary>
        [JsonPropertyName("risk_medium")]
        public int Risk_Medium { get; set; }
        /// <summary>
        /// phát hiện rủi ro thấp
        /// </summary>
        [JsonPropertyName("risk_low")]
        public int Risk_Low { get; set; }
    }


    /// <summary>
    /// Thống kê các kiến nghị kiểm toán
    /// </summary>
    public class AudittedUnit3Model
    {

        [JsonPropertyName("id")]
        public int? Id { get; set; }

        [JsonPropertyName("auditdetectcode")]
        public string AuditDetectCode { get; set; }

        [JsonPropertyName("auditrequestcode")]
        public string AuditRequestCode { get; set; }

        [JsonPropertyName("processstatus")]
        public int? ProcessStatus { get; set; }

        [JsonPropertyName("timestatus")]
        public int? TimeStatus { get; set; }

        [JsonPropertyName("completeat")]
        public string CompleteAt { get; set; }
        [JsonPropertyName("username")]
        public string Username { get; set; }
    }

    public class SearchModel
    {
        [JsonPropertyName("facility")]
        public int? Facility { get; set; }

        [JsonPropertyName("activity")]
        public int? Activity { get; set; }

        [JsonPropertyName("fromdate")]
        public string FromDate { get; set; }

        [JsonPropertyName("todate")]
        public string ToDate { get; set; }

        [JsonPropertyName("auditid")]
        public int AuditId { get; set; }

        [JsonPropertyName("year")]
        public string Year { get; set; }

        [JsonPropertyName("start_number")]
        public int StartNumber { get; set; }

        [JsonPropertyName("page_size")]
        public int PageSize { get; set; }
    }

    /// <summary>
    /// Các cuộc kiểm toán được phân công trong năm hiện tại
    /// </summary>
    public class AuditorWork1Model
    {
        /// <summary>
        /// Tên cuộc kiểm toán
        /// </summary>
        [JsonPropertyName("audit_name")]
        public string AuditName { get; set; }
        /// <summary>
        /// Thời gian bắt đầu dự kiến
        /// </summary>
        [JsonPropertyName("audit_start_date")]
        public string AuditStartDate { get; set; }
        /// <summary>
        /// Thời gian kết thúc dự kiến
        /// </summary>
        [JsonPropertyName("audit_end_date")]
        public string AuditEndDate { get; set; }
        /// <summary>
        /// Trạng thái cuộc kiểm toán
        /// </summary>
        [JsonPropertyName("status")]
        public string Status { get; set; }
    }
    /// <summary>
    /// Các giấy tờ làm việc đang xử lý
    /// </summary>
    public class AuditorWork2Model
    {
        /// <summary>
        /// Tên cuộc kiểm toán
        /// </summary>
        [JsonPropertyName("audit_name")]
        public string AuditName { get; set; }
        /// <summary>
        /// Ngày kết thúc thực địa dự kiến
        /// </summary>
        [JsonPropertyName("audit_end_field")]
        public string AuditEndField { get; set; }
        /// <summary>
        /// Mã giấy tờ làm việc
        /// </summary>
        [JsonPropertyName("paper_code")]
        public string PaperCode { get; set; }
        [JsonPropertyName("paper_id")]
        public int PaperId { get; set; }
        /// <summary>
        /// Trạng thái giấy tờ làm việc
        /// </summary>
        [JsonPropertyName("status")]
        public string Status { get; set; }
    }

    /// <summary>
    /// Các phát hiện của kiểm toán
    /// </summary>
    public class AuditorWork3Model
    {
        /// <summary>
        /// Tên cuộc kiểm toán
        /// </summary>
        [JsonPropertyName("audit_name")]
        public string AuditName { get; set; }
        /// <summary>
        /// Mã phát hiện
        /// </summary>
        [JsonPropertyName("audit_detect_code")]
        public string AuditDetectCode { get; set; }
        [JsonPropertyName("audit_detect_id")]
        public int AuditDetectId { get; set; }
        /// <summary>
        /// Trạng thái phát hiện
        /// </summary>
        [JsonPropertyName("status")]
        public string Status { get; set; }
        [JsonPropertyName("statuscode")]
        public string StatusCode { get; set; }
    }

    /// <summary>
    /// Các kiến nghị được phân công theo dõi
    /// </summary>
    public class AuditorWork4Model
    {
        /// <summary>
        /// Mã kiến nghị
        /// </summary>
        [JsonPropertyName("requestcode")]
        public string RequestCode { get; set; }
        [JsonPropertyName("requestid")]
        public int RequestId { get; set; }
        /// <summary>
        /// Trạng thái thực hiện
        /// </summary>
        [JsonPropertyName("processstatus")]
        public int? ProcessStatus { get; set; }
        /// <summary>
        /// Trạng thái thời gian
        /// </summary>
        [JsonPropertyName("timestatus")]
        public int? TimeStatus { get; set; }
        /// <summary>
        /// Người chịu trách nhiệm
        /// </summary>
        [JsonPropertyName("useredit")]
        public string UserEdit { get; set; }
        /// <summary>
        /// Đơn vị đầu mối
        /// </summary>
        [JsonPropertyName("unitedit")]
        public string UnitEdit { get; set; }
        /// <summary>
        /// thời hạn hoàn thành
        /// </summary>
        [JsonPropertyName("completeat")]
        public string CompleteAt { get; set; }
    }

    public class RiskHeatMap
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }
        [JsonPropertyName("rating")]
        public string Rating { get; set; }
    }

    public class AuditWorkItem
    {
        [JsonPropertyName("person")]
        public string Person { get; set; }
        [JsonPropertyName("numofauditor")]
        public int NumOfAuditor { get; set; }
        [JsonPropertyName("releasedate")]
        public string ReleaseDate { get; set; }
        [JsonPropertyName("status")]
        public string Status { get; set; }
    }

    public class UnitDetect
    {
        /// <summary>
        /// Tên đơn vị của cuộc kiểm toán   
        /// </summary>
        [JsonPropertyName("facilityname")]
        public string FacilityName { get; set; }

        /// <summary>
        /// Phát hiện rủi ro cao
        /// </summary>
        [JsonPropertyName("risk_high")]
        public int Risk_High { get; set; }
        /// <summary>
        /// phát hiện rủi ro trung bình
        /// </summary>
        [JsonPropertyName("risk_medium")]
        public int Risk_Medium { get; set; }
        /// <summary>
        /// phát hiện rủi ro thấp
        /// </summary>
        [JsonPropertyName("risk_low")]
        public int Risk_Low { get; set; }
    }

    public class RiskSummary
    {
        [JsonPropertyName("total_risk")]
        public int Total_Risk { get; set; }
        [JsonPropertyName("total_risk_high")]
        public int Risk_High { get; set; }
        [JsonPropertyName("total_risk_medium")]
        public int Risk_Medium { get; set; }
        [JsonPropertyName("total_risk_low")]
        public int Risk_Low { get; set; }
    }

    /// <summary>
    /// Công việc của các kiểm toán viên
    /// </summary>
    public class AuditorWork
    {
        /// <summary>
        /// Tên kiểm toán viên
        /// </summary>
        [JsonPropertyName("username")]
        public string UserName { get; set; }

        /// <summary>
        /// Số lượng chương trình KT được phân công
        /// </summary>
        [JsonPropertyName("totalwork")]
        public int TotalWork { get; set; }

        /// <summary>
        /// Số giấy tờ làm việc
        /// </summary>
        [JsonPropertyName("totalpaper")]
        public int TotalPaper { get; set; }
    }

    public class AuditSchedule
    {
        [JsonPropertyName("work")]
        public string Work { get; set; }
        [JsonPropertyName("id")]
        public int Id { get; set; }
        [JsonPropertyName("expected_date_schedule")]
        public string Expected_Date_Schedule { get; set; }
        [JsonPropertyName("actual_date_schedule")]
        public string Actual_Date_Schedule { get; set; }
        [JsonPropertyName("deviating_plan")]
        public string DeviatingPlan { get; set; }
        [JsonIgnore]
        public DateTime? Expected_Date { get; set; }
        [JsonIgnore]
        public DateTime? Actual_Date { get; set; }
    }
}
