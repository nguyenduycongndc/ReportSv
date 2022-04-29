﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels
{
    public class AuditProgramListModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }
        [JsonPropertyName("year")]
        public string Year { get; set; }
        [JsonPropertyName("audit_work")]
        public string AuditWork { get; set; }
        [JsonPropertyName("audit_work_id")]
        public int? AuditWorkID { get; set; }
        [JsonPropertyName("audit_process")]
        public string AuditProcess { get; set; }
        [JsonPropertyName("audit_process_id")]
        public int? AuditProcessID { get; set; }
        [JsonPropertyName("audit_activity")]
        public string AuditActivity { get; set; }
        [JsonPropertyName("audit_activity_id")]
        public int? AuditActivityID { get; set; }
        [JsonPropertyName("audit_facility")]
        public string AuditFacility { get; set; }
        [JsonPropertyName("audit_facility_id")]
        public int? AuditFacilityID { get; set; }
        [JsonPropertyName("status")]
        public int? Status { get; set; }
    }

    public class AuditProgramDetailModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }
        [JsonPropertyName("year")]
        public string Year { get; set; }
        [JsonPropertyName("audit_work")]
        public string AuditWork { get; set; }
        [JsonPropertyName("audit_process")]
        public string AuditProcess { get; set; }
        [JsonPropertyName("audit_activity")]
        public string AuditActivity { get; set; }
        [JsonPropertyName("audit_facility")]
        public string AuditFacility { get; set; }
        [JsonPropertyName("person_in_charge")]
        public string PersonInCharge { get; set; }

        [JsonPropertyName("audit_risk_list")]
        public List<AuditRiskModel> AuditRiskList { get; set; }

    }
    public class AuditRiskModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }
        [JsonPropertyName("cat_risk_code")]
        public string CatRiskCode { get; set; }

        [JsonPropertyName("cat_risk_name")]
        public string CatRiskName { get; set; }

        [JsonPropertyName("potential_possibility")]
        public int? Potential_Possibility { get; set; } // Khả năng xảy ra

        [JsonPropertyName("potential_infulence_level")]
        public int? Potential_InfulenceLevel { get; set; } // Mức độ ảnh hưởng

        [JsonPropertyName("potential_risk_rating")]
        public int? Potential_RiskRating { get; set; } // Xếp hạng rủi ro

        [JsonPropertyName("potential_risk_rating_name")]
        public string Potential_RiskRating_Name { get; set; } // Xếp hạng rủi ro name

        [JsonPropertyName("score_potential_risk")]
        public int? ScorePotentialRiskRating { get; set; } // Điểm RR tiềm tàng

        [JsonPropertyName("score_remaining_risk")]
        public int? ScoreRemainingRisk { get; set; } // Điểm RR còn lại

        [JsonPropertyName("audit_proposal")]
        public int? AuditProposal { get; set; } // Đề xuất kiểm toán
    }
    public class AuditRiskDetailModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }

        [JsonPropertyName("audiit_scope_id")]
        public int? AuditScopeId { get; set; }
        [JsonPropertyName("cat_risk_id")]
        public int? CatRiskID { get; set; }
        [JsonPropertyName("cat_risk_code")]
        public string CatRiskCode { get; set; }

        [JsonPropertyName("cat_risk_name")]
        public string CatRiskName { get; set; }
        [JsonPropertyName("description_cat_risk")]
        public string DescriptionCatRisk { get; set; }

        [JsonPropertyName("potential_possibility")]
        public int? Potential_Possibility { get; set; } // Khả năng xảy ra

        [JsonPropertyName("potential_infulence_level")]
        public int? Potential_InfulenceLevel { get; set; } // Mức độ ảnh hưởng

        [JsonPropertyName("potential_risk_rating")]
        public int? Potential_RiskRating { get; set; } // Xếp hạng rủi ro

        [JsonPropertyName("potential_reason_rating")]
        public string Potential_ReasonRating { get; set; } // Lý do

        [JsonPropertyName("audit_proposal")]
        public int? AuditProposal { get; set; } // Đề xuất kiểm toán

        [JsonPropertyName("remaining_possibility")]
        public int? Remaining_Possibility { get; set; } // Khả năng xảy ra

        [JsonPropertyName("remaining_infulence_level")]
        public int? Remaining_InfulenceLevel { get; set; } // Mức độ ảnh hưởng

        [JsonPropertyName("remaining_risk_rating")]
        public int? Remaining_RiskRating { get; set; } // Xếp hạng rủi ro

        [JsonPropertyName("remaining_reason_rating")]
        public string Remaining_ReasonRating { get; set; } // lý do
        [JsonPropertyName("list_program_procedures")]
        public List<ProgramProceduresDetailModel> listProgramProcedures { get; set; }

        [JsonPropertyName("cat_control_list")]
        public List<CatControlDetailModel> CatControlList { get; set; }
        

    }
    public class ProgramProceduresDetailModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }
        [JsonPropertyName("code")]
        public string Code { get; set; }
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("control_id")]
        public int? ControlID { get; set; }
        [JsonPropertyName("control_code")]
        public string ControlCode { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
        [JsonPropertyName("lst_auditor")]
        public string lstAuditor { get; set; }
    }
    public class AuditRiskModifyModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }

        [JsonPropertyName("audiit_scope_id")]
        public int? AuditScopeId { get; set; }
        [JsonPropertyName("cat_risk_id")]
        public int? CatRiskID { get; set; }

        [JsonPropertyName("potential_possibility")]
        public int? Potential_Possibility { get; set; } // Khả năng xảy ra

        [JsonPropertyName("potential_infulence_level")]
        public int? Potential_InfulenceLevel { get; set; } // Mức độ ảnh hưởng

        [JsonPropertyName("potential_risk_rating")]
        public int? Potential_RiskRating { get; set; } // Xếp hạng rủi ro

        [JsonPropertyName("potential_risk_rating_name")]
        public string PotentialRiskRatingName { get; set; } // Xếp hạng rủi ro name

        [JsonPropertyName("potential_reason_rating")]
        public string Potential_ReasonRating { get; set; } // Lý do

        [JsonPropertyName("audit_proposal")]
        public int? AuditProposal { get; set; } // Đề xuất kiểm toán

        [JsonPropertyName("remaining_possibility")]
        public int? Remaining_Possibility { get; set; } // Khả năng xảy ra

        [JsonPropertyName("remaining_infulence_level")]
        public int? Remaining_InfulenceLevel { get; set; } // Mức độ ảnh hưởng

        [JsonPropertyName("remaining_risk_rating")]
        public int? Remaining_RiskRating { get; set; } // Xếp hạng rủi ro

        [JsonPropertyName("remaining_risk_rating_name")]
        public string RemainingRiskRatingName { get; set; } // Xếp hạng rủi ro name

        [JsonPropertyName("remaining_reason_rating")]
        public string Remaining_ReasonRating { get; set; } // lý do
    }
    public class ProgramCatRiskCreateModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }

        [JsonPropertyName("code")]
        public string Code { get; set; }
        [JsonPropertyName("name")]
        public string Name { get; set; }
        [JsonPropertyName("unitid")]
        public int? Unitid { get; set; }
        [JsonPropertyName("activationid")]
        public int? Activationid { get; set; }
        [JsonPropertyName("processid")]
        public int? Processid { get; set; }
        [JsonPropertyName("status")]
        public int? Status { get; set; }

        [JsonPropertyName("relatestep")]
        public string RelateStep { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("audit_work_scope_id")]
        public int? AuditWorkScopeId { get; set; }
    }
    public class ProgramProceduresModifyModel
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }
        [JsonPropertyName("code")]
        public string Code { get; set; }
        [JsonPropertyName("name")]
        public string Name { get; set; }
        [JsonPropertyName("unitid")]
        public int? Unitid { get; set; }
        [JsonPropertyName("activationid")]
        public int? Activationid { get; set; }
        [JsonPropertyName("processid")]
        public int? Processid { get; set; }
        [JsonPropertyName("status")]
        public int? Status { get; set; }

        [JsonPropertyName("control_id")]
        public int? ControlID { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("risk_scoring_id")]
        public int? RiskScoringId { get; set; }
    }
}
