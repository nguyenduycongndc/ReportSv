namespace Report_service.Models.MigrationsModels
{
    using System.ComponentModel.DataAnnotations.Schema;
    using System;
    using System.ComponentModel.DataAnnotations;

    [Table("ASSESSMENT_RESULT")]
    public class AssessmentResult
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public bool Status { get; set; } = true;
        public bool Deleted { get; set; } = false;
        public string Description { get; set; }
        public DateTime CreateDate { get; set; } = DateTime.Now;
        public int? UserCreate { get; set; }
        public DateTime? LastModified { get; set; }
        public int? ModifiedBy { get; set; }
        public int DomainId { get; set; }
        public int ScoreBoardId { get; set; } = 0;
        public int StageStatus { get; set; } = 0;
        public int RiskLevel { get; set; } = 0;
        public int RiskLevelChange { get; set; } = 0;
        public string RiskLevelChangeName { get; set; }
        public bool Audit { get; set; } = false;
        public int? AuditReason { get; set; }
        public string PassAuditReason { get; set; }
        public DateTime? AuditDate { get; set; }
        public DateTime? LastAudit { get; set; }
        public string LastRiskLevel { get; set; }
        public int? AssessmentStatus { get; set; }
    }
}
