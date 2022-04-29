using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.MigrationsModels
{
    [Table("AUDIT_FACILITY")]
    public class AuditFacility
    {
        [Key]
        [Column("ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [Column("Code")]
        [JsonPropertyName("code")]
        public string Code { get; set; }

        [Column("Name")]
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [Column("Status")]
        [JsonPropertyName("is_active")]
        public bool? IsActive { get; set; }

        [Column("Description")]
        [JsonPropertyName("description")]
        public string Description { get; set; }

        [Column("CreateDate")]
        [JsonPropertyName("created_at")]
        public DateTime? CreatedAt { get; set; }

        [Column("UserCreate")]
        [JsonPropertyName("created_by")]
        public int? CreatedBy { get; set; }

        [Column("LastModified")]
        [JsonPropertyName("modified_at")]
        public DateTime? ModifiedAt { get; set; }

        [Column("ModifiedBy")]
        [JsonPropertyName("modified_by")]
        public int? ModifiedBy { get; set; }

        [Column("ObjectClassId")]
        [JsonPropertyName("object_class_id")]
        public int? ObjectClassId { get; set; }

        [Column("DomainId")]
        [JsonPropertyName("domain_id")]
        public int? DomainId { get; set; }

        [Column("ParentId")]
        [JsonPropertyName("parent_id")]
        public int? ParentId { get; set; }

        [Column("ParentName")]
        [JsonPropertyName("parent_name")]
        public string ParentName { get; set; }

        [Column("ObjectClassName")]
        [JsonPropertyName("object_class_name")]
        public string ObjectClassName { get; set; }
    }
}
