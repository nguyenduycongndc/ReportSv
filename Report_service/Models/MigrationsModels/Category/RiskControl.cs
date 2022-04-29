using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.MigrationsModels
{
    [Table("RISK_CONTROL")]
    public class RiskControl
    {
        [Key]
        [Column("id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("riskid")]
        public int riskid { get; set; }

        [ForeignKey("riskid")]
        public virtual CatRisk CatRisk { get; set; }

        [JsonPropertyName("controlid")]
        public int controlid { get; set; }

        [ForeignKey("controlid")]
        public virtual CatControl CatControl { get; set; }
    }
}
