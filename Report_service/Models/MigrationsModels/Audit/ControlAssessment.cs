using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.MigrationsModels.Audit
{
    [Table("CONTROL_ASSESSMENT")]
    public class ControlAssessment
    {
        [Key]
        [Column("id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonPropertyName("id")]
        public int Id { get; set; }

        //đây là cột quy định đánh giá kiểm soát thuộc giấy tờ nào
        [Column("working_paper_id")]
        [JsonPropertyName("workingpaperid")]
        public int workingpaperid { get; set; }

        //đây Đây là cột quy định đánh giá thuộc rủi ro nào
        [Column("risk_id")]
        [JsonPropertyName("riskid")]
        public int riskid { get; set; }

        //đây Đây là cột quy định đánh giá thuộc kiểm soát nào
        [Column("control_id")]
        [JsonPropertyName("controlid")]
        public int controlid { get; set; }

    }
}
