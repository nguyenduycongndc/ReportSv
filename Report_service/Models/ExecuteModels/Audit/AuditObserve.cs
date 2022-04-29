using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels.Audit
{
    public class AuditObserveSearchModel
    {
        [JsonPropertyName("year")]
        public int? year { get; set; }
        [JsonPropertyName("auditwork_id")]
        public int? auditwork_id { get; set; }
        [JsonPropertyName("code")]
        public string code { get; set; }
        [JsonPropertyName("name")]
        public string name { get; set; }
        [JsonPropertyName("working_paper_code")]
        public string working_paper_code { get; set; }
        [JsonPropertyName("discoverer")]
        public int? discoverer { get; set; }
        [JsonPropertyName("start_number")]
        public int StartNumber { get; set; }
        [JsonPropertyName("page_size")]
        public int PageSize { get; set; }
    }
    public class AuditObserveModel
    {
        [JsonPropertyName("id")]
        public int id { get; set; }
        [JsonPropertyName("year")]
        public int? year { get; set; }
        [JsonPropertyName("code")]
        public string code { get; set; }
        [JsonPropertyName("name")]
        public string name { get; set; }
        [JsonPropertyName("auditwork_id")]
        public int? auditwork_id { get; set; }
        [JsonPropertyName("auditwork_name")]
        public string auditwork_name { get; set; }
        [JsonPropertyName("working_paper_code")]
        public string working_paper_code { get; set; }
        [JsonPropertyName("discoverer_name")]
        public string discoverer_name { get; set; }
    }
}
