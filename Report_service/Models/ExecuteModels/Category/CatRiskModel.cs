using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels
{
    public class CatRiskModel
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
        [JsonPropertyName("isdeleted")]
        public bool? IsDeleted { get; set; }

        [JsonPropertyName("relatestep")]
        public string RelateStep { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }
    public class CatRiskSearchModel
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("code")]
        public string Code { get; set; }
        [JsonPropertyName("name")]
        public string Name { get; set; }
        [JsonPropertyName("unitid")]
        public string Unitid { get; set; }
        [JsonPropertyName("activationid")]
        public string Activationid { get; set; }
        [JsonPropertyName("processid")]
        public string Processid { get; set; }
        [JsonPropertyName("status")]
        public string Status { get; set; }

        [JsonPropertyName("start_number")]
        public int StartNumber { get; set; }

        [JsonPropertyName("page_size")]
        public int PageSize { get; set; }

        [JsonPropertyName("isdeleted")]
        public bool? IsDeleted { get; set; }

        [JsonPropertyName("relatestep")]
        public string RelateStep { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }

    public class CatRiskModifiModel
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

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
        public int Status { get; set; }        

        [JsonPropertyName("isdeleted")]
        public bool? IsDeleted { get; set; }

        [JsonPropertyName("relatestep")]
        public string RelateStep { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }
    public class CatRiskCreateModel
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

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
        public int Status { get; set; }        

        [JsonPropertyName("isdeleted")]
        public bool? IsDeleted { get; set; }

        [JsonPropertyName("relatestep")]
        public string RelateStep { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }

    public class CatRiskDetailModel
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

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

        [JsonPropertyName("isdeleted")]
        public bool? IsDeleted { get; set; }

        [JsonPropertyName("relatestep")]
        public string RelateStep { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("unitname")]
        public string Unitname { get; set; }
        //[JsonPropertyName("activationid")]
        //public string Activationname { get; set; }
        //[JsonPropertyName("processid")]
        //public string Processname { get; set; }

       
    }
}