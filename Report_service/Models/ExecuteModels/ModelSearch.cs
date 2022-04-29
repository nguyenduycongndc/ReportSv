using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels
{
    public class ModelSearch
    {
        [JsonPropertyName("full_name")]
        public string FullName { get; set; }
        [JsonPropertyName("email")]
        public string Email { get; set; }
        [JsonPropertyName("user_name")]
        public string UserName { get; set; }
        [JsonPropertyName("status")]
        public string Status { get; set; }
        [JsonPropertyName("description")]
        public string Description { get; set; }
        [JsonPropertyName("department_id")]
        public string DepartmentId { get; set; }
        [JsonPropertyName("users_type")]
        public string UsersType { get; set; }
        [JsonPropertyName("start_number")]
        public int StartNumber { get; set; }
        [JsonPropertyName("page_size")]
        public int PageSize { get; set; }
    }
    public class ActiveClass
    {
        public int id { get; set; }
        public int status { get; set; }
    }
    public class ChangePass
    {
        public int id { get; set; }
        public string password { get; set; }
    }
    public class ChangeWorkplace
    {
        public int id { get; set; }
        public int departmentId { get; set; }
        public string dateofjoining { get; set; }
    }
    public class DeleteAll
    {
        public string listID { get; set; }
    }
    public class AuditProgramModelSearch
    {
        [JsonPropertyName("year")]
        public string Year { get; set; }
        [JsonPropertyName("audit_process")]
        public int? AuditProcess { get; set; }
        [JsonPropertyName("audit_work")]
        public int? AuditWork { get; set; }
        [JsonPropertyName("facility")]
        public int? Facility { get; set; }
        [JsonPropertyName("status")]
        public string Status { get; set; }
        [JsonPropertyName("activity")]
        public int? Activity { get; set; }
        [JsonPropertyName("start_number")]
        public int StartNumber { get; set; }
        [JsonPropertyName("page_size")]
        public int PageSize { get; set; }
    }
}
