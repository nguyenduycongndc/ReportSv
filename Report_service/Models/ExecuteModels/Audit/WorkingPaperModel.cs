using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Report_service.Models.ExecuteModels.Audit
{
    public class WorkingPaperModel
    {
        [JsonPropertyName("id")]
        public int id { get; set; }

        //code : Mã giấy tờ
        [JsonPropertyName("code")]
        public string code { get; set; }

        //year : Năm
        [JsonPropertyName("year")]
        public string year { get; set; }

        //auditworkid : id cuộc kiểm toán
        [JsonPropertyName("auditworkid")]
        public string auditworkid { get; set; }
        

        //processid : id quy trình được kiểm toán
        [JsonPropertyName("processid")]
        public string processid { get; set; }
        

        //unitid : id đơn vị
        [JsonPropertyName("unitid")]
        public string unitid { get; set; }        
        

        //status : trạng thái
        [JsonPropertyName("status")]
        public string status { get; set; }

        //asigneeid : id người thực hiện
        [JsonPropertyName("asigneeid")]
        public string asigneeid { get; set; }

        //reviewerid : id người rà soát
        [JsonPropertyName("reviewerid")]
        public string reviewerid { get; set; }

        

        //riskid : id của rủi ro
        [JsonPropertyName("riskid")]
        public string riskid { get; set; }

        //prototype : mẫu
        [JsonPropertyName("prototype")]
        public string prototype { get; set; }

        //conclusion : Kết luận
        [JsonPropertyName("conclusion")]
        public string conclusion { get; set; }

        [JsonPropertyName("target")]
        public string target { get; set; }
    }

    public class WorkingPaperSearchModel
    {     
        //code : Mã giấy tờ
        [JsonPropertyName("code")]
        public string code { get; set; }

        //year : Năm
        [JsonPropertyName("year")]
        public string year { get; set; }

        //auditworkid : id cuộc kiểm toán
        [JsonPropertyName("auditworkid")]
        public string auditworkid { get; set; }

        //processid : id quy trình được kiểm toán
        [JsonPropertyName("processid")]
        public string processid { get; set; }

        //unitid : id đơn vị
        [JsonPropertyName("unitid")]
        public string unitid { get; set; }

        //status : trạng thái
        [JsonPropertyName("status")]
        public string status { get; set; }   
    }

    public class WorkingPaperModifyModel
    {
        [JsonPropertyName("id")]
        public int id { get; set; }

        //code : Mã giấy tờ
        [JsonPropertyName("code")]
        public string code { get; set; }

        //year : Năm
        [JsonPropertyName("year")]
        public int? year { get; set; }

        //auditworkid : id cuộc kiểm toán
        [JsonPropertyName("auditworkid")]
        public int? auditworkid { get; set; }
        [JsonPropertyName("auditworkname")]
        public string auditworkname { get; set; }

        //processid : id quy trình được kiểm toán
        [JsonPropertyName("processid")]
        public int? processid { get; set; }
        [JsonPropertyName("processname")]
        public string processname { get; set; }

        //unitid : id đơn vị
        [JsonPropertyName("unitid")]
        public int? unitid { get; set; }
        [JsonPropertyName("unitname")]
        public string unitname { get; set; }

        //status : trạng thái
        [JsonPropertyName("status")]
        public int? status { get; set; }

        //asigneeid : id người thực hiện
        [JsonPropertyName("asigneeid")]
        public int? asigneeid { get; set; }

        //reviewerid : id người rà soát
        [JsonPropertyName("reviewerid")]
        public int? reviewerid { get; set; }

        //approvedate : ngày duyệt
        [JsonPropertyName("approvedate")]
        public DateTime? approvedate { get; set; }

        //riskid : id của rủi ro
        [JsonPropertyName("riskid")]
        public int? riskid { get; set; }

        //prototype : mẫu
        [JsonPropertyName("prototype")]
        public string prototype { get; set; }

        //conclusion : Kết luận
        [JsonPropertyName("conclusion")]
        public string conclusion { get; set; }

        [JsonPropertyName("target")]
        public string target { get; set; }
    }

    public class WorkingPaperDetailModel
    {
        [JsonPropertyName("id")]
        public int id { get; set; }

        //code : Mã giấy tờ
        [JsonPropertyName("code")]
        public string code { get; set; }

        //year : Năm
        [JsonPropertyName("year")]
        public int? year { get; set; }

        //auditworkid : id cuộc kiểm toán
        [JsonPropertyName("auditworkid")]
        public int? auditworkid { get; set; }
        [JsonPropertyName("auditworkname")]
        public string auditworkname { get; set; }

        //processid : id quy trình được kiểm toán
        [JsonPropertyName("processid")]
        public int? processid { get; set; }
        [JsonPropertyName("processname")]
        public string processname { get; set; }

        //unitid : id đơn vị
        [JsonPropertyName("unitid")]
        public int? unitid { get; set; }
        [JsonPropertyName("unitname")]
        public string unitname { get; set; }

        //status : trạng thái
        [JsonPropertyName("status")]
        public int? status { get; set; }

        //asigneeid : id người thực hiện
        [JsonPropertyName("asigneeid")]
        public int? asigneeid { get; set; }

        //reviewerid : id người rà soát
        [JsonPropertyName("reviewerid")]
        public int? reviewerid { get; set; }

        //approvedate : ngày duyệt
        [JsonPropertyName("approvedate")]
        public DateTime? approvedate { get; set; }

        //riskid : id của rủi ro
        [JsonPropertyName("riskid")]
        public int? riskid { get; set; }

        //prototype : mẫu
        [JsonPropertyName("prototype")]
        public string prototype { get; set; }

        //conclusion : Kết luận
        [JsonPropertyName("conclusion")]
        public string conclusion { get; set; }

        [JsonPropertyName("target")]
        public string target { get; set; }
    }

    public class WorkingPaperSearchResultModel
    {
        [JsonPropertyName("id")]
        public int id { get; set; }

        //code : Mã giấy tờ
        [JsonPropertyName("code")]
        public string code { get; set; }

        //year : Năm
        [JsonPropertyName("year")]
        public string year { get; set; }

        //auditworkid : id cuộc kiểm toán
        [JsonPropertyName("auditworkid")]
        public string auditworkid { get; set; }
        [JsonPropertyName("auditworkname")]
        public string auditworkname { get; set; }

        //processid : id quy trình được kiểm toán
        [JsonPropertyName("processid")]
        public string processid { get; set; }
        [JsonPropertyName("processname")]
        public string processname { get; set; }

        //unitid : id đơn vị
        [JsonPropertyName("unitid")]
        public string unitid { get; set; }
        [JsonPropertyName("unitname")]
        public string unitname { get; set; }

        //status : trạng thái
        [JsonPropertyName("status")]
        public string status { get; set; }

        //asigneeid : id người thực hiện
        [JsonPropertyName("asigneeid")]
        public string asigneeid { get; set; }

        //reviewerid : id người rà soát
        [JsonPropertyName("reviewerid")]
        public string reviewerid { get; set; }

        //approvedate : ngày duyệt
        [JsonPropertyName("approvedate")]
        public DateTime? approvedate { get; set; }

        //riskid : id của rủi ro
        [JsonPropertyName("riskid")]
        public string riskid { get; set; }

        //prototype : mẫu
        [JsonPropertyName("prototype")]
        public string prototype { get; set; }

        //conclusion : Kết luận
        [JsonPropertyName("conclusion")]
        public string conclusion { get; set; }

        [JsonPropertyName("target")]
        public string target { get; set; }
    }
}