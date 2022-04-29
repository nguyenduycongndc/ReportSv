using Report_service.DataAccess;
using Report_service.Repositories;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Text.Json;
using Report_service.Models.ExecuteModels;
using Report_service.Models.MigrationsModels;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Web;
using NPOI.XWPF.UserModel;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;
using Spire.Doc;
using NPOI.OpenXmlFormats.Wordprocessing;
using StackExchange.Redis;
using Document = Spire.Doc.Document;
using Spire.Doc.Documents;
using Spire.Doc.Reporting;
using Spire.Doc.Fields;
using System.Drawing;

namespace Report_service.Controllers.Report
{
    [Route("[controller]")]
    [ApiController]
    public class ReportAuditWorkYearController : BaseController
    {
        private const string level1 = "Kiểm soát tốt";
        private const string level2 = "Chập nhận được";
        private const string level3 = "Cần cải thiện";
        private const string level4 = "Không đạt yêu cầu";
        private const string process1 = "Chưa hoàn thành";
        private const string process2 = "Hoàn thành một phần";
        private const string process3 = "Đã hoàn thành";

        public ReportAuditWorkYearController(ILoggerManager logger, IUnitOfWork uow) : base(logger, uow)
        {
        }

        [HttpGet("Search")]
        public IActionResult Search(string jsonData)
        {
            try
            {
                var obj = JsonSerializer.Deserialize<ReportAuditWorkYearSearchModel>(jsonData);
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var status = string.IsNullOrEmpty(obj.Status) ? "" : obj.Status;

                var report = from a in _uow.Repository<ReportAuditWorkYear>().Find(x => (string.IsNullOrEmpty(obj.Year) || x.year.Value.ToString() == obj.Year) && !(x.IsDeleted ?? false)).OrderByDescending(x => x.CreatedAt)
                             join b in _uow.Repository<ApprovalFunction>().Find(x => x.function_code == "M_RAP") on a.Id equals b.item_id into g
                             from s in g.DefaultIfEmpty()
                             where (string.IsNullOrEmpty(obj.Status) || (s.StatusCode ?? "1.0") == status)
                             select a;
                var approval_config = _uow.Repository<ApprovalConfig>().GetAll(a => a.item_code == "M_RAP").OrderBy(x => x.StatusCode).ToArray();
                var approval_status = _uow.Repository<ApprovalFunction>().Find(a => a.function_code == "M_RAP").ToArray();

                IEnumerable<ReportAuditWorkYear> data = report;
                var count = data.Count();
                if (obj.StartNumber >= 0 && obj.PageSize > 0)
                {
                    data = data.Skip(obj.StartNumber).Take(obj.PageSize);
                }
                var result = data.Select(a => new ReportAuditWorkYearModel()
                {
                    Id = a.Id,
                    year = a.year,
                    name = a.name,
                    Status = approval_status.FirstOrDefault(x => x.item_id == a.Id)?.StatusCode ?? "1.0",
                    StatusName = approval_config.FirstOrDefault(x => x.StatusCode == (approval_status.FirstOrDefault(x => x.item_id == a.Id)?.StatusCode ?? "1.0"))?.StatusName ?? "",
                    //Reviewerid = (approval_status.FirstOrDefault(x => x.item_id == a.Id)?.approver_last ?? 0) != 0 ? approval_status.FirstOrDefault(x => x.item_id == a.Id)?.approver_last : approval_status.FirstOrDefault(x => x.item_id == a.Id)?.approver,
                    approval_user = approval_status.FirstOrDefault(x => x.item_id == a.Id)?.approver,
                    approval_user_last = approval_status.FirstOrDefault(x => x.item_id == a.Id)?.approver_last,
                });
                return Ok(new { code = "1", msg = "success", data = result, total = count });
            }
            catch
            {
                return BadRequest();
            }
        }
        [HttpPost]
        public IActionResult Create([FromBody] ReportAuditWorkYearModel reportauditworkyearinfo)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel _userInfo)
                {
                    return Unauthorized();
                }
                var _allcheck = _uow.Repository<ReportAuditWorkYear>().Find(a => !(a.IsDeleted ?? false) && a.year == reportauditworkyearinfo.year).ToArray();
                if (_allcheck.Length > 0)
                {
                    return Ok(new { code = "0", msg = "success" });
                }

                var report = new ReportAuditWorkYear
                {
                    year = reportauditworkyearinfo.year,
                    name = reportauditworkyearinfo.name,
                    CreatedAt = DateTime.Now,
                    CreatedBy = _userInfo.Id,
                };
                _uow.Repository<ReportAuditWorkYear>().Add(report);
                return Ok(new { code = "1", data = report.Id, msg = "success" });
            }
            catch
            {
                return BadRequest();
            }
        }
        [HttpDelete("{id}")]
        public IActionResult DeleteConfirmed(int id)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var _report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == id);
                if (_report == null)
                {
                    return NotFound();
                }
                _report.IsDeleted = true;
                _report.DeletedAt = DateTime.Now;
                _report.DeletedBy = userInfo.Id;
                _uow.Repository<ReportAuditWorkYear>().Update(_report);
                return Ok(new { code = "1", msg = "success" });
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }
        [HttpGet("{id}")]
        public IActionResult Details(int id)
        {
            try
            {
                var current_year = 0;
                var previous_year = 0;
                var now = DateTime.Now.Date;
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }

                var report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == id);
                var approval_status = _uow.Repository<ApprovalFunction>().Include(x => x.Users).FirstOrDefault(a => a.function_code == "M_RAP" && report.Id == a.item_id);

                if (report != null)
                {

                    var result = new ReportAuditWorkYearModel()
                    {
                        Id = report.Id,
                        year = report.year,
                        name = report.name,
                        evaluation = report.evaluation,
                        concerns = report.concerns,
                        reason = report.reason,
                        note = report.note,
                        quality = report.quality,
                        OverCome = report.overcome,
                        Status = approval_status?.StatusCode ?? "1.0",
                    };
                    //prepare
                    current_year = report.year.Value;
                    previous_year = report.year.Value - 1;

                    var getreport = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                     where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                     select a).ToList();
                    //end 
                    #region Tình hình thực hiện kế hoạch kiểm toán nội bộ năm
                    var auditworkplan = _uow.Repository<AuditWorkPlan>().Include(a => a.AuditPlan).Where(a => (a.Year == current_year.ToString() || a.Year == previous_year.ToString()) && a.IsDeleted != true).AsEnumerable().GroupBy(p => p.AuditPlan.Year).SelectMany(x => x.Where(v => v.AuditPlan.Version == x.Max(p => p.AuditPlan.Version)));
                    var prepareauditplan = auditworkplan.GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var reportauditwork = getreport.Where(p => p.AuditWork.Classify == 1).GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var auditworkexpected = (from a in _uow.Repository<AuditWork>().GetAll()
                                             join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_PAP") on a.Id equals b.item_id
                                             where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && a.Classify == 2 && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                             select a).GroupBy(p => p.Year).Select(p => new
                                             {
                                                 year = p.Key,
                                                 total = p.Count()
                                             }).ToList();

                    var audit_plan_current = prepareauditplan.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var audit_completed_current = reportauditwork.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var completed_current = Math.Round(audit_completed_current * 1.0 / ((audit_plan_current > 0 ? audit_plan_current : 1) * 1.0) * 100, 2);
                    var audit_expected_current = auditworkexpected.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;

                    var audit_plan_previous = prepareauditplan.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var audit_completed_previous = reportauditwork.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var completed_previous = Math.Round(audit_completed_previous * 1.0 / ((audit_plan_previous > 0 ? audit_plan_previous : 1) * 1.0) * 100, 2);
                    var audit_expected_previous = auditworkexpected.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;

                    var report1data = new ReportTable1
                    {
                        audit_plan_current = audit_plan_current,
                        audit_completed_current = audit_completed_current,
                        completed_current = completed_current,
                        audit_expected_current = audit_expected_current,
                        audit_plan_previous = audit_plan_previous,
                        audit_completed_previous = audit_completed_previous,
                        completed_previous = completed_previous,
                        audit_expected_previous = audit_expected_previous
                    };
                    var xx = (audit_plan_current - audit_plan_previous) * 1.0 / (audit_plan_previous > 0 ? audit_plan_previous : 1) * 1.0 * 100;
                    report1data.audit_plan_volatility = Math.Round((audit_plan_current - audit_plan_previous) * 1.0 / ((audit_plan_previous > 0 ? audit_plan_previous : 1) * 1.0) * 100, 2);
                    report1data.audit_completed_volatility = Math.Round((audit_completed_current - audit_completed_previous) * 1.0 / ((audit_completed_previous > 0 ? audit_completed_previous : 1) * 1.0) * 100, 2);
                    #endregion

                    #region Kết quả cuộc kiểm toán đã phát hành báo cáo trong năm
                    var report2data = getreport.Where(p => p.Year == current_year.ToString()).Select(p => new ReportTable2
                    {
                        audit_name = p.AuditWork.Name,
                        audit_time = (p.StartDateField.HasValue ? p.StartDateField.Value.ToString("dd/MM/yyyy") : "") + "-" + (p.EndDateField.HasValue ? p.EndDateField.Value.ToString("dd/MM/yyyy") : ""),
                        report_date = p.ReportDate.HasValue ? p.ReportDate.Value.ToString("dd/MM/yyyy") : "",
                        level = GetLevel(p.AuditRatingLevelTotal),
                        risk_high = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 1).ToString(),
                        risk_medium = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 2).ToString(),
                    }).ToList();
                    #endregion

                    #region Các phát hiện kiểm toán trọng yếu
                    var report3data = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                       join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                       where b.StatusCode == "3.1" && !(a.AuditDetect.IsDeleted ?? false) && a.AuditDetect.rating_risk == 1 && a.AuditDetect.year == current_year
                                       select new ReportTable3
                                       {
                                           audit_name = a.AuditDetect.auditwork_name,
                                           audit_summary = a.AuditDetect.summary_audit_detect,
                                           audit_request_content = a.Content,
                                           audit_request_status = GetProcess(a.ProcessStatus),
                                       }).ToList();
                    #endregion

                    #region Tình hình thực hiện các kiến nghị của kiểm toán nội 

                    var auditrequestmonitor = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                               join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                               where b.StatusCode == "3.1" && a.AuditDetect != null && !(a.AuditDetect.IsDeleted ?? false) && (a.AuditDetect.rating_risk == 1 || a.AuditDetect.rating_risk == 2) && a.AuditDetect.year <= current_year && !(a.is_deleted ?? false)
                                               let time = (a.extend_at.HasValue ? a.extend_at.Value : a.CompleteAt.HasValue ? a.CompleteAt.Value : DateTime.MinValue)
                                               select new ReportAuditRequest
                                               {
                                                   code = a.Code,
                                                   extendat = a.extend_at,
                                                   rating = a.AuditDetect.rating_risk,
                                                   completed = a.CompleteAt,
                                                   actualcompleted = a.ActualCompleteAt,
                                                   conclusion = a.Conclusion,
                                                   timestatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false))) ? 2 : 1,
                                                   processstatus = a.ProcessStatus,
                                                   year = a.AuditDetect.year,
                                                   day = ((a.ProcessStatus ?? 1) == 3 && a.ActualCompleteAt.HasValue) ? (a.ActualCompleteAt - time).Value.TotalDays : (now - time).TotalDays
                                               }
                                               ).ToList();
                    var type1 = GetData(auditrequestmonitor.Where(p => p.timestatus == 1), 1, current_year);
                    var type2 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2), 2, current_year);
                    var type3 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && p.day < 30), 3, current_year);
                    var type4 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 30 <= p.day && p.day < 60), 4, current_year);
                    var type5 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 60 <= p.day && p.day <= 90), 5, current_year);
                    var type6 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 90 < p.day), 6, current_year);
                    var type7 = GetData(auditrequestmonitor.Where(p => p.extendat.HasValue), 7, current_year);
                    List<ReportTable4> report4data = new()
                    {
                        type1,
                        type2,
                        type3,
                        type4,
                        type5,
                        type6,
                        type7
                    };
                    var type8 = new ReportTable4
                    {
                        type = 8,
                        beginning_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_high),
                        notclose_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_high),
                        close_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_high),
                        ending_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_high),
                        beginning_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_medium),
                        notclose_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_medium),
                        close_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_medium),
                        ending_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_medium),
                    };
                    report4data.Add(type8);
                    #endregion
                    return Ok(new { code = "1", msg = "success", data = result, report1 = report1data, report2 = report2data, report3 = report3data, report4 = report4data });
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception)
            {

                return BadRequest();
            }
        }
        [HttpPut]
        public IActionResult Update([FromBody] ReportAuditWorkYearModel reportauditworkyearinfo)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel _userInfo)
                {
                    return Unauthorized();
                }
                var report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == reportauditworkyearinfo.Id);
                if (report == null)
                {
                    return NotFound();
                }

                report.evaluation = reportauditworkyearinfo.evaluation;
                report.concerns = reportauditworkyearinfo.concerns;
                report.reason = reportauditworkyearinfo.reason;
                report.note = reportauditworkyearinfo.note;
                report.quality = reportauditworkyearinfo.quality;
                report.overcome = reportauditworkyearinfo.OverCome;
                report.ModifiedAt = DateTime.Now;
                report.ModifiedBy = _userInfo.Id;
                _uow.Repository<ReportAuditWorkYear>().Update(report);
                return Ok(new { code = "1", data = report.Id, msg = "success" });
            }
            catch
            {
                return BadRequest();
            }
        }

        private static string GetLevel(int? value)
        {
            return value switch
            {
                1 => level1,
                2 => level2,
                3 => level3,
                _ => level4,
            };
        }

        private static string GetProcess(int? value)
        {
            return value switch
            {
                1 => process1,
                2 => process2,
                3 => process3,
                _ => process1,
            };
        }

        private static ReportTable4 GetData(IEnumerable<ReportAuditRequest> list, int type, int current_year)
        {
            var item = new ReportTable4()
            {
                type = type,
                beginning_high = list.Count(x => x.rating == 1 && x.year < current_year),
                notclose_high = list.Count(x => x.rating == 1 && x.conclusion != 2 && x.year == current_year),
                close_high = list.Count(x => x.rating == 1 && x.conclusion == 2 && x.year == current_year),
                beginning_medium = list.Count(x => x.rating == 2 && x.year < current_year),
                notclose_medium = list.Count(x => x.rating == 2 && x.conclusion != 2 && x.year == current_year),
                close_medium = list.Count(x => x.rating == 2 && x.conclusion == 2 && x.year == current_year),
            };
            item.ending_high = item.beginning_high + item.notclose_high;
            item.ending_medium = item.beginning_medium + item.notclose_medium;
            return item;
        }

        [HttpGet("ListStatusReportYear")]
        public IActionResult ListStatusReportYear()
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var approval_status = _uow.Repository<ApprovalConfig>().GetAll(a => a.item_code == "M_RAP").ToArray().OrderBy(x => x.StatusCode);

                var dt = approval_status.Select(a => new StatusApprove()
                {
                    id = a.id,
                    status_code = a.StatusCode,
                    status_name = a.StatusName,
                });
                return Ok(new { code = "1", msg = "success", data = dt });
            }
            catch (Exception)
            {
                return Ok(new { code = "0", msg = "fail", data = new StatusApprove() });
            }

        }
        [HttpPost("ExportWord")]
        public IActionResult ExportWord()
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                Spire.Doc.Document doc = new Spire.Doc.Document();
                var path = "D:/PROJECT_TINHVAN/KitanoApplication/API_backend/report-service/Report_service/";
                string output = Path.Combine(path, "MicrosoftOffice.docx");
                doc.LoadFromFile(output, FileFormat.Docx);

                //Merge the specified value into file
                string[] Datevalues = { string.Format("{0:d}", System.DateTime.Now), string.Format("{0:d}", System.DateTime.Now), string.Format("{0:d}", System.DateTime.Now), string.Format("{0:d}", System.DateTime.Now), string.Format("{0:d}", System.DateTime.Now) };
                string[] FieldName = { "SubGrantPAStartDateValue", "SubGrantPAEndDateValue", "SubGrantPAExtensionDateValue", "SubGrantPSStartDateValue", "SubGrantPSEndDateValue" };
                doc.MailMerge.Execute(FieldName, Datevalues);
                MemoryStream fs = new MemoryStream();
                fs.Position = 0;
                doc.SaveToStream(fs, FileFormat.Docx);
                byte[] bytes = fs.ToArray();
                return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Grid.doc");
            }
            catch (Exception ex)
            {
                return Ok(new { code = "0", msg = "fail", data = new StatusApprove() });
            }

        }
        public class SystemParam
        {
            public int? id { get; set; }
            public string name { get; set; }
            public string value { get; set; }
        }
        [HttpPost("ExportWordold/{id}")]
        public IActionResult ExportWordold(int id)
        {

            var _paramInfoPrefix = "Param.SystemInfo";
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var iCache = (IConnectionMultiplexer)HttpContext.RequestServices
                      .GetService(typeof(IConnectionMultiplexer));
                if (!iCache.IsConnected)
                {
                    return BadRequest();
                }
                var redisDb = iCache.GetDatabase();
                var value_get = redisDb.StringGet(_paramInfoPrefix);
                var company = "";
                if (value_get.HasValue)
                {
                    var list_param = JsonSerializer.Deserialize<List<SystemParam>>(value_get);
                    company = list_param.FirstOrDefault(a => a.name == "COMPANY_NAME")?.value;
                }
                var report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == id);
                if (report != null)
                {
                    var current_year = 0;
                    var previous_year = 0;
                    var now = DateTime.Now.Date;
                    current_year = report.year.Value;
                    previous_year = report.year.Value - 1;

                    var getreport = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                     where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                     select a).ToList();

                    #region Tình hình thực hiện kế hoạch kiểm toán nội bộ năm
                    var auditworkplan = _uow.Repository<AuditWorkPlan>().Include(a => a.AuditPlan).Where(a => (a.Year == current_year.ToString() || a.Year == previous_year.ToString()) && a.IsDeleted != true).AsEnumerable().GroupBy(p => p.AuditPlan.Year).SelectMany(x => x.Where(v => v.AuditPlan.Version == x.Max(p => p.AuditPlan.Version)));
                    var prepareauditplan = auditworkplan.GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var reportauditwork = getreport.Where(p => p.AuditWork.Classify == 1).GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var auditworkexpected = (from a in _uow.Repository<AuditWork>().GetAll()
                                             join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_PAP") on a.Id equals b.item_id
                                             where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && a.Classify == 2 && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                             select a).GroupBy(p => p.Year).Select(p => new
                                             {
                                                 year = p.Key,
                                                 total = p.Count()
                                             }).ToList();

                    var audit_plan_current = prepareauditplan.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var audit_completed_current = reportauditwork.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var completed_current = Math.Round((double)(audit_completed_current / (audit_plan_current > 0 ? audit_plan_current : 1)) * 100, 2);
                    var audit_expected_current = auditworkexpected.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;

                    var audit_plan_previous = prepareauditplan.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var audit_completed_previous = reportauditwork.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var completed_previous = Math.Round((double)(audit_completed_previous / (audit_plan_previous > 0 ? audit_plan_previous : 1)) * 100, 2);
                    var audit_expected_previous = auditworkexpected.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;

                    var report1data = new ReportTable1
                    {
                        audit_plan_current = audit_plan_current,
                        audit_completed_current = audit_completed_current,
                        completed_current = completed_current,
                        audit_expected_current = audit_expected_current,
                        audit_plan_previous = audit_plan_previous,
                        audit_completed_previous = audit_completed_previous,
                        completed_previous = completed_previous,
                        audit_expected_previous = audit_expected_previous
                    };
                    report1data.audit_plan_volatility = Math.Round((audit_plan_current - audit_plan_previous) * 1.0 / (audit_plan_previous > 0 ? audit_plan_previous : 1) * 1.0 * 100, 2);
                    report1data.audit_completed_volatility = Math.Round((audit_completed_current - audit_completed_previous) * 1.0 / (audit_completed_previous > 0 ? audit_completed_previous : 1) * 1.0 * 100, 2);
                    #endregion

                    #region Kết quả cuộc kiểm toán đã phát hành báo cáo trong năm
                    var report2data = getreport.Where(p => p.Year == current_year.ToString()).Select(p => new ReportTable2
                    {
                        audit_name = p.AuditWork.Name,
                        audit_time = (p.StartDateField.HasValue ? p.StartDateField.Value.ToString("dd/MM/yyyy") : "") + "-" + (p.EndDateField.HasValue ? p.EndDateField.Value.ToString("dd/MM/yyyy") : ""),
                        report_date = p.ReportDate.HasValue ? p.ReportDate.Value.ToString("dd/MM/yyyy") : "",
                        level = GetLevel(p.AuditRatingLevelTotal),
                        risk_high = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 1).ToString(),
                        risk_medium = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 2).ToString(),
                    }).ToList();
                    #endregion

                    #region Các phát hiện kiểm toán trọng yếu
                    var report3data = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                       join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                       where b.StatusCode == "3.1" && !(a.AuditDetect.IsDeleted ?? false) && a.AuditDetect.rating_risk == 1 && a.AuditDetect.year == current_year
                                       select new ReportTable3
                                       {
                                           audit_name = a.AuditDetect.auditwork_name,
                                           audit_summary = a.AuditDetect.summary_audit_detect,
                                           audit_request_content = a.Content,
                                           audit_request_status = GetProcess(a.ProcessStatus),
                                       }).ToList();
                    #endregion

                    #region Tình hình thực hiện các kiến nghị của kiểm toán nội 

                    var auditrequestmonitor = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                               join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                               where b.StatusCode == "3.1" && a.AuditDetect != null && !(a.AuditDetect.IsDeleted ?? false) && (a.AuditDetect.rating_risk == 1 || a.AuditDetect.rating_risk == 2) && a.AuditDetect.year <= current_year && !(a.is_deleted ?? false)
                                               let time = (a.extend_at.HasValue ? a.extend_at.Value : a.CompleteAt.HasValue ? a.CompleteAt.Value : DateTime.MinValue)
                                               select new ReportAuditRequest
                                               {
                                                   code = a.Code,
                                                   extendat = a.extend_at,
                                                   rating = a.AuditDetect.rating_risk,
                                                   completed = a.CompleteAt,
                                                   actualcompleted = a.ActualCompleteAt,
                                                   conclusion = a.Conclusion,
                                                   timestatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false))) ? 2 : 1,
                                                   processstatus = a.ProcessStatus,
                                                   year = a.AuditDetect.year,
                                                   day = ((a.ProcessStatus ?? 1) == 3 && a.ActualCompleteAt.HasValue) ? (a.ActualCompleteAt - time).Value.TotalDays : (now - time).TotalDays
                                               }
                                               ).ToList();
                    var type1 = GetData(auditrequestmonitor.Where(p => p.timestatus == 1), 1, current_year);
                    var type2 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2), 2, current_year);
                    var type3 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && p.day < 30), 3, current_year);
                    var type4 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 30 <= p.day && p.day < 60), 4, current_year);
                    var type5 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 60 <= p.day && p.day <= 90), 5, current_year);
                    var type6 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 90 < p.day), 6, current_year);
                    var type7 = GetData(auditrequestmonitor.Where(p => p.extendat.HasValue), 7, current_year);
                    List<ReportTable4> report4data = new()
                    {
                        type1,
                        type2,
                        type3,
                        type4,
                        type5,
                        type6,
                        type7
                    };
                    var type8 = new ReportTable4
                    {
                        type = 8,
                        beginning_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_high),
                        notclose_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_high),
                        close_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_high),
                        ending_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_high),
                        beginning_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_medium),
                        notclose_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_medium),
                        close_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_medium),
                        ending_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_medium),
                    };
                    report4data.Add(type8);
                    #endregion
                    var _config = (IConfiguration)HttpContext.RequestServices
                    .GetService(typeof(IConfiguration));
                    var path = _config["Template:ReporDocsTemplate"];
                    string output = Path.Combine(path, "Kitano_BaoCaoKiemToanNam.docx");
                    var stream = new FileStream(output, FileMode.Open);
                    XWPFDocument doc = new XWPFDocument(stream);
                    var tt = new
                    {
                        year = current_year,
                        yearreport = "Ngày " + now.Day + " tháng " + now.Month + " năm " + now.Year,
                        danhgiatongquanktnb = report.evaluation ?? "",
                        quanngaichinhktnb = report.concerns ?? "",
                        reason = report.reason ?? "",
                        note = report.note ?? "",
                        congtactheodoikhacphuc = report.overcome ?? "",
                        congtacquantrichatluonghoatdong = string.IsNullOrEmpty(report.quality) ? "" : HttpUtility.HtmlDecode(report.quality),
                        company = company,
                    };
                    var table2field = new
                    {
                        auditplancurrent = report1data.audit_plan_current,
                        auditcompletedcurrent = report1data.audit_completed_current,
                        completedcurrent = report1data.completed_current,
                        auditexpectedcurrent = report1data.audit_expected_current,
                        auditplanprevious = report1data.audit_plan_previous,
                        auditcompletedprevious = report1data.audit_completed_previous,
                        completedprevious = report1data.completed_previous,
                        auditexpectedprevious = report1data.audit_expected_previous,
                        auditplanvolatility = report1data.audit_plan_volatility,
                        auditcompletedvolatility = report1data.audit_completed_volatility,
                    };
                    var typy1 = report4data.FirstOrDefault(a => a.type == 1);
                    var typy2 = report4data.FirstOrDefault(a => a.type == 2);
                    var typy3 = report4data.FirstOrDefault(a => a.type == 3);
                    var typy4 = report4data.FirstOrDefault(a => a.type == 4);
                    var typy5 = report4data.FirstOrDefault(a => a.type == 5);
                    var typy6 = report4data.FirstOrDefault(a => a.type == 6);
                    var typy7 = report4data.FirstOrDefault(a => a.type == 7);
                    var typy8 = report4data.FirstOrDefault(a => a.type == 8);
                    var table5field = new
                    {
                        dauky11 = typy1.beginning_high,
                        dauky21 = typy2.beginning_high,
                        dauky31 = typy3.beginning_high,
                        dauky41 = typy4.beginning_high,
                        dauky51 = typy5.beginning_high,
                        dauky61 = typy6.beginning_high,
                        dauky71 = typy7.beginning_high,
                        dauky81 = typy8.beginning_high,
                        dauky12 = typy1.beginning_medium,
                        dauky22 = typy2.beginning_medium,
                        dauky32 = typy3.beginning_medium,
                        dauky42 = typy4.beginning_medium,
                        dauky52 = typy5.beginning_medium,
                        dauky62 = typy6.beginning_medium,
                        dauky72 = typy7.beginning_medium,
                        dauky82 = typy8.beginning_medium,
                        chuadong11 = typy1.notclose_high,
                        chuadong21 = typy2.notclose_high,
                        chuadong31 = typy3.notclose_high,
                        chuadong41 = typy4.notclose_high,
                        chuadong51 = typy5.notclose_high,
                        chuadong61 = typy6.notclose_high,
                        chuadong71 = typy7.notclose_high,
                        chuadong81 = typy8.notclose_high,
                        chuadong12 = typy1.notclose_medium,
                        chuadong22 = typy2.notclose_medium,
                        chuadong32 = typy3.notclose_medium,
                        chuadong42 = typy4.notclose_medium,
                        chuadong52 = typy5.notclose_medium,
                        chuadong62 = typy6.notclose_medium,
                        chuadong72 = typy7.notclose_medium,
                        chuadong82 = typy8.notclose_medium,
                        dadong11 = typy1.close_high,
                        dadong21 = typy2.close_high,
                        dadong31 = typy3.close_high,
                        dadong41 = typy4.close_high,
                        dadong51 = typy5.close_high,
                        dadong61 = typy6.close_high,
                        dadong71 = typy7.close_high,
                        dadong81 = typy8.close_high,
                        dadong12 = typy1.close_medium,
                        dadong22 = typy2.close_medium,
                        dadong32 = typy3.close_medium,
                        dadong42 = typy4.close_medium,
                        dadong52 = typy5.close_medium,
                        dadong62 = typy6.close_medium,
                        dadong72 = typy7.close_medium,
                        dadong82 = typy8.close_medium,
                        cuoiky11 = typy1.ending_high,
                        cuoiky21 = typy2.ending_high,
                        cuoiky31 = typy3.ending_high,
                        cuoiky41 = typy4.ending_high,
                        cuoiky51 = typy5.ending_high,
                        cuoiky61 = typy6.ending_high,
                        cuoiky71 = typy7.ending_high,
                        cuoiky81 = typy8.ending_high,
                        cuoiky12 = typy1.ending_medium,
                        cuoiky22 = typy2.ending_medium,
                        cuoiky32 = typy3.ending_medium,
                        cuoiky42 = typy4.ending_medium,
                        cuoiky52 = typy5.ending_medium,
                        cuoiky62 = typy6.ending_medium,
                        cuoiky72 = typy7.ending_medium,
                        cuoiky82 = typy8.ending_medium,
                    };
                    //Traverse paragraphs                  
                    foreach (var para in doc.Paragraphs)
                    {
                        ReplaceKey(para, tt);
                    }
                    //Traverse the table      
                    var tables = doc.Tables;
                    foreach (var table in tables)
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.GetTableCells())
                            {
                                foreach (var para in cell.Paragraphs)
                                {
                                    ReplaceKey(para, tt);
                                }
                            }
                        }
                    }
                    #region [Add row table 2]
                    var oprTable2 = tables[1];
                    foreach (var row in oprTable2.Rows)
                    {
                        foreach (var cell in row.GetTableCells())
                        {
                            foreach (var para in cell.Paragraphs)
                            {
                                ReplaceKey(para, table2field);
                            }
                        }
                    }
                    #endregion
                    #region [Add row table 3]

                    var oprTable3 = tables[2];
                    if (report2data.Count == 0)
                    {
                        XWPFTableRow m_Row3 = oprTable3.CreateRow();
                        m_Row3.AddNewTableCell().SetText("");
                        //m_Row3.GetCell(0).SetText("");
                        //m_Row3.GetCell(1).SetText("");
                        //m_Row3.GetCell(2).SetText("");
                        //m_Row3.GetCell(3).SetText("");
                        //m_Row3.GetCell(4).SetText("");
                        //m_Row3.GetCell(5).SetText("");
                        //m_Row3.GetCell(6).SetText("");
                    }
                    else
                    {
                        var i3 = 1;
                        foreach (var item in report2data)
                        {
                            XWPFTableRow m_Row3 = oprTable3.CreateRow();
                            m_Row3.GetCell(0)?.SetText(i3 + "");
                            m_Row3.GetCell(1)?.SetText(item.audit_name);
                            m_Row3.GetCell(2)?.SetText(item.audit_time);
                            m_Row3.GetCell(3)?.SetText(item.report_date);
                            m_Row3.GetCell(4)?.SetText(item.level);
                            m_Row3.GetCell(5)?.SetText(item.risk_high);
                            m_Row3.CreateCell().SetText(item.risk_medium);
                            //m_Row3.GetCell(6)?.SetText(item.risk_medium);
                            i3++;
                        }
                    }

                    #endregion
                    #region [Add row table 4]

                    var oprTable4 = tables[3];
                    if (report3data.Count == 0)
                    {
                        XWPFTableRow m_Row4 = oprTable4.CreateRow();
                        m_Row4.GetCell(0)?.SetText("");
                        m_Row4.GetCell(1)?.SetText("");
                        m_Row4.GetCell(2)?.SetText("");
                        m_Row4.GetCell(3)?.SetText("");
                    }
                    else
                    {
                        foreach (var item in report3data)
                        {
                            XWPFTableRow m_Row4 = oprTable4.CreateRow();
                            m_Row4.GetCell(0)?.SetText(item.audit_name);
                            m_Row4.GetCell(1)?.SetText(item.audit_summary);
                            m_Row4.GetCell(2)?.SetText(item.audit_request_content);
                            m_Row4.GetCell(3)?.SetText(item.audit_request_status);
                        }
                    }

                    #endregion
                    #region [Add row table 5]
                    var oprTable5 = tables[4];
                    foreach (var row in oprTable5.Rows)
                    {
                        foreach (var cell in row.GetTableCells())
                        {
                            foreach (var para in cell.Paragraphs)
                            {
                                ReplaceKey(para, table5field);
                            }
                        }
                    }
                    #endregion

                    MemoryStream fs = new MemoryStream();
                    fs.Position = 0;
                    doc.Write(fs);
                    byte[] bytes = fs.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Grid.doc");
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                return Ok(new { code = "0", msg = "fail", data = new StatusApprove() });
            }
        }
        [HttpPost("ExportWordnewMIC/{id}")]
        public IActionResult ExportWordNewMIC(int id)
        {

            var _paramInfoPrefix = "Param.SystemInfo";
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var iCache = (IConnectionMultiplexer)HttpContext.RequestServices
                      .GetService(typeof(IConnectionMultiplexer));
                if (!iCache.IsConnected)
                {
                    return BadRequest();
                }
                var redisDb = iCache.GetDatabase();
                var value_get = redisDb.StringGet(_paramInfoPrefix);
                var company = "";
                if (value_get.HasValue)
                {
                    var list_param = JsonSerializer.Deserialize<List<SystemParam>>(value_get);
                    company = list_param.FirstOrDefault(a => a.name == "COMPANY_NAME")?.value;
                }
                var report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == id);
                if (report != null)
                {
                    var current_year = 0;
                    var previous_year = 0;
                    var now = DateTime.Now.Date;
                    current_year = report.year.Value;
                    previous_year = report.year.Value - 1;

                    var getreport = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                     where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                     select a).ToList();

                    #region Tình hình thực hiện kế hoạch kiểm toán nội bộ năm
                    var auditworkplan = _uow.Repository<AuditWorkPlan>().Include(a => a.AuditPlan).Where(a => (a.Year == current_year.ToString() || a.Year == previous_year.ToString()) && a.IsDeleted != true).AsEnumerable().GroupBy(p => p.AuditPlan.Year).SelectMany(x => x.Where(v => v.AuditPlan.Version == x.Max(p => p.AuditPlan.Version)));
                    var prepareauditplan = auditworkplan.GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var reportauditwork = getreport.Where(p => p.AuditWork.Classify == 1).GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var auditworkexpected = (from a in _uow.Repository<AuditWork>().GetAll()
                                             join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_PAP") on a.Id equals b.item_id
                                             where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && a.Classify == 2 && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                             select a).GroupBy(p => p.Year).Select(p => new
                                             {
                                                 year = p.Key,
                                                 total = p.Count()
                                             }).ToList();

                    var audit_plan_current = prepareauditplan.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var audit_completed_current = reportauditwork.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var completed_current = Math.Round((double)(audit_completed_current / (audit_plan_current > 0 ? audit_plan_current : 1)) * 100, 2);
                    var audit_expected_current = auditworkexpected.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;

                    var audit_plan_previous = prepareauditplan.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var audit_completed_previous = reportauditwork.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var completed_previous = Math.Round((double)(audit_completed_previous / (audit_plan_previous > 0 ? audit_plan_previous : 1)) * 100, 2);
                    var audit_expected_previous = auditworkexpected.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;

                    var report1data = new ReportTable1
                    {
                        audit_plan_current = audit_plan_current,
                        audit_completed_current = audit_completed_current,
                        completed_current = completed_current,
                        audit_expected_current = audit_expected_current,
                        audit_plan_previous = audit_plan_previous,
                        audit_completed_previous = audit_completed_previous,
                        completed_previous = completed_previous,
                        audit_expected_previous = audit_expected_previous
                    };
                    report1data.audit_plan_volatility = Math.Round((audit_plan_current - audit_plan_previous) * 1.0 / (audit_plan_previous > 0 ? audit_plan_previous : 1) * 1.0 * 100, 2);
                    report1data.audit_completed_volatility = Math.Round((audit_completed_current - audit_completed_previous) * 1.0 / (audit_completed_previous > 0 ? audit_completed_previous : 1) * 1.0 * 100, 2);
                    #endregion

                    #region Kết quả cuộc kiểm toán đã phát hành báo cáo trong năm
                    var report2data = getreport.Where(p => p.Year == current_year.ToString()).Select(p => new ReportTable2
                    {
                        audit_name = p.AuditWork.Name,
                        audit_time = (p.StartDateField.HasValue ? p.StartDateField.Value.ToString("dd/MM/yyyy") : "") + "-" + (p.EndDateField.HasValue ? p.EndDateField.Value.ToString("dd/MM/yyyy") : ""),
                        report_date = p.ReportDate.HasValue ? p.ReportDate.Value.ToString("dd/MM/yyyy") : "",
                        level = GetLevel(p.AuditRatingLevelTotal),
                        risk_high = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 1).ToString(),
                        risk_medium = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 2).ToString(),
                    }).ToList();
                    #endregion

                    #region Các phát hiện kiểm toán trọng yếu
                    var report3data = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                       join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                       where b.StatusCode == "3.1" && !(a.AuditDetect.IsDeleted ?? false) && a.AuditDetect.rating_risk == 1 && a.AuditDetect.year == current_year
                                       select new ReportTable3
                                       {
                                           audit_name = a.AuditDetect.auditwork_name,
                                           audit_summary = a.AuditDetect.summary_audit_detect,
                                           audit_request_content = a.Content,
                                           audit_request_status = GetProcess(a.ProcessStatus),
                                       }).ToList();
                    #endregion

                    #region Tình hình thực hiện các kiến nghị của kiểm toán nội 

                    var auditrequestmonitor = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                               join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                               where b.StatusCode == "3.1" && a.AuditDetect != null && !(a.AuditDetect.IsDeleted ?? false) && (a.AuditDetect.rating_risk == 1 || a.AuditDetect.rating_risk == 2) && a.AuditDetect.year <= current_year && !(a.is_deleted ?? false)
                                               let time = (a.extend_at.HasValue ? a.extend_at.Value : a.CompleteAt.HasValue ? a.CompleteAt.Value : DateTime.MinValue)
                                               select new ReportAuditRequest
                                               {
                                                   code = a.Code,
                                                   extendat = a.extend_at,
                                                   rating = a.AuditDetect.rating_risk,
                                                   completed = a.CompleteAt,
                                                   actualcompleted = a.ActualCompleteAt,
                                                   conclusion = a.Conclusion,
                                                   timestatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false))) ? 2 : 1,
                                                   processstatus = a.ProcessStatus,
                                                   year = a.AuditDetect.year,
                                                   day = ((a.ProcessStatus ?? 1) == 3 && a.ActualCompleteAt.HasValue) ? (a.ActualCompleteAt - time).Value.TotalDays : (now - time).TotalDays
                                               }
                                               ).ToList();
                    var type1 = GetData(auditrequestmonitor.Where(p => p.timestatus == 1), 1, current_year);
                    var type2 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2), 2, current_year);
                    var type3 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && p.day < 30), 3, current_year);
                    var type4 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 30 <= p.day && p.day < 60), 4, current_year);
                    var type5 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 60 <= p.day && p.day <= 90), 5, current_year);
                    var type6 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 90 < p.day), 6, current_year);
                    var type7 = GetData(auditrequestmonitor.Where(p => p.extendat.HasValue), 7, current_year);
                    List<ReportTable4> report4data = new()
                    {
                        type1,
                        type2,
                        type3,
                        type4,
                        type5,
                        type6,
                        type7
                    };
                    var type8 = new ReportTable4
                    {
                        type = 8,
                        beginning_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_high),
                        notclose_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_high),
                        close_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_high),
                        ending_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_high),
                        beginning_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_medium),
                        notclose_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_medium),
                        close_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_medium),
                        ending_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_medium),
                    };
                    report4data.Add(type8);
                    #endregion
                    var _config = (IConfiguration)HttpContext.RequestServices
                    .GetService(typeof(IConfiguration));
                    var path = _config["Template:ReporDocsTemplate"];
                    var fullPath = Path.Combine(path, "Kitano_BaoCaoKiemToanNam.docx");
                    fullPath = fullPath.ToString().Replace("\\", "/");
                    Document doc = new Document(fullPath);
                    var tt = new
                    {
                        year = current_year + "",
                        yearreport = "Ngày " + now.Day + " tháng " + now.Month + " năm " + now.Year,
                        danhgiatongquanktnb = report.evaluation ?? "",
                        quanngaichinhktnb = report.concerns ?? "",
                        reason = report.reason ?? "",
                        note = report.note ?? "",
                        congtactheodoikhacphuc = report.overcome ?? "",
                        congtacquantrichatluonghoatdong = "",
                        company = company,
                    };
                    var table2field = new
                    {
                        auditplancurrent = report1data.audit_plan_current + "",
                        auditcompletedcurrent = report1data.audit_completed_current + "",
                        completedcurrent = report1data.completed_current + "",
                        auditexpectedcurrent = report1data.audit_expected_current + "",
                        auditplanprevious = report1data.audit_plan_previous + "",
                        auditcompletedprevious = report1data.audit_completed_previous + "",
                        completedprevious = report1data.completed_previous + "",
                        auditexpectedprevious = report1data.audit_expected_previous + "",
                        auditplanvolatility = report1data.audit_plan_volatility + "",
                        auditcompletedvolatility = report1data.audit_completed_volatility + "",
                    };
                    var typy1 = report4data.FirstOrDefault(a => a.type == 1);
                    var typy2 = report4data.FirstOrDefault(a => a.type == 2);
                    var typy3 = report4data.FirstOrDefault(a => a.type == 3);
                    var typy4 = report4data.FirstOrDefault(a => a.type == 4);
                    var typy5 = report4data.FirstOrDefault(a => a.type == 5);
                    var typy6 = report4data.FirstOrDefault(a => a.type == 6);
                    var typy7 = report4data.FirstOrDefault(a => a.type == 7);
                    var typy8 = report4data.FirstOrDefault(a => a.type == 8);
                    var table5field = new
                    {
                        dauky11 = typy1.beginning_high + "",
                        dauky21 = typy2.beginning_high + "",
                        dauky31 = typy3.beginning_high + "",
                        dauky41 = typy4.beginning_high + "",
                        dauky51 = typy5.beginning_high + "",
                        dauky61 = typy6.beginning_high + "",
                        dauky71 = typy7.beginning_high + "",
                        dauky81 = typy8.beginning_high + "",
                        dauky12 = typy1.beginning_medium + "",
                        dauky22 = typy2.beginning_medium + "",
                        dauky32 = typy3.beginning_medium + "",
                        dauky42 = typy4.beginning_medium + "",
                        dauky52 = typy5.beginning_medium + "",
                        dauky62 = typy6.beginning_medium + "",
                        dauky72 = typy7.beginning_medium + "",
                        dauky82 = typy8.beginning_medium + "",
                        chuadong11 = typy1.notclose_high + "",
                        chuadong21 = typy2.notclose_high + "",
                        chuadong31 = typy3.notclose_high + "",
                        chuadong41 = typy4.notclose_high + "",
                        chuadong51 = typy5.notclose_high + "",
                        chuadong61 = typy6.notclose_high + "",
                        chuadong71 = typy7.notclose_high + "",
                        chuadong81 = typy8.notclose_high + "",
                        chuadong12 = typy1.notclose_medium + "",
                        chuadong22 = typy2.notclose_medium + "",
                        chuadong32 = typy3.notclose_medium + "",
                        chuadong42 = typy4.notclose_medium + "",
                        chuadong52 = typy5.notclose_medium + "",
                        chuadong62 = typy6.notclose_medium + "",
                        chuadong72 = typy7.notclose_medium + "",
                        chuadong82 = typy8.notclose_medium + "",
                        dadong11 = typy1.close_high + "",
                        dadong21 = typy2.close_high + "",
                        dadong31 = typy3.close_high + "",
                        dadong41 = typy4.close_high + "",
                        dadong51 = typy5.close_high + "",
                        dadong61 = typy6.close_high + "",
                        dadong71 = typy7.close_high + "",
                        dadong81 = typy8.close_high + "",
                        dadong12 = typy1.close_medium + "",
                        dadong22 = typy2.close_medium + "",
                        dadong32 = typy3.close_medium + "",
                        dadong42 = typy4.close_medium + "",
                        dadong52 = typy5.close_medium + "",
                        dadong62 = typy6.close_medium + "",
                        dadong72 = typy7.close_medium + "",
                        dadong82 = typy8.close_medium + "",
                        cuoiky11 = typy1.ending_high + "",
                        cuoiky21 = typy2.ending_high + "",
                        cuoiky31 = typy3.ending_high + "",
                        cuoiky41 = typy4.ending_high + "",
                        cuoiky51 = typy5.ending_high + "",
                        cuoiky61 = typy6.ending_high + "",
                        cuoiky71 = typy7.ending_high + "",
                        cuoiky81 = typy8.ending_high + "",
                        cuoiky12 = typy1.ending_medium + "",
                        cuoiky22 = typy2.ending_medium + "",
                        cuoiky32 = typy3.ending_medium + "",
                        cuoiky42 = typy4.ending_medium + "",
                        cuoiky52 = typy5.ending_medium + "",
                        cuoiky62 = typy6.ending_medium + "",
                        cuoiky72 = typy7.ending_medium + "",
                        cuoiky82 = typy8.ending_medium + "",
                    };
                    TextSelection selectionss = doc.FindAllString("congtacquantrichatluonghoatdong", false, true)?.FirstOrDefault();
                    if (selectionss != null)
                    {
                        if (!string.IsNullOrEmpty(report.quality))
                        {
                            var paragraph = selectionss?.GetAsOneRange()?.OwnerParagraph;
                            var object_arr = paragraph.ChildObjects;
                            if (object_arr.Count > 0)
                            {
                                for (int j = 0; j < object_arr.Count; j++)
                                {
                                    paragraph.ChildObjects.RemoveAt(j);
                                }
                                paragraph.AppendHTML(report.quality);
                            }
                        }
                    }
                    Table table1 = doc.Sections[0].Tables[2] as Table;
                    var ii = 1;
                    foreach (var item in report2data)
                    {
                        var row = table1.AddRow();
                        var txtTable1 = row.Cells[0].AddParagraph().AppendText(ii + "");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[1].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[2].AddParagraph().AppendText(item.audit_time);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[3].AddParagraph().AppendText(item.report_date);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                        txtTable1 = row.Cells[4].AddParagraph().AppendText(item.level);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                        txtTable1 = row.Cells[5].AddParagraph().AppendText(item.risk_high);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                        txtTable1 = row.Cells[6].AddParagraph().AppendText(item.risk_medium);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        ii++;
                    }
                    Table table2 = doc.Sections[0].Tables[3] as Table;
                    foreach (var item in report3data)
                    {
                        var row = table2.AddRow();
                        var txtTable1 = row.Cells[0].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[1].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[2].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[3].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                    }

                    doc.MailMerge.Execute(
                    new string[] { "year" ,"year_1" ,"year_2" , "yearreport" , "danhgiatongquanktnb" ,
                        "quanngaichinhktnb",
                        "reason",
                        "note",
                        "congtactheodoikhacphuc",
                        "congtacquantrichatluonghoatdong",
                        "company",
                        "company_1",
                        "auditplancurrent",
                        "auditcompletedcurrent",
                        "completedcurrent",
                        "auditexpectedcurrent",
                        "auditplanprevious",
                        "auditcompletedprevious",
                        "completedprevious",
                        "auditexpectedprevious",
                        "auditplanvolatility",
                        "auditcompletedvolatility",
                        "dauky11",
                        "dauky21",
                        "dauky31",
                        "dauky41",
                        "dauky51",
                        "dauky61",
                        "dauky71",
                        "dauky81",
                        "dauky12",
                        "dauky22",
                        "dauky32",
                        "dauky42",
                        "dauky52",
                        "dauky62",
                        "dauky72",
                        "dauky82",
                        "chuadong11",
                        "chuadong21",
                        "chuadong31",
                        "chuadong41",
                        "chuadong51",
                        "chuadong61",
                        "chuadong71",
                        "chuadong81",
                        "chuadong12",
                        "chuadong22",
                        "chuadong32",
                        "chuadong42",
                        "chuadong52",
                        "chuadong62",
                        "chuadong72",
                        "chuadong82",
                        "dadong11",
                        "dadong21",
                        "dadong31",
                        "dadong41",
                        "dadong51",
                        "dadong61",
                        "dadong71",
                        "dadong81",
                        "dadong12",
                        "dadong22",
                        "dadong32",
                        "dadong42",
                        "dadong52",
                        "dadong62",
                        "dadong72",
                        "dadong82",
                        "cuoiky11",
                        "cuoiky21",
                        "cuoiky31",
                        "cuoiky41",
                        "cuoiky51",
                        "cuoiky61",
                        "cuoiky71",
                        "cuoiky81",
                        "cuoiky12",
                        "cuoiky22",
                        "cuoiky32",
                        "cuoiky42",
                        "cuoiky52",
                        "cuoiky62",
                        "cuoiky72",
                        "cuoiky82"
                    },
                    new string[] { tt.year ,tt.year,tt.year, tt.yearreport ,tt.danhgiatongquanktnb , tt.quanngaichinhktnb ,tt.reason,
                       tt.note ,tt.congtactheodoikhacphuc,tt.congtacquantrichatluonghoatdong,tt.company,tt.company,
                       table2field.auditplancurrent,table2field.auditcompletedcurrent,table2field.completedcurrent,table2field.auditexpectedcurrent,table2field.auditplanprevious,
                       table2field.auditcompletedprevious,table2field.completedprevious,table2field.auditexpectedprevious,table2field.auditplanvolatility,table2field.auditcompletedvolatility, table5field.dauky11,table5field.dauky21,table5field.dauky31,table5field.dauky41,table5field.dauky51,table5field.dauky61,table5field.dauky71,table5field.dauky81,table5field.dauky12,table5field.dauky22,table5field.dauky32,table5field.dauky42,table5field.dauky52,table5field.dauky62,table5field.dauky72,table5field.dauky82,         table5field.chuadong11,table5field.chuadong21,table5field.chuadong31,table5field.chuadong41,table5field.chuadong51,table5field.chuadong61,table5field.chuadong71,table5field.chuadong81,table5field.chuadong12,table5field.chuadong22,table5field.chuadong32,table5field.chuadong42,table5field.chuadong52,table5field.chuadong62,table5field.chuadong72,table5field.chuadong82,        table5field.dadong11,table5field.dadong21,table5field.dadong31,table5field.dadong41,table5field.dadong51,table5field.dadong61,table5field.dadong71,table5field.dadong81,table5field.dadong12,table5field.dadong22,table5field.dadong32,table5field.dadong42,table5field.dadong52,table5field.dadong62,table5field.dadong72,table5field.dadong82,               table5field.cuoiky11,table5field.cuoiky21,table5field.cuoiky31,table5field.cuoiky41,table5field.cuoiky51,table5field.cuoiky61,table5field.cuoiky71,table5field.cuoiky81,table5field.cuoiky12,table5field.cuoiky22,table5field.cuoiky32,table5field.cuoiky42,table5field.cuoiky52,table5field.cuoiky62,table5field.cuoiky72,table5field.cuoiky82
                        }
                    );

                    //foreach (Section section1 in doc.Sections)
                    //{
                    //    for (int i = 0; i < section1.Body.ChildObjects.Count; i++)
                    //    {
                    //        if (section1.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
                    //        {
                    //            if (String.IsNullOrEmpty((section1.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                    //            {
                    //                section1.Body.ChildObjects.Remove(section1.Body.ChildObjects[i]);
                    //                i--;
                    //            }
                    //        }

                    //    }
                    //}

                    MemoryStream stream = new MemoryStream();
                    stream.Position = 0;
                    doc.SaveToFile(stream, FileFormat.Docx);
                    byte[] bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Grid.docx");
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                return Ok(new { code = "0", msg = "fail", data = new StatusApprove() });
            }
        }

        private static string ConvertPreciousNew(DateTime datatime)
        {
            string text = "";
            var get_month = datatime.Month;
            switch (get_month)
            {
                case 1:
                case 2:
                case 3:
                    text = "Quý I";
                    break;
                case 4:
                case 5:
                case 6:
                    text = "Quý II";
                    break;
                case 7:
                case 8:
                case 9:
                    text = "Quý II";
                    break;
                case 10:
                case 11:
                case 12:
                    text = "Quý IV";
                    break;
            }
            return text;
        }
        [HttpPost("ExportWordnewMBS/{id}")]
        public IActionResult ExportWordNewMBS(int id)
        {

            var _paramInfoPrefix = "Param.SystemInfo";
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var iCache = (IConnectionMultiplexer)HttpContext.RequestServices
                      .GetService(typeof(IConnectionMultiplexer));
                if (!iCache.IsConnected)
                {
                    return BadRequest();
                }
                var hearderSystem = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "REPORT_HEADER");
                var dataCt = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "COMPANY_NAME");
                var redisDb = iCache.GetDatabase();
                var value_get = redisDb.StringGet(_paramInfoPrefix);
                var company = "";
                if (value_get.HasValue)
                {
                    var list_param = JsonSerializer.Deserialize<List<SystemParam>>(value_get);
                    company = list_param.FirstOrDefault(a => a.name == "COMPANY_NAME")?.value;
                }
                var report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == id);
                if (report != null)
                {
                    var current_year = 0;
                    var previous_year = 0;
                    var now = DateTime.Now.Date;
                    current_year = report.year.Value;
                    previous_year = report.year.Value - 1;

                    var getreport = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                     where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                     select a).ToList();

                    #region Tình hình thực hiện kế hoạch kiểm toán nội bộ năm
                    var auditworkplan = _uow.Repository<AuditWorkPlan>().Include(a => a.AuditPlan).Where(a => (a.Year == current_year.ToString() || a.Year == previous_year.ToString()) && a.IsDeleted != true).AsEnumerable().GroupBy(p => p.AuditPlan.Year).SelectMany(x => x.Where(v => v.AuditPlan.Version == x.Max(p => p.AuditPlan.Version)));
                    var prepareauditplan = auditworkplan.GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var reportauditwork = getreport.Where(p => p.AuditWork.Classify == 1).GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var auditworkexpected = (from a in _uow.Repository<AuditWork>().GetAll()
                                             join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_PAP") on a.Id equals b.item_id
                                             where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && a.Classify == 2 && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                             select a).GroupBy(p => p.Year).Select(p => new
                                             {
                                                 year = p.Key,
                                                 total = p.Count()
                                             }).ToList();

                    var audit_plan_current = prepareauditplan.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var audit_completed_current = reportauditwork.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var completed_current = Math.Round((double)(audit_completed_current / (audit_plan_current > 0 ? audit_plan_current : 1)) * 100, 2);
                    var audit_expected_current = auditworkexpected.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;

                    var audit_plan_previous = prepareauditplan.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var audit_completed_previous = reportauditwork.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var completed_previous = Math.Round((double)(audit_completed_previous / (audit_plan_previous > 0 ? audit_plan_previous : 1)) * 100, 2);
                    var audit_expected_previous = auditworkexpected.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;

                    var report1data = new ReportTable1
                    {
                        audit_plan_current = audit_plan_current,
                        audit_completed_current = audit_completed_current,
                        completed_current = completed_current,
                        audit_expected_current = audit_expected_current,
                        audit_plan_previous = audit_plan_previous,
                        audit_completed_previous = audit_completed_previous,
                        completed_previous = completed_previous,
                        audit_expected_previous = audit_expected_previous
                    };
                    report1data.audit_plan_volatility = Math.Round((audit_plan_current - audit_plan_previous) * 1.0 / (audit_plan_previous > 0 ? audit_plan_previous : 1) * 1.0 * 100, 2);
                    report1data.audit_completed_volatility = Math.Round((audit_completed_current - audit_completed_previous) * 1.0 / (audit_completed_previous > 0 ? audit_completed_previous : 1) * 1.0 * 100, 2);
                    #endregion

                    #region Kết quả cuộc kiểm toán đã phát hành báo cáo trong năm
                    var report2data = getreport.Where(p => p.Year == current_year.ToString()).Select(p => new ReportTable2
                    {
                        audit_name = p.AuditWork.Name,
                        audit_time = (p.StartDateField.HasValue ? p.StartDateField.Value.ToString("dd/MM/yyyy") : "") + "-" + (p.EndDateField.HasValue ? p.EndDateField.Value.ToString("dd/MM/yyyy") : ""),
                        report_date = p.ReportDate.HasValue ? p.ReportDate.Value.ToString("dd/MM/yyyy") : "",
                        level = GetLevel(p.AuditRatingLevelTotal),
                        risk_high = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 1).ToString(),
                        risk_medium = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 2).ToString(),
                    }).ToList();
                    #endregion

                    #region Các phát hiện kiểm toán trọng yếu
                    var report3data = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                       join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                       where b.StatusCode == "3.1" && !(a.AuditDetect.IsDeleted ?? false) && a.AuditDetect.rating_risk == 1 && a.AuditDetect.year == current_year
                                       select new ReportTable3
                                       {
                                           audit_name = a.AuditDetect.auditwork_name,
                                           audit_summary = a.AuditDetect.summary_audit_detect,
                                           audit_request_content = a.Content,
                                           audit_request_status = GetProcess(a.ProcessStatus),
                                       }).ToList();
                    #endregion

                    #region Tình hình thực hiện các kiến nghị của kiểm toán nội 

                    var auditrequestmonitor = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                               join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                               where b.StatusCode == "3.1" && a.AuditDetect != null && !(a.AuditDetect.IsDeleted ?? false) && (a.AuditDetect.rating_risk == 1 || a.AuditDetect.rating_risk == 2) && a.AuditDetect.year <= current_year && !(a.is_deleted ?? false)
                                               let time = (a.extend_at.HasValue ? a.extend_at.Value : a.CompleteAt.HasValue ? a.CompleteAt.Value : DateTime.MinValue)
                                               select new ReportAuditRequest
                                               {
                                                   code = a.Code,
                                                   extendat = a.extend_at,
                                                   rating = a.AuditDetect.rating_risk,
                                                   completed = a.CompleteAt,
                                                   actualcompleted = a.ActualCompleteAt,
                                                   conclusion = a.Conclusion,
                                                   timestatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false))) ? 2 : 1,
                                                   processstatus = a.ProcessStatus,
                                                   year = a.AuditDetect.year,
                                                   day = ((a.ProcessStatus ?? 1) == 3 && a.ActualCompleteAt.HasValue) ? (a.ActualCompleteAt - time).Value.TotalDays : (now - time).TotalDays
                                               }
                                               ).ToList();
                    var type1 = GetData(auditrequestmonitor.Where(p => p.timestatus == 1), 1, current_year);
                    var type2 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2), 2, current_year);
                    var type3 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && p.day < 30), 3, current_year);
                    var type4 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 30 <= p.day && p.day < 60), 4, current_year);
                    var type5 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 60 <= p.day && p.day <= 90), 5, current_year);
                    var type6 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 90 < p.day), 6, current_year);
                    var type7 = GetData(auditrequestmonitor.Where(p => p.extendat.HasValue), 7, current_year);
                    List<ReportTable4> report4data = new()
                    {
                        type1,
                        type2,
                        type3,
                        type4,
                        type5,
                        type6,
                        type7
                    };
                    var type8 = new ReportTable4
                    {
                        type = 8,
                        beginning_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_high),
                        notclose_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_high),
                        close_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_high),
                        ending_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_high),
                        beginning_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_medium),
                        notclose_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_medium),
                        close_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_medium),
                        ending_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_medium),
                    };
                    report4data.Add(type8);
                    #endregion
                    var _config = (IConfiguration)HttpContext.RequestServices
                    .GetService(typeof(IConfiguration));

                    #region Data mới
                    var list_auditwork = (from a in _uow.Repository<AuditWork>().GetAll()
                                          where !(a.IsDeleted ?? false) && (a.Year == current_year.ToString())
                                          select a).ToList();
                    var dataTKH = (from a in _uow.Repository<AuditWork>().GetAll()
                                   where !(a.IsDeleted ?? false) && a.Classify == 1 && (a.Year == current_year.ToString())
                                   select a).ToList();
                    var dataDX = (from a in _uow.Repository<AuditWork>().GetAll()
                                  where !(a.IsDeleted ?? false) && a.Classify == 2 && (a.Year == current_year.ToString())
                                  select a).ToList();
                    var BC = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                              join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                              where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString())
                              select a).ToList();
                    #endregion
                    //var path = _config["Template:ReporDocsTemplate"];
                    //var fullPath = Path.Combine(path, "Kitano_BaoCaoKiemToanNam.docx");
                    //fullPath = fullPath.ToString().Replace("\\", "/");
                    //Document doc = new Document(fullPath);
                    Document doc = new Document(@"D:\test\MBS_Kitano_BaoCaoKiemToanNam_v0.1.docx");
                    var tt = new
                    {
                        year = current_year + "",
                        yearreport = "Ngày " + now.Day + " tháng " + now.Month + " năm " + now.Year,
                        danhgiatongquanktnb = report.evaluation ?? "",
                        quanngaichinhktnb = report.concerns ?? "",
                        reason = report.reason ?? "",
                        note = report.note ?? "",
                        congtactheodoikhacphuc = report.overcome ?? "",
                        congtacquantrichatluonghoatdong = "",
                        company = company,
                    };


                    doc.MailMerge.Execute(
                    new string[] { "year" ,"cong_ty_1",

                    },
                    new string[] { tt.year,dataCt.Value,

                        }
                    );
                    #region Table1
                    Table table1 = doc.Sections[0].Tables[0] as Table;
                    String[] Header = { "STT", "Tên chương trình", "Thời gian thực hiện", "Đã hoàn thành" };
                    table1.ResetCells(list_auditwork.Count() + 4, 4);

                    //Header Row
                    TableRow FRow = table1.Rows[0];
                    FRow.Height = 23;
                    //Header Format
                    FRow.RowFormat.BackColor = ColorTranslator.FromHtml("#c0d9ff");
                    for (int i = 0; i < Header.Length; i++)
                    {
                        //Cell Alignment
                        Paragraph p = FRow.Cells[i].AddParagraph();
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        //Data Format
                        TextRange TR = p.AppendText(Header[i]);
                        TR.CharacterFormat.FontName = "Times New Roman";
                        TR.CharacterFormat.FontSize = 12;
                        TR.CharacterFormat.Bold = true;
                        FRow.Cells[0].SetCellWidth(10, CellWidthType.Percentage);
                    }
                    TableRow _RowI = table1.Rows[1];
                    _RowI.Height = 23;
                    Paragraph prI = _RowI.Cells[0].AddParagraph();
                    _RowI.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prI.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TextRange txtTable1 = prI.AppendText("I");
                    table1[1, 0].CellFormat.BackColor = ColorTranslator.FromHtml("#BDBDBD");
                    txtTable1.CharacterFormat.FontName = "Times New Roman";
                    txtTable1.CharacterFormat.FontSize = 12;
                    txtTable1.CharacterFormat.Bold = true;
                    _RowI.Cells[0].SetCellWidth(5, CellWidthType.Percentage);

                    _RowI.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    txtTable1 = table1[1, 1].AddParagraph().AppendText("Các chương trình trong kế hoạch được TBKS MB giao");
                    table1[1, 1].CellFormat.BackColor = ColorTranslator.FromHtml("#BDBDBD");
                    txtTable1.CharacterFormat.FontName = "Times New Roman";
                    txtTable1.CharacterFormat.FontSize = 12;
                    txtTable1.CharacterFormat.Bold = true;
                    table1.ApplyHorizontalMerge(1, 1, 3);
                    for (int x = 0; x < dataTKH.Count(); x++)
                    {
                        var quy = dataTKH[x].StartDate.HasValue ? ConvertPreciousNew(dataTKH[x].StartDate.Value) : "";
                        var BCD = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                                   join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                   where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.auditwork_id == dataTKH[x].Id)
                                   select a).FirstOrDefault();

                        TableRow _Row = table1.Rows[x + 2];
                        _Row.Height = 23;
                        Paragraph pr10 = _Row.Cells[0].AddParagraph();
                        _Row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr10.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr10.AppendText((x + 1).ToString() + ".");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        _Row.Cells[0].SetCellWidth(10, CellWidthType.Percentage);

                        Paragraph pr11 = _Row.Cells[1].AddParagraph();
                        _Row.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr11.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr11.AppendText(dataTKH[x].Name);
                        //_Row.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        //pr.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        //txtTable1 = table1[x + 2, 1].AddParagraph().AppendText(dataTKH[x].Name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;

                        Paragraph pr12 = _Row.Cells[2].AddParagraph();
                        _Row.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr12.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr12.AppendText(quy);
                        //txtTable1 = table1[x + 2, 2].AddParagraph().AppendText(quy);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;

                        Paragraph pr13 = _Row.Cells[3].AddParagraph();
                        _Row.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr13.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr13.AppendText(BCD != null ? "x" : "");
                        //txtTable1 = table1[x + 2, 3].AddParagraph().AppendText(BCD != null ? "x" : "");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                    }
                    TableRow _RowII = table1.Rows[dataTKH.Count() + 2];
                    _RowII.Height = 23;
                    Paragraph prII = _RowII.Cells[0].AddParagraph();
                    _RowII.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prII.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable1 = prII.AppendText("II");
                    table1[dataTKH.Count() + 2, 0].CellFormat.BackColor = ColorTranslator.FromHtml("#BDBDBD");
                    txtTable1.CharacterFormat.FontName = "Times New Roman";
                    txtTable1.CharacterFormat.FontSize = 12;
                    txtTable1.CharacterFormat.Bold = true;
                    _RowII.Cells[0].SetCellWidth(10, CellWidthType.Percentage);

                    _RowII.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    txtTable1 = table1[dataTKH.Count() + 2, 1].AddParagraph().AppendText("Các chương trình phát sinh ngoài kế hoạch được giao");
                    table1[dataTKH.Count() + 2, 1].CellFormat.BackColor = ColorTranslator.FromHtml("#BDBDBD");
                    txtTable1.CharacterFormat.FontName = "Times New Roman";
                    txtTable1.CharacterFormat.FontSize = 12;
                    txtTable1.CharacterFormat.Bold = true;
                    table1.ApplyHorizontalMerge(dataTKH.Count() + 2, 1, 3);
                    for (int z = 0; z < dataDX.Count(); z++)
                    {
                        var quy = dataDX[z].StartDate.HasValue ? ConvertPreciousNew(dataDX[z].StartDate.Value) : "";
                        var BCDX = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                                    join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                    where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.auditwork_id == dataDX[z].Id)
                                    select a).FirstOrDefault();
                        TableRow _Row2 = table1.Rows[dataTKH.Count() + z + 3];
                        Paragraph pr20 = _Row2.Cells[0].AddParagraph();
                        _Row2.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr20.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr20.AppendText((z + 1).ToString() + ".");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;

                        Paragraph pr21 = _Row2.Cells[1].AddParagraph();
                        _Row2.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr21.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr21.AppendText(dataDX[z].Name);
                        //txtTable1 = table1[dataTKH.Count() + z + 3, 1].AddParagraph().AppendText(dataDX[z].Name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;

                        Paragraph pr22 = _Row2.Cells[2].AddParagraph();
                        _Row2.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr22.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr22.AppendText(quy);
                        //txtTable1 = table1[dataTKH.Count() + z + 3, 2].AddParagraph().AppendText(quy);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;

                        Paragraph pr23 = _Row2.Cells[3].AddParagraph();
                        _Row2.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        pr23.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtTable1 = pr23.AppendText(BCDX != null ? "x" : "");
                        //txtTable1 = table1[dataTKH.Count() + z + 3, 3].AddParagraph().AppendText(BCDX != null ? "x" : "");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                    }

                    TableRow _Row3 = table1.Rows[list_auditwork.Count() + 3];
                    _Row3.Height = 23;
                    _Row3.RowFormat.BackColor = ColorTranslator.FromHtml("#c0d9ff");
                    Paragraph pr31 = _Row3.Cells[1].AddParagraph();
                    _Row3.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    pr31.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable1 = pr31.AppendText("Tổng cộng");
                    //txtTable1 = table1[list_auditwork.Count() + 3, 1].AddParagraph().AppendText("Tổng cộng");
                    txtTable1.CharacterFormat.FontName = "Times New Roman";
                    txtTable1.CharacterFormat.FontSize = 12;
                    _Row3.Cells[0].SetCellWidth(10, CellWidthType.Percentage);

                    Paragraph pr32 = _Row3.Cells[2].AddParagraph();
                    _Row3.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    pr32.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable1 = pr32.AppendText(list_auditwork.Count().ToString());
                    //txtTable1 = table1[list_auditwork.Count() + 3, 2].AddParagraph().AppendText(list_auditwork.Count().ToString());
                    txtTable1.CharacterFormat.FontName = "Times New Roman";
                    txtTable1.CharacterFormat.FontSize = 12;

                    Paragraph pr33 = _Row3.Cells[3].AddParagraph();
                    _Row3.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    pr33.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable1 = pr33.AppendText(BC.Count().ToString());
                    //txtTable1 = table1[list_auditwork.Count() + 3, 3].AddParagraph().AppendText(BC.Count().ToString());
                    txtTable1.CharacterFormat.FontName = "Times New Roman";
                    txtTable1.CharacterFormat.FontSize = 12;
                    #endregion
                    #region Table2
                    Table table2 = doc.Sections[0].Tables[1] as Table;

                    table2.ResetCells(3, 7);
                    table2.ApplyVerticalMerge(0, 0, 1);
                    table2.ApplyVerticalMerge(4, 0, 1);
                    table2.ApplyVerticalMerge(5, 0, 1);
                    table2.ApplyVerticalMerge(6, 0, 1);
                    table2.ApplyHorizontalMerge(0, 1, 3);

                    TableRow RowTb2_0 = table2.Rows[0];
                    RowTb2_0.Height = 23;
                    //Header Format
                    RowTb2_0.RowFormat.BackColor = ColorTranslator.FromHtml("#c0d9ff");
                    Paragraph prTb2_00 = RowTb2_0.Cells[0].AddParagraph();
                    RowTb2_0.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_00.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TextRange txtTable2 = prTb2_00.AppendText("Kế hoạch");

                    //TextRange txtTable2 = table2[0, 0].AddParagraph().AppendText("Kế hoạch");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    Paragraph prTb2_01 = RowTb2_0.Cells[1].AddParagraph();
                    RowTb2_0.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_01.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable2 = prTb2_01.AppendText("Thực hiện");
                    //txtTable2 = table2[0, 1].AddParagraph().AppendText("Thực hiện");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    Paragraph prTb2_04 = RowTb2_0.Cells[4].AddParagraph();
                    RowTb2_0.Cells[4].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_04.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable2 = prTb2_04.AppendText("Năm liền trước");
                    //txtTable2 = table2[0, 4].AddParagraph().AppendText("Năm liền trước");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    Paragraph prTb2_05 = RowTb2_0.Cells[5].AddParagraph();
                    RowTb2_0.Cells[5].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_05.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable2 = prTb2_05.AppendText("%Thực hiện/Kế hoạch");
                    //txtTable2 = table2[0, 5].AddParagraph().AppendText("%Thực hiện/Kế hoạch");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    Paragraph prTb2_06 = RowTb2_0.Cells[6].AddParagraph();
                    RowTb2_0.Cells[6].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_06.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable2 = prTb2_06.AppendText("% Thực hiện so với năm liền trước");
                    //txtTable2 = table2[0, 6].AddParagraph().AppendText("% Thực hiện so với năm liền trước");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    TableRow RowTb2_1 = table2.Rows[1];
                    RowTb2_1.Height = 23;
                    //Header Format
                    RowTb2_1.RowFormat.BackColor = ColorTranslator.FromHtml("#c0d9ff");
                    Paragraph prTb2_11 = RowTb2_1.Cells[1].AddParagraph();
                    RowTb2_1.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_11.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable2 = prTb2_11.AppendText("Tổng");
                    //txtTable2 = table2[1, 1].AddParagraph().AppendText("Tổng");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    Paragraph prTb2_12 = RowTb2_1.Cells[2].AddParagraph();
                    RowTb2_1.Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_12.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable2 = prTb2_12.AppendText("Trong kế hoạch");
                    //txtTable2 = table2[1, 2].AddParagraph().AppendText("Trong kế hoạch");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    Paragraph prTb2_13 = RowTb2_1.Cells[3].AddParagraph();
                    RowTb2_1.Cells[3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    prTb2_13.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    txtTable2 = prTb2_13.AppendText("Đột xuất");
                    //txtTable2 = table2[1, 3].AddParagraph().AppendText("Đột xuất");
                    txtTable2.CharacterFormat.FontName = "Times New Roman";
                    txtTable2.CharacterFormat.FontSize = 12;
                    txtTable2.CharacterFormat.Bold = true;

                    #endregion
                    //foreach (Section section1 in doc.Sections)
                    //{
                    //    for (int i = 0; i < section1.Body.ChildObjects.Count; i++)
                    //    {
                    //        if (section1.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
                    //        {
                    //            if (String.IsNullOrEmpty((section1.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                    //            {
                    //                section1.Body.ChildObjects.Remove(section1.Body.ChildObjects[i]);
                    //                i--;
                    //            }
                    //        }

                    //    }
                    //}

                    MemoryStream stream = new MemoryStream();
                    stream.Position = 0;
                    doc.SaveToFile(stream, FileFormat.Docx);
                    byte[] bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Grid.docx");
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                return Ok(new { code = "0", msg = "fail", data = new StatusApprove() });
            }
        }
        [HttpPost("ExportWordnewAMC/{id}")]
        public IActionResult ExportWordNewAMC(int id)
        {

            var _paramInfoPrefix = "Param.SystemInfo";
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var iCache = (IConnectionMultiplexer)HttpContext.RequestServices
                      .GetService(typeof(IConnectionMultiplexer));
                if (!iCache.IsConnected)
                {
                    return BadRequest();
                }
                var redisDb = iCache.GetDatabase();
                var value_get = redisDb.StringGet(_paramInfoPrefix);
                var company = "";
                if (value_get.HasValue)
                {
                    var list_param = JsonSerializer.Deserialize<List<SystemParam>>(value_get);
                    company = list_param.FirstOrDefault(a => a.name == "COMPANY_NAME")?.value;
                }
                var report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == id);
                if (report != null)
                {
                    var current_year = 0;
                    var previous_year = 0;
                    var now = DateTime.Now.Date;
                    current_year = report.year.Value;
                    previous_year = report.year.Value - 1;

                    var getreport = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect)
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                     where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                     select a).ToList();

                    #region Tình hình thực hiện kế hoạch kiểm toán nội bộ năm
                    var auditworkplan = _uow.Repository<AuditWorkPlan>().Include(a => a.AuditPlan).Where(a => (a.Year == current_year.ToString() || a.Year == previous_year.ToString()) && a.IsDeleted != true).AsEnumerable().GroupBy(p => p.AuditPlan.Year).SelectMany(x => x.Where(v => v.AuditPlan.Version == x.Max(p => p.AuditPlan.Version)));
                    var prepareauditplan = auditworkplan.GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var reportauditwork = getreport.Where(p => p.AuditWork.Classify == 1).GroupBy(p => p.Year).Select(p => new
                    {
                        year = p.Key,
                        total = p.Count()
                    }).ToList();
                    var auditworkexpected = (from a in _uow.Repository<AuditWork>().GetAll()
                                             join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_PAP") on a.Id equals b.item_id
                                             where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && a.Classify == 2 && (a.Year == current_year.ToString() || a.Year == previous_year.ToString())
                                             select a).GroupBy(p => p.Year).Select(p => new
                                             {
                                                 year = p.Key,
                                                 total = p.Count()
                                             }).ToList();

                    var audit_plan_current = prepareauditplan.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var audit_completed_current = reportauditwork.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;
                    var completed_current = Math.Round((double)(audit_completed_current / (audit_plan_current > 0 ? audit_plan_current : 1)) * 100, 2);
                    var audit_expected_current = auditworkexpected.FirstOrDefault(p => p.year == current_year.ToString())?.total ?? 0;

                    var audit_plan_previous = prepareauditplan.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var audit_completed_previous = reportauditwork.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;
                    var completed_previous = Math.Round((double)(audit_completed_previous / (audit_plan_previous > 0 ? audit_plan_previous : 1)) * 100, 2);
                    var audit_expected_previous = auditworkexpected.FirstOrDefault(p => p.year == previous_year.ToString())?.total ?? 0;

                    var report1data = new ReportTable1
                    {
                        audit_plan_current = audit_plan_current,
                        audit_completed_current = audit_completed_current,
                        completed_current = completed_current,
                        audit_expected_current = audit_expected_current,
                        audit_plan_previous = audit_plan_previous,
                        audit_completed_previous = audit_completed_previous,
                        completed_previous = completed_previous,
                        audit_expected_previous = audit_expected_previous
                    };
                    report1data.audit_plan_volatility = Math.Round((audit_plan_current - audit_plan_previous) * 1.0 / (audit_plan_previous > 0 ? audit_plan_previous : 1) * 1.0 * 100, 2);
                    report1data.audit_completed_volatility = Math.Round((audit_completed_current - audit_completed_previous) * 1.0 / (audit_completed_previous > 0 ? audit_completed_previous : 1) * 1.0 * 100, 2);
                    #endregion

                    #region Kết quả cuộc kiểm toán đã phát hành báo cáo trong năm
                    var report2data = getreport.Where(p => p.Year == current_year.ToString()).Select(p => new ReportTable2
                    {
                        audit_name = p.AuditWork.Name,
                        audit_time = (p.StartDateField.HasValue ? p.StartDateField.Value.ToString("dd/MM/yyyy") : "") + "-" + (p.EndDateField.HasValue ? p.EndDateField.Value.ToString("dd/MM/yyyy") : ""),
                        report_date = p.ReportDate.HasValue ? p.ReportDate.Value.ToString("dd/MM/yyyy") : "",
                        level = GetLevel(p.AuditRatingLevelTotal),
                        risk_high = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 1).ToString(),
                        risk_medium = p.AuditWork.AuditDetect.Count(x => x.rating_risk == 2).ToString(),
                    }).ToList();
                    #endregion

                    #region Các phát hiện kiểm toán trọng yếu
                    var report3data = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                       join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                       where b.StatusCode == "3.1" && !(a.AuditDetect.IsDeleted ?? false) && a.AuditDetect.rating_risk == 1 && a.AuditDetect.year == current_year
                                       select new ReportTable3
                                       {
                                           audit_name = a.AuditDetect.auditwork_name,
                                           audit_summary = a.AuditDetect.summary_audit_detect,
                                           audit_request_content = a.Content,
                                           audit_request_status = GetProcess(a.ProcessStatus),
                                       }).ToList();
                    #endregion

                    #region Tình hình thực hiện các kiến nghị của kiểm toán nội 

                    var auditrequestmonitor = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                               join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                               where b.StatusCode == "3.1" && a.AuditDetect != null && !(a.AuditDetect.IsDeleted ?? false) && (a.AuditDetect.rating_risk == 1 || a.AuditDetect.rating_risk == 2) && a.AuditDetect.year <= current_year && !(a.is_deleted ?? false)
                                               let time = (a.extend_at.HasValue ? a.extend_at.Value : a.CompleteAt.HasValue ? a.CompleteAt.Value : DateTime.MinValue)
                                               select new ReportAuditRequest
                                               {
                                                   code = a.Code,
                                                   extendat = a.extend_at,
                                                   rating = a.AuditDetect.rating_risk,
                                                   completed = a.CompleteAt,
                                                   actualcompleted = a.ActualCompleteAt,
                                                   conclusion = a.Conclusion,
                                                   timestatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false))) ? 2 : 1,
                                                   processstatus = a.ProcessStatus,
                                                   year = a.AuditDetect.year,
                                                   day = ((a.ProcessStatus ?? 1) == 3 && a.ActualCompleteAt.HasValue) ? (a.ActualCompleteAt - time).Value.TotalDays : (now - time).TotalDays
                                               }
                                               ).ToList();
                    var type1 = GetData(auditrequestmonitor.Where(p => p.timestatus == 1), 1, current_year);
                    var type2 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2), 2, current_year);
                    var type3 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && p.day < 30), 3, current_year);
                    var type4 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 30 <= p.day && p.day < 60), 4, current_year);
                    var type5 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 60 <= p.day && p.day <= 90), 5, current_year);
                    var type6 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 90 < p.day), 6, current_year);
                    var type7 = GetData(auditrequestmonitor.Where(p => p.extendat.HasValue), 7, current_year);
                    List<ReportTable4> report4data = new()
                    {
                        type1,
                        type2,
                        type3,
                        type4,
                        type5,
                        type6,
                        type7
                    };
                    var type8 = new ReportTable4
                    {
                        type = 8,
                        beginning_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_high),
                        notclose_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_high),
                        close_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_high),
                        ending_high = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_high),
                        beginning_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.beginning_medium),
                        notclose_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.notclose_medium),
                        close_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.close_medium),
                        ending_medium = report4data.Where(p => p.type != 2 && p.type != 7).Sum(p => p.ending_medium),
                    };
                    report4data.Add(type8);
                    #endregion
                    var _config = (IConfiguration)HttpContext.RequestServices
                    .GetService(typeof(IConfiguration));
                    var path = _config["Template:ReporDocsTemplate"];
                    var fullPath = Path.Combine(path, "Kitano_BaoCaoKiemToanNam.docx");
                    fullPath = fullPath.ToString().Replace("\\", "/");
                    Document doc = new Document(fullPath);
                    var tt = new
                    {
                        year = current_year + "",
                        yearreport = "Ngày " + now.Day + " tháng " + now.Month + " năm " + now.Year,
                        danhgiatongquanktnb = report.evaluation ?? "",
                        quanngaichinhktnb = report.concerns ?? "",
                        reason = report.reason ?? "",
                        note = report.note ?? "",
                        congtactheodoikhacphuc = report.overcome ?? "",
                        congtacquantrichatluonghoatdong = "",
                        company = company,
                    };
                    var table2field = new
                    {
                        auditplancurrent = report1data.audit_plan_current + "",
                        auditcompletedcurrent = report1data.audit_completed_current + "",
                        completedcurrent = report1data.completed_current + "",
                        auditexpectedcurrent = report1data.audit_expected_current + "",
                        auditplanprevious = report1data.audit_plan_previous + "",
                        auditcompletedprevious = report1data.audit_completed_previous + "",
                        completedprevious = report1data.completed_previous + "",
                        auditexpectedprevious = report1data.audit_expected_previous + "",
                        auditplanvolatility = report1data.audit_plan_volatility + "",
                        auditcompletedvolatility = report1data.audit_completed_volatility + "",
                    };
                    var typy1 = report4data.FirstOrDefault(a => a.type == 1);
                    var typy2 = report4data.FirstOrDefault(a => a.type == 2);
                    var typy3 = report4data.FirstOrDefault(a => a.type == 3);
                    var typy4 = report4data.FirstOrDefault(a => a.type == 4);
                    var typy5 = report4data.FirstOrDefault(a => a.type == 5);
                    var typy6 = report4data.FirstOrDefault(a => a.type == 6);
                    var typy7 = report4data.FirstOrDefault(a => a.type == 7);
                    var typy8 = report4data.FirstOrDefault(a => a.type == 8);
                    var table5field = new
                    {
                        dauky11 = typy1.beginning_high + "",
                        dauky21 = typy2.beginning_high + "",
                        dauky31 = typy3.beginning_high + "",
                        dauky41 = typy4.beginning_high + "",
                        dauky51 = typy5.beginning_high + "",
                        dauky61 = typy6.beginning_high + "",
                        dauky71 = typy7.beginning_high + "",
                        dauky81 = typy8.beginning_high + "",
                        dauky12 = typy1.beginning_medium + "",
                        dauky22 = typy2.beginning_medium + "",
                        dauky32 = typy3.beginning_medium + "",
                        dauky42 = typy4.beginning_medium + "",
                        dauky52 = typy5.beginning_medium + "",
                        dauky62 = typy6.beginning_medium + "",
                        dauky72 = typy7.beginning_medium + "",
                        dauky82 = typy8.beginning_medium + "",
                        chuadong11 = typy1.notclose_high + "",
                        chuadong21 = typy2.notclose_high + "",
                        chuadong31 = typy3.notclose_high + "",
                        chuadong41 = typy4.notclose_high + "",
                        chuadong51 = typy5.notclose_high + "",
                        chuadong61 = typy6.notclose_high + "",
                        chuadong71 = typy7.notclose_high + "",
                        chuadong81 = typy8.notclose_high + "",
                        chuadong12 = typy1.notclose_medium + "",
                        chuadong22 = typy2.notclose_medium + "",
                        chuadong32 = typy3.notclose_medium + "",
                        chuadong42 = typy4.notclose_medium + "",
                        chuadong52 = typy5.notclose_medium + "",
                        chuadong62 = typy6.notclose_medium + "",
                        chuadong72 = typy7.notclose_medium + "",
                        chuadong82 = typy8.notclose_medium + "",
                        dadong11 = typy1.close_high + "",
                        dadong21 = typy2.close_high + "",
                        dadong31 = typy3.close_high + "",
                        dadong41 = typy4.close_high + "",
                        dadong51 = typy5.close_high + "",
                        dadong61 = typy6.close_high + "",
                        dadong71 = typy7.close_high + "",
                        dadong81 = typy8.close_high + "",
                        dadong12 = typy1.close_medium + "",
                        dadong22 = typy2.close_medium + "",
                        dadong32 = typy3.close_medium + "",
                        dadong42 = typy4.close_medium + "",
                        dadong52 = typy5.close_medium + "",
                        dadong62 = typy6.close_medium + "",
                        dadong72 = typy7.close_medium + "",
                        dadong82 = typy8.close_medium + "",
                        cuoiky11 = typy1.ending_high + "",
                        cuoiky21 = typy2.ending_high + "",
                        cuoiky31 = typy3.ending_high + "",
                        cuoiky41 = typy4.ending_high + "",
                        cuoiky51 = typy5.ending_high + "",
                        cuoiky61 = typy6.ending_high + "",
                        cuoiky71 = typy7.ending_high + "",
                        cuoiky81 = typy8.ending_high + "",
                        cuoiky12 = typy1.ending_medium + "",
                        cuoiky22 = typy2.ending_medium + "",
                        cuoiky32 = typy3.ending_medium + "",
                        cuoiky42 = typy4.ending_medium + "",
                        cuoiky52 = typy5.ending_medium + "",
                        cuoiky62 = typy6.ending_medium + "",
                        cuoiky72 = typy7.ending_medium + "",
                        cuoiky82 = typy8.ending_medium + "",
                    };
                    TextSelection selectionss = doc.FindAllString("congtacquantrichatluonghoatdong", false, true)?.FirstOrDefault();
                    if (selectionss != null)
                    {
                        if (!string.IsNullOrEmpty(report.quality))
                        {
                            var paragraph = selectionss?.GetAsOneRange()?.OwnerParagraph;
                            var object_arr = paragraph.ChildObjects;
                            if (object_arr.Count > 0)
                            {
                                for (int j = 0; j < object_arr.Count; j++)
                                {
                                    paragraph.ChildObjects.RemoveAt(j);
                                }
                                paragraph.AppendHTML(report.quality);
                            }
                        }
                    }
                    Table table1 = doc.Sections[0].Tables[2] as Table;
                    var ii = 1;
                    foreach (var item in report2data)
                    {
                        var row = table1.AddRow();
                        var txtTable1 = row.Cells[0].AddParagraph().AppendText(ii + "");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[1].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[2].AddParagraph().AppendText(item.audit_time);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[3].AddParagraph().AppendText(item.report_date);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                        txtTable1 = row.Cells[4].AddParagraph().AppendText(item.level);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                        txtTable1 = row.Cells[5].AddParagraph().AppendText(item.risk_high);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                        txtTable1 = row.Cells[6].AddParagraph().AppendText(item.risk_medium);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        ii++;
                    }
                    Table table2 = doc.Sections[0].Tables[3] as Table;
                    foreach (var item in report3data)
                    {
                        var row = table2.AddRow();
                        var txtTable1 = row.Cells[0].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[1].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[2].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;

                        txtTable1 = row.Cells[3].AddParagraph().AppendText(item.audit_name);
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 11;
                    }

                    doc.MailMerge.Execute(
                    new string[] { "year" ,"year_1" ,"year_2" , "yearreport" , "danhgiatongquanktnb" ,
                        "quanngaichinhktnb",
                        "reason",
                        "note",
                        "congtactheodoikhacphuc",
                        "congtacquantrichatluonghoatdong",
                        "company",
                        "company_1",
                        "auditplancurrent",
                        "auditcompletedcurrent",
                        "completedcurrent",
                        "auditexpectedcurrent",
                        "auditplanprevious",
                        "auditcompletedprevious",
                        "completedprevious",
                        "auditexpectedprevious",
                        "auditplanvolatility",
                        "auditcompletedvolatility",
                        "dauky11",
                        "dauky21",
                        "dauky31",
                        "dauky41",
                        "dauky51",
                        "dauky61",
                        "dauky71",
                        "dauky81",
                        "dauky12",
                        "dauky22",
                        "dauky32",
                        "dauky42",
                        "dauky52",
                        "dauky62",
                        "dauky72",
                        "dauky82",
                        "chuadong11",
                        "chuadong21",
                        "chuadong31",
                        "chuadong41",
                        "chuadong51",
                        "chuadong61",
                        "chuadong71",
                        "chuadong81",
                        "chuadong12",
                        "chuadong22",
                        "chuadong32",
                        "chuadong42",
                        "chuadong52",
                        "chuadong62",
                        "chuadong72",
                        "chuadong82",
                        "dadong11",
                        "dadong21",
                        "dadong31",
                        "dadong41",
                        "dadong51",
                        "dadong61",
                        "dadong71",
                        "dadong81",
                        "dadong12",
                        "dadong22",
                        "dadong32",
                        "dadong42",
                        "dadong52",
                        "dadong62",
                        "dadong72",
                        "dadong82",
                        "cuoiky11",
                        "cuoiky21",
                        "cuoiky31",
                        "cuoiky41",
                        "cuoiky51",
                        "cuoiky61",
                        "cuoiky71",
                        "cuoiky81",
                        "cuoiky12",
                        "cuoiky22",
                        "cuoiky32",
                        "cuoiky42",
                        "cuoiky52",
                        "cuoiky62",
                        "cuoiky72",
                        "cuoiky82"
                    },
                    new string[] { tt.year ,tt.year,tt.year, tt.yearreport ,tt.danhgiatongquanktnb , tt.quanngaichinhktnb ,tt.reason,
                       tt.note ,tt.congtactheodoikhacphuc,tt.congtacquantrichatluonghoatdong,tt.company,tt.company,
                       table2field.auditplancurrent,table2field.auditcompletedcurrent,table2field.completedcurrent,table2field.auditexpectedcurrent,table2field.auditplanprevious,
                       table2field.auditcompletedprevious,table2field.completedprevious,table2field.auditexpectedprevious,table2field.auditplanvolatility,table2field.auditcompletedvolatility, table5field.dauky11,table5field.dauky21,table5field.dauky31,table5field.dauky41,table5field.dauky51,table5field.dauky61,table5field.dauky71,table5field.dauky81,table5field.dauky12,table5field.dauky22,table5field.dauky32,table5field.dauky42,table5field.dauky52,table5field.dauky62,table5field.dauky72,table5field.dauky82,         table5field.chuadong11,table5field.chuadong21,table5field.chuadong31,table5field.chuadong41,table5field.chuadong51,table5field.chuadong61,table5field.chuadong71,table5field.chuadong81,table5field.chuadong12,table5field.chuadong22,table5field.chuadong32,table5field.chuadong42,table5field.chuadong52,table5field.chuadong62,table5field.chuadong72,table5field.chuadong82,        table5field.dadong11,table5field.dadong21,table5field.dadong31,table5field.dadong41,table5field.dadong51,table5field.dadong61,table5field.dadong71,table5field.dadong81,table5field.dadong12,table5field.dadong22,table5field.dadong32,table5field.dadong42,table5field.dadong52,table5field.dadong62,table5field.dadong72,table5field.dadong82,               table5field.cuoiky11,table5field.cuoiky21,table5field.cuoiky31,table5field.cuoiky41,table5field.cuoiky51,table5field.cuoiky61,table5field.cuoiky71,table5field.cuoiky81,table5field.cuoiky12,table5field.cuoiky22,table5field.cuoiky32,table5field.cuoiky42,table5field.cuoiky52,table5field.cuoiky62,table5field.cuoiky72,table5field.cuoiky82
                        }
                    );

                    //foreach (Section section1 in doc.Sections)
                    //{
                    //    for (int i = 0; i < section1.Body.ChildObjects.Count; i++)
                    //    {
                    //        if (section1.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
                    //        {
                    //            if (String.IsNullOrEmpty((section1.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                    //            {
                    //                section1.Body.ChildObjects.Remove(section1.Body.ChildObjects[i]);
                    //                i--;
                    //            }
                    //        }

                    //    }
                    //}

                    MemoryStream stream = new MemoryStream();
                    stream.Position = 0;
                    doc.SaveToFile(stream, FileFormat.Docx);
                    byte[] bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Grid.docx");
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                return Ok(new { code = "0", msg = "fail", data = new StatusApprove() });
            }
        }
        private static void ReplaceKey(XWPFParagraph para, object model)
        {
            string text = para.ParagraphText;
            var runs = para.Runs;
            string styleid = para.Style;
            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                text = run.ToString();
                Type t = model.GetType();
                PropertyInfo[] pi = t.GetProperties();
                foreach (PropertyInfo p in pi)
                {
                    //$$ corresponds to $$ in the template, and can also be changed to other symbols, such as {$name}, must be unique
                    if (text.Contains("[[" + p.Name + "]]"))
                    {
                        text = text.Replace("[[" + p.Name + "]]", p.GetValue(model, null).ToString());
                    }
                }
                runs[i].SetText(text, 0);
            }
        }

        [HttpPost("ExportWordNewMCREDIT/{id}")]
        public IActionResult ExportWordNewMCREDIT(int id)
        {

            var _paramInfoPrefix = "Param.SystemInfo";
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var iCache = (IConnectionMultiplexer)HttpContext.RequestServices
                      .GetService(typeof(IConnectionMultiplexer));
                if (!iCache.IsConnected)
                {
                    return BadRequest();
                }
                var redisDb = iCache.GetDatabase();
                var value_get = redisDb.StringGet(_paramInfoPrefix);
                var company = "";
                if (value_get.HasValue)
                {
                    var list_param = JsonSerializer.Deserialize<List<SystemParam>>(value_get);
                    company = list_param.FirstOrDefault(a => a.name == "COMPANY_NAME")?.value;
                }
                var report = _uow.Repository<ReportAuditWorkYear>().FirstOrDefault(a => a.Id == id);
                if (report != null)
                {
                    var current_year = 0;
                    var previous_year = 0;
                    var now = DateTime.Now.Date;
                    current_year = report.year.Value;
                    previous_year = report.year.Value - 1;
                    var lastDayOfPrevious_year = new DateTime(previous_year, 12, 31);
                    var lastDayOfCurrent_year = new DateTime(current_year, 12, 31);
                    var hearderSystem = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "REPORT_HEADER");

                    var getreport = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork, p => p.AuditWork.AuditDetect, p => p.AuditWork.MainStage)
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW") on a.Id equals b.item_id
                                     where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString())
                                     select a).ToArray();

                    var perpareAuditPlan = (from a in _uow.Repository<AuditWork>().Include()
                                            join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_PAP") on a.Id equals b.item_id
                                            where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && (a.Year == current_year.ToString())
                                            select a).ToArray();

                    var auditRequestMonitor = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect, p => p.CatAuditRequest, p => p.FacilityRequestMonitorMapping)
                                               join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD" && p.StatusCode == "3.1") on a.AuditDetect.id equals b.item_id.Value
                                               let checkstatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false)))
                                                && (a.is_deleted != true)
                                               select a).OrderByDescending(p => p.created_at).ToList();

                    var previousYearsAuditRequestMonitor = new List<AuditRequestMonitor>();
                    var currentYearAuditRequestMonitor = new List<AuditRequestMonitor>();


                    var currentYearsAuditDetect = (from a in _uow.Repository<AuditDetect>().Include(x => x.AuditRequestMonitor, x => x.AuditWork.AuditPlan)
                                                   join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.id equals b.item_id
                                                   where b.StatusCode == "3.1" && !(a.IsDeleted ?? false) && a.year == current_year && a.audit_report
                                                   orderby a.auditprocess_name
                                                   select a).ToArray();

                    var closedCurrentYearsAuditRequestMonitor = new List<AuditRequestMonitor>();

                    auditRequestMonitor.ForEach(o =>
                    {
                        if (o.Code.Split(".")[1] == current_year.ToString())
                        {
                            currentYearAuditRequestMonitor.Add(o);
                            if (o.Conclusion == 2)
                            {
                                closedCurrentYearsAuditRequestMonitor.Add(o);
                            }
                        }
                        else if (o.Conclusion != 2 && int.TryParse(o.Code.Split(".")[1], out int result) ? int.Parse(o.Code.Split(".")[1]) < current_year : false)
                        {
                            previousYearsAuditRequestMonitor.Add(o);
                        }
                    }
                    );



                    var previousYearsARM = previousYearsAuditRequestMonitor.Count();
                    var currentYearARM = currentYearAuditRequestMonitor.Count();
                    var closedCurrentYearARM = closedCurrentYearsAuditRequestMonitor.Count();
                    var remainingARM = previousYearsARM + currentYearARM - closedCurrentYearARM;

                    var _config = (IConfiguration)HttpContext.RequestServices
                    .GetService(typeof(IConfiguration));
                    var path = _config["Template:ReporDocsTemplate"];
                    var fullPath = Path.Combine(path, "Mcredit_Kitano_BaoCaoKiemToanNam_v0.1.docx");
                    fullPath = fullPath.ToString().Replace("\\", "/");
                    Document doc = new Document(fullPath);

                    doc.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true;
                    Paragraph paragraph_header = doc.Sections[0].HeadersFooters.FirstPageHeader.AddParagraph();
                    paragraph_header.AppendHTML(hearderSystem.Value);

                    //Remove header 2
                    doc.Sections[1].HeadersFooters.Header.ChildObjects.Clear();
                    //Add line breaks
                    Paragraph headerParagraph = doc.Sections[1].HeadersFooters.FirstPageHeader.AddParagraph();
                    headerParagraph.AppendBreak(Spire.Doc.Documents.BreakType.LineBreak);

                    Table table1 = doc.Sections[0].Tables[1] as Table;
                    foreach (var (item, i) in getreport.Select((item, i) => (item, i)))
                    {
                        var schedule = item.AuditWork.MainStage.Where(x => x.index == 5).FirstOrDefault().actual_date;
                        var row = table1.AddRow();
                        var paraTable1 = row.Cells[0].AddParagraph();
                        var txtTable1 = paraTable1.AppendText(i + 1 + ".");
                        paraTable1.ApplyStyle(BuiltinStyle.Subtitle);
                        paraTable1.Format.HorizontalAlignment = HorizontalAlignment.Center;

                        paraTable1 = row.Cells[1].AddParagraph();
                        txtTable1 = paraTable1.AppendText(item.AuditWorkName);
                        paraTable1.ApplyStyle(BuiltinStyle.Subtitle);

                        paraTable1 = row.Cells[2].AddParagraph();
                        txtTable1 = paraTable1.AppendText("Đã hoàn thành");
                        paraTable1.ApplyStyle(BuiltinStyle.Subtitle);
                        paraTable1.Format.HorizontalAlignment = HorizontalAlignment.Center;

                        paraTable1 = row.Cells[3].AddParagraph();
                        txtTable1 = paraTable1.AppendText(schedule != null ? schedule.Value.ToString("dd/MM/yyyy") : "");
                        paraTable1.ApplyStyle(BuiltinStyle.Subtitle);
                        paraTable1.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    }

                    Table table2 = doc.Sections[1].Tables[0] as Table;
                    foreach (var (item, i) in currentYearsAuditDetect.Select((item, i) => (item, i)))
                    {
                        var row = table2.AddRow();
                        var paraTable1 = row.Cells[0].AddParagraph();
                        var txtTable1 = paraTable1.AppendText(i + 1 + ".");
                        paraTable1.ApplyStyle(BuiltinStyle.Subtitle);
                        paraTable1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        paraTable1.Format.LineSpacing = 14.4360902256f;

                        paraTable1 = row.Cells[1].AddParagraph();
                        txtTable1 = paraTable1.AppendText(item.auditprocess_name);
                        paraTable1.ApplyStyle(BuiltinStyle.Subtitle);
                        paraTable1.Format.LineSpacing = 14.4360902256f;

                        paraTable1 = row.Cells[2].AddParagraph();
                        txtTable1 = paraTable1.AppendText(item.title);
                        paraTable1.ApplyStyle(BuiltinStyle.Subtitle);
                        paraTable1.Format.LineSpacing = 14.4360902256f;

                        paraTable1 = row.Cells[3].AddParagraph();
                        foreach (var (requestMonitor, j) in item.AuditRequestMonitor.Select((requestMonitor, j) => (requestMonitor, j)))
                        {
                            txtTable1 = paraTable1.AppendText("- " + requestMonitor.Content);
                            paraTable1.ApplyStyle(BuiltinStyle.Subtitle);
                            if (j + 1 < item.AuditRequestMonitor.Count())
                            {
                                txtTable1 = paraTable1.AppendText("\n");
                            }
                            paraTable1.Format.LineSpacing = 14.4360902256f;

                        }

                    }


                    table2.Rows.RemoveAt(1);
                    table1.Rows.RemoveAt(1);

                    doc.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);
                    //remove empty paragraphs
                    doc.MailMerge.HideEmptyParagraphs = true;
                    //remove empty group
                    doc.MailMerge.HideEmptyGroup = true;
                    doc.MailMerge.Execute(
                    new string[] { "year" ,"ngay_bc_1", "thang_bc_1", "nam_bc_1",
                        "audit_total",
                        "audit_done",
                        "orther_info",
                        "pre_year_last_day",
                        "cur_year_last_day",
                        "pre_year_arm",
                        "cur_year_arm",
                        "closed_cur_year_arm",
                        "remaining_arm",
                    },
                    new string[] { current_year + "", DateTime.Now.Day.ToString() , DateTime.Now.Month.ToString() , DateTime.Now.Year.ToString(),
                        perpareAuditPlan.Count().ToString(),
                        getreport.Count().ToString(),
                        report.quality,
                        lastDayOfPrevious_year.ToString("dd/MM/yyyy"),
                        lastDayOfCurrent_year.ToString("dd/MM/yyyy"),
                        previousYearsARM.ToString(),
                        currentYearARM.ToString(),
                        closedCurrentYearARM.ToString(),
                        remainingARM.ToString(),
                    }
                    );

                    MemoryStream stream = new MemoryStream();
                    stream.Position = 0;
                    doc.SaveToFile(stream, FileFormat.Docx);
                    byte[] bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Grid.docx");
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                return Ok(new { code = "0", msg = "fail", data = new StatusApprove() });
            }

        }



        /// <summary>
        /// Hàm này dùng để giải quyết dữ liệu html cho vào docs, vì MailMerge không hỗ trợ html như hàm AppendHtml()
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            if (args.FieldName == "orther_info")
            {
                args.CurrentMergeField.OwnerParagraph.AppendHTML(args.FieldValue != null ? args.FieldValue.ToString() : "");
                args.Text = "";
                args.CurrentMergeField.OwnerParagraph.Format.LeftIndent = 0;
                args.CurrentMergeField.OwnerParagraph.Format.BeforeSpacing = 0;
                args.CurrentMergeField.OwnerParagraph.Format.AfterSpacing = 0;
                args.CurrentMergeField.OwnerParagraph.Format.BeforeAutoSpacing = false;
                args.CurrentMergeField.OwnerParagraph.Format.AfterAutoSpacing = false;
            }
        }
    }
}
