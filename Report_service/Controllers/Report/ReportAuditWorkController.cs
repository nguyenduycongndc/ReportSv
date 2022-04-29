using Report_service.DataAccess;
using Report_service.Models.ExecuteModels;
using Report_service.Models.ExecuteModels.Audit;
using Report_service.Models.MigrationsModels;
using Report_service.Repositories;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Authorization;
using Report_service.Models.MigrationsModels.Category;
using Spire.Doc;
using Spire.Pdf;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Web;
using Audit_service.Models.MigrationsModels;

namespace Report_service.Controllers.Report
{
    [Route("[controller]")]
    [ApiController]
    public class ReportAuditWorkController : BaseController
    {
        protected readonly IConfiguration _config;
        public ReportAuditWorkController(ILoggerManager logger, IUnitOfWork uow, IConfiguration config) : base(logger, uow)
        {
            _config = config;
        }

        [HttpGet("Search")]
        public IActionResult Search(string jsonData)
        {
            try
            {
                var obj = JsonSerializer.Deserialize<ReportAuditWorkSearchModel>(jsonData);
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var status = obj.Status;
                var approval_status = _uow.Repository<ApprovalFunction>().Find(a => a.function_code == "M_RAW" && (string.IsNullOrEmpty(status) || a.StatusCode == status)).ToArray();
                var list_appoval_id = approval_status.Select(a => a.item_id).ToList();
                var auditplan_file = _uow.Repository<ApprovalFunctionFile>().Find(a => a.function_code == "M_RAW" && a.IsDeleted != true).ToArray();
                var report = _uow.Repository<ReportAuditWork>().Include(a => a.AuditWork).Where(a => (string.IsNullOrEmpty(obj.Year) || a.Year == obj.Year.Trim())
                                              && (string.IsNullOrEmpty(status) || list_appoval_id.Contains(a.Id) || (status == "1.0" && !list_appoval_id.Contains(a.Id)))
                                              && (string.IsNullOrEmpty(obj.Code) || a.AuditWorkCode.ToLower().Contains(obj.Code.ToLower().Trim()))
                                              && (string.IsNullOrEmpty(obj.Name) || a.AuditWorkName.ToLower().Contains(obj.Name.ToLower().Trim()))
                                              && (!obj.PersonInCharge.HasValue || (a.AuditWork.person_in_charge.HasValue && a.AuditWork.person_in_charge == obj.PersonInCharge))
                                              && a.IsDeleted != true).ToArray();
                IEnumerable<ReportAuditWork> data = report;
                var count = data.Count();
                if (obj.StartNumber >= 0 && obj.PageSize > 0)
                {
                    data = data.Skip(obj.StartNumber).Take(obj.PageSize);
                }
                var MainStage = _uow.Repository<MainStage>().GetAll().ToArray();
                var startdate = MainStage.FirstOrDefault(a => a.index == 1);
                var enddate = MainStage.FirstOrDefault(a => a.index == 4);
                var user = _uow.Repository<Users>().Find(a => a.IsDeleted != true && a.IsActive == true).ToArray();
                var result = data.Select(a => new ReportAuditWorkListModel()
                {
                    Id = a.Id,
                    Year = a.Year,
                    Name = a.AuditWorkName,
                    Code = a.AuditWorkCode,
                    StartDate = MainStage.FirstOrDefault(x => x.index == 1 && x.auditwork_id == a.auditwork_id)?.actual_date,
                    EndDate = MainStage.FirstOrDefault(x => x.index == 4 && x.auditwork_id == a.auditwork_id)?.actual_date,
                    str_person_in_charge = user.FirstOrDefault(x => a.AuditWork.person_in_charge == x.Id)?.FullName,
                    Status = approval_status.FirstOrDefault(x => x.item_id == a.Id)?.StatusCode ?? "1.0",
                    approval_user = approval_status.FirstOrDefault(x => x.item_id == a.Id)?.approver,
                    approval_user_last = approval_status.FirstOrDefault(x => x.item_id == a.Id)?.approver_last,
                    Path = auditplan_file.FirstOrDefault(x => x.item_id == a.Id)?.Path,
                    FileId = auditplan_file.FirstOrDefault(x => x.item_id == a.Id)?.id,
                });
                return Ok(new { code = "1", msg = "success", data = result, total = count });
            }
            catch (Exception ex)
            {
                return BadRequest();
            }
        }
        [HttpPost]
        public IActionResult Create([FromBody] ReportAuditWorkCreateModel reportauditworkinfo)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel _userInfo)
                {
                    return Unauthorized();
                }
                var _allcheck = _uow.Repository<ReportAuditWork>().Find(a => a.IsDeleted != true && a.auditwork_id == reportauditworkinfo.AuditWorkId).ToArray();
                if (_allcheck.Length > 0)
                {
                    return Ok(new { code = "0", msg = "success" });
                }
                var auditwork = _uow.Repository<AuditWork>().FirstOrDefault(a => a.Id == reportauditworkinfo.AuditWorkId);
                if (auditwork != null)
                {
                    var report = new ReportAuditWork();
                    report.Year = reportauditworkinfo.Year;
                    report.Status = reportauditworkinfo.Status;
                    report.auditwork_id = auditwork.Id;
                    report.AuditWorkName = auditwork.Name;
                    report.AuditWorkCode = auditwork.Code;
                    report.IsActive = true;
                    report.IsDeleted = false;
                    report.CreatedAt = DateTime.Now;
                    report.CreatedBy = _userInfo.Id;
                    _uow.Repository<ReportAuditWork>().Add(report);
                    return Ok(new { code = "1", data = report.Id, msg = "success" });
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
        public IActionResult Edit([FromBody] ReportAuditWorkModifyModel reportauditworkinfo)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel _userInfo)
                {
                    return Unauthorized();
                }
                var report = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == reportauditworkinfo.Id);
                if (report == null)
                {
                    return NotFound();
                }
                //report.NumOfWorkdays = reportauditworkinfo.NumOfWorkdays;
                report.AuditRatingLevelTotal = reportauditworkinfo.AuditRatingLevelTotal;
                report.BaseRatingTotal = reportauditworkinfo.BaseRatingTotal;
                report.GeneralConclusions = reportauditworkinfo.GeneralConclusions;
                report.OtherContent = reportauditworkinfo.OtherContent;
                //report.StartDate = reportauditworkinfo.StartDate;
                //report.EndDate = reportauditworkinfo.EndDate;
                //report.StartDateField = reportauditworkinfo.StartDateField;
                //report.EndDateField = reportauditworkinfo.EndDateField;
                //report.ReportDate = reportauditworkinfo.ReportDate;
                //if (!string.IsNullOrEmpty(reportauditworkinfo.StartDate))
                //    report.StartDate = DateTime.ParseExact(reportauditworkinfo.StartDate, "dd/MM/yyyy", null);
                //if (!string.IsNullOrEmpty(reportauditworkinfo.EndDate))
                //    report.EndDate = DateTime.ParseExact(reportauditworkinfo.EndDate, "dd/MM/yyyy", null);
                //if (!string.IsNullOrEmpty(reportauditworkinfo.StartDateField))
                //    report.StartDateField = DateTime.ParseExact(reportauditworkinfo.StartDateField, "dd/MM/yyyy", null);
                //if (!string.IsNullOrEmpty(reportauditworkinfo.EndDateField))
                //    report.EndDateField = DateTime.ParseExact(reportauditworkinfo.EndDateField, "dd/MM/yyyy", null);
                //if (!string.IsNullOrEmpty(reportauditworkinfo.ReportDate))
                //    report.ReportDate = DateTime.ParseExact(reportauditworkinfo.ReportDate, "dd/MM/yyyy", null);
                report.ModifiedAt = DateTime.Now;
                report.ModifiedBy = _userInfo.Id;
                var lstFacilityScope = new List<AuditWorkScopeFacility>();
                if (reportauditworkinfo.LisFacilityScope.Count > 0)
                {
                    var listScopeID = reportauditworkinfo.LisFacilityScope.Select(a => a.ScopeId).ToArray();
                    var auditscope = _uow.Repository<AuditWorkScopeFacility>().Find(a => listScopeID.Contains(a.Id)).ToArray();
                    foreach (var item in reportauditworkinfo.LisFacilityScope)
                    {
                        var auditscopeitem = auditscope.FirstOrDefault(a => a.Id == item.ScopeId);
                        if (auditscopeitem != null)
                        {
                            auditscopeitem.AuditRatingLevelReport = item.AuditRatingLevelItem;
                            auditscopeitem.BaseRatingReport = item.BaseRatingItem;
                            lstFacilityScope.Add(auditscopeitem);
                        }
                    }
                }
                foreach (var item in lstFacilityScope)
                {
                    _uow.Repository<AuditWorkScopeFacility>().UpdateWithoutSave(item);
                }
                _uow.Repository<ReportAuditWork>().UpdateWithoutSave(report);
                _uow.SaveChanges();
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
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }

                var report = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == id);

                if (report != null)
                {
                    var approval_status = _uow.Repository<ApprovalFunction>().FirstOrDefault(a => a.item_id == id && a.function_code == "M_RAW");
                    var auditwork = _uow.Repository<AuditWork>().Include(a => a.AuditWorkScope, a => a.Users).FirstOrDefault(a => a.Id == report.auditwork_id);
                    //var auditdetect = (from a in _uow.Repository<AuditDetect>().Find(x => x.IsDeleted != true && x.auditwork_id == report.auditwork_id && x.audit_report == true)
                    //                   join b in _uow.Repository<ApprovalFunction>().Find(x => x.function_code == "M_AD" && x.StatusCode == "3.1") on a.id equals b.item_id
                    //                   select a).ToArray();
                    var auditdetect = _uow.Repository<AuditDetect>().Find(x => x.IsDeleted != true && x.auditwork_id == report.auditwork_id && x.audit_report == true).ToArray();
                    //var auditdetect = _uow.Repository<auditdetect>().Find(x => x.IsDeleted != true && x.auditwork_id == report.auditwork_id && x.status == 3).ToArray();
                    var detecttype = _uow.Repository<CatDetectType>().Find(x => x.IsDeleted != true).ToArray();
                    var auditassign = _uow.Repository<AuditAssignment>().Include(a => a.Users).Where(a => a.auditwork_id == auditwork.Id && a.user_id.HasValue && a.IsDeleted != true).Select(a => a.Users.FullName).ToArray();
                    var auditworkscope = _uow.Repository<AuditWorkScope>().Find(a => a.auditwork_id == auditwork.Id && a.IsDeleted != true).AsEnumerable().GroupBy(a => new { a.auditprocess_id, a.auditfacilities_id, a.bussinessactivities_id }).Select(g => g.FirstOrDefault()).ToArray();

                    var lstauditworkscope = auditworkscope.Where(a => a.auditprocess_id.HasValue).Select(a => new AuditWorkScopeModel()
                    {
                        ID = a.Id,
                        AuditProcess = a.auditprocess_name,
                        AuditFacility = a.auditfacilities_name,
                        AuditActivity = a.bussinessactivities_name,
                    }).ToList();

                    var listfacility = _uow.Repository<AuditWorkScopeFacility>().Find(a => a.auditfacilities_id.HasValue && a.IsDeleted != true && a.auditwork_id == report.auditwork_id).ToArray();

                    var listdetect = auditdetect.Where(a => a.classify_audit_detect.HasValue).AsEnumerable().GroupBy(a => a.classify_audit_detect).Select(g => g.FirstOrDefault());
                    var auditminutes = _uow.Repository<AuditMinutes>().Find(a => a.auditfacilities_id.HasValue && a.IsDeleted != true && a.auditwork_id == report.auditwork_id).ToArray();
                    var lstAuditworkfacility = listfacility.Select(a => new AuditWorkFacilityModel()
                    {
                        ID = a.Id,
                        AuditFacilityId = a.auditfacilities_id,
                        AuditFacility = a.auditfacilities_name,
                        AuditRatingLevelItem = auditminutes.FirstOrDefault(x => x.auditfacilities_id == a.auditfacilities_id)?.rating_level_total,
                        BaseRatingItem = a.BaseRatingReport,

                    }).ToList();

                    var lstDetectType = listdetect.Select(a => new SumaryDetectModel()
                    {
                        classify_audit_detect = a.classify_audit_detect,
                        classify_audit_detect_name = detecttype.Where(x => x.Id == a.classify_audit_detect).Select(c => c.Name).FirstOrDefault().ToString(),
                        Hight = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 1),
                        Middle = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 2),
                        Low = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 3),
                        Agree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is true),
                        NotAgree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is false),
                    }).ToList();

                    var str_name = auditassign.Length > 0 ? string.Join(",", auditassign) : "";
                    var result = new ReportAuditWorkDetailModel()
                    {
                        Id = report.Id,
                        Year = report.Year,
                        Code = report.AuditWorkCode,
                        Name = report.AuditWorkName,
                        Target = auditwork?.Target,
                        Status = approval_status?.StatusCode ?? "1.0",
                        NumOfAuditor = auditwork?.NumOfAuditor,
                        AuditScope = auditwork?.AuditScope,
                        str_person_in_charge = auditwork.person_in_charge.HasValue ? auditwork.Users.FullName : "",
                        StrAuditorName = str_name,
                        AuditWorkScopeList = lstauditworkscope,
                        OutOfAuditScope = auditwork?.AuditScopeOutside,
                        AuditWorkFacilityList = lstAuditworkfacility,
                        StartDatePlaning = auditwork?.from_date,
                        EndDatePlaning = auditwork?.to_date,
                        StartDateField = report.StartDateField,
                        EndDateField = report.EndDateField,
                        ReportDate = report.ReportDate,
                        NumOfWorkdays = report.NumOfWorkdays,
                        Classify = auditwork?.Classify,
                        RatingLevelAudit = report.AuditRatingLevelTotal,
                        RatingBaseAudit = report.BaseRatingTotal,
                        GeneralConclusions = report.GeneralConclusions,
                        OtherContent = report.OtherContent,
                        SumaryFacilityList = lstDetectType,
                        AuditDetectList = auditdetect.Select(a => new AuditDetectInfoModel()
                        {
                            ID = a.id,
                            code = a.code,
                            title = a.title,
                            Description = a.description,
                            classify_audit_detect_name = lstDetectType.Select(c => c.classify_audit_detect_name).FirstOrDefault().ToString(),
                            rating_risk = a.rating_risk,
                            str_classify_audit_detect = a.CatDetectType.Name,
                            opinion_audit = a.opinion_audit,
                        }).ToList(),

                    };
                    return Ok(new { code = "1", msg = "success", data = result });
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
        [HttpDelete("{id}")]
        public IActionResult DeleteConfirmed(int id)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var _report = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == id);
                if (_report == null)
                {
                    return NotFound();
                }
                _report.IsDeleted = true;
                _report.DeletedAt = DateTime.Now;
                _report.DeletedBy = userInfo.Id;
                _uow.Repository<ReportAuditWork>().Update(_report);
                return Ok(new { code = "1", msg = "success" });
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }
        [HttpPost("RequestApproval")]
        public IActionResult RequestApproval(ApprovalModel model)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var checkReport = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == model.ReportAuditWorkid);
                if (checkReport == null)
                {
                    return NotFound();
                }
                if (checkReport.Status == 1 || checkReport.Status == 4)
                {
                    checkReport.Status = 2;
                    checkReport.approval_user = model.approvaluser;
                    _uow.Repository<ReportAuditWork>().Update(checkReport);
                    return Ok(new { code = "1", msg = "success" });
                }
                else
                {
                    return BadRequest();
                }

            }
            catch (Exception)
            {
                return BadRequest();
            }
        }
        [HttpPost("SubmitApproval/{id}")]
        public IActionResult SubmitApproval(int id)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var checkReport = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == id);
                if (checkReport == null)
                {
                    return NotFound();
                }
                checkReport.Status = 3;
                checkReport.ApprovalDate = DateTime.Now;
                _uow.Repository<ReportAuditWork>().Update(checkReport);
                return Ok(new { code = "1", msg = "success" });
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }

        [HttpPost("RejectApproval")]
        public IActionResult RejectApproval(RejectApprovalModel model)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var checkReport = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == model.id);
                if (checkReport == null)
                {
                    return NotFound();
                }
                checkReport.Status = 4;
                checkReport.ReasonReject = model.reason_note;
                checkReport.ApprovalDate = DateTime.Now;
                _uow.Repository<ReportAuditWork>().Update(checkReport);
                return Ok(new { code = "1", msg = "success" });
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }
        [HttpPost("ReportpdateStatus")]
        public IActionResult updateStatus()
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel _userInfo)
                {
                    return Unauthorized();
                }
                var _changeStatus = new ReportUpdateStatusModel();
                var data = Request.Form["data"];
                var pathSave = "";
                var file_type = "";
                if (!string.IsNullOrEmpty(data))
                {
                    _changeStatus = JsonSerializer.Deserialize<ReportUpdateStatusModel>(data);
                    var file = Request.Form.Files;
                    file_type = file.FirstOrDefault()?.ContentType;
                    pathSave = CreateUploadFile(file, "ReportAuditWork");
                }
                else
                {
                    return BadRequest();
                }
                var updateStatus = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == _changeStatus.Id && a.IsDeleted != true);
                if (updateStatus == null)
                {
                    return NotFound();
                }

                updateStatus.Status = _changeStatus.Status;

                if (!string.IsNullOrEmpty(_changeStatus.Browsedate))
                {
                    updateStatus.ApprovalDate = DateTime.ParseExact(_changeStatus.Browsedate, "yyyy-MM-dd", null);
                }
                //updateStatus.Browsedate = _changeStatus.Browsedate;
                updateStatus.Path = pathSave;
                updateStatus.FileType = file_type;
                _uow.Repository<ReportAuditWork>().Update(updateStatus);
                return Ok(new { code = "1", msg = "success" });
            }
            catch (Exception ex)
            {
                return BadRequest();
            }
        }
        protected string CreateUploadFile(IFormFileCollection list_ImageFile, string folder = "")
        {
            var lstStr = new List<string>();
            foreach (var item in list_ImageFile)
            {
                string imgURL = CreateUploadURL(item, folder);
                if (!string.IsNullOrEmpty(imgURL)) lstStr.Add(imgURL);
            }
            return string.Join(",", lstStr);
        }
        protected string CreateUploadURL(IFormFile imageFile, string folder = "")
        {
            var pathSave = "";
            var pathconfig = _config["Upload:ReporDocsPath"];
            if (imageFile != null)
            {
                if (string.IsNullOrEmpty(folder)) folder = "Public";
                var pathToSave = Path.Combine(pathconfig, folder);
                var fileName = Path.GetFileName(imageFile.FileName)?.Trim();
                var new_folder = DateTime.Now.ToString("yyyyMM");
                var fullPathroot = Path.Combine(pathToSave, new_folder);
                if (!Directory.Exists(fullPathroot))
                {
                    Directory.CreateDirectory(fullPathroot);
                }
                pathSave = Path.Combine(folder, new_folder, fileName);
                var filePath = Path.Combine(fullPathroot, fileName);
                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    imageFile.CopyTo(fileStream);
                }
            }
            return pathSave;
        }
        [AllowAnonymous]
        [HttpGet("DownloadAttach")]
        public IActionResult DownloadFile(int id)
        {
            try
            {
                var report = _uow.Repository<ReportAuditWork>().FirstOrDefault(o => o.Id == id);
                if (report == null)
                {
                    return NotFound();
                }
                var self = _uow.Repository<ApprovalFunctionFile>().FirstOrDefault(a => a.item_id == report.Id && a.function_code == "M_RAW" && a.IsDeleted != true);
                if (self == null)
                {
                    return NotFound();
                }
                var fullPath = Path.Combine(_config["Upload:ReporDocsPath"], self.Path);
                var name = "DownLoadFile";
                if (!string.IsNullOrEmpty(self.Path))
                {
                    var _array = self.Path.Replace("/", "\\").Split("\\");
                    name = _array[_array.Length - 1];
                }
                var fs = new FileStream(fullPath, FileMode.Open);

                return File(fs, self.FileType, name);
            }
            catch (Exception)
            {
                return NotFound();
            }
        }

        //xuất word
        [HttpGet("ExportFileWordMIC/{id}")]
        public IActionResult ExportFileWordMIC(int id)
        {
            byte[] Bytes = null;
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var checkReport = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == id && a.IsDeleted.Equals(false));

                var auditWorkScope = _uow.Repository<AuditWorkScope>().GetAll().Where(a => a.auditwork_id == checkReport.auditwork_id && a.IsDeleted != true).ToArray();

                var lstauditworkscope = auditWorkScope.Where(a => a.auditprocess_id.HasValue).Select(a => new AuditWorkScopeModel()
                {
                    ID = a.Id,
                    AuditProcess = a.auditprocess_name,
                    AuditFacility = a.auditfacilities_name,
                    AuditActivity = a.bussinessactivities_name,
                }).ToList();

                var hearderSystem = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "REPORT_HEADER");
                var dataCt = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "COMPANY_NAME");
                var day = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("dd") : "...";
                var month = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("MM") : "...";
                var year = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("yyyy") : "...";

                var approval_status = _uow.Repository<ApprovalFunction>().FirstOrDefault(a => a.item_id == id && a.function_code == "M_RAW");
                var auditwork = _uow.Repository<AuditWork>().Include(a => a.AuditWorkScope, a => a.Users).FirstOrDefault(a => a.Id == checkReport.auditwork_id);
                //var auditdetect = (from a in _uow.Repository<AuditDetect>().Find(x => x.IsDeleted != true && x.auditwork_id == checkReport.auditwork_id && x.audit_report == true)
                //                   join b in _uow.Repository<ApprovalFunction>().Find(x => x.function_code == "M_AD" && x.StatusCode == "3.1") on a.id equals b.item_id
                //                   select a).ToArray();
                var auditdetect = _uow.Repository<AuditDetect>().Find(x => x.IsDeleted != true && x.auditwork_id == checkReport.auditwork_id && x.audit_report == true).ToArray();

                //var auditdetect = _uow.Repository<auditdetect>().Find(x => x.IsDeleted != true && x.auditwork_id == report.auditwork_id && x.status == 3).ToArray();
                var detecttype = _uow.Repository<CatDetectType>().Find(x => x.IsDeleted != true).ToArray();
                var auditworkscope = auditWorkScope.AsEnumerable().GroupBy(a => new { a.auditprocess_id, a.auditfacilities_id, a.bussinessactivities_id }).Select(g => g.FirstOrDefault()).ToArray();

                var listfacility = _uow.Repository<AuditWorkScopeFacility>().Find(a => a.auditfacilities_id.HasValue && a.IsDeleted != true && a.auditwork_id == checkReport.auditwork_id).ToArray();

                var listdetect = auditdetect.Where(a => a.classify_audit_detect.HasValue).AsEnumerable().GroupBy(a => a.classify_audit_detect).Select(g => g.FirstOrDefault());

                var _auditRequestMonitorList = _uow.Repository<AuditRequestMonitor>().Include(x => x.Users).Where(a => a.is_deleted != true).ToArray();

                var sytemCategory = _uow.Repository<SystemCategory>().Find(o => o.ParentGroup == "MucXepHangKiemToan" && o.Status && !o.Deleted).ToArray();
                var auditminutes = _uow.Repository<AuditMinutes>().Find(a => a.auditwork_id == checkReport.auditwork_id && a.IsDeleted != true).ToArray();
                var lstAuditworkfacility = listfacility.Select(a => new AuditWorkFacilityModel()
                {
                    ID = a.Id,
                    AuditFacilityId = a.auditfacilities_id,
                    AuditFacility = a.auditfacilities_name,
                    AuditRatingLevelItem = auditminutes.FirstOrDefault(x => x.auditfacilities_id == a.auditfacilities_id)?.rating_level_total,
                    BaseRatingItem = a.BaseRatingReport,
                }).ToList();

                var lstDetectType = listdetect.Select(a => new SumaryDetectModel()
                {
                    classify_audit_detect = a.classify_audit_detect,
                    classify_audit_detect_name = detecttype.Where(x => x.Id == a.classify_audit_detect).Select(c => c.Name).FirstOrDefault().ToString(),
                    Hight = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 1),
                    Middle = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 2),
                    Low = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 3),
                    Agree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is true),
                    NotAgree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is false),

                }).ToList();

                var countRiskRatingHight = lstDetectType.Select(a => a.Hight).ToArray();

                var countRiskRatingMedium = lstDetectType.Select(a => a.Middle).ToArray();

                var countRiskRatingLow = lstDetectType.Select(a => a.Low).ToArray();

                var countOpinionAuditTrue = lstDetectType.Select(a => a.Agree).ToArray();

                var countOpinionAuditFalse = lstDetectType.Select(a => a.NotAgree).ToArray();

                var ListStatisticsOfDetections = listdetect.Select(a => new ListStatisticsOfDetections
                {
                    id = a.id,
                    audit_detect_code = a.code,
                    auditwork_id = a.auditwork_id,
                    auditfacilities_id = a.auditfacilities_id,
                    auditfacilities_name = a.auditfacilities_name,//đơn vị kiểm toán
                    title = a.title,//Tiêu đề phát hiện
                    year = a.year,
                    risk_rating = a.rating_risk,
                    str_classify_audit_detect = a.CatDetectType.Name,
                    reason = a.reason,
                    opinion_audit_true = countOpinionAuditTrue.Count(),
                    opinion_audit_false = countOpinionAuditFalse.Count(),
                    risk_rating_hight = countRiskRatingHight.Count(),
                    risk_rating_medium = countRiskRatingMedium.Count(),
                    risk_rating_low = countRiskRatingLow.Count(),
                }).GroupBy(a => a.str_classify_audit_detect).Select(z => z.FirstOrDefault()).ToList();

                //ẩn hiện
                var checkHiden = _uow.Repository<ConfigDocument>().GetAll().Where(a => a.item_name == "Báo cáo cuộc kiểm toán"
                && a.isShow == true).ToArray();
                var MTKT = checkHiden.Where(a => a.item_code == "MTKT" && a.status == true).Count();
                var PVKT = checkHiden.Where(a => a.item_code == "PVKT" && a.status == true).Count();
                var GHKT = checkHiden.Where(a => a.item_code == "GHKT" && a.status == true).Count();
                var THKT = checkHiden.Where(a => a.item_code == "THKT" && a.status == true).Count();
                var TGKT = checkHiden.Where(a => a.item_code == "TGKT" && a.status == true).Count();
                var DGDVKT = checkHiden.Where(a => a.item_code == "DGDVKT" && a.status == true).Count();
                var DGCCKT = checkHiden.Where(a => a.item_code == "DGCCKT" && a.status == true).Count();
                var DDD = checkHiden.Where(a => a.item_code == "DDD" && a.status == true).Count();
                var DCHT = checkHiden.Where(a => a.item_code == "DCHT" && a.status == true).Count();
                var YKKT = checkHiden.Where(a => a.item_code == "YKKT" && a.status == true).Count();
                var THPHKT = checkHiden.Where(a => a.item_code == "THPHKT" && a.status == true).Count();
                var KNKN = checkHiden.Where(a => a.item_code == "KNKN" && a.status == true).Count();
                var PLKT = checkHiden.Where(a => a.item_code == "PLKT" && a.status == true).Count();
                //ẩn hiện

                //// Export word
                //var fullPath = Path.Combine(_config["Template:ReporDocsTemplate"], "Kitano_BaoCaoKiemToan_v03.docx");
                //fullPath = fullPath.ToString().Replace("\\", "/");
                //using (Document doc = new Document(fullPath))
                using (Document doc = new Document(@"D:\test\Kitano_BaoCaoKiemToan_v03.docx"))
                {
                    //Header
                    doc.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true;
                    Paragraph paragraph_header = doc.Sections[0].HeadersFooters.FirstPageHeader.AddParagraph();
                    paragraph_header.AppendHTML(hearderSystem.Value);

                    //Remove header 2
                    doc.Sections[1].HeadersFooters.Header.ChildObjects.Clear();
                    //Add line breaks
                    Paragraph headerParagraph = doc.Sections[1].HeadersFooters.FirstPageHeader.AddParagraph();
                    headerParagraph.AppendBreak(BreakType.LineBreak);

                    if (MTKT != 1)
                    {
                        doc.Replace("Cuộc kiểm toán nhằm đảm bảo các mục tiêu sau:", "", false, true);
                    }
                    if (PVKT != 1)
                    {
                        doc.Replace("Phạm vi cuộc kiểm toán như sau:", "", false, true);
                    }
                    if (GHKT != 1)
                    {
                        doc.Replace("Giới hạn của cuộc kiểm toán:", "", false, true);
                    }
                    if (THKT != 1)
                    {
                        doc.Replace("Thời hiệu kiểm toán từ ngày", "", false, true);
                        doc.Replace("đến ngày", "", false, true);
                    }
                    if (TGKT != 1)
                    {
                        doc.Replace("Thời gian thực hiện kiểm toán từ ngày", "", false, true);
                    }
                    if (DGCCKT != 1)
                    {
                        doc.Replace("Mức xếp hạng kiểm toán:", "", false, true);
                        doc.Replace("Cơ sở xếp hạng kiểm toán:", "", false, true);
                        doc.Replace("Kết luận chung:", "", false, true);
                    }
                    if (DCHT != 1)
                    {
                        doc.Replace("Bảng tổng hợp số lượng phát hiện", "", false, true);
                    }
                    if ((DGDVKT != 1 && DGCCKT != 1 && DDD != 1 && DCHT != 1 && YKKT != 1))
                    {
                        doc.Replace("Dựa trên kết quả kiểm toán, đoàn Kiểm toán đưa ra đánh giá tổng quan và kết luận về:", "", false, true);
                    }


                    var MainStage = _uow.Repository<MainStage>().Find(a => a.auditwork_id == checkReport.auditwork_id).ToArray();
                    var startdate = MainStage.FirstOrDefault(a => a.index == 1);
                    var enddate = MainStage.FirstOrDefault(a => a.index == 4);
                    var mucdoxephang = checkReport.AuditRatingLevelTotal.HasValue ? sytemCategory.FirstOrDefault(a => a.Code == checkReport.AuditRatingLevelTotal.ToString())?.Name : "";
                    var cosoxephang = checkReport.BaseRatingTotal;
                    var ketluanchung = checkReport.GeneralConclusions;
                    var vandekiemsoat = "";
                    var ykienkiemsoat = "";
                    foreach (var item in lstAuditworkfacility)
                    {
                        vandekiemsoat += item.AuditFacility + ": " + auditminutes.FirstOrDefault(x => x.auditfacilities_id == item.AuditFacilityId)?.problem + "\r\n";
                        ykienkiemsoat += item.AuditFacility + ": " + auditminutes.FirstOrDefault(x => x.auditfacilities_id == item.AuditFacilityId)?.idea + "\r\n";
                    }
                    if (!string.IsNullOrEmpty(vandekiemsoat))
                    {
                        vandekiemsoat = vandekiemsoat.Substring(0, vandekiemsoat.Length - 2).Replace("\n", "");
                    }
                    if (!string.IsNullOrEmpty(ykienkiemsoat))
                    {
                        ykienkiemsoat = ykienkiemsoat.Substring(0, ykienkiemsoat.Length - 2).Replace("\n", "");
                    }

                    TextSelection selectionss = doc.FindAllString("noi_dung_khac_1", false, true)?.FirstOrDefault();
                    if (selectionss != null)
                    {
                        if (!string.IsNullOrEmpty(checkReport.OtherContent) && KNKN == 1)
                        {
                            var paragraph = selectionss?.GetAsOneRange()?.OwnerParagraph;
                            var object_arr = paragraph.ChildObjects;
                            if (object_arr.Count > 0)
                            {
                                for (int j = 0; j < object_arr.Count; j++)
                                {
                                    paragraph.ChildObjects.RemoveAt(j);
                                }
                                paragraph.AppendHTML(checkReport.OtherContent);
                            }
                        }
                    }

                    doc.MailMerge.Execute(
                    new string[] { "ngay_bc_1" , "thang_bc_1"  , "nam_bc_1" ,
                        "td_tong_quan_1",
                        "td_ket_luan_1",
                        "td_muc_tieu_kt_1",
                        "td_pham_vi_kt_1",
                        "td_gioi_han_pham_vi_kt_1",
                        "td_thoi_hieu_kt_1",
                        "td_thoi_gian_kt_1",
                        "den_ngay_1",
                        "td_danh_gia_dvkt",
                        "td_danh_gia_chung_ckt",
                        "td_diem_dat_duoc",
                        "td_diem_can_ht",
                        "td_y_kien_dvkt_1",
                        "td_tong_hop",
                        "td_kien_nghi_khuyen_nghi",//11
                        "td_phu_luc",
                        "muc_dich_kt_1",
                        "pham_vi_kt_1",
                        "ngoai_pham_vi_kt_1",
                        "thoi_hieu_kt_tu_1",
                        "thoi_hieu_kt_den_1",
                        "thoi_gian_kt_tu_1",
                        "thoi_gian_kt_den_1",
                        "muc_xep_hang_kiem_toan_1",
                        "co_so_xep_hang_1",
                        "ket_luan_chung_1"
                        ,"van_de_kiem_soat_1","y_kien_kiem_soat_1","noi_dung_khac_1"
                    },
                    new string[] { DateTime.Now.Day.ToString() , DateTime.Now.Month.ToString() , DateTime.Now.Year.ToString() ,
                        (MTKT != 1 && PVKT != 1 && GHKT != 1 && THKT != 1) ? "" : "TỔNG QUAN",
                        (DGDVKT != 1 && DGCCKT != 1 && DDD != 1 && DCHT != 1 && YKKT != 1) ? "" : "KẾT LUẬN",
                        (MTKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "MTKT")?.content),
                        (PVKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "PVKT")?.content),
                        (GHKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "GHKT")?.content),
                        (THKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "THKT")?.content),
                        (TGKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "TGKT")?.content),
                        (TGKT != 1) ? "" : "đến ngày",
                        (DGDVKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DGDVKT")?.content),
                        (DGCCKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DGCCKT")?.content),
                        (DDD != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DDD")?.content),
                        (DCHT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DCHT")?.content),
                        (YKKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "YKKT")?.content),
                        (THPHKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "THPHKT")?.content),
                        (KNKN != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "KNKN")?.content),//11
                        (PLKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "PLKT")?.content),

                        MTKT == 1 ? checkReport.AuditWork.Target: "",
                        PVKT == 1 ? checkReport.AuditWork.AuditScope: "",
                        GHKT == 1 ? checkReport.AuditWork.AuditScopeOutside: "",
                        (THKT == 1 ? (auditwork.from_date != null ? auditwork.from_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (THKT == 1 ? (auditwork.to_date != null ?  auditwork.to_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (TGKT == 1 ? (startdate != null && startdate.actual_date.HasValue ? startdate.actual_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (TGKT == 1 ? (enddate != null && enddate.actual_date.HasValue ? enddate.actual_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (DGCCKT == 1 ? mucdoxephang: ""),
                        (DGCCKT == 1 ? cosoxephang: ""),
                        (DGCCKT == 1 ? ketluanchung: ""),
                        (DDD == 1 ? vandekiemsoat: ""),
                        (DDD == 1 ? ykienkiemsoat:""),
                        ""
                    }
                    );
                    if (DGDVKT == 1)
                    {
                        Table table2 = doc.Sections[0].Tables[1] as Table;
                        //table2.ResetCells(lstAuditworkfacility.Count() + 1, 2);
                        //TextRange txtTable2 = table2[0, 0].AddParagraph().AppendText("Đơn vị được KT");
                        //txtTable2.CharacterFormat.FontName = "Times New Roman";
                        //txtTable2.CharacterFormat.FontSize = 12;
                        //txtTable2.CharacterFormat.Bold = true;

                        //txtTable2 = table2[0, 1].AddParagraph().AppendText("Mức xếp hạng kiểm toán");
                        //txtTable2.CharacterFormat.FontName = "Times New Roman";
                        //txtTable2.CharacterFormat.FontSize = 12;
                        //txtTable2.CharacterFormat.Bold = true;
                        for (int x = 0; x < lstAuditworkfacility.Count(); x++)
                        {
                            var row = table2.AddRow();
                            var txtTable2 = row.Cells[0].AddParagraph().AppendText(lstAuditworkfacility[x].AuditFacility);
                            txtTable2.CharacterFormat.FontName = "Times New Roman";
                            txtTable2.CharacterFormat.FontSize = 12;

                            var mucdo = lstAuditworkfacility[x].AuditRatingLevelItem.HasValue ? sytemCategory.FirstOrDefault(a => a.Id == lstAuditworkfacility[x].AuditRatingLevelItem)?.Name : "";
                            txtTable2 = row.Cells[1].AddParagraph().AppendText(mucdo);
                            txtTable2.CharacterFormat.FontName = "Times New Roman";
                            txtTable2.CharacterFormat.FontSize = 12;
                        }
                    }
                    else
                    {
                        doc.Sections[0].Tables.Remove(doc.Sections[0].Tables[1]);
                    }
                    if (DCHT == 1)
                    {
                        Table table1 = doc.Sections[0].Tables[2] as Table;
                        table1.ResetCells(lstDetectType.Count() + 2, 7);
                        table1.ApplyVerticalMerge(0, 0, 1);
                        table1.ApplyVerticalMerge(1, 0, 1);
                        table1.ApplyHorizontalMerge(0, 2, 4);
                        table1.ApplyHorizontalMerge(0, 5, 6);
                        TextRange txtTable1 = table1[0, 0].AddParagraph().AppendText("Loại phát hiện");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[0, 1].AddParagraph().AppendText("Tổng cộng số lượng");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[0, 2].AddParagraph().AppendText("Số lượng phát hiện");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[0, 5].AddParagraph().AppendText("Ý kiến của đơn vị được kiểm toán");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 2].AddParagraph().AppendText("Cao/Quan trọng");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 3].AddParagraph().AppendText("Trung bình");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 4].AddParagraph().AppendText("Thấp/Ít quan trọng");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 5].AddParagraph().AppendText("Đồng ý");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 6].AddParagraph().AppendText("Không đồng ý");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        for (int z = 0; z < lstDetectType.Count(); z++)
                        {
                            txtTable1 = table1[z + 2, 0].AddParagraph().AppendText(lstDetectType[z].classify_audit_detect_name);
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 1].AddParagraph().AppendText((lstDetectType[z].Hight + lstDetectType[z].Middle + lstDetectType[z].Low).ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 2].AddParagraph().AppendText(lstDetectType[z].Hight.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            table1[z + 2, 2].CellFormat.BackColor = ColorTranslator.FromHtml("#FF0000");
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 3].AddParagraph().AppendText(lstDetectType[z].Middle.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            table1[z + 2, 3].CellFormat.BackColor = ColorTranslator.FromHtml("#FFC000");
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 4].AddParagraph().AppendText(lstDetectType[z].Low.ToString());
                            table1[z + 2, 4].CellFormat.BackColor = ColorTranslator.FromHtml("#92D050");
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 5].AddParagraph().AppendText(lstDetectType[z].Agree.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 6].AddParagraph().AppendText(lstDetectType[z].NotAgree.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;
                        }
                    }
                    else
                    {
                        doc.Sections[0].Tables.Remove(doc.Sections[0].Tables[2]);
                    }
                    if (THPHKT == 1)
                    {
                        Table table11 = doc.Sections[1].Tables[0] as Table;
                        //table11.ResetCells(auditdetect.Count() + 1, 4);
                        //TextRange txtTable11 = table11[0, 0].AddParagraph().AppendText("Khung quản trị");
                        //txtTable11.CharacterFormat.FontName = "Times New Roman";
                        //txtTable11.CharacterFormat.FontSize = 12;
                        //txtTable11.CharacterFormat.Bold = true;

                        //txtTable11 = table11[0, 1].AddParagraph().AppendText("Mã phát hiện");
                        //txtTable11.CharacterFormat.FontName = "Times New Roman";
                        //txtTable11.CharacterFormat.FontSize = 12;
                        //txtTable11.CharacterFormat.Bold = true;

                        //txtTable11 = table11[0, 2].AddParagraph().AppendText("Loại phát hiện");
                        //txtTable11.CharacterFormat.FontName = "Times New Roman";
                        //txtTable11.CharacterFormat.FontSize = 12;
                        //txtTable11.CharacterFormat.Bold = true;

                        //txtTable11 = table11[0, 3].AddParagraph().AppendText("Mô tả phát hiện");
                        //txtTable11.CharacterFormat.FontName = "Times New Roman";
                        //txtTable11.CharacterFormat.FontSize = 12;
                        //txtTable11.CharacterFormat.Bold = true;
                        var listauditdetectRiskHight = auditdetect.Where(a => a.audit_report == true).OrderBy(a => a.admin_framework).ToArray();
                        for (int x = 0; x < listauditdetectRiskHight.Length; x++)
                        {
                            var auditdetectitem = listauditdetectRiskHight[x];
                            var adminFramework = auditdetectitem.admin_framework;
                            var str_adminFramework = "";
                            switch (adminFramework)
                            {
                                case 1:
                                    str_adminFramework = "Quản trị/Tổ chức/Số liệu tài chính";
                                    break;
                                case 2:
                                    str_adminFramework = "Hoạt động/Quy trình/Quy định";
                                    break;
                                case 3:
                                    str_adminFramework = "Kiểm sát vận hành/Thực thi";
                                    break;
                                case 4:
                                    str_adminFramework = "Công nghệ thông tinh";
                                    break;
                            }
                            var row = table11.AddRow();
                            var txtTable11 = row.Cells[0].AddParagraph().AppendText(str_adminFramework);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[0].CellFormat.BackColor = ColorTranslator.FromHtml("#FFFFFF");


                            var risk_level = "";
                            var color_ = "#FFFFFF";
                            switch (auditdetectitem.rating_risk)
                            {
                                case 1:
                                    risk_level = "Cao/Quan trọng";
                                    color_ = "#ff0000";
                                    break;
                                case 2:
                                    risk_level = "Trung bình";
                                    color_ = "#ffbf00";
                                    break;
                                case 3:
                                    risk_level = "Thấp/Ít quan trọng";
                                    color_ = "#92d050";
                                    break;
                            }

                            txtTable11 = row.Cells[1].AddParagraph().AppendText(auditdetectitem.code);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[1].CellFormat.BackColor = ColorTranslator.FromHtml(color_);


                            txtTable11 = row.Cells[2].AddParagraph().AppendText(risk_level);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[2].CellFormat.BackColor = ColorTranslator.FromHtml(color_);

                            var type_item = detecttype.FirstOrDefault(a => a.Id == auditdetectitem.classify_audit_detect)?.Name;
                            txtTable11 = row.Cells[3].AddParagraph().AppendText(type_item);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[3].CellFormat.BackColor = ColorTranslator.FromHtml("#FFFFFF");

                            txtTable11 = row.Cells[4].AddParagraph().AppendText(auditdetectitem.short_title);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[4].CellFormat.BackColor = ColorTranslator.FromHtml("#FFFFFF");
                        }
                    }
                    else
                    {
                        doc.Sections[1].Tables.Remove(doc.Sections[1].Tables[0]);
                        //doc.Sections.Remove(doc.Sections[1]);
                    }
                    if (PLKT == 1)
                    {
                        TextSelection selections = doc.FindAllString("PHỤ LỤC DANH SÁCH PHÁT HIỆN KIỂM TOÁN", false, true)?.FirstOrDefault();
                        if (selections != null)
                        {
                            var Paragraph = selections?.GetAsOneRange()?.OwnerParagraph;
                            var sestionOwner = Paragraph.Owner.Owner as Section;
                            var listdetectNew = auditdetect.Where(a => a.classify_audit_detect.HasValue).AsEnumerable().GroupBy(a => a.classify_audit_detect);
                            var no_index = 0;
                            foreach (var item in listdetectNew)
                            {
                                var auditdetectitem = item;
                                var type_name = detecttype.FirstOrDefault(a => a.Id == auditdetectitem.Key)?.Name;
                                if (!string.IsNullOrEmpty(type_name))
                                {
                                    no_index++;

                                    var textheader = sestionOwner.AddParagraph().AppendText(no_index + ". " + type_name);
                                    textheader.CharacterFormat.FontName = "Times New Roman";
                                    textheader.CharacterFormat.FontSize = 11;
                                    textheader.CharacterFormat.Bold = true;

                                    Table table_new = sestionOwner.AddTable(true);
                                    table_new.ResetCells(item.Count() + 2, 5);
                                    table_new.ApplyVerticalMerge(0, 0, 1);
                                    table_new.ApplyHorizontalMerge(0, 1, 2);
                                    table_new.ApplyHorizontalMerge(0, 3, 4);

                                    var paragraph = table_new[0, 0].AddParagraph();
                                    table_new[0, 0].Width = 90;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    TextRange txtTable_new = paragraph.AppendText("Mã phát hiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;
                                    table_new[0, 0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                                    table_new[0, 0].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[0, 1].AddParagraph();
                                    table_new[0, 1].Width = 500;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Phát hiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[0, 1].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");
                                    table_new[0, 1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                                    paragraph = table_new[0, 3].AddParagraph();
                                    table_new[0, 3].Width = 500;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Kiến nghị/Khuyến nghị");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[0, 3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[0, 3].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 1].AddParagraph();
                                    table_new[1, 1].Width = 370;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Nội dung phát hiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[1, 1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 1].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 2].AddParagraph();
                                    table_new[1, 2].Width = 130;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Mức độ rủi ro");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[1, 2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 2].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 3].AddParagraph();
                                    table_new[1, 3].Width = 350;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Nội dung");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;
                                    table_new[1, 3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 3].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 4].AddParagraph();
                                    table_new[1, 4].Width = 150;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Thời hạn hoàn thiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[1, 4].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 4].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");
                                    var zz = 0;
                                    foreach (var item_row in item.ToList())
                                    {
                                        txtTable_new = table_new[zz + 2, 0].AddParagraph().AppendText(item_row.code);
                                        table_new[zz + 2, 0].Width = 90;
                                        txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                        txtTable_new.CharacterFormat.FontSize = 11;

                                        txtTable_new = table_new[zz + 2, 1].AddParagraph().AppendText(item_row.summary_audit_detect);
                                        table_new[zz + 2, 1].Width = 370;
                                        txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                        txtTable_new.CharacterFormat.FontSize = 11;

                                        var risk_level = "";
                                        switch (item_row.rating_risk)
                                        {
                                            case 1:
                                                risk_level = "Cao/Quan trọng";
                                                break;
                                            case 2:
                                                risk_level = "Trung bình";
                                                break;
                                            case 3:
                                                risk_level = "Thấp/Ít quan trọng";
                                                break;
                                        }
                                        txtTable_new = table_new[zz + 2, 2].AddParagraph().AppendText(risk_level);
                                        table_new[zz + 2, 2].Width = 130;
                                        txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                        txtTable_new.CharacterFormat.FontSize = 11;

                                        var _auditRequestMonitor = _auditRequestMonitorList.Where(a => a.detectid == item_row.id).ToArray();
                                        var paragraph1 = table_new[zz + 2, 3].AddParagraph();
                                        var paragraph2 = table_new[zz + 2, 4].AddParagraph();
                                        for (int d = 0; d < _auditRequestMonitor.Length; d++)
                                        {
                                            var _auditRequestMonitorItem = _auditRequestMonitor[d];
                                            var txtTable_item = paragraph1.AppendText("- " + _auditRequestMonitorItem.Code + ": " + _auditRequestMonitorItem.Content + "\n");
                                            txtTable_item.CharacterFormat.FontName = "Times New Roman";
                                            txtTable_item.CharacterFormat.FontSize = 11;

                                            var date = _auditRequestMonitorItem.CompleteAt.HasValue ? _auditRequestMonitorItem.CompleteAt.Value.ToString("dd/MM/yyyy") : "";
                                            if (!string.IsNullOrEmpty(date))
                                            {
                                                txtTable_item = paragraph2.AppendText("- " + _auditRequestMonitorItem.Code + ": " + date + "\n");
                                                txtTable_item.CharacterFormat.FontName = "Times New Roman";
                                                txtTable_item.CharacterFormat.FontSize = 11;
                                            }
                                        }
                                        table_new[zz + 2, 3].Width = 350;
                                        table_new[zz + 2, 4].Width = 150;
                                        zz++;
                                    }
                                }
                                sestionOwner.AddParagraph();
                            }
                        }
                    }
                    else
                    {
                        doc.Sections.Remove(doc.Sections[4]);
                    }
                    //if (KNKN != 1)
                    //{
                    //    doc.Sections.Remove(doc.Sections[2]);
                    //}
                    foreach (Section section1 in doc.Sections)
                    {
                        for (int i = 0; i < section1.Body.ChildObjects.Count; i++)
                        {
                            if (section1.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
                            {
                                if (String.IsNullOrEmpty((section1.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                                {
                                    section1.Body.ChildObjects.Remove(section1.Body.ChildObjects[i]);
                                    i--;
                                }
                            }

                        }
                    }

                    MemoryStream stream = new MemoryStream();
                    stream.Position = 0;
                    doc.SaveToFile(stream, Spire.Doc.FileFormat.Docx);
                    Bytes = stream.ToArray();
                    return File(Bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Kitano_BaoCaoKiemToan.docx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"ExportAudit - {DateTime.Now} : {ex.Message}!");
                return BadRequest();
            }
        }

        [HttpGet("ExportFileWordMBS/{id}")]
        public IActionResult ExportFileWordMBS(int id)
        {
            byte[] Bytes = null;
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var checkReport = _uow.Repository<ReportAuditWork>().FirstOrDefault(a => a.Id == id && a.IsDeleted.Equals(false));

                var auditWorkScope = _uow.Repository<AuditWorkScope>().GetAll().Where(a => a.auditwork_id == checkReport.auditwork_id && a.IsDeleted != true).ToArray();

                var lstauditworkscope = auditWorkScope.Where(a => a.auditprocess_id.HasValue).Select(a => new AuditWorkScopeModel()
                {
                    ID = a.Id,
                    AuditProcess = a.auditprocess_name,
                    AuditFacility = a.auditfacilities_name,
                    AuditActivity = a.bussinessactivities_name,
                }).ToList();

                var hearderSystem = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "REPORT_HEADER");
                var dataCt = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "COMPANY_NAME");
                var day = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("dd") : "...";
                var month = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("MM") : "...";
                var year = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("yyyy") : "...";

                var approval_status = _uow.Repository<ApprovalFunction>().FirstOrDefault(a => a.item_id == id && a.function_code == "M_RAW");
                var auditwork = _uow.Repository<AuditWork>().Include(a => a.AuditWorkScope, a => a.Users).FirstOrDefault(a => a.Id == checkReport.auditwork_id);

                var auditdetect = _uow.Repository<AuditDetect>().Find(x => x.IsDeleted != true && x.auditwork_id == checkReport.auditwork_id && x.audit_report == true).ToArray();

                //var auditdetect = _uow.Repository<auditdetect>().Find(x => x.IsDeleted != true && x.auditwork_id == report.auditwork_id && x.status == 3).ToArray();
                var detecttype = _uow.Repository<CatDetectType>().Find(x => x.IsDeleted != true).ToArray();
                var auditworkscope = auditWorkScope.AsEnumerable().GroupBy(a => new { a.auditprocess_id, a.auditfacilities_id, a.bussinessactivities_id }).Select(g => g.FirstOrDefault()).ToArray();

                var listfacility = _uow.Repository<AuditWorkScopeFacility>().Find(a => a.auditfacilities_id.HasValue && a.IsDeleted != true && a.auditwork_id == checkReport.auditwork_id).ToArray();

                var listdetect = auditdetect.Where(a => a.classify_audit_detect.HasValue).AsEnumerable().GroupBy(a => a.classify_audit_detect).Select(g => g.FirstOrDefault());

                var _auditRequestMonitorList = _uow.Repository<AuditRequestMonitor>().Include(x => x.Users).Where(a => a.is_deleted != true).ToArray();

                var sytemCategory = _uow.Repository<SystemCategory>().Find(o => o.ParentGroup == "MucXepHangKiemToan" && o.Status && !o.Deleted).ToArray();
                var auditminutes = _uow.Repository<AuditMinutes>().Find(a => a.auditwork_id == checkReport.auditwork_id && a.IsDeleted != true).ToArray();
                var lstAuditworkfacility = listfacility.Select(a => new AuditWorkFacilityModel()
                {
                    ID = a.Id,
                    AuditFacilityId = a.auditfacilities_id,
                    AuditFacility = a.auditfacilities_name,
                    AuditRatingLevelItem = auditminutes.FirstOrDefault(x => x.auditfacilities_id == a.auditfacilities_id)?.rating_level_total,
                    BaseRatingItem = a.BaseRatingReport,
                }).ToList();

                var lstDetectType = listdetect.Select(a => new SumaryDetectModel()
                {
                    classify_audit_detect = a.classify_audit_detect,
                    classify_audit_detect_name = detecttype.Where(x => x.Id == a.classify_audit_detect).Select(c => c.Name).FirstOrDefault().ToString(),
                    Hight = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 1),
                    Middle = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 2),
                    Low = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 3),
                    Agree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is true),
                    NotAgree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is false),

                }).ToList();

                var countRiskRatingHight = lstDetectType.Select(a => a.Hight).ToArray();

                var countRiskRatingMedium = lstDetectType.Select(a => a.Middle).ToArray();

                var countRiskRatingLow = lstDetectType.Select(a => a.Low).ToArray();

                var countOpinionAuditTrue = lstDetectType.Select(a => a.Agree).ToArray();

                var countOpinionAuditFalse = lstDetectType.Select(a => a.NotAgree).ToArray();

                var ListStatisticsOfDetections = listdetect.Select(a => new ListStatisticsOfDetections
                {
                    id = a.id,
                    audit_detect_code = a.code,
                    auditwork_id = a.auditwork_id,
                    auditfacilities_id = a.auditfacilities_id,
                    auditfacilities_name = a.auditfacilities_name,//đơn vị kiểm toán
                    title = a.title,//Tiêu đề phát hiện
                    year = a.year,
                    risk_rating = a.rating_risk,
                    str_classify_audit_detect = a.CatDetectType.Name,
                    reason = a.reason,
                    opinion_audit_true = countOpinionAuditTrue.Count(),
                    opinion_audit_false = countOpinionAuditFalse.Count(),
                    risk_rating_hight = countRiskRatingHight.Count(),
                    risk_rating_medium = countRiskRatingMedium.Count(),
                    risk_rating_low = countRiskRatingLow.Count(),
                }).GroupBy(a => a.str_classify_audit_detect).Select(z => z.FirstOrDefault()).ToList();

                //ẩn hiện
                var checkHiden = _uow.Repository<ConfigDocument>().GetAll().Where(a => a.item_name == "Báo cáo cuộc kiểm toán"
                && a.isShow == true).ToArray();
                var MTKT = checkHiden.Where(a => a.item_code == "MTKT" && a.status == true).Count();
                var PVKT = checkHiden.Where(a => a.item_code == "PVKT" && a.status == true).Count();
                var GHKT = checkHiden.Where(a => a.item_code == "GHKT" && a.status == true).Count();
                var THKT = checkHiden.Where(a => a.item_code == "THKT" && a.status == true).Count();
                var TGKT = checkHiden.Where(a => a.item_code == "TGKT" && a.status == true).Count();
                var DGDVKT = checkHiden.Where(a => a.item_code == "DGDVKT" && a.status == true).Count();
                var DGCCKT = checkHiden.Where(a => a.item_code == "DGCCKT" && a.status == true).Count();
                var DDD = checkHiden.Where(a => a.item_code == "DDD" && a.status == true).Count();
                var DCHT = checkHiden.Where(a => a.item_code == "DCHT" && a.status == true).Count();
                var YKKT = checkHiden.Where(a => a.item_code == "YKKT" && a.status == true).Count();
                var THPHKT = checkHiden.Where(a => a.item_code == "THPHKT" && a.status == true).Count();
                var KNKN = checkHiden.Where(a => a.item_code == "KNKN" && a.status == true).Count();
                var PLKT = checkHiden.Where(a => a.item_code == "PLKT" && a.status == true).Count();
                //ẩn hiện

                // Export word
                var fullPath = Path.Combine(_config["Template:ReporDocsTemplate"], "MBS_Kitano_BaoCaoCuocKiemToan_v0.1.docx");
                fullPath = fullPath.ToString().Replace("\\", "/");
                using (Document doc = new Document(fullPath))
                //using (Document doc = new Document(@"D:\test\MBS_Kitano_BaoCaoCuocKiemToan_v0.1.docx"))
                {
                    //Header
                    doc.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true;
                    Paragraph paragraph_header = doc.Sections[0].HeadersFooters.FirstPageHeader.AddParagraph();
                    paragraph_header.AppendHTML(hearderSystem.Value);

                    //Remove header 2
                    doc.Sections[1].HeadersFooters.Header.ChildObjects.Clear();
                    //Add line breaks
                    Paragraph headerParagraph = doc.Sections[1].HeadersFooters.FirstPageHeader.AddParagraph();
                    headerParagraph.AppendBreak(BreakType.LineBreak);

                    if (MTKT != 1)
                    {
                        doc.Replace("Cuộc kiểm toán nhằm đảm bảo các mục tiêu sau:", "", false, true);
                    }
                    if (PVKT != 1)
                    {
                        doc.Replace("Phạm vi cuộc kiểm toán như sau:", "", false, true);
                    }
                    if (GHKT != 1)
                    {
                        doc.Replace("Giới hạn của cuộc kiểm toán:", "", false, true);
                    }
                    if (THKT != 1)
                    {
                        doc.Replace("Thời hiệu kiểm toán từ ngày", "", false, true);
                        doc.Replace("đến ngày", "", false, true);
                    }
                    if (TGKT != 1)
                    {
                        doc.Replace("Thời gian thực hiện kiểm toán từ ngày", "", false, true);
                    }
                    if (DGCCKT != 1)
                    {
                        doc.Replace("Mức xếp hạng kiểm toán:", "", false, true);
                        doc.Replace("Cơ sở xếp hạng kiểm toán:", "", false, true);
                        doc.Replace("Kết luận chung:", "", false, true);
                    }

                    if (DCHT != 1)
                    {
                        doc.Replace("Bảng tổng hợp số lượng phát hiện", "", false, true);
                    }
                    if ((DGDVKT != 1 && DGCCKT != 1 && DDD != 1 && DCHT != 1 && YKKT != 1))
                    {
                        doc.Replace("Dựa trên kết quả kiểm toán, đoàn Kiểm toán đưa ra đánh giá tổng quan và kết luận về:", "", false, true);
                    }


                    var MainStage = _uow.Repository<MainStage>().Find(a => a.auditwork_id == checkReport.auditwork_id).ToArray();
                    var startdate = MainStage.FirstOrDefault(a => a.index == 1);
                    var enddate = MainStage.FirstOrDefault(a => a.index == 4);
                    var mucdoxephang = checkReport.AuditRatingLevelTotal.HasValue ? sytemCategory.FirstOrDefault(a => a.Code == checkReport.AuditRatingLevelTotal.ToString())?.Name : "";
                    var cosoxephang = checkReport.BaseRatingTotal;
                    var ketluanchung = checkReport.GeneralConclusions;
                    var vandekiemsoat = "";
                    var ykienkiemsoat = "";
                    foreach (var item in lstAuditworkfacility)
                    {
                        vandekiemsoat += item.AuditFacility + ": " + auditminutes.FirstOrDefault(x => x.auditfacilities_id == item.AuditFacilityId)?.problem + "\r\n";
                        ykienkiemsoat += item.AuditFacility + ": " + auditminutes.FirstOrDefault(x => x.auditfacilities_id == item.AuditFacilityId)?.idea + "\r\n";
                    }
                    if (!string.IsNullOrEmpty(vandekiemsoat))
                    {
                        vandekiemsoat = vandekiemsoat.Substring(0, vandekiemsoat.Length - 2).Replace("\n", "");
                    }
                    if (!string.IsNullOrEmpty(ykienkiemsoat))
                    {
                        ykienkiemsoat = ykienkiemsoat.Substring(0, ykienkiemsoat.Length - 2).Replace("\n", "");
                    }

                    TextSelection selectionss = doc.FindAllString("noi_dung_khac_1", false, true)?.FirstOrDefault();
                    if (selectionss != null)
                    {
                        if (!string.IsNullOrEmpty(checkReport.OtherContent) && KNKN == 1)
                        {
                            var paragraph = selectionss?.GetAsOneRange()?.OwnerParagraph;
                            var object_arr = paragraph.ChildObjects;
                            if (object_arr.Count > 0)
                            {
                                for (int j = 0; j < object_arr.Count; j++)
                                {
                                    paragraph.ChildObjects.RemoveAt(j);
                                }
                                paragraph.AppendHTML(checkReport.OtherContent);
                            }
                        }
                    }

                    doc.MailMerge.Execute(
                    new string[] { "ngay_bc_1" , "thang_bc_1"  , "nam_bc_1" , "ten_ckt",
                        //"td_tong_quan_1",
                        //"td_ket_luan_1",
                        //"td_muc_tieu_kt_1",
                        //"td_pham_vi_kt_1",
                        //"td_gioi_han_pham_vi_kt_1",
                        //"td_thoi_hieu_kt_1",
                        //"td_thoi_gian_kt_1",
                        //"den_ngay_1",
                        //"td_danh_gia_dvkt",
                        //"td_danh_gia_chung_ckt",
                        //"td_diem_dat_duoc",
                        //"td_diem_can_ht",
                        //"td_y_kien_dvkt_1",
                        //"td_tong_hop",
                        //"td_kien_nghi_khuyen_nghi",//11
                        //"td_phu_luc",
                        "muc_dich_kt_1",
                        "pham_vi_kt_1",
                        "ngoai_pham_vi_kt_1",
                        "thoi_hieu_kt_tu_1",
                        "thoi_hieu_kt_den_1",
                        "thoi_gian_kt_tu_1",
                        "thoi_gian_kt_den_1",
                        "muc_xep_hang_kiem_toan_1",
                        "co_so_xep_hang_1",
                        "ket_luan_chung_1"
                        ,"van_de_kiem_soat_1","y_kien_kiem_soat_1","noi_dung_khac_1",
                        "nguoi_phu_trach_1"
                    },
                    new string[] { DateTime.Now.Day.ToString() , DateTime.Now.Month.ToString() , DateTime.Now.Year.ToString() ,checkReport.AuditWorkName,
                        //(MTKT != 1 && PVKT != 1 && GHKT != 1 && THKT != 1) ? "" : "TỔNG QUAN",
                        //(DGDVKT != 1 && DGCCKT != 1 && DDD != 1 && DCHT != 1 && YKKT != 1) ? "" : "KẾT LUẬN",
                        //(MTKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "MTKT")?.content),
                        //(PVKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "PVKT")?.content),
                        //(GHKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "GHKT")?.content),
                        //(THKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "THKT")?.content),
                        //(TGKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "TGKT")?.content),
                        //(TGKT != 1) ? "" : "đến ngày",
                        //(DGDVKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DGDVKT")?.content),
                        //(DGCCKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DGCCKT")?.content),
                        //(DDD != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DDD")?.content),
                        //(DCHT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "DCHT")?.content),
                        //(YKKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "YKKT")?.content),
                        //(THPHKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "THPHKT")?.content),
                        //(KNKN != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "KNKN")?.content),//11
                        //(PLKT != 1) ? "" : (checkHiden.FirstOrDefault(a=>a.item_code == "PLKT")?.content),

                        MTKT == 1 ? checkReport.AuditWork.Target: "",
                        PVKT == 1 ? checkReport.AuditWork.AuditScope: "",
                        GHKT == 1 ? checkReport.AuditWork.AuditScopeOutside: "",
                        (THKT == 1 ? (auditwork.from_date != null ? auditwork.from_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (THKT == 1 ? (auditwork.to_date != null ?  auditwork.to_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (TGKT == 1 ? (startdate != null && startdate.actual_date.HasValue ? startdate.actual_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (TGKT == 1 ? (enddate != null && enddate.actual_date.HasValue ? enddate.actual_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy"): ""),
                        (DGCCKT == 1 ? mucdoxephang: ""),
                        (DGCCKT == 1 ? cosoxephang: ""),
                        (DGCCKT == 1 ? ketluanchung: ""),
                        (DDD == 1 ? vandekiemsoat: ""),
                        (DDD == 1 ? ykienkiemsoat:""),
                        "",
                        auditwork.Users != null ? auditwork.Users.FullName : ""
                    }
                    );
                    if (DCHT == 1)
                    {
                        Table table1 = doc.Sections[0].Tables[3] as Table;
                        table1.ResetCells(lstDetectType.Count() + 3, 7);
                        table1.ApplyVerticalMerge(0, 0, 1);
                        table1.ApplyVerticalMerge(1, 0, 1);
                        table1.ApplyHorizontalMerge(0, 2, 4);
                        table1.ApplyHorizontalMerge(0, 5, 6);
                        TextRange txtTable1 = table1[0, 0].AddParagraph().AppendText("Loại phát hiện");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[0, 1].AddParagraph().AppendText("Tổng cộng số lượng");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[0, 2].AddParagraph().AppendText("Số lượng phát hiện");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[0, 5].AddParagraph().AppendText("Ý kiến của đơn vị được kiểm toán");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 2].AddParagraph().AppendText("Cao/Quan trọng");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 3].AddParagraph().AppendText("Trung bình");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 4].AddParagraph().AppendText("Thấp/Ít quan trọng");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 5].AddParagraph().AppendText("Đồng ý");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        txtTable1 = table1[1, 6].AddParagraph().AppendText("Không đồng ý");
                        txtTable1.CharacterFormat.FontName = "Times New Roman";
                        txtTable1.CharacterFormat.FontSize = 12;
                        txtTable1.CharacterFormat.Bold = true;

                        for (int z = 0; z < lstDetectType.Count(); z++)
                        {

                            txtTable1 = table1[z + 2, 0].AddParagraph().AppendText(lstDetectType[z].classify_audit_detect_name);
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 1].AddParagraph().AppendText((lstDetectType[z].Hight + lstDetectType[z].Middle + lstDetectType[z].Low).ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 2].AddParagraph().AppendText(lstDetectType[z].Hight.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            table1[z + 2, 2].CellFormat.BackColor = ColorTranslator.FromHtml("#FF0000");
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 3].AddParagraph().AppendText(lstDetectType[z].Middle.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            table1[z + 2, 3].CellFormat.BackColor = ColorTranslator.FromHtml("#FFC000");
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 4].AddParagraph().AppendText(lstDetectType[z].Low.ToString());
                            table1[z + 2, 4].CellFormat.BackColor = ColorTranslator.FromHtml("#92D050");
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 5].AddParagraph().AppendText(lstDetectType[z].Agree.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;

                            txtTable1 = table1[z + 2, 6].AddParagraph().AppendText(lstDetectType[z].NotAgree.ToString());
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;
                        }

                            txtTable1 = table1[lstDetectType.Count() + 2, 0].AddParagraph().AppendText("Mức đánh giá");
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;
                            txtTable1 = table1[lstDetectType.Count() + 2, 1].AddParagraph().AppendText("");
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;
                            txtTable1 = table1[lstDetectType.Count() + 2, 2].AddParagraph().AppendText(mucdoxephang);
                            table1[lstDetectType.Count() + 2, 2].CellFormat.BackColor = ColorTranslator.FromHtml("#FFC000");
                            txtTable1.CharacterFormat.FontName = "Times New Roman";
                            txtTable1.CharacterFormat.FontSize = 12;
                            table1.ApplyHorizontalMerge(lstDetectType.Count() + 2, 2, 6);
                            table1.Rows[lstDetectType.Count() + 2].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    }
                    else
                    {
                        doc.Sections[0].Tables.Remove(doc.Sections[0].Tables[1]);
                    }
                    if (THPHKT == 1)
                    {
                        Table table11 = doc.Sections[1].Tables[1] as Table;
                        var listauditdetectRiskHight = auditdetect.Where(a => a.audit_report == true).OrderBy(a => a.admin_framework).ToArray();
                        for (int x = 0; x < listauditdetectRiskHight.Length; x++)
                        {
                            var auditdetectitem = listauditdetectRiskHight[x];
                            var adminFramework = auditdetectitem.admin_framework;
                            var str_adminFramework = "";
                            switch (adminFramework)
                            {
                                case 1:
                                    str_adminFramework = "Quản trị/Tổ chức/Số liệu tài chính";
                                    break;
                                case 2:
                                    str_adminFramework = "Hoạt động/Quy trình/Quy định";
                                    break;
                                case 3:
                                    str_adminFramework = "Kiểm sát vận hành/Thực thi";
                                    break;
                                case 4:
                                    str_adminFramework = "Công nghệ thông tinh";
                                    break;
                            }
                            var row = table11.AddRow();
                            var txtTable11 = row.Cells[0].AddParagraph().AppendText(str_adminFramework);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[0].CellFormat.BackColor = ColorTranslator.FromHtml("#FFFFFF");


                            var risk_level = "";
                            var color_ = "#FFFFFF";
                            switch (auditdetectitem.rating_risk)
                            {
                                case 1:
                                    risk_level = "Cao/Quan trọng";
                                    color_ = "#ff0000";
                                    break;
                                case 2:
                                    risk_level = "Trung bình";
                                    color_ = "#ffbf00";
                                    break;
                                case 3:
                                    risk_level = "Thấp/Ít quan trọng";
                                    color_ = "#92d050";
                                    break;
                            }

                            txtTable11 = row.Cells[1].AddParagraph().AppendText(auditdetectitem.code);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[1].CellFormat.BackColor = ColorTranslator.FromHtml(color_);


                            txtTable11 = row.Cells[2].AddParagraph().AppendText(risk_level);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[2].CellFormat.BackColor = ColorTranslator.FromHtml(color_);

                            var type_item = detecttype.FirstOrDefault(a => a.Id == auditdetectitem.classify_audit_detect)?.Name;
                            txtTable11 = row.Cells[3].AddParagraph().AppendText(type_item);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[3].CellFormat.BackColor = ColorTranslator.FromHtml("#FFFFFF");

                            txtTable11 = row.Cells[4].AddParagraph().AppendText(auditdetectitem.short_title);
                            txtTable11.CharacterFormat.FontName = "Times New Roman";
                            txtTable11.CharacterFormat.FontSize = 12;
                            row.Cells[4].CellFormat.BackColor = ColorTranslator.FromHtml("#FFFFFF");
                        }
                    }
                    else
                    {
                        doc.Sections[1].Tables.Remove(doc.Sections[1].Tables[1]);
                        //doc.Sections.Remove(doc.Sections[1]);
                    }
                    if (PLKT == 1)
                    {
                        TextSelection selections = doc.FindAllString("PLKQKT1", false, true)?.FirstOrDefault();
                        if (selections != null)
                        {
                            var Paragraph = selections?.GetAsOneRange()?.OwnerParagraph;
                            var sestionOwner = Paragraph.Owner.Owner as Section;
                            var listdetectNew = auditdetect.Where(a => a.classify_audit_detect.HasValue).AsEnumerable().GroupBy(a => a.classify_audit_detect);
                            var no_index = 0;
                            foreach (var item in listdetectNew)
                            {
                                var auditdetectitem = item;
                                var type_name = detecttype.FirstOrDefault(a => a.Id == auditdetectitem.Key)?.Name;
                                if (!string.IsNullOrEmpty(type_name))
                                {
                                    no_index++;

                                    var textheader = sestionOwner.AddParagraph().AppendText(no_index + ". " + type_name);
                                    textheader.CharacterFormat.FontName = "Times New Roman";
                                    textheader.CharacterFormat.FontSize = 11;
                                    textheader.CharacterFormat.Bold = true;

                                    Table table_new = sestionOwner.AddTable(true);
                                    table_new.ResetCells(item.Count() + 2, 5);
                                    table_new.ApplyVerticalMerge(0, 0, 1);
                                    table_new.ApplyHorizontalMerge(0, 1, 2);
                                    table_new.ApplyHorizontalMerge(0, 3, 4);

                                    var paragraph = table_new[0, 0].AddParagraph();
                                    table_new[0, 0].Width = 90;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    TextRange txtTable_new = paragraph.AppendText("Mã phát hiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;
                                    table_new[0, 0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                                    table_new[0, 0].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[0, 1].AddParagraph();
                                    table_new[0, 1].Width = 500;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Phát hiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[0, 1].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");
                                    table_new[0, 1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                                    paragraph = table_new[0, 3].AddParagraph();
                                    table_new[0, 3].Width = 500;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Kiến nghị/Khuyến nghị");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[0, 3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[0, 3].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 1].AddParagraph();
                                    table_new[1, 1].Width = 370;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Nội dung phát hiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[1, 1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 1].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 2].AddParagraph();
                                    table_new[1, 2].Width = 130;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Mức độ rủi ro");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[1, 2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 2].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 3].AddParagraph();
                                    table_new[1, 3].Width = 350;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Nội dung");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;
                                    table_new[1, 3].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 3].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");

                                    paragraph = table_new[1, 4].AddParagraph();
                                    table_new[1, 4].Width = 150;
                                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    txtTable_new = paragraph.AppendText("Thời hạn hoàn thiện");
                                    txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                    txtTable_new.CharacterFormat.FontSize = 11;
                                    txtTable_new.CharacterFormat.Bold = true;

                                    table_new[1, 4].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    table_new[1, 4].CellFormat.BackColor = ColorTranslator.FromHtml("#dbdbdb");
                                    var zz = 0;
                                    foreach (var item_row in item.ToList())
                                    {
                                        txtTable_new = table_new[zz + 2, 0].AddParagraph().AppendText(item_row.code);
                                        table_new[zz + 2, 0].Width = 90;
                                        txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                        txtTable_new.CharacterFormat.FontSize = 11;

                                        txtTable_new = table_new[zz + 2, 1].AddParagraph().AppendText(item_row.summary_audit_detect);
                                        table_new[zz + 2, 1].Width = 370;
                                        txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                        txtTable_new.CharacterFormat.FontSize = 11;

                                        var risk_level = "";
                                        switch (item_row.rating_risk)
                                        {
                                            case 1:
                                                risk_level = "Cao/Quan trọng";
                                                break;
                                            case 2:
                                                risk_level = "Trung bình";
                                                break;
                                            case 3:
                                                risk_level = "Thấp/Ít quan trọng";
                                                break;
                                        }
                                        txtTable_new = table_new[zz + 2, 2].AddParagraph().AppendText(risk_level);
                                        table_new[zz + 2, 2].Width = 130;
                                        txtTable_new.CharacterFormat.FontName = "Times New Roman";
                                        txtTable_new.CharacterFormat.FontSize = 11;

                                        var _auditRequestMonitor = _auditRequestMonitorList.Where(a => a.detectid == item_row.id).ToArray();
                                        var paragraph1 = table_new[zz + 2, 3].AddParagraph();
                                        var paragraph2 = table_new[zz + 2, 4].AddParagraph();
                                        for (int d = 0; d < _auditRequestMonitor.Length; d++)
                                        {
                                            var _auditRequestMonitorItem = _auditRequestMonitor[d];
                                            var txtTable_item = paragraph1.AppendText("- " + _auditRequestMonitorItem.Code + ": " + _auditRequestMonitorItem.Content + "\n");
                                            txtTable_item.CharacterFormat.FontName = "Times New Roman";
                                            txtTable_item.CharacterFormat.FontSize = 11;

                                            var date = _auditRequestMonitorItem.CompleteAt.HasValue ? _auditRequestMonitorItem.CompleteAt.Value.ToString("dd/MM/yyyy") : "";
                                            if (!string.IsNullOrEmpty(date))
                                            {
                                                txtTable_item = paragraph2.AppendText("- " + _auditRequestMonitorItem.Code + ": " + date + "\n");
                                                txtTable_item.CharacterFormat.FontName = "Times New Roman";
                                                txtTable_item.CharacterFormat.FontSize = 11;
                                            }
                                        }
                                        table_new[zz + 2, 3].Width = 350;
                                        table_new[zz + 2, 4].Width = 150;
                                        zz++;
                                    }
                                }
                                sestionOwner.AddParagraph();
                            }
                        }
                            doc.Replace("PLKQKT1", "", false, true);
                    }
                    else
                    {
                        doc.Sections.Remove(doc.Sections[5]);
                    }
                    foreach (Section section1 in doc.Sections)
                    {
                        for (int i = 0; i < section1.Body.ChildObjects.Count; i++)
                        {
                            if (section1.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
                            {
                                if (String.IsNullOrEmpty((section1.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                                {
                                    section1.Body.ChildObjects.Remove(section1.Body.ChildObjects[i]);
                                    i--;
                                }
                            }

                        }
                    }

                    MemoryStream stream = new MemoryStream();
                    stream.Position = 0;
                    doc.SaveToFile(stream, Spire.Doc.FileFormat.Docx);
                    Bytes = stream.ToArray();
                    return File(Bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Kitano_BaoCaoKiemToan.docx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"ExportAudit - {DateTime.Now} : {ex.Message}!");
                return BadRequest();
            }
        }

        [HttpGet("ExportFileWordAMC/{id}")]
        public IActionResult ExportFileWordAMC(int id)
        {
            byte[] Bytes = null;
            //try
            //{
            if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
            {
                return Unauthorized();
            }
            var checkReport = _uow.Repository<ReportAuditWork>().Include(x => x.AuditWork).FirstOrDefault(a => a.Id == id && a.IsDeleted.Equals(false));

            var auditWorkScope = _uow.Repository<AuditWorkScope>().GetAll().Where(a => a.auditwork_id == checkReport.auditwork_id && a.IsDeleted != true).ToArray();

            var lstauditworkscope = auditWorkScope.Where(a => a.auditprocess_id.HasValue).Select(a => new AuditWorkScopeModel()
            {
                ID = a.Id,
                AuditProcess = a.auditprocess_name,
                AuditFacility = a.auditfacilities_name,
                AuditActivity = a.bussinessactivities_name,
            }).ToList();

            var approval_status = _uow.Repository<ApprovalFunction>().FirstOrDefault(a => a.item_id == id && a.function_code == "M_RAW");
            var auditwork = _uow.Repository<AuditWork>().Include(a => a.AuditWorkScope, a => a.Users).FirstOrDefault(a => a.Id == checkReport.auditwork_id);
            //var auditdetect = (from a in _uow.Repository<AuditDetect>().Find(x => x.IsDeleted != true && x.auditwork_id == checkReport.auditwork_id && x.audit_report == true)
            //                   join b in _uow.Repository<ApprovalFunction>().Find(x => x.function_code == "M_AD" && x.StatusCode == "3.1") on a.id equals b.item_id
            //                   select a).ToArray();
            var auditdetect = _uow.Repository<AuditDetect>().Include(x => x.CatDetectType).Where(x => x.IsDeleted != true && x.auditwork_id == checkReport.auditwork_id && x.audit_report == true).OrderBy(x => x.admin_framework).ToArray();
            var allAuditDetectId = auditdetect.Select(x => x.id.ToString());
            var auditrequestMonitor = _uow.Repository<AuditRequestMonitor>().Include(x => x.AuditFacility, x => x.FacilityRequestMonitorMapping, x => x.CatAuditRequest).Where(x => x.detectid.HasValue ? allAuditDetectId.Contains(x.detectid.ToString()) : false && x.is_deleted != true).ToArray();
            var grpAuditrequestMonitor = auditrequestMonitor.GroupBy(x => x.CatAuditRequest).Select(grpType => grpType).ToArray();

            var detecttype = _uow.Repository<CatDetectType>().Find(x => x.IsDeleted != true).ToArray();
            var auditworkscope = auditWorkScope.AsEnumerable().GroupBy(a => new { a.auditprocess_id, a.auditfacilities_id, a.bussinessactivities_id }).Select(g => g.FirstOrDefault()).ToArray();

            var listfacility = _uow.Repository<AuditWorkScopeFacility>().Find(a => a.auditfacilities_id.HasValue && a.IsDeleted != true && a.auditwork_id == checkReport.auditwork_id).ToArray();

            var listdetect = auditdetect.Where(a => a.classify_audit_detect.HasValue).AsEnumerable().GroupBy(a => a.classify_audit_detect).Select(g => g.FirstOrDefault());

         
            var auditminutes = _uow.Repository<AuditMinutes>().Find(a => a.auditwork_id == checkReport.auditwork_id && a.IsDeleted != true).ToArray();
            var lstAuditworkfacility = listfacility.Select(a => new AuditWorkFacilityModel()
            {
                ID = a.Id,
                AuditFacilityId = a.auditfacilities_id,
                AuditFacility = a.auditfacilities_name,
                AuditRatingLevelItem = auditminutes.FirstOrDefault(x => x.auditfacilities_id == a.auditfacilities_id)?.rating_level_total,
                BaseRatingItem = a.BaseRatingReport,
                Idea = auditminutes.FirstOrDefault(x => x.auditfacilities_id == a.auditfacilities_id)?.idea,
            }).ToList();

            var lstDetectType = listdetect.Select(a => new SumaryDetectModel()
            {
                classify_audit_detect = a.classify_audit_detect,
                classify_audit_detect_name = detecttype.Where(x => x.Id == a.classify_audit_detect).Select(c => c.Name).FirstOrDefault().ToString(),
                Hight = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 1),
                Middle = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 2),
                Low = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.rating_risk == 3),
                Agree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is true),
                NotAgree = auditdetect.Where(x => x.classify_audit_detect == a.classify_audit_detect).Count(x => x.opinion_audit is false),

            }).ToList();

            var countRiskRatingHight = lstDetectType.Select(a => a.Hight).ToArray();

            var countRiskRatingMedium = lstDetectType.Select(a => a.Middle).ToArray();

            var countRiskRatingLow = lstDetectType.Select(a => a.Low).ToArray();

            var countOpinionAuditTrue = lstDetectType.Select(a => a.Agree).ToArray();

            var countOpinionAuditFalse = lstDetectType.Select(a => a.NotAgree).ToArray();

            var ListStatisticsOfDetections = listdetect.Select(a => new ListStatisticsOfDetections
            {
                id = a.id,
                audit_detect_code = a.code,
                auditwork_id = a.auditwork_id,
                auditfacilities_id = a.auditfacilities_id,
                auditfacilities_name = a.auditfacilities_name,//đơn vị kiểm toán
                title = a.title,//Tiêu đề phát hiện
                year = a.year,
                risk_rating = a.rating_risk,
                str_classify_audit_detect = a.CatDetectType.Name,
                reason = a.reason,
                opinion_audit_true = countOpinionAuditTrue.Count(),
                opinion_audit_false = countOpinionAuditFalse.Count(),
                risk_rating_hight = countRiskRatingHight.Count(),
                risk_rating_medium = countRiskRatingMedium.Count(),
                risk_rating_low = countRiskRatingLow.Count(),
            }).GroupBy(a => a.str_classify_audit_detect).Select(z => z.FirstOrDefault()).ToList();


            // Export word
            var fullPath = Path.Combine(_config["Template:ReporDocsTemplate"], "AMC_Kitano_BaoCaoCuocKiemToan_v0.1.docx");
            fullPath = fullPath.ToString().Replace("\\", "/");
            using (Document doc = new Document(fullPath))
            //using (Document doc = new Document(@"D:\test\Kitano_BaoCaoKiemToan_v0.2.docx"))
            {
                var font = "Times New Roman";

                ParagraphStyle contentStyle = new ParagraphStyle(doc);
                contentStyle.Name = "contentstyle";
                contentStyle.CharacterFormat.FontName = font;
                contentStyle.CharacterFormat.FontSize = 12;
                contentStyle.ParagraphFormat.BeforeSpacing = 0;
                contentStyle.ParagraphFormat.AfterSpacing = 0;
                contentStyle.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                contentStyle.ParagraphFormat.LineSpacing = 14.4360902256f;//line height 1.2 trong word
                doc.Styles.Add(contentStyle);

                var risk_color_high = "#FF0000";
                var risk_color_medium = "#ffff00";
                var risk_color_low = "#91d0b4";

                ListStyle listStyle = new ListStyle(doc, ListType.Numbered);
                listStyle.Name = "levelstyle";
                listStyle.Levels[0].PatternType = ListPatternType.Arabic;
                listStyle.Levels[1].NumberPrefix = "\x0000.";
                listStyle.Levels[1].PatternType = ListPatternType.Arabic;
                listStyle.Levels[0].CharacterFormat.FontName = font;
                listStyle.Levels[0].CharacterFormat.FontSize = 12;
                listStyle.Levels[0].CharacterFormat.Bold = true;
                listStyle.Levels[0].ParagraphFormat.BeforeSpacing = 0;
                listStyle.Levels[0].ParagraphFormat.AfterSpacing = 0;
                listStyle.Levels[0].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                listStyle.Levels[0].ParagraphFormat.LineSpacing = 14.4360902256f;
                listStyle.Levels[1].CharacterFormat.FontName = font;
                listStyle.Levels[1].CharacterFormat.FontSize = 12;
                listStyle.Levels[1].ParagraphFormat.BeforeSpacing = 0;
                listStyle.Levels[1].ParagraphFormat.AfterSpacing = 0;
                listStyle.Levels[1].ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                listStyle.Levels[1].ParagraphFormat.LineSpacing = 14.4360902256f;
                listStyle.Levels[1].CharacterFormat.Bold = true;

                doc.ListStyles.Add(listStyle);

                ListStyle bulletStyle = new ListStyle(doc, ListType.Bulleted);
                bulletStyle.Name = "bulletstyle";
                bulletStyle.Levels[0].BulletCharacter = "-";
                bulletStyle.Levels[0].CharacterFormat.FontName = font;
                bulletStyle.Levels[0].CharacterFormat.FontSize = 12;
                bulletStyle.Levels[0].TextPosition = 28.08f;//28.08 = 0.39 inch trong word
                bulletStyle.Levels[0].NumberPosition = 0;
                doc.ListStyles.Add(bulletStyle);

                var MainStage = _uow.Repository<MainStage>().Find(a => a.auditwork_id == checkReport.auditwork_id).ToArray();
                var startdate = MainStage.FirstOrDefault(a => a.index == 1);
                var enddate = MainStage.FirstOrDefault(a => a.index == 4);
                var mucdoxephang = getRatingLevelName(checkReport.AuditRatingLevelTotal);
                var cosoxephang = checkReport.BaseRatingTotal;
                var ketluanchung = checkReport.GeneralConclusions;



                Section sectionRequestMonitor = doc.Sections[3];

                Table table1 = doc.Sections[0].Tables[3] as Table;
                Table table2 = doc.Sections[0].Tables[4] as Table;
                Table table3 = doc.Sections[1].Tables[1] as Table;
                Table table4 = doc.Sections[2].Tables[1] as Table;

                List<string> facilityInRequestMonitor = new List<string>();


                for (int i = 0; i < sectionRequestMonitor.Body.ChildObjects.Count; i++)
                {
                    if (sectionRequestMonitor.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        if (String.IsNullOrEmpty((sectionRequestMonitor.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                        {
                            sectionRequestMonitor.Body.ChildObjects.Remove(sectionRequestMonitor.Body.ChildObjects[i]);
                            i--;
                        }
                    }

                }


                int? total = 0;
                for (int z = 0; z < lstDetectType.Count(); z++)
                {
                    total = total + lstDetectType[z].Hight + lstDetectType[z].Middle + lstDetectType[z].Low;

                    var row = table1.AddRow();
                    var para = row.Cells[0].AddParagraph();
                    var txtTable1 = para.AppendText(lstDetectType[z].classify_audit_detect_name);
                    para.ApplyStyle("contentstyle");
                    txtTable1.CharacterFormat.FontSize = 10.5f;

                    para = row.Cells[1].AddParagraph();
                    txtTable1 = para.AppendText((lstDetectType[z].Hight + lstDetectType[z].Middle + lstDetectType[z].Low).ToString());
                    para.ApplyStyle("contentstyle");
                    txtTable1.CharacterFormat.FontSize = 10.5f;
                    para.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    para = row.Cells[2].AddParagraph();
                    txtTable1 = para.AppendText(lstDetectType[z].Hight.ToString());
                    para.ApplyStyle("contentstyle");
                    row.Cells[2].CellFormat.BackColor = ColorTranslator.FromHtml(risk_color_high);
                    txtTable1.CharacterFormat.FontSize = 10.5f;
                    para.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    para = row.Cells[3].AddParagraph();
                    txtTable1 = para.AppendText(lstDetectType[z].Middle.ToString());
                    para.ApplyStyle("contentstyle");
                    row.Cells[3].CellFormat.BackColor = ColorTranslator.FromHtml(risk_color_medium);
                    txtTable1.CharacterFormat.FontSize = 10.5f;
                    para.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    para = row.Cells[4].AddParagraph();
                    txtTable1 = para.AppendText(lstDetectType[z].Low.ToString());
                    row.Cells[4].CellFormat.BackColor = ColorTranslator.FromHtml(risk_color_low);
                    para.ApplyStyle("contentstyle");
                    txtTable1.CharacterFormat.FontSize = 10.5f;
                    para.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    para = row.Cells[5].AddParagraph();
                    txtTable1 = para.AppendText(lstDetectType[z].Agree.ToString());
                    para.ApplyStyle("contentstyle");
                    txtTable1.CharacterFormat.FontSize = 10.5f;
                    para.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    para = row.Cells[6].AddParagraph();
                    txtTable1 = para.AppendText(lstDetectType[z].NotAgree.ToString());
                    para.ApplyStyle("contentstyle");
                    txtTable1.CharacterFormat.FontSize = 10.5f;
                    para.Format.HorizontalAlignment = HorizontalAlignment.Center;
                }

                var rowTotal = table1.AddRow();
                var paraTotal = rowTotal.Cells[0].AddParagraph();
                var txtTable1Total = paraTotal.AppendText("Mức đánh giá");
                paraTotal.ApplyStyle("contentstyle");
                txtTable1Total.CharacterFormat.FontSize = 10.5f;
                txtTable1Total.CharacterFormat.Bold = true;

                paraTotal = rowTotal.Cells[1].AddParagraph();
                txtTable1Total = paraTotal.AppendText(total.ToString());
                paraTotal.ApplyStyle("contentstyle");
                txtTable1Total.CharacterFormat.FontSize = 10.5f;
                txtTable1Total.CharacterFormat.Bold = true;
                paraTotal.Format.HorizontalAlignment = HorizontalAlignment.Center;

                paraTotal = rowTotal.Cells[2].AddParagraph();
                txtTable1Total = paraTotal.AppendText(mucdoxephang);
                paraTotal.ApplyStyle("contentstyle");
                txtTable1Total.CharacterFormat.FontSize = 10.5f;
                txtTable1Total.CharacterFormat.Bold = true;
                paraTotal.Format.HorizontalAlignment = HorizontalAlignment.Center;
                rowTotal.Cells[2].CellFormat.BackColor = ColorTranslator.FromHtml("#ffff00");

                table1.Rows.RemoveAt(2);
                table1.ApplyHorizontalMerge(lstDetectType.Count + 2, 2, 6);


                foreach (var item in lstAuditworkfacility)
                {
                    var row = table2.AddRow();

                    var paraTable2 = row.Cells[0].AddParagraph();
                    var txtParaTable2 = paraTable2.AppendText(item.AuditFacility);
                    paraTable2.ApplyStyle("contentstyle");
                    paraTable2.Format.HorizontalAlignment = HorizontalAlignment.Left;

                    paraTable2 = row.Cells[1].AddParagraph();
                    txtParaTable2 = paraTable2.AppendText(item.Idea);
                    paraTable2.ApplyStyle("contentstyle");
                    paraTable2.Format.HorizontalAlignment = HorizontalAlignment.Left;
                }
                table2.Rows.RemoveAt(0);

                foreach (var item in auditdetect)
                {
                    var str_adminFramework = "";
                    var risk_level = "";
                    var color_ = "";

                    switch (item.admin_framework)
                    {
                        case 1:
                            str_adminFramework = "Quản trị/Tổ chức/Số liệu tài chính";
                            break;
                        case 2:
                            str_adminFramework = "Hoạt động/Quy trình/Quy định";
                            break;
                        case 3:
                            str_adminFramework = "Kiểm sát vận hành/Thực thi";
                            break;
                        case 4:
                            str_adminFramework = "Công nghệ thông tin";
                            break;
                    }

                    switch (item.rating_risk)
                    {
                        case 1:
                            risk_level = "Cao/Quan trọng";
                            color_ = risk_color_high;
                            break;
                        case 2:
                            risk_level = "Trung bình";
                            color_ = risk_color_medium;
                            break;
                        case 3:
                            risk_level = "Thấp/ít quan trọng";
                            color_ = risk_color_low;
                            break;
                    }

                    var row = table3.AddRow();

                    var paraTable3 = row.Cells[0].AddParagraph();
                    var txtParaTable3 = paraTable3.AppendText(str_adminFramework);
                    paraTable3.ApplyStyle("contentstyle");
                    paraTable3.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable3.CharacterFormat.FontSize = 11;

                    paraTable3 = row.Cells[1].AddParagraph();
                    txtParaTable3 = paraTable3.AppendText(item.code);
                    paraTable3.ApplyStyle("contentstyle");
                    paraTable3.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable3.CharacterFormat.FontSize = 11;
                    row.Cells[1].CellFormat.BackColor = ColorTranslator.FromHtml(color_);

                    paraTable3 = row.Cells[2].AddParagraph();
                    txtParaTable3 = paraTable3.AppendText(risk_level);
                    paraTable3.ApplyStyle("contentstyle");
                    paraTable3.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable3.CharacterFormat.FontSize = 11;
                    row.Cells[2].CellFormat.BackColor = ColorTranslator.FromHtml(color_);

                    paraTable3 = row.Cells[3].AddParagraph();
                    txtParaTable3 = paraTable3.AppendText(item.CatDetectType.Name);
                    paraTable3.ApplyStyle("contentstyle");
                    paraTable3.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable3.CharacterFormat.FontSize = 11;

                    paraTable3 = row.Cells[4].AddParagraph();
                    txtParaTable3 = paraTable3.AppendText(item.short_title);
                    paraTable3.ApplyStyle("contentstyle");
                    paraTable3.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable3.CharacterFormat.FontSize = 11;

                    row = table4.AddRow();

                    var paraTable4 = row.Cells[0].AddParagraph();
                    var txtParaTable4 = paraTable4.AppendText(str_adminFramework);
                    paraTable4.ApplyStyle("contentstyle");
                    paraTable4.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable4.CharacterFormat.FontSize = 11;

                    paraTable4 = row.Cells[1].AddParagraph();
                    txtParaTable4 = paraTable4.AppendText(item.code);
                    paraTable4.ApplyStyle("contentstyle");
                    paraTable4.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable4.CharacterFormat.FontSize = 11;
                    row.Cells[1].CellFormat.BackColor = ColorTranslator.FromHtml(color_);

                    paraTable4 = row.Cells[2].AddParagraph();
                    txtParaTable4 = paraTable4.AppendText(risk_level);
                    paraTable4.ApplyStyle("contentstyle");
                    paraTable4.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable4.CharacterFormat.FontSize = 11;
                    row.Cells[2].CellFormat.BackColor = ColorTranslator.FromHtml(color_);

                    paraTable4 = row.Cells[3].AddParagraph();
                    txtParaTable4 = paraTable4.AppendText(item.CatDetectType.Name);
                    paraTable4.ApplyStyle("contentstyle");
                    paraTable4.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable4.CharacterFormat.FontSize = 11;


                    paraTable4 = row.Cells[4].AddParagraph();
                    txtParaTable4 = paraTable4.AppendText(item.summary_audit_detect);
                    paraTable4.ApplyStyle("contentstyle");
                    paraTable4.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    txtParaTable4.CharacterFormat.FontSize = 11;

                }

                table3.Rows.RemoveAt(1);
                table4.Rows.RemoveAt(1);

                foreach (var monitorGrp in grpAuditrequestMonitor)
                {
                    var paraSectionRequestMonitor = sectionRequestMonitor.AddParagraph();
                    var txtParaSectionRequestMonitor = paraSectionRequestMonitor.AppendText(monitorGrp?.Key?.Name ?? "");
                    paraSectionRequestMonitor.ApplyStyle("contentstyle");
                    paraSectionRequestMonitor.ListFormat.ApplyStyle("levelstyle");
                    txtParaSectionRequestMonitor.CharacterFormat.Bold = true;
                    paraSectionRequestMonitor.Format.FirstLineIndent = -18;
                    paraSectionRequestMonitor.Format.LeftIndent = 25.2f;

                    var monitorGrpWithFacilityGrp = monitorGrp.GroupBy(x => x.FacilityRequestMonitorMapping.Where(x => x.type == 1).FirstOrDefault().audit_facility_name).Select(grpFac => grpFac.ToArray()).ToArray();
                    foreach (var monitorFacGrp in monitorGrpWithFacilityGrp)
                    {
                        var facility = monitorFacGrp?.FirstOrDefault()?.FacilityRequestMonitorMapping?.FirstOrDefault(x => x.type == 1)?.audit_facility_name ?? "";
                        if (!string.IsNullOrEmpty(facility))
                        {
                            facilityInRequestMonitor.Add(facility);
                        }
                        paraSectionRequestMonitor = sectionRequestMonitor.AddParagraph();
                        txtParaSectionRequestMonitor = paraSectionRequestMonitor.AppendText(" " + facility);
                        paraSectionRequestMonitor.ListFormat.ListLevelNumber = 1;
                        paraSectionRequestMonitor.ListFormat.ApplyStyle("levelstyle");
                        paraSectionRequestMonitor.ApplyStyle("contentstyle");
                        txtParaSectionRequestMonitor.CharacterFormat.Bold = true;
                        paraSectionRequestMonitor.Format.FirstLineIndent = 0;
                        paraSectionRequestMonitor.Format.LeftIndent = 25.2f;
                        foreach (var monitor in monitorFacGrp)
                        {
                            paraSectionRequestMonitor = sectionRequestMonitor.AddParagraph();
                            txtParaSectionRequestMonitor = paraSectionRequestMonitor.AppendText(monitor.Content);
                            paraSectionRequestMonitor.ApplyStyle("contentstyle");
                            paraSectionRequestMonitor.ListFormat.ApplyStyle("bulletstyle");
                            paraSectionRequestMonitor.Format.FirstLineIndent = -18;
                            paraSectionRequestMonitor.Format.LeftIndent = 25.2f;
                        }
                    }
                }
                sectionRequestMonitor.AddParagraph();

                facilityInRequestMonitor = facilityInRequestMonitor.Distinct().ToList();

                //remove empty paragraphs
                doc.MailMerge.HideEmptyParagraphs = true;
                //remove empty group
                doc.MailMerge.HideEmptyGroup = true;
                doc.MailMerge.Execute(
                new string[] { "ngay_kt", "thang_kt", "nam_kt",
                    "ten_kt",
                    "muc_tieu_kt",
                    "pham_vi_kt",
                    "gioi_han_kt",
                    "xep_hang_kt",
                    "co_so_xep_hang",
                    "ket_luan_chung",
                    "don_vi_thuc_hien",

                },
                new string[] { DateTime.Now.Day.ToString() , DateTime.Now.Month.ToString() , DateTime.Now.Year.ToString(),
                    checkReport.AuditWorkName,
                    checkReport.AuditWork?.Target,
                    checkReport.AuditWork?.AuditScope,
                    checkReport.AuditWork?.AuditScopeOutside,
                    mucdoxephang,
                    cosoxephang,
                    ketluanchung,
                    string.Join( ", ",facilityInRequestMonitor)
                }
                );

                MemoryStream stream = new MemoryStream();
                stream.Position = 0;
                doc.SaveToFile(stream, Spire.Doc.FileFormat.Docx);
                Bytes = stream.ToArray();
                return File(Bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Kitano_BaoCaoKiemToan.docx");
            }
            //}
            //catch (Exception ex)
            //{
            //    _logger.LogError($"ExportAudit - {DateTime.Now} : {ex.Message}!");
            //    return BadRequest();
            //}
        }

        [HttpGet("ExportFileWordMCREDIT/{id}")]
        public IActionResult ExportFileWordMCREDIT(int id)
        {
            byte[] Bytes = null;
            try
            {
                var fullPath = Path.Combine(_config["Template:ReporDocsTemplate"], "MCredit_Kitano_BaoCaoKiemToan_v0.1.docx");
                fullPath = fullPath.ToString().Replace("\\", "/");
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var checkReport = _uow.Repository<ReportAuditWork>().Include(x => x.AuditWork).FirstOrDefault(a => a.Id == id && a.IsDeleted.Equals(false));

                var auditWorkScope = _uow.Repository<AuditWorkScope>().GetAll().Where(a => a.auditwork_id == checkReport.auditwork_id && a.IsDeleted != true).ToArray();


                var hearderSystem = _uow.Repository<SystemParameter>().FirstOrDefault(x => x.Parameter_Name == "REPORT_HEADER");
                var day = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("dd") : "...";
                var month = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("MM") : "...";
                var year = checkReport.CreatedAt != null ? checkReport.CreatedAt.Value.ToString("yyyy") : "...";

                var approval_status = _uow.Repository<ApprovalFunction>().FirstOrDefault(a => a.item_id == id && a.function_code == "M_RAW");
                var auditwork = _uow.Repository<AuditWork>().Include(a => a.AuditWorkScope, a => a.Users).FirstOrDefault(a => a.Id == checkReport.auditwork_id);
                //var auditdetect = (from a in _uow.Repository<AuditDetect>().Find(x => x.IsDeleted != true && x.auditwork_id == checkReport.auditwork_id && x.audit_report == true)
                //                   join b in _uow.Repository<ApprovalFunction>().Find(x => x.function_code == "M_AD" && x.StatusCode == "3.1") on a.id equals b.item_id
                //                   select a).ToArray();
                var auditdetect = _uow.Repository<AuditDetect>().Include(x => x.AuditRequestMonitor).Where(x => x.IsDeleted != true & x.auditwork_id == checkReport.auditwork_id).ToArray();
                var auditDetectId = auditdetect.Select(x => x.id).ToArray();
                var facilityRequestMonitorMapping = _uow.Repository<FacilityRequestMonitorMapping>().Find(x => x.AuditRequestMonitor.detectid != null && auditDetectId.Contains((int)x.AuditRequestMonitor.detectid)).ToArray();

                // Export word

                using (Document doc = new Document(fullPath))
                //using (Document doc = new Document(@"D:\test\Kitano_BaoCaoKiemToan_v0.2.docx"))
                {
                    //Header
                    doc.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true;
                    Paragraph paragraph_header = doc.Sections[0].HeadersFooters.FirstPageHeader.AddParagraph();
                    paragraph_header.AppendHTML(hearderSystem.Value);

                    //Remove header 2
                    doc.Sections[1].HeadersFooters.Header.ChildObjects.Clear();
                    //Add line breaks
                    Paragraph headerParagraph = doc.Sections[1].HeadersFooters.FirstPageHeader.AddParagraph();
                    headerParagraph.AppendBreak(BreakType.LineBreak);
                    //Remove header 3
                    doc.Sections[2].HeadersFooters.Header.ChildObjects.Clear();
                    //Add line breaks
                    Paragraph headerParagraph2 = doc.Sections[2].HeadersFooters.FirstPageHeader.AddParagraph();
                    headerParagraph2.AppendBreak(BreakType.LineBreak);

                    var MainStage = _uow.Repository<MainStage>().Find(a => a.auditwork_id == checkReport.auditwork_id).ToArray();
                    var startdate = MainStage.FirstOrDefault(a => a.index == 1);
                    var enddate = MainStage.FirstOrDefault(a => a.index == 4);

                    //remove empty paragraphs
                    doc.MailMerge.HideEmptyParagraphs = true;
                    //remove empty group
                    doc.MailMerge.HideEmptyGroup = true;
                    doc.MailMerge.Execute(
                    new string[] { "ngay_bc_1" , "thang_bc_1"  , "nam_bc_1" ,
                        "nam_kt_1",
                        "ten_kt_up_1",
                        "ten_kt_1",
                        "muc_dich_kt_1",
                        "pham_vi_kt_1",
                        "ngoai_pham_vi_kt_1",
                        "thoi_hieu_kt_tu_1",
                        "thoi_hieu_kt_den_1",
                        "thoi_gian_kt_tu_1",
                        "thoi_gian_kt_den_1",
                        "muc_xep_hang_kiem_toan_1",
                        "co_so_xep_hang_1",
                        "ket_luan_chung_1",
                    },
                    new string[] { DateTime.Now.Day.ToString() , DateTime.Now.Month.ToString() , DateTime.Now.Year.ToString() ,
                    checkReport.Year,
                    checkReport.AuditWorkName != null ? checkReport.AuditWorkName.ToUpper() : "",
                    checkReport.AuditWorkName,
                    checkReport.AuditWork.Target,
                    checkReport.AuditWork.AuditScope,
                    checkReport.AuditWork.AuditScopeOutside,
                    auditwork.from_date != null ? auditwork.from_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy",
                    auditwork.to_date != null ?  auditwork.to_date.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy",
                    checkReport.AuditWork.StartDate != null && checkReport.AuditWork.StartDate.HasValue ? checkReport.AuditWork.StartDate.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy",
                    checkReport.AuditWork.EndDate != null && checkReport.AuditWork.EndDate.HasValue ? checkReport.AuditWork.EndDate.Value.ToString("dd/MM/yyyy"): "dd/MM/yyyy",
                    getRatingLevelName(checkReport.AuditRatingLevelTotal),
                    checkReport.BaseRatingTotal,
                    checkReport.GeneralConclusions
                    }
                    );

                    var table1 = doc.Sections[0].Tables[1];
                    var font = "Be Vietnam Pro";
                    var sectionDetectDetails = doc.Sections[1];

                    ListStyle listStyle = new ListStyle(doc, ListType.Numbered);
                    listStyle.Name = "levelstyle";
                    listStyle.Levels[0].PatternType = ListPatternType.Arabic;
                    listStyle.Levels[0].TextPosition = 28.08f;//28.08 = 0.39 inch trong word
                    listStyle.Levels[0].NumberPosition = 0;
                    doc.ListStyles.Add(listStyle);

                    var listauditdetectRiskHight = auditdetect.Where(a => a.audit_report == true && a.IsDeleted != true).OrderBy(a => a.admin_framework).ToArray();
                    for (int x = 0; x < listauditdetectRiskHight.Length; x++)
                    {
                        var row = table1.AddRow();
                        var parTxtTable1 = row.Cells[0].AddParagraph();
                        var txtTable1 = parTxtTable1.AppendText(x + 1 + ".");
                        parTxtTable1.ApplyStyle(BuiltinStyle.Normal);
                        parTxtTable1.Format.HorizontalAlignment = HorizontalAlignment.Center;

                        parTxtTable1 = row.Cells[1].AddParagraph();
                        txtTable1 = parTxtTable1.AppendText(listauditdetectRiskHight[x].title);
                        parTxtTable1.ApplyStyle(BuiltinStyle.Normal);
                        parTxtTable1.Format.HorizontalAlignment = HorizontalAlignment.Left;

                        var risk_level = "";
                        var color_ = "#FFFFFF";
                        switch (listauditdetectRiskHight[x].rating_risk)
                        {
                            case 1:
                                risk_level = "Cao";
                                color_ = "#e10000";
                                break;
                            case 2:
                                risk_level = "Trung bình";
                                color_ = "#e1bf00";
                                break;
                            case 3:
                                risk_level = "Thấp";
                                color_ = "#70ad47";
                                break;
                        }

                        parTxtTable1 = row.Cells[2].AddParagraph();
                        txtTable1 = parTxtTable1.AppendText(risk_level);
                        parTxtTable1.ApplyStyle(BuiltinStyle.Normal);
                        parTxtTable1.Format.HorizontalAlignment = HorizontalAlignment.Center;

                        var paraSectionDetectDetails = sectionDetectDetails.AddParagraph();
                        var txtParaSectionDetectDetails = paraSectionDetectDetails.AppendText(listauditdetectRiskHight[x].title + " - ");
                        txtParaSectionDetectDetails = paraSectionDetectDetails.AppendText(risk_level);
                        txtParaSectionDetectDetails.CharacterFormat.TextColor = ColorTranslator.FromHtml(color_);
                        txtParaSectionDetectDetails.CharacterFormat.Italic = true;
                        paraSectionDetectDetails.ApplyStyle(BuiltinStyle.Heading2);
                        paraSectionDetectDetails.ListFormat.ApplyStyle("levelstyle");
                        paraSectionDetectDetails.Format.FirstLineIndent = -28.08f;//28.08 = 0.39 inch trong word

                        paraSectionDetectDetails = sectionDetectDetails.AddParagraph();
                        txtParaSectionDetectDetails = paraSectionDetectDetails.AppendText("Phát hiện:");
                        paraSectionDetectDetails.ApplyStyle(BuiltinStyle.Title);

                        paraSectionDetectDetails = sectionDetectDetails.AddParagraph();
                        txtParaSectionDetectDetails = paraSectionDetectDetails.AppendText(listauditdetectRiskHight[x].description);
                        paraSectionDetectDetails.ApplyStyle(BuiltinStyle.Subtitle);

                        paraSectionDetectDetails = sectionDetectDetails.AddParagraph();
                        txtParaSectionDetectDetails = paraSectionDetectDetails.AppendText("Rủi ro:");
                        paraSectionDetectDetails.ApplyStyle(BuiltinStyle.Title);

                        paraSectionDetectDetails = sectionDetectDetails.AddParagraph();
                        txtParaSectionDetectDetails = paraSectionDetectDetails.AppendText(listauditdetectRiskHight[x].affect);
                        paraSectionDetectDetails.ApplyStyle(BuiltinStyle.Subtitle);

                        paraSectionDetectDetails = sectionDetectDetails.AddParagraph();
                        txtParaSectionDetectDetails = paraSectionDetectDetails.AppendText("Kiến nghị:");
                        paraSectionDetectDetails.ApplyStyle(BuiltinStyle.Title);

                        var cellcolor = "#9cc2e5";

                        Table table2 = sectionDetectDetails.AddTable(true);
                        table2.ResetCells(listauditdetectRiskHight[x].AuditRequestMonitor.Count() + 1, 4);

                        var cell = table2.Rows[0].Cells[0];
                        var paraCellTable2 = cell.AddParagraph();
                        var txtParaCellTable2 = paraCellTable2.AppendText("Kiến nghị của KTNB");
                        cell.CellFormat.BackColor = ColorTranslator.FromHtml(cellcolor);
                        paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtParaCellTable2.CharacterFormat.FontName = font;
                        txtParaCellTable2.CharacterFormat.FontSize = 10;
                        txtParaCellTable2.CharacterFormat.Bold = true;

                        cell = table2.Rows[0].Cells[1];
                        paraCellTable2 = cell.AddParagraph();
                        txtParaCellTable2 = paraCellTable2.AppendText("Đơn vị thực hiện");
                        cell.CellFormat.BackColor = ColorTranslator.FromHtml(cellcolor);
                        paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtParaCellTable2.CharacterFormat.FontName = font;
                        txtParaCellTable2.CharacterFormat.FontSize = 10;
                        txtParaCellTable2.CharacterFormat.Bold = true;

                        cell = table2.Rows[0].Cells[2];
                        paraCellTable2 = cell.AddParagraph();
                        txtParaCellTable2 = paraCellTable2.AppendText("Ý kiến của đơn vị");
                        cell.CellFormat.BackColor = ColorTranslator.FromHtml(cellcolor);
                        paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtParaCellTable2.CharacterFormat.FontName = font;
                        txtParaCellTable2.CharacterFormat.FontSize = 10;
                        txtParaCellTable2.CharacterFormat.Bold = true;

                        cell = table2.Rows[0].Cells[3];
                        paraCellTable2 = cell.AddParagraph();
                        txtParaCellTable2 = paraCellTable2.AppendText("Thời hạn hoàn thành");
                        cell.CellFormat.BackColor = ColorTranslator.FromHtml(cellcolor);
                        paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        txtParaCellTable2.CharacterFormat.FontName = font;
                        txtParaCellTable2.CharacterFormat.FontSize = 10;
                        txtParaCellTable2.CharacterFormat.Bold = true;



                        foreach (var (item, i) in listauditdetectRiskHight[x].AuditRequestMonitor.Select((item, i) => (item, i)))
                        {
                            var facilitymapping = facilityRequestMonitorMapping.Where(x => item.Id == x.audit_request_monitor_id).Select(x => x.audit_facility_name).Distinct().ToArray();

                            cell = table2.Rows[i + 1].Cells[0];
                            paraCellTable2 = cell.AddParagraph();
                            txtParaCellTable2 = paraCellTable2.AppendText(item.Content);
                            cell.CellFormat.BackColor = ColorTranslator.FromHtml("#ffffff");
                            txtParaCellTable2.CharacterFormat.FontName = font;
                            txtParaCellTable2.CharacterFormat.FontSize = 10;
                            paraCellTable2.Format.LineSpacing = 15.66f;
                            paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Left;

                            cell = table2.Rows[i + 1].Cells[1];
                            paraCellTable2 = cell.AddParagraph();
                            txtParaCellTable2 = paraCellTable2.AppendText(string.Join(", ", facilitymapping));
                            cell.CellFormat.BackColor = ColorTranslator.FromHtml("#ffffff");
                            txtParaCellTable2.CharacterFormat.FontName = font;
                            txtParaCellTable2.CharacterFormat.FontSize = 10;
                            paraCellTable2.Format.LineSpacing = 15.66f;
                            paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Left;

                            cell = table2.Rows[i + 1].Cells[2];
                            paraCellTable2 = cell.AddParagraph();
                            txtParaCellTable2 = paraCellTable2.AppendText(item.note);
                            cell.CellFormat.BackColor = ColorTranslator.FromHtml("#ffffff");
                            txtParaCellTable2.CharacterFormat.FontName = font;
                            txtParaCellTable2.CharacterFormat.FontSize = 10;
                            paraCellTable2.Format.LineSpacing = 15.66f;
                            paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Left;

                            cell = table2.Rows[i + 1].Cells[3];
                            paraCellTable2 = cell.AddParagraph();
                            txtParaCellTable2 = paraCellTable2.AppendText(item.CompleteAt != null ? item.CompleteAt.Value.ToString("dd/MM/yyyy") : "");
                            cell.CellFormat.BackColor = ColorTranslator.FromHtml("#ffffff");
                            txtParaCellTable2.CharacterFormat.FontName = font;
                            txtParaCellTable2.CharacterFormat.FontSize = 10;
                            paraCellTable2.Format.LineSpacing = 15.66f;
                            paraCellTable2.Format.HorizontalAlignment = HorizontalAlignment.Left;
                        }
                        var breakSection = sectionDetectDetails.AddParagraph();
                        var txtSectionDetectDetails = breakSection.AppendText("");
                    }
                    table1.Rows.RemoveAt(1);


                    MemoryStream stream = new MemoryStream();
                    stream.Position = 0;
                    doc.SaveToFile(stream, Spire.Doc.FileFormat.Docx);
                    Bytes = stream.ToArray();
                    return File(Bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Kitano_BaoCaoKiemToan.docx");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"ExportAudit - {DateTime.Now} : {ex.Message}!");
                return BadRequest();
            }


        }

        private string getRatingLevelName(int? rating)
        {
            switch (rating)
            {
                case 1:
                    return "Kiểm soát tốt";
                case 2:
                    return "Chấp nhận được";
                case 3:
                    return "Cần cải thiện";
                case 4:
                    return "Không đạt yêu cầu";
                default:
                    return "";
            }
        }
    }
}