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
using Report_service.Models.ExecuteModels.Dashboard;

namespace Report_service.Controllers.Dashboard
{
    [Route("[controller]")]
    [ApiController]
    public class AuditedUnitDashboardController : BaseController
    {
        public AuditedUnitDashboardController(ILoggerManager logger, IUnitOfWork uow) : base(logger, uow)
        {
        }

        [HttpGet("auditedunitdashboard")]
        public IActionResult GetReport(string jsonData)
        {
            if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
            {
                return Unauthorized();
            }
            var obj = JsonSerializer.Deserialize<SearchModel>(jsonData);
            var Year = DateTime.Now.Year;
            #region Lịch sử đánh giá rủi ro
            var _report1 = (from a in _uow.Repository<ScoreBoard>().GetAll()
                            join b in _uow.Repository<AssessmentResult>().GetAll() on a.ID equals b.ScoreBoardId
                            where /*a.Status && a.CurrentStatus ==1 &&*/  !a.Deleted && a.ApplyFor == "DV" && a.ObjectId == obj.Facility && a.Year <= Year && a.Stage == 1
                            select new
                            {
                                assessment_risklevel = string.IsNullOrEmpty(b.RiskLevelChangeName) ? a.RiskLevel : b.RiskLevelChangeName,
                                Year = a.Year,
                            })?.OrderByDescending(p => p.Year).Take(3).Select(p => new AudittedUnit1Model
                            {
                                Year = p.Year,
                                Risk = p.assessment_risklevel,
                            });

            #endregion

            #region Thống kê số lượng phát hiện kiểm toán 
            var auditdetect = (from a in _uow.Repository<AuditDetect>().Find(p => (obj.Facility == 0 || p.auditfacilities_id == obj.Facility) && !(p.IsDeleted ?? false))
                               join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD" && p.StatusCode == "3.1") on a.id equals b.item_id
                               select a);
            var _report2 = new AudittedUnit2Model
            {
                Risk_High = auditdetect.Count(x => x.rating_risk == 1),
                Risk_Medium = auditdetect.Count(x => x.rating_risk == 2),
                Risk_Low = auditdetect.Count(x => x.rating_risk == 3)
            };
            #endregion
            return Ok(new { code = "1", report1 = _report1, report2 = _report2 });

        }
        [HttpGet("auditrequest")]
        public IActionResult GetAuditRequest(string jsonData)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var now = DateTime.Now.Date;

                var obj = JsonSerializer.Deserialize<SearchModel>(jsonData);

                var list_auditrequest = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect, p => p.FacilityRequestMonitorMapping, p => p.Users)
                                         join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD" && p.StatusCode == "3.1") on a.AuditDetect.id equals b.item_id
                                         select a).Where(p => p.Conclusion != 2 && !(p.AuditDetect.IsDeleted ?? false) && !(p.is_deleted ?? false) && (obj.Facility == 0 || p.FacilityRequestMonitorMapping.FirstOrDefault(x => x.type == 1).audit_facility_id == obj.Facility)).OrderByDescending(p => p.Code);

                var count = list_auditrequest.Count();
                IEnumerable<AuditRequestMonitor> data = list_auditrequest;

                if (obj.StartNumber >= 0 && obj.PageSize > 0)
                {
                    data = data.Skip(obj.StartNumber).Take(obj.PageSize);
                }
                var lst = data.Select(a => new AudittedUnit3Model
                {
                    Id = a.Id,
                    AuditRequestCode = a.Code,
                    ProcessStatus = a.ProcessStatus ?? 1,
                    TimeStatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false))) ? 2 : 1,
                    CompleteAt = a.CompleteAt.HasValue ? a.CompleteAt.Value.ToString("dd/MM/yyyy") : "",
                    Username = a.Users?.FullName
                });
                return Ok(new { code = "1", msg = "success", data = lst, total = count });
            }
            catch
            {
                return BadRequest();
            }

        }
    }
}
