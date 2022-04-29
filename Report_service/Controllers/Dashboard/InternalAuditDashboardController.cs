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
    public class InternalAuditDashboardController : BaseController
    {
        public InternalAuditDashboardController(ILoggerManager logger, IUnitOfWork uow) : base(logger, uow)
        {
        }

        [HttpGet("getdata")]
        public IActionResult GetReport()
        {
            var current_year = 0;
            var now = DateTime.Now.Date;
            if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
            {
                return Unauthorized();
            }

            current_year = DateTime.Now.Year;
            var audit_detect = (from a in _uow.Repository<AuditDetect>().Find(p => !(p.IsDeleted ?? false) && p.year == current_year)
                                join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD" && p.StatusCode == "3.1") on a.id equals b.item_id
                                select a);
            #region tình hình thực hiện kế hoạch kiểm toán năm (hiện tại)

            var auditworkcomplete = (from a in _uow.Repository<ReportAuditWork>().Include(p => p.AuditWork)
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_RAW" && p.StatusCode == "3.1") on a.Id equals b.item_id
                                     where a.Year == current_year.ToString() && !(a.IsDeleted ?? false)
                                     select a).Count();

            var prepareauditplan = _uow.Repository<AuditWorkPlan>().Include(a => a.AuditPlan).Where(a => a.Year == current_year.ToString() && a.IsDeleted != true).AsEnumerable().GroupBy(p => p.AuditPlan.Year).SelectMany(x => x.Where(v => v.AuditPlan.Version == x.Max(p => p.AuditPlan.Version))).Count();

            var auditworkexpected = (from a in _uow.Repository<AuditWork>().GetAll()
                                     join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_PAP") on a.Id equals b.item_id
                                     where !(a.IsDeleted ?? false) && b.StatusCode == "3.1" && a.Classify == 2 && (a.Year == current_year.ToString())
                                     select a).Count();

            var _report1 = new InternalAudit1Model
            {
                Total_Audit_Completed = auditworkcomplete,
                Total_Audit_Pending = prepareauditplan + auditworkexpected - auditworkcomplete,
                Total_Audit_Plan = prepareauditplan,
                Total_Audit_Expect = auditworkexpected,
            };
            #endregion

            #region Tình hình thực hiện kiến nghị của kiểm toán
            var auditrequestmonitor = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect)
                                       join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.AuditDetect.id equals b.item_id
                                       where b.StatusCode == "3.1" && a.AuditDetect != null && !(a.AuditDetect.IsDeleted ?? false) && (a.AuditDetect.rating_risk == 1 || a.AuditDetect.rating_risk == 2) && a.AuditDetect.year == current_year && !(a.is_deleted ?? false)
                                       let time = (a.extend_at.HasValue ? a.extend_at.Value : a.CompleteAt.HasValue ? a.CompleteAt.Value : DateTime.MinValue)
                                       select new ReportAuditRequest
                                       {
                                           extendat = a.extend_at,
                                           conclusion = a.Conclusion,
                                           year = a.AuditDetect.year,
                                           timestatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue ? a.CompleteAt < now : false)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue ? a.ActualCompleteAt > a.CompleteAt : false))) ? 2 : 1,
                                           day = ((a.ProcessStatus ?? 1) == 3 && a.ActualCompleteAt.HasValue) ? (a.ActualCompleteAt - time).Value.TotalDays : (now - time).TotalDays
                                       }
                                               ).ToList();
            var type1 = GetData(auditrequestmonitor.Where(p => p.timestatus == 1), 1);
            var type2 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && p.day < 30), 2);
            var type3 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 30 <= p.day && p.day < 60), 3);
            var type4 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 60 <= p.day && p.day <= 90), 4);
            var type5 = GetData(auditrequestmonitor.Where(p => p.timestatus == 2 && 90 < p.day), 5);
            var type6 = GetData(auditrequestmonitor.Where(p => p.extendat.HasValue), 6);
            List<InternalAudit2Model> _report2 = new()
            {

                type6,
                type5,
                type4,
                type3,
                type2,
                type1
            };

            #endregion

            #region Số lượng phát hiện kiểm toán trong năm theo mức rủi ro
            var _report3 = audit_detect.GroupBy(p => p.rating_risk).Select(p => new InternalAudit3Model
            {
                Ratingrisk = p.Key.Value,
                Total = p.Count(),
            });
            #endregion
            return Ok(new
            {
                code = "1",
                report1 = _report1,
                report2 = _report2,
                report3 = _report3
            });

        }

        private static InternalAudit2Model GetData(IEnumerable<ReportAuditRequest> list, int type)
        {
            var item = new InternalAudit2Model()
            {
                Type = type,
                Close = list.Count(p => p.conclusion == 2),
                Open = list.Count(p => p.conclusion != 2)
            };
            return item;
        }


    }
}
