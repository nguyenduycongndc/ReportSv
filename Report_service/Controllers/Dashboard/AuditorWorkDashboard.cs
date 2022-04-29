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
    public class AuditorWorkDashboard : BaseController
    {
        public AuditorWorkDashboard(ILoggerManager logger, IUnitOfWork uow) : base(logger, uow)
        {
        }

        [HttpGet("getauditwork")]
        public IActionResult AuditWorkReport()
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var now = DateTime.Now.Date;
                var auditwork = _uow.Repository<AuditWork>().Include(p => p.AuditAssignment, p => p.MainStage).Where(a => a.AuditAssignment.Any(x => x.user_id == userInfo.Id) && a.Year == now.Year.ToString() && !(a.IsDeleted ?? false)).Select(a => new AuditorWork1Model
                {
                    AuditName = a.Name,
                    AuditStartDate = a.StartDate.HasValue ? a.StartDate.Value.ToString("dd/MM/yyyy") : "",
                    AuditEndDate = a.EndDate.HasValue ? a.EndDate.Value.ToString("dd/MM/yyyy") : "",
                    Status = a.MainStage.Where(p => p.auditwork_id == a.Id && p.actual_date.HasValue).OrderByDescending(x => x.index).FirstOrDefault() == null ? "Chưa thực hiện" : a.MainStage.Where(p => p.auditwork_id == a.Id && p.actual_date.HasValue).OrderByDescending(x => x.index).FirstOrDefault().status,
                });

                var count = auditwork.Count();
                IEnumerable<AuditorWork1Model> data = auditwork;

                ////if (obj.StartNumber >= 0 && obj.PageSize > 0)
                ////{
                ////    data = data.Skip(obj.StartNumber).Take(obj.PageSize);
                ////}
                var lst = data;

                return Ok(new { code = "1", msg = "success", data = lst, total = count });
            }
            catch
            {
                return BadRequest();
            }

        }

        [HttpGet("getworkingpaper")]
        public IActionResult AuditorWorkingPaper()
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var approve = _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_WP");
                var workingpaper = (from a in _uow.Repository<WorkingPaper>().GetAll()
                                    join b in _uow.Repository<AuditWork>().GetAll() on a.auditworkid equals b.Id
                                    where a.asigneeid == userInfo.Id && !(a.isdelete ?? false)
                                    select new AuditorWork2Model
                                    {
                                        AuditName = b.Name,
                                        AuditEndField = b.EndDate.HasValue ? b.EndDate.Value.ToString("dd/MM/yyyy") : "",
                                        PaperCode = a.code,
                                        PaperId = a.id,
                                        Status = approve.FirstOrDefault(x => x.item_id == a.id).StatusCode ?? "1.0"
                                    }).Where(p => (p.Status == "3.2" || p.Status == "1.0" || p.Status == "2.2"));
                var count = workingpaper.Count();
                IEnumerable<AuditorWork2Model> data = workingpaper;

                ////if (obj.StartNumber >= 0 && obj.PageSize > 0)
                ////{
                ////    data = data.Skip(obj.StartNumber).Take(obj.PageSize);
                ////}
                var lst = data;

                return Ok(new { code = "1", msg = "success", data = lst, total = count });
            }
            catch
            {
                return BadRequest();
            }

        }

        [HttpGet("getauditdetect")]
        public IActionResult AuditDetectReport()
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var auditdetect = (from a in _uow.Repository<AuditDetect>().Include(p => p.AuditWork)
                                   join b in _uow.Repository<ApprovalFunction>().Find(p => p.function_code == "M_AD") on a.id equals b.item_id into g
                                   from e in g.DefaultIfEmpty()
                                   where a.CreatedBy == userInfo.Id && !(a.IsDeleted ?? false)
                                   let status = (e.StatusCode ?? "1.0") == "1.0" ? "Bản nháp" : "Từ chối duyệt"
                                   select new AuditorWork3Model
                                   {
                                       AuditName = a.AuditWork.Name,
                                       AuditDetectCode = a.code,
                                       AuditDetectId = a.id,
                                       StatusCode = e.StatusCode ?? "1.0",
                                       Status = status,
                                   }).Where(p => p.StatusCode == "1.0" || p.StatusCode == "2.2" || p.StatusCode == "3.2");
                var count = auditdetect.Count();
                IEnumerable<AuditorWork3Model> data = auditdetect;

                ////if (obj.StartNumber >= 0 && obj.PageSize > 0)
                ////{
                ////    data = data.Skip(obj.StartNumber).Take(obj.PageSize);
                ////}
                var lst = data;

                return Ok(new { code = "1", msg = "success", data = lst, total = count });
            }
            catch
            {
                return BadRequest();
            }

        }

        [HttpGet("getauditrequest")]
        public IActionResult AuditRequestReport()
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var now = DateTime.Now.Date;
                var _unit = _uow.Repository<AuditFacility>().Find(a => a.IsActive == true).AsEnumerable();
                var auditrequest = (from a in _uow.Repository<AuditRequestMonitor>().Include(p => p.AuditDetect, p => p.CatAuditRequest, p => p.FacilityRequestMonitorMapping)
                                    where (a.is_deleted != true) && (a.Conclusion ?? 1) == 1 && a.AuditDetect.followers == userInfo.Id
                                    select new AuditorWork4Model
                                    {
                                        RequestCode = a.Code,
                                        RequestId = a.Id,
                                        ProcessStatus = a.ProcessStatus,
                                        TimeStatus = (((a.ProcessStatus ?? 1) != 3 && (a.extend_at.HasValue ? a.extend_at < now : a.CompleteAt.HasValue && a.CompleteAt < now)) || (a.ActualCompleteAt.HasValue && (a.extend_at.HasValue ? a.ActualCompleteAt > a.extend_at : a.CompleteAt.HasValue && a.ActualCompleteAt > a.CompleteAt))) ? 2 : 1,
                                        UserEdit = a.Users.FullName,
                                        UnitEdit = _unit.FirstOrDefault(p => p.Id == a.FacilityRequestMonitorMapping.FirstOrDefault(x => x.type == 1).audit_facility_id).Name,
                                        CompleteAt = a.CompleteAt.HasValue ? a.CompleteAt.Value.ToString("dd/MM/yyyy") : ""
                                    });
                var count = auditrequest.Count();
                IEnumerable<AuditorWork4Model> data = auditrequest;

                ////if (obj.StartNumber >= 0 && obj.PageSize > 0)
                ////{
                ////    data = data.Skip(obj.StartNumber).Take(obj.PageSize);
                ////}
                var lst = data;

                return Ok(new { code = "1", msg = "success", data = lst, total = count });
            }
            catch
            {
                return BadRequest();
            }

        }
    }
}
