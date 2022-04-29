using Microsoft.AspNetCore.Mvc;
using Report_service.DataAccess;
using Report_service.Models.ExecuteModels;
using Report_service.Models.ExecuteModels.Dashboard;
using Report_service.Models.MigrationsModels;
using Report_service.Models.MigrationsModels.Audit;
using Report_service.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace Report_service.Controllers.Dashboard
{
    [Route("[controller]")]
    [ApiController]
    public class AuditWorkDashboard : BaseController
    {
        public AuditWorkDashboard(ILoggerManager logger, IUnitOfWork uow) : base(logger, uow)
        {
        }
        [HttpGet("getdata")]
        public IActionResult GetReport(string jsonData)
        {
            try
            {
                if (HttpContext.Items["UserInfo"] is not CurrentUserModel userInfo)
                {
                    return Unauthorized();
                }
                var obj = JsonSerializer.Deserialize<SearchModel>(jsonData);
                var _audit = _uow.Repository<AuditWork>().Include(x => x.Users).Where(p => p.Id == obj.AuditId).Select(a => new AuditWorkItem
                {
                    Person = a.person_in_charge.HasValue ? a.Users.FullName : "",
                    Status = "",
                }).FirstOrDefault();
                if (_audit != null)
                    _audit.NumOfAuditor = _uow.Repository<AuditAssignment>().Find(x => x.auditwork_id == obj.AuditId).Count();
                var auditMainStage = _uow.Repository<MainStage>().GetAll().Where(a => a.auditwork_id == obj.AuditId);
                
                if (auditMainStage.Any() && auditMainStage.Where(x => x.actual_date.HasValue).OrderByDescending(x => x.index)?.FirstOrDefault() != null)
                {
                    _audit.Status = auditMainStage.Where(x => x.actual_date.HasValue).OrderByDescending(x => x.index)?.FirstOrDefault().status;
                }
                else
                {
                    _audit.Status = "Chưa thực hiện";
                }
                if (auditMainStage.Any() && auditMainStage.OrderByDescending(x => x.index).FirstOrDefault().actual_date.HasValue)
                {
                    _audit.ReleaseDate = auditMainStage.OrderByDescending(x => x.index).FirstOrDefault().actual_date.Value.ToString("dd/MM/yyyy");
                }
                else
                {
                    _audit.ReleaseDate = "";
                }
                #region Tiến trình cuộc kiểm toán

                var _auditschedule = _uow.Repository<Schedule>().Include(a => a.Users).Where(a => a.auditwork_id == obj.AuditId && a.is_deleted != true && a.actual_date.HasValue).OrderBy(x => x.work).Select(p => new AuditSchedule
                {
                    Work = p.work,
                    Actual_Date_Schedule = p.actual_date.HasValue ? p.actual_date.Value.ToString("dd/MM/yyyy") : null,
                    Expected_Date_Schedule = p.expected_date.HasValue ? p.expected_date.Value.ToString("dd/MM/yyyy") : null,
                    DeviatingPlan = p.actual_date.HasValue && p.expected_date.HasValue ? (p.actual_date - p.expected_date).Value.TotalDays.ToString() : "0",
                    Actual_Date = p.actual_date,
                    Expected_Date = p.expected_date,
                });
                var _min = _auditschedule.Min(p => p.Actual_Date);
                var _max = _auditschedule.Max(p => p.Actual_Date);
                var _total = _max.HasValue && _min.HasValue ? (_max - _min).Value.TotalDays : 0;
                #endregion
                #region Risk Heat Map cho kế hoạch cuộc kiểm toán
                var _riskheatmap = _uow.Repository<AuditStrategyRisk>().Include(a => a.AuditWorkScopeFacility).Where(a => a.AuditWorkScopeFacility.auditwork_id == obj.AuditId && a.is_deleted != true).Select(p => new RiskHeatMap
                {
                    Name = p.name_risk,
                    Rating = p.risk_level == 1 ? "Cao" : (p.risk_level == 2 ? "Trung bình" : (p.risk_level == 3 ? "Thấp" : ""))
                }).ToList();
                #endregion

                #region Thống kê số lượng phát hiện kiểm toán
                var _unitdetect = _uow.Repository<AuditDetect>().Include(p => p.AuditFacility).Where(p => (obj.Facility == 0 || p.auditwork_id == obj.AuditId) && !(p.IsDeleted ?? false)).AsEnumerable().GroupBy(p => p.AuditFacility).Select(p => new UnitDetect
                {
                    FacilityName = p.Key.Name,
                    Risk_High = p.Count(x => x.rating_risk == 1),
                    Risk_Medium = p.Count(x => x.rating_risk == 2),
                    Risk_Low = p.Count(x => x.rating_risk == 3)
                }).ToList();

                var _risksummary = new RiskSummary
                {
                    Risk_High = _unitdetect.Sum(p => p.Risk_High),
                    Risk_Medium = _unitdetect.Sum(p => p.Risk_Medium),
                    Risk_Low = _unitdetect.Sum(p => p.Risk_Low),
                    Total_Risk = _unitdetect.Sum(p => p.Risk_High + p.Risk_Low + p.Risk_Medium),
                };
                #endregion

                #region Công việc của các kiểm toán viên
                var _auditorwork = _uow.Repository<AuditWorkScopeUserMapping>().Include(p => p.Users, p => p.AuditWorkScope).Where(p => p.AuditWorkScope.auditwork_id == obj.AuditId).AsEnumerable().GroupBy(p => new { p.user_id, p.Users.FullName }).Select(p => new AuditorWork
                {
                    UserName = p.Key.FullName,
                    TotalWork = p.Select(x => x.auditwork_scope_id).Distinct().Count(),
                    TotalPaper = _uow.Repository<WorkingPaper>().Find(x => x.auditworkid == obj.AuditId && x.asigneeid == p.Key.user_id).Count(),
                });
                #endregion
                return Ok(new
                {
                    code = "1",
                    audit = _audit,
                    riskheatmap = _riskheatmap,
                    unitdetect = _unitdetect,
                    risksummary = _risksummary,
                    auditorwork = _auditorwork,
                    auditschedule = _auditschedule,
                    min = _min,
                    max = _max,
                    total = _total
                });
            }
            catch
            {
                return BadRequest();
            }

        }
    }
}
