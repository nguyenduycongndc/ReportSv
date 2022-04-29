using Report_service.Models.MigrationsModels;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System;
using Report_service.Models.MigrationsModels.Audit;
using Audit_service.Models.MigrationsModels;

namespace Report_service.DataAccess
{
    public class KitanoSqlContext : DbContext
    {
        public KitanoSqlContext()
        {
        }
        public KitanoSqlContext(DbContextOptions<KitanoSqlContext> options) : base(options)
        {
        }

        // TABLE
        //===================================================================
        public virtual DbSet<Users> Users { get; set; }
        public virtual DbSet<UsersGroup> UsersGroup { get; set; }
        public virtual DbSet<UsersGroupMapping> UsersGroupMapping { get; set; }
        public virtual DbSet<AuditFacility> Department { get; set; }
        public virtual DbSet<UnitType> UnitType { get; set; }
        public virtual DbSet<UsersWorkHistory> UsersWorkHistory { get; set; }
        public virtual DbSet<Roles> Roles { get; set; }
        public virtual DbSet<UsersRoles> UsersRoles { get; set; }
        public virtual DbSet<UsersGroupRoles> UsersGroupRoles { get; set; }
        public virtual DbSet<SystemParameter> SystemParameter { get; set; }
        public virtual DbSet<AuditPlan> AuditPlan { get; set; }
        public virtual DbSet<AuditWork> AuditWork { get; set; }
        public virtual DbSet<AuditAssignment> AuditAssignment { get; set; }
        public virtual DbSet<CatAuditProcedures> CatAuditProcedures { get; set; }
        public virtual DbSet<CatControl> CatControl { get; set; }
        public virtual DbSet<CatRisk> CatRisk { get; set; }

        public virtual DbSet<AuditWorkScope> AuditWorkScope { get; set; }
        public virtual DbSet<ProcessLevelRiskScoring> ProcessLevelRiskScoring { get; set; }
        public virtual DbSet<RiskScoringProcedures> RiskScoringProcedures { get; set; }
        public virtual DbSet<RiskControl> RiskControl { get; set; }
        public virtual DbSet<AuditRequestMonitor> AuditRequest { get; set; }

        public virtual DbSet<AuditDetect> AuditDetect { get; set; }
        public virtual DbSet<AuditObserve> AuditObserve { get; set; }
        public virtual DbSet<ReportAuditWork> ReportAuditWork { get; set; }
        public virtual DbSet<UnitComment> UnitComment { get; set; }
        public virtual DbSet<Schedule> Schedule { get; set; }
        public virtual DbSet<ReportAuditWorkYear> ReportAuditWorkYear { get; set; }
        public virtual DbSet<ApprovalFunction> ApprovalFunction { get; set; }
        public virtual DbSet<AuditWorkPlan> AuditWorkPlan { get; set; }
        public virtual DbSet<ApprovalConfig> ApprovalConfig { get; set; }
        public virtual DbSet<ScoreBoard> ScoreBoard { get; set; }
        public virtual DbSet<AssessmentResult> AssessmentResult { get; set; }
        public virtual DbSet<CatRiskLevel> CatRiskLevel { get; set; }
        public virtual DbSet<AuditProgram> AuditProgram { get; set; }
        public virtual DbSet<AuditWorkScopeFacility> AuditWorkScopeFacility { get; set; }
        public virtual DbSet<ReportAuditWorkFile> ReportAuditWorkFile { get; set; }
        public virtual DbSet<MainStage> MainStage { get; set; }
        public virtual DbSet<ApprovalFunctionFile> ApprovalFunctionFile { get; set; }

        public virtual DbSet<SystemCategory> SystemCategory { get; set; }
        public virtual DbSet<AuditMinutes> AuditMinutes { get; set; }

        public virtual DbSet<AuditStrategyRisk> AuditStrategyRisk { get; set; }
        public virtual DbSet<ConfigDocument> ConfigDocument { get; set; }

        //===================================================================
        protected override void OnModelCreating(ModelBuilder builder)
        {
            base.OnModelCreating(builder);
        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionBuilder)
        {
            optionBuilder.UseNpgsql(ConnectionService.connstring);
            optionBuilder.UseLoggerFactory(GetLoggerFactory());       // bật logger
        }
        public override int SaveChanges()
        {
            ChangeTracker.DetectChanges();
            return base.SaveChanges();
        }
        private ILoggerFactory GetLoggerFactory()
        {
            IServiceCollection serviceCollection = new ServiceCollection();
            serviceCollection.AddLogging(builder =>
                    builder.AddConsole()
                           .AddFilter(DbLoggerCategory.Database.Command.Name,
                                    LogLevel.Information));
            return serviceCollection.BuildServiceProvider()
                    .GetService<ILoggerFactory>();
        }
    }
}
