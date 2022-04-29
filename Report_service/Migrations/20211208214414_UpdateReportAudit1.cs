using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace Report_service.Migrations
{
    public partial class UpdateReportAudit1 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<DateTime>(
                name: "approval_date",
                table: "REPORT_AUDIT_WORK",
                type: "timestamp without time zone",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "approval_date",
                table: "REPORT_AUDIT_WORK");
        }
    }
}
