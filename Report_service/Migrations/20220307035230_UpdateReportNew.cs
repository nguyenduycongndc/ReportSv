using Microsoft.EntityFrameworkCore.Migrations;

namespace Report_service.Migrations
{
    public partial class UpdateReportNew : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "other_content",
                table: "REPORT_AUDIT_WORK",
                type: "text",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "other_content",
                table: "REPORT_AUDIT_WORK");
        }
    }
}
