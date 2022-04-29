using Microsoft.EntityFrameworkCore.Migrations;

namespace Report_service.Migrations
{
    public partial class AddColumnOvercome : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {

            migrationBuilder.AddColumn<string>(
                name: "overcome",
                table: "REPORT_AUDIT_WORK_YEAR",
                type: "text",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "overcome",
                table: "REPORT_AUDIT_WORK_YEAR");
        }
    }
}
