using Microsoft.EntityFrameworkCore.Migrations;

namespace Report_service.Migrations
{
    public partial class UpdateReportAudit : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "approval_user",
                table: "REPORT_AUDIT_WORK",
                type: "integer",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "file_type",
                table: "REPORT_AUDIT_WORK",
                type: "text",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "path",
                table: "REPORT_AUDIT_WORK",
                type: "text",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "reason_reject",
                table: "REPORT_AUDIT_WORK",
                type: "text",
                nullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_REPORT_AUDIT_WORK_approval_user",
                table: "REPORT_AUDIT_WORK",
                column: "approval_user");

            migrationBuilder.AddForeignKey(
                name: "FK_REPORT_AUDIT_WORK_USERS_approval_user",
                table: "REPORT_AUDIT_WORK",
                column: "approval_user",
                principalTable: "USERS",
                principalColumn: "id",
                onDelete: ReferentialAction.Restrict);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_REPORT_AUDIT_WORK_USERS_approval_user",
                table: "REPORT_AUDIT_WORK");

            migrationBuilder.DropIndex(
                name: "IX_REPORT_AUDIT_WORK_approval_user",
                table: "REPORT_AUDIT_WORK");

            migrationBuilder.DropColumn(
                name: "approval_user",
                table: "REPORT_AUDIT_WORK");

            migrationBuilder.DropColumn(
                name: "file_type",
                table: "REPORT_AUDIT_WORK");

            migrationBuilder.DropColumn(
                name: "path",
                table: "REPORT_AUDIT_WORK");

            migrationBuilder.DropColumn(
                name: "reason_reject",
                table: "REPORT_AUDIT_WORK");
        }
    }
}
