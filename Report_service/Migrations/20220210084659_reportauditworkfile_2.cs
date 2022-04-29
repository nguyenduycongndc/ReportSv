using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

namespace Report_service.Migrations
{
    public partial class reportauditworkfile_2 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "REPORT_AUDIT_WORK_FILE",
                columns: table => new
                {
                    id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    report_auditwork_id = table.Column<int>(type: "integer", nullable: false),
                    path = table.Column<string>(type: "text", nullable: true),
                    file_type = table.Column<string>(type: "text", nullable: true),
                    isdelete = table.Column<bool>(type: "boolean", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_REPORT_AUDIT_WORK_FILE", x => x.id);
                    table.ForeignKey(
                        name: "FK_REPORT_AUDIT_WORK_FILE_REPORT_AUDIT_WORK_report_auditwork_id",
                        column: x => x.report_auditwork_id,
                        principalTable: "REPORT_AUDIT_WORK",
                        principalColumn: "id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_REPORT_AUDIT_WORK_FILE_report_auditwork_id",
                table: "REPORT_AUDIT_WORK_FILE",
                column: "report_auditwork_id");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "REPORT_AUDIT_WORK_FILE");
        }
    }
}
