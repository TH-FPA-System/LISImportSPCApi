using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;
using LISImportSPCApi.Models;
using System.Globalization;

namespace LISImportSPCApi.Controllers
{
    [ApiController]
    [Route("api/import")]
    public class ImportController : ControllerBase
    {
        private readonly string _connStr;

        public ImportController(IConfiguration config)
        {
            _connStr = config.GetConnectionString("DefaultConnection")
                       ?? throw new Exception("Connection string missing");
        }

        [HttpPost("excel")]
        public async Task<IActionResult> ImportExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File is empty");

            var rows = new List<RawExcelRowDto>();
            var errors = new List<string>();

            // 1️⃣ Read Excel file
            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                using var workbook = new XLWorkbook(stream);
                var sheet = workbook.Worksheets.FirstOrDefault();
                if (sheet == null)
                    return BadRequest("No worksheet found in Excel file");

                int rowCount = sheet.LastRowUsed()?.RowNumber() ?? 0;
                if (rowCount < 2)
                    return BadRequest("No data rows found in Excel file");

                for (int row = 2; row <= rowCount; row++)
                {
                    var xlRow = sheet.Row(row);

                    // 1️⃣ Task
                    int task = 0;
                    if (!int.TryParse(xlRow.Cell(1).GetString(), out task))
                        errors.Add($"Row {row}: Invalid Task value '{xlRow.Cell(1).GetString()}'");

                    // 2️⃣ TaskName
                    string taskName = xlRow.Cell(2).GetString() ?? "";

                    // 3️⃣ TestPart
                    string testPart = xlRow.Cell(3).GetString() ?? "";
                    if (string.IsNullOrWhiteSpace(testPart))
                        errors.Add($"Row {row}: TestPart is empty");

                    // 4️⃣ TestPartDesc
                    string testPartDesc = xlRow.Cell(4).GetString() ?? "";

                    // 5️⃣ Value
                    double value = 0;
                    if (!double.TryParse(xlRow.Cell(5).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                        errors.Add($"Row {row}: Invalid Value '{xlRow.Cell(5).GetString()}'");

                    // 6️⃣ Unit
                    string unit = xlRow.Cell(6).GetString() ?? "";

                    // 7️⃣ TestDateTime
                    DateTime testDateTime;
                    var dateCell = xlRow.Cell(7).GetString();
                    if (!DateTime.TryParseExact(dateCell, "dd/MM/yyyy H:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out testDateTime))
                    {
                        if (!DateTime.TryParse(dateCell, out testDateTime))
                            errors.Add($"Row {row}: Invalid date format '{dateCell}'");
                    }

                    // 8️⃣ Part
                    string part = xlRow.Cell(8).GetString() ?? "";

                    // 9️⃣ Serial
                    string serial = xlRow.Cell(9).GetString() ?? "";

                    // Only add row if no critical errors
                    if (!errors.Any(e => e.StartsWith($"Row {row}:")))
                    {
                        rows.Add(new RawExcelRowDto
                        {
                            Task = task,
                            TaskName = taskName,
                            TestPart = testPart,
                            TestPartDesc = testPartDesc,
                            Value = value,
                            Unit = unit,
                            TestDateTime = testDateTime,
                            Part = part,
                            Serial = serial
                        });
                    }
                }

                if (errors.Count > 0)
                    return BadRequest(string.Join("; ", errors));
            }
            catch (Exception ex)
            {
                return BadRequest($"Error reading Excel file: {ex.Message}");
            }

            // 2️⃣ Insert into DB with duplicate check
            var previewList = new List<dynamic>();
            int insertedCount = 0;
            int duplicateCount = 0;

            try
            {
                using var conn = new SqlConnection(_connStr);
                await conn.OpenAsync();
                using var tran = conn.BeginTransaction();

                foreach (var r in rows)
                {
                    // Duplicate check
                    var checkCmd = new SqlCommand(@"
                        SELECT TOP 1 test_result_id
                        FROM test_result_clean
                        WHERE task=@task AND test_part=@test_part AND date_tested=@date_tested
                    ", conn, tran);

                    checkCmd.Parameters.Add("@task", SqlDbType.Int).Value = r.Task;
                    checkCmd.Parameters.Add("@test_part", SqlDbType.VarChar, 50).Value = r.TestPart;
                    checkCmd.Parameters.Add("@date_tested", SqlDbType.DateTime2).Value = r.TestDateTime;

                    var existingId = await checkCmd.ExecuteScalarAsync();
                    bool isDuplicate = existingId != null;

                    string statusText = isDuplicate ? "DUPLICATE" : (r.Value < 0 ? "FAIL" : "PASS");
                    string statusChar = r.Value < 0 ? "F" : "P";
                    string duplicateReason = isDuplicate ? "Duplicate: task + test_part + date_tested" : null;

                    previewList.Add(new
                    {
                        Task = r.Task,
                        TestPart = r.TestPart,
                        TestDate = r.TestDateTime,
                        Value = r.Value,
                        Status = statusText
                    });

                    // Insert row
                    var insertCmd = new SqlCommand(@"
                        INSERT INTO test_result_clean
                        (part, serial, task, run_number, test_part, date_tested, 
                         test_value, test_unit, result_status, result_text, 
                         test_info1, test_info2, is_duplicate, duplicate_reason)
                        VALUES
                        (@part, @serial, @task, 1, @test_part, @date_tested,
                         @test_value, @test_unit, @status_char, @status_text,
                         @info1, @info2, @is_duplicate, @duplicate_reason)
                    ", conn, tran);

                    insertCmd.Parameters.Add("@part", SqlDbType.VarChar, 50).Value = r.Part;
                    insertCmd.Parameters.Add("@serial", SqlDbType.VarChar, 50).Value = r.Serial;
                    insertCmd.Parameters.Add("@task", SqlDbType.Int).Value = r.Task;
                    insertCmd.Parameters.Add("@test_part", SqlDbType.VarChar, 50).Value = r.TestPart;
                    insertCmd.Parameters.Add("@date_tested", SqlDbType.DateTime2).Value = r.TestDateTime;
                    insertCmd.Parameters.Add("@test_value", SqlDbType.Float).Value = r.Value;
                    insertCmd.Parameters.Add("@test_unit", SqlDbType.VarChar, 16).Value = r.Unit;
                    insertCmd.Parameters.Add("@status_char", SqlDbType.Char, 1).Value = statusChar;
                    insertCmd.Parameters.Add("@status_text", SqlDbType.VarChar, 16).Value = statusText;
                    insertCmd.Parameters.Add("@info1", SqlDbType.VarChar, 50).Value = r.TaskName;
                    insertCmd.Parameters.Add("@info2", SqlDbType.VarChar, 50).Value = r.TestPartDesc;
                    insertCmd.Parameters.Add("@is_duplicate", SqlDbType.Bit).Value = isDuplicate;
                    insertCmd.Parameters.Add("@duplicate_reason", SqlDbType.VarChar, 100).Value = (object?)duplicateReason ?? DBNull.Value;

                    await insertCmd.ExecuteNonQueryAsync();

                    if (isDuplicate) duplicateCount++;
                    else insertedCount++;
                }

                tran.Commit();
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Database error: {ex.Message}");
            }

            return Ok(new
            {
                ImportedRows = insertedCount,
                DuplicateRows = duplicateCount,
                Preview = previewList
            });
        }
    }
}
