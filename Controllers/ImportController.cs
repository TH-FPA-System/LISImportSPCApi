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

                    // Parse Excel columns
                    int task = int.TryParse(xlRow.Cell(1).GetString(), out var t) ? t : 0;
                    string part = xlRow.Cell(2).GetString() ?? "";
                    string serial = xlRow.Cell(3).GetString() ?? "";
                    string taskName = xlRow.Cell(4).GetString() ?? "";
                    string testPart = xlRow.Cell(5).GetString() ?? "";
                    string testPartDesc = xlRow.Cell(6).GetString() ?? "";
                    double value = double.TryParse(xlRow.Cell(7).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0;
                    string unit = xlRow.Cell(8).GetString() ?? "";
                    string dateCell = xlRow.Cell(9).GetString();                 
                    string storeLocation = xlRow.Cell(10).GetString() ?? "";
                    string created_by = xlRow.Cell(11).GetString() ?? "";

                    // Validate
                    if (task == 0) errors.Add($"Row {row}: Invalid Task");
                    if (string.IsNullOrWhiteSpace(testPart)) errors.Add($"Row {row}: TestPart empty");

                    DateTime testDateTime;
                    if (!DateTime.TryParseExact(dateCell, "dd/MM/yyyy H:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out testDateTime))
                    {
                        if (!DateTime.TryParse(dateCell, out testDateTime))
                            errors.Add($"Row {row}: Invalid date format '{dateCell}'");
                    }

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
                            Serial = serial,
                            StoreLocation = storeLocation,
                            Created_by = created_by
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

            // 2️⃣ Insert into DB
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
                    // Duplicate check: task + test_part + date_tested
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

                    string statusChar = r.Value < 0 ? "F" : "P";
                    //FORCE NOCHECK DUPLICATE
                    isDuplicate = false;
                    string statusText = isDuplicate ? "DUPLICATE" : (r.Value < 0 ? "FAIL" : "PASS");

                   
                    previewList.Add(new
                    {
                        Task = r.Task,
                        TestPart = r.TestPart,
                        TestDate = r.TestDateTime,
                        Value = r.Value,
                        Status = statusText
                    });

                    if (!isDuplicate)
                    {
                        var insertCmd = new SqlCommand(@"
                    INSERT INTO test_result_clean
                    (part, serial, task, run_number, test_part, test_value, test_unit, result_status, result_text, test_info1, test_info2, store_location, date_tested,created_by)
                    VALUES
                    (@part, @serial, @task, DEFAULT, @test_part, @test_value, @test_unit, @status_char, @status_text, @info1, @info2, @store_location, @date_tested, @created_by)
                ", conn, tran);

                        insertCmd.Parameters.Add("@part", SqlDbType.VarChar, 50).Value = r.Part;
                        insertCmd.Parameters.Add("@serial", SqlDbType.VarChar, 50).Value = r.Serial;
                        insertCmd.Parameters.Add("@task", SqlDbType.Int).Value = r.Task;
                        insertCmd.Parameters.Add("@test_part", SqlDbType.VarChar, 50).Value = r.TestPart;
                        insertCmd.Parameters.Add("@test_value", SqlDbType.Float).Value = r.Value;
                        insertCmd.Parameters.Add("@test_unit", SqlDbType.VarChar, 16).Value = r.Unit;
                        insertCmd.Parameters.Add("@status_char", SqlDbType.Char, 1).Value = statusChar;
                        insertCmd.Parameters.Add("@status_text", SqlDbType.VarChar, 16).Value = statusText;
                        insertCmd.Parameters.Add("@info1", SqlDbType.VarChar, 50).Value = r.TaskName;
                        insertCmd.Parameters.Add("@info2", SqlDbType.VarChar, 200).Value = r.TestPartDesc;
                        insertCmd.Parameters.Add("@store_location", SqlDbType.VarChar, 50).Value = string.IsNullOrEmpty(r.StoreLocation) ? DBNull.Value : r.StoreLocation;
                        insertCmd.Parameters.Add("@date_tested", SqlDbType.DateTime2).Value = r.TestDateTime;
                        insertCmd.Parameters.Add("@created_by", SqlDbType.VarChar, 50).Value = r.Created_by;

                        await insertCmd.ExecuteNonQueryAsync();
                        insertedCount++;
                    }
                    else
                    {
                        duplicateCount++;
                    }
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
