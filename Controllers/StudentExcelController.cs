using System.Data;
using ExcelDataReader;
using ExcelToDatabase.Data;
using ExcelToDatabase.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace ExcelToDatabase.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class StudentExcelController : ControllerBase
    {
        private readonly ExcelToDatabaseDbcontext _context;
        public StudentExcelController(ExcelToDatabaseDbcontext context)
        {
            _context = context;
        }
        [HttpPost("UploadExcelFile")]
        public async Task<IActionResult> UploadExcelFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No File Uploaded");

            var extension = Path.GetExtension(file.FileName).ToLower();
            if (extension != ".xlsx" && extension != ".xls")
                return BadRequest("Invalid file format. Only .xlsx and .xls are allowed.");

            var students = new List<StudentExcel>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = file.OpenReadStream())
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
                {
                    LeaveOpen = false
                }))
                {
                    bool isFirst = true;
                    while (reader.Read())
                    {
                        if (isFirst)
                        {
                            isFirst = false;
                            continue;
                        }
                        string firstName = reader.GetValue(1)?.ToString();
                        string lastName = reader.GetValue(2)?.ToString();
                        object dobValue = reader.GetValue(3);
                        object marksValue = reader.GetValue(4);

                        if (string.IsNullOrEmpty(firstName) || string.IsNullOrEmpty(lastName))
                        {
                            continue;
                        }

                        dobValue = reader.GetValue(3); // Change index if needed

                        // Log the actual data type and value
                        Console.WriteLine($"DOB Column Type: {dobValue?.GetType()} | Value: {dobValue}");

                        if (dobValue == null || string.IsNullOrWhiteSpace(dobValue.ToString()))
                        {
                            Console.WriteLine($"Skipping row - DOB is empty.");
                            continue;
                        }

                        DateOnly dob;

                        if (dobValue is double excelDate)  // Case: Excel serial date
                        {
                            dob = DateOnly.FromDateTime(DateTime.FromOADate(excelDate));
                        }
                        else if (dobValue is string dobString)  // Case: String format
                        {
                            dobString = dobString.Trim();

                            if (DateTime.TryParseExact(dobString, new[] { "dd/MM/yyyy", "yyyy-MM-dd", "M/d/yyyy", "yyyy/MM/dd" },
                                                       CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                            {
                                dob = DateOnly.FromDateTime(parsedDate);
                            }
                            else
                            {
                                Console.WriteLine($"Skipping row - Invalid DOB format: {dobString}");
                                continue;
                            }
                        }
                        else if (dobValue is DateTime dateTimeValue)  //  Case: System.DateTime
                        {
                            dob = DateOnly.FromDateTime(dateTimeValue);
                        }
                        else
                        {
                            Console.WriteLine($"Skipping row - DOB is neither double, string, nor DateTime.");
                            continue;
                        }

                        //  Successfully parsed DOB
                        Console.WriteLine($"Valid DOB: {dob}");
                        int marks;

                        if (marksValue is double numericMarks)
                        {
                            marks = Convert.ToInt32(numericMarks);
                        }
                        else if (marksValue is string marksString && int.TryParse(marksString, out marks))
                        {
                            // Parsed successfully
                        }
                        else
                        {
                            Console.WriteLine($"Skipping row - Invalid marks: {marksValue}");
                            continue; // Skip if marks are invalid
                        }

                        var student = new StudentExcel
                        {
                            FirstName = firstName,
                            LastName = lastName,
                            DOB = dob,
                            marks = marks
                        };
                        students.Add(student);

                        if (students.Count >= 500)
                        {
                            _context.StudentExcels.AddRange(students);
                            await _context.SaveChangesAsync();
                            students.Clear();
                        }
                    }
                    if (students.Count > 0)
                    {
                        _context.StudentExcels.AddRange(students);
                        await _context.SaveChangesAsync();
                    }
                }
            }
            return Ok(new { Message = "Student Excel file uploaded successfully!" });
        }
    }
}


        //    [HttpPost("UploadExcelFile")]
        //    public IActionResult UploadExcelFile([FromForm] IFormFile file)
        //    {
        //        try
        //        {
        //            if (file == null || file.Length == 0)
        //            {
        //                return BadRequest("No File Uploaded");
        //            }
        //            var UploadsFolder = $"{Directory.GetCurrentDirectory()}\\Uploads";
        //            if (!Directory.Exists(UploadsFolder))
        //            {
        //                Directory.CreateDirectory(UploadsFolder);
        //            }
        //            var filePath = Path.Combine(UploadsFolder, file.FileName);
        //            using (var stream = new FileStream(filePath, FileMode.Create))
        //            {
        //                file.CopyTo(stream);
        //            }
        //            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
        //            {
        //                using (var reader = ExcelReaderFactory.CreateReader(stream))
        //                {
        //                    bool isHadderSkipped = false;
        //                    do
        //                    {
        //                        while (reader.Read())
        //                        {
        //                            if (!isHadderSkipped)
        //                            {
        //                                isHadderSkipped = true;
        //                                continue;
        //                            }
        //                            StudentExcel s = new StudentExcel();
        //                            s.Id = reader.IsDBNull(1) ? 0 : Convert.ToInt32(reader.GetValue(1).ToString());
        //                            s.FirstName = reader.IsDBNull(2) ? "" : reader.GetValue(2).ToString();
        //                            s.LastName = reader.IsDBNull(3) ? "" : reader.GetValue(3).ToString();
        //                            s.DOB = reader.IsDBNull(4) || reader.GetValue(4) == null
        //                                ? DateOnly.MinValue
        //                                : DateOnly.FromDateTime(Convert.ToDateTime(reader.GetValue(4)));
        //                            s.marks = reader.IsDBNull(5) ? 0 : Convert.ToInt32(reader.GetValue(5).ToString());


        //                            _context.Add(s);
        //                            _context.SaveChanges();
        //                        }
        //                    } while (reader.NextResult());


        //                }
        //            }
        //            return Ok("Succefully Inserted");
        //        }
        //        catch (Exception ex)
        //        {
        //            var message = ex.Message;
        //            return BadRequest(message);
        //        }
        //    }


