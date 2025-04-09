using ClosedXML.Excel;
using Excel_Project.Models;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace Excel_Project.Controllers
{
    public class UserController : Controller
    {
        private readonly string _excelFilePath;

        public UserController()
        {
            _excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "ExcelFiles", "ClientData.xlsx");
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult SubmitForm(UserInfo form)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return Json(new { success = false, error = "Please fill all required fields correctly" });
                }

                Directory.CreateDirectory(Path.GetDirectoryName(_excelFilePath));

                if (!System.IO.File.Exists(_excelFilePath))
                {
                    CreateNewExcelFile(form);
                }
                else
                {
                    AppendToExistingExcelFile(form);
                }

                return Json(new
                {
                    success = true,
                    message = "Data saved successfully!",
                    filePath = $"/ExcelFiles/ClientData.xlsx"
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    success = false,
                    error = "Error saving data: " + ex.Message
                });
            }
        }

        private void CreateNewExcelFile(UserInfo form)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Client Data");

                // Add headers
                worksheet.Cell(1, 1).Value = "Full Name";
                worksheet.Cell(1, 2).Value = "Date";
                worksheet.Cell(1, 3).Value = "Address";
                worksheet.Cell(1, 4).Value = "Governorate";
                worksheet.Cell(1, 5).Value = "Mobile Number";
                worksheet.Cell(1, 6).Value = "Additional Number";
                worksheet.Cell(1, 7).Value = "Price";
                worksheet.Cell(1, 8).Value = "Product Code";
                worksheet.Cell(1, 9).Value = "Product Name";
                worksheet.Cell(1, 10).Value = "Quantity";

                // Add first data row
                worksheet.Cell(2, 1).Value = form.FullName;
                worksheet.Cell(2, 2).Value = form.Date.ToString("yyyy-MM-dd");
                worksheet.Cell(2, 3).Value = form.Address;
                worksheet.Cell(2, 4).Value = form.Governorate;
                worksheet.Cell(2, 5).Value = form.MobileNumber;
                worksheet.Cell(2, 6).Value = form.AdditionalNumber;
                worksheet.Cell(2, 7).Value = form.Price;
                worksheet.Cell(2, 8).Value = form.ProductCode;
                worksheet.Cell(2, 9).Value = form.ProductName;
                worksheet.Cell(2, 10).Value = form.Quantity;

                workbook.SaveAs(_excelFilePath);
            }
        }

        private void AppendToExistingExcelFile(UserInfo form)
        {
            try
            {
                using (var workbook = new XLWorkbook(_excelFilePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    int lastUsedRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    int newRow = lastUsedRow + 1;

                    if (lastUsedRow == 1 && worksheet.Cell(1, 1).IsEmpty())
                    {
                        // Add headers if file is empty
                        worksheet.Cell(1, 1).Value = "Full Name";
                        worksheet.Cell(1, 2).Value = "Date";
                        worksheet.Cell(1, 3).Value = "Address";
                        worksheet.Cell(1, 4).Value = "Governorate";
                        worksheet.Cell(1, 5).Value = "Mobile Number";
                        worksheet.Cell(1, 6).Value = "Additional Number";
                        worksheet.Cell(1, 7).Value = "Price";
                        worksheet.Cell(1, 8).Value = "Product Code";
                        worksheet.Cell(1, 9).Value = "Product Name";
                        worksheet.Cell(1, 10).Value = "Quantity";
                        newRow = 2;
                    }

                    // Add new data
                    worksheet.Cell(newRow, 1).Value = form.FullName;
                    worksheet.Cell(newRow, 2).Value = form.Date.ToString("yyyy-MM-dd");
                    worksheet.Cell(newRow, 3).Value = form.Address;
                    worksheet.Cell(newRow, 4).Value = form.Governorate;
                    worksheet.Cell(newRow, 5).Value = form.MobileNumber;
                    worksheet.Cell(newRow, 6).Value = form.AdditionalNumber;
                    worksheet.Cell(newRow, 7).Value = form.Price;
                    worksheet.Cell(newRow, 8).Value = form.ProductCode;
                    worksheet.Cell(newRow, 9).Value = form.ProductName;
                    worksheet.Cell(newRow, 10).Value = form.Quantity;

                    workbook.Save();
                }
            }
            catch (Exception)
            {
                System.IO.File.Delete(_excelFilePath);
                CreateNewExcelFile(form);
            }
        }

        [HttpGet]
        public IActionResult DownloadFile()
        {
            if (!System.IO.File.Exists(_excelFilePath))
            {
                return NotFound();
            }

            var memory = new MemoryStream();
            using (var stream = new FileStream(_excelFilePath, FileMode.Open))
            {
                stream.CopyTo(memory);
            }
            memory.Position = 0;

            return File(memory,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "ClientData.xlsx");
        }
    }
}