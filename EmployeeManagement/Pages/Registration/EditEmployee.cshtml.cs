using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using EmplyoeeManagement.Models;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using System.IO;

namespace EmplyoeeManagement.Pages.Registration
{
    public class EditModel : PageModel
    {
        [BindProperty]
        public Employee Employee { get; set; }

        public List<SelectListItem> GradeOptions { get; set; }
        public List<SelectListItem> BUOptions { get; set; }

        public void OnGet(int id)
        {
            LoadDropdownOptions();
            // Fetch the employee to be edited using the ID
            Employee = GetEmployeeById(id);
        }

        public IActionResult OnPost(int id)
        {
            if (!ModelState.IsValid)
            {
                LoadDropdownOptions();
                return Page();
            }

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["Employees"];

                var row = FindEmployeeRow(worksheet, id);
                if (row != -1)
                {
                    worksheet.Cells[row, 2].Value = Employee.FirstName;
                    worksheet.Cells[row, 3].Value = Employee.LastName;
                    worksheet.Cells[row, 4].Value = Employee.Email;
                    worksheet.Cells[row, 5].Value = Employee.Phone;
                    worksheet.Cells[row, 6].Value = Employee.Grade;
                    worksheet.Cells[row, 7].Value = Employee.BU;
                    worksheet.Cells[row, 8].Value = Employee.DateOfHire.ToShortDateString();
                }

                package.Save();
            }

            return RedirectToPage("EmployeeList");
        }

        private Employee GetEmployeeById(int id)
        {
            // Fetch employee data from the Excel sheet based on the ID
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["Employees"];
                var row = FindEmployeeRow(worksheet, id);
                if (row != -1)
                {
                    return new Employee
                    {
                        EmployeeId = id,
                        FirstName = worksheet.Cells[row, 2].Text,
                        LastName = worksheet.Cells[row, 3].Text,
                        Email = worksheet.Cells[row, 4].Text,
                        Phone = worksheet.Cells[row, 5].Text,
                        Grade = worksheet.Cells[row, 6].Text,
                        BU = worksheet.Cells[row, 7].Text,
                        DateOfHire = DateTime.Parse(worksheet.Cells[row, 8].Text)
                    };
                }
            }
            return null;
        }

        private int FindEmployeeRow(ExcelWorksheet worksheet, int id)
        {
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                if (worksheet.Cells[row, 1].Text == id.ToString())
                {
                    return row;
                }
            }
            return -1;
        }

        private void LoadDropdownOptions()
        {
            GradeOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "Manager", Text = "Manager" },
                new SelectListItem { Value = "Developer", Text = "Developer" },
                new SelectListItem { Value = "Designer", Text = "Designer" },
                new SelectListItem { Value = "Analyst", Text = "Analyst" },
                new SelectListItem { Value = "Sales", Text = "Sales" }
            };

            BUOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "IT", Text = "IT" },
                new SelectListItem { Value = "HR", Text = "HR" },
                new SelectListItem { Value = "Finance", Text = "Finance" },
                new SelectListItem { Value = "Marketing", Text = "Marketing" },
                new SelectListItem { Value = "Operations", Text = "Operations" }
            };
        }
    }
}
