using OfficeOpenXml;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using EmplyoeeManagement.Models;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace EmplyoeeManagement.Pages.Registration
{
    public class RegistrationModel : PageModel
    {
        [BindProperty]
        public Employee Employee { get; set; }

        public List<SelectListItem> GradeOptions { get; set; }
        public List<SelectListItem> BUOptions { get; set; }

        public void OnGet(int? id = null)
        {
            LoadDropdownOptions();

            if (id.HasValue)
            {
                var employee = GetEmployeeById(id.Value);
                if (employee != null)
                {
                    Employee = employee;
                }
            }
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                OnGet(); // Reload dropdown options
                return Page();
            }

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");

            try
            {
                using (var package = new ExcelPackage())
                {
                    if (System.IO.File.Exists(filePath))
                    {
                        // Load the existing file
                        using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                        {
                            package.Load(fileStream);
                        }
                    }
                    else
                    {
                        // Create a new worksheet if the file doesn't exist
                        package.Workbook.Worksheets.Add("Employees");
                    }

                    var worksheet = package.Workbook.Worksheets["Employees"] ?? package.Workbook.Worksheets.Add("Employees");

                    // Create header row if the worksheet is empty
                    if (worksheet.Dimension == null)
                    {
                        worksheet.Cells[1, 1].Value = "ID";
                        worksheet.Cells[1, 2].Value = "First Name";
                        worksheet.Cells[1, 3].Value = "Last Name";
                        worksheet.Cells[1, 4].Value = "Email";
                        worksheet.Cells[1, 5].Value = "Phone";
                        worksheet.Cells[1, 6].Value = "Position";
                        worksheet.Cells[1, 7].Value = "Department";
                        worksheet.Cells[1, 8].Value = "Date of Hire";
                    }

                    // Find the next empty row
                    var row = worksheet.Dimension?.Rows + 1 ?? 2;

                    // Write employee data
                    worksheet.Cells[row, 1].Value = Employee.EmployeeId;
                    worksheet.Cells[row, 2].Value = Employee.FirstName;
                    worksheet.Cells[row, 3].Value = Employee.LastName;
                    worksheet.Cells[row, 4].Value = Employee.Email;
                    worksheet.Cells[row, 5].Value = Employee.Phone;
                    worksheet.Cells[row, 6].Value = Employee.Grade;
                    worksheet.Cells[row, 7].Value = Employee.BU;
                    worksheet.Cells[row, 8].Value = Employee.DateOfHire.ToShortDateString();

                    // Save changes to the file
                    await package.SaveAsAsync(new FileInfo(filePath));
                }

                return RedirectToPage("EmployeeList");
            }
            catch (Exception ex)
            {
                // Log the exception if needed (e.g., ILogger)
                ModelState.AddModelError(string.Empty, "An error occurred while saving the data.");
                OnGet(); // Reload dropdown options
                return Page();
            }
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

        private Employee GetEmployeeById(int id)
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");

            if (!System.IO.File.Exists(filePath))
                return null;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["Employees"];
                if (worksheet == null || worksheet.Dimension == null)
                    return null;

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Value != null &&
                        int.TryParse(worksheet.Cells[row, 1].Value.ToString(), out int employeeId) &&
                        employeeId == id)
                    {
                        return new Employee
                        {
                            EmployeeId = employeeId,
                            FirstName = worksheet.Cells[row, 2].Value?.ToString(),
                            LastName = worksheet.Cells[row, 3].Value?.ToString(),
                            Email = worksheet.Cells[row, 4].Value?.ToString(),
                            Phone = worksheet.Cells[row, 5].Value?.ToString(),
                            Grade = worksheet.Cells[row, 6].Value?.ToString(),
                            BU = worksheet.Cells[row, 7].Value?.ToString(),
                            DateOfHire = DateTime.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out var date) ? date : DateTime.Now
                        };
                    }
                }
            }

            return null;
        }

        private bool UpdateEmployeeInWorksheet(ExcelWorksheet worksheet)
        {
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                if (worksheet.Cells[row, 1].Value != null &&
                    int.TryParse(worksheet.Cells[row, 1].Value.ToString(), out int employeeId) &&
                    employeeId == Employee.EmployeeId)
                {
                    // Update the employee record
                    worksheet.Cells[row, 2].Value = Employee.FirstName;
                    worksheet.Cells[row, 3].Value = Employee.LastName;
                    worksheet.Cells[row, 4].Value = Employee.Email;
                    worksheet.Cells[row, 5].Value = Employee.Phone;
                    worksheet.Cells[row, 6].Value = Employee.Grade;
                    worksheet.Cells[row, 7].Value = Employee.BU;
                    worksheet.Cells[row, 8].Value = Employee.DateOfHire.ToShortDateString();
                    return true;
                }
            }

            return false;
        }
    }
}
