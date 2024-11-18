using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using EmplyoeeManagement.Models;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;


namespace EmplyoeeManagement.Pages.Registration
{
    public class RegistrationModel : PageModel
    {
        [BindProperty]
        public Employee Employee { get; set; }

        public List<SelectListItem> GradeOptions { get; set; }
        public List<SelectListItem> BUOptions { get; set; }

       public void OnGet()
{
    LoadDropdownOptions();

    // Ensure the Employee object is initialized
    Employee ??= new Employee();
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
                        using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                        {
                            package.Load(fileStream);
                        }
                    }
                    else
                    {
                        package.Workbook.Worksheets.Add("Employees");
                    }

                    var worksheet = package.Workbook.Worksheets["Employees"];

                    if (worksheet.Dimension == null)
                    {
                        worksheet.Cells[1, 1].Value = "ID";
                        worksheet.Cells[1, 2].Value = "First Name";
                        worksheet.Cells[1, 3].Value = "Last Name";
                        worksheet.Cells[1, 4].Value = "Email";
                        worksheet.Cells[1, 5].Value = "Phone";
                        worksheet.Cells[1, 6].Value = "Grade";
                        worksheet.Cells[1, 7].Value = "BU";
                        worksheet.Cells[1, 8].Value = "Date of Hire";
                    }

                    var row = worksheet.Dimension?.Rows + 1 ?? 2;

                    worksheet.Cells[row, 1].Value = Employee.EmployeeId;
                    worksheet.Cells[row, 2].Value = Employee.FirstName;
                    worksheet.Cells[row, 3].Value = Employee.LastName;
                    worksheet.Cells[row, 4].Value = Employee.Email;
                    worksheet.Cells[row, 5].Value = Employee.Phone;
                    worksheet.Cells[row, 6].Value = Employee.Grade;
                    worksheet.Cells[row, 7].Value = Employee.BU;
                    worksheet.Cells[row, 8].Value = Employee.DateOfHire.ToShortDateString();

                    await package.SaveAsAsync(new FileInfo(filePath));
                }

                return RedirectToPage("EmployeeList");
            }
            catch (Exception)
            {
                ModelState.AddModelError(string.Empty, "An error occurred while saving the data.");
                OnGet(); // Reload dropdown options
                return Page();
            }
        }

        private void LoadDropdownOptions()
        {
            GradeOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "A3", Text = "A3" },
                new SelectListItem { Value = "A4", Text = "A4" },
                new SelectListItem { Value = "A5", Text = "A5" },
                new SelectListItem { Value = "B1", Text = "B1" },
                new SelectListItem { Value = "B2", Text = "B2" },
                new SelectListItem { Value = "C1", Text = "C1" },
                new SelectListItem { Value = "C2", Text = "C2" },
                new SelectListItem { Value = "D1", Text = "D1" },
                new SelectListItem { Value = "D2", Text = "D2" }
            };

            BUOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "C&CA", Text = "C&CA" },
                new SelectListItem { Value = "FS", Text = "FS" },
                new SelectListItem { Value = "Sogeti", Text = "Sogeti" },
                new SelectListItem { Value = "Infra", Text = "Infra" },
                new SelectListItem { Value = "Digital", Text = "Digital" }
            };
        }
    }
}
