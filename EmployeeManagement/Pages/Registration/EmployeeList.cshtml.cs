using EmplyoeeManagement.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;


namespace EmplyoeeManagement.Pages.Registration
{
    public class EmployeeListModel : PageModel
    {
        public List<Employee> Employees { get; set; }
        private string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");

        public void OnGet()
        {
            // Initialize the list of employees
            Employees = new List<Employee>();

            // Check if the Excel file exists
            if (System.IO.File.Exists(filePath))
            {
                // Load the Excel file using EPPlus
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets["Employees"];

                    // Check if worksheet is available
                    if (worksheet != null)
                    {
                        // Start reading from the second row to skip headers
                        int row = 2;
                        while (worksheet.Cells[row, 1].Value != null)
                        {
                            // Map Excel cells to Employee properties
                            Employees.Add(new Employee
                            {
                                EmployeeId = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                                FirstName = worksheet.Cells[row, 2].Value.ToString(),
                                LastName = worksheet.Cells[row, 3].Value.ToString(),
                                Email = worksheet.Cells[row, 4].Value.ToString(),
                                Phone = worksheet.Cells[row, 5].Value.ToString(),
                                Grade = worksheet.Cells[row, 6].Value.ToString(),
                                BU = worksheet.Cells[row, 7].Value.ToString(),
                                DateOfHire = DateTime.Parse(worksheet.Cells[row, 8].Value.ToString())
                            });

                            row++; // Move to the next row
                        }
                    }
                }
            }
        }

        public object GetEmployees()
        {
            return Employees;
        }

public Employee Employee { get; set; }
        

       public async Task<IActionResult> OnPostDeleteAsync(int id)
{
    var tempFilePath = Path.Combine(Path.GetTempPath(), "EmployeeData_temp.xlsx");
    System.IO.File.Copy(filePath, tempFilePath, overwrite: true);

    try
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets["Employees"];
            var rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                if (int.Parse(worksheet.Cells[row, 1].Text) == id)
                {
                    worksheet.DeleteRow(row);
                    await package.SaveAsync();
                    break;
                }
            }
        }

        // After successful deletion, redirect to the same page to reload the employee list
        return RedirectToPage("/Registration/EmployeeList");
    }
    catch (Exception ex)
    {
        // Handle any errors (optional: log the error)
        ModelState.AddModelError(string.Empty, "An error occurred while deleting the employee.");
        return Page(); // Stay on the same page to show the error
    }
    finally
    {
        if (System.IO.File.Exists(tempFilePath))
        {
            System.IO.File.Delete(tempFilePath); // Clean up temp file
        }
    }
}


    
   }

}

