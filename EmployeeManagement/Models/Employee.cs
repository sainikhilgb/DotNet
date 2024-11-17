using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace EmplyoeeManagement.Models
{

    public class Employee
    {
        [Key]       
        public int EmployeeId {get; set;}
        [Required(ErrorMessage = "Please enter your first name")]
        [StringLength(50)]
        [DisplayName("First name")]
        public string FirstName {get; set;} 
         [Required(ErrorMessage = "Please enter your Last name")]
        [StringLength(50)]
        [DisplayName("Last name")]
        public string LastName {get; set;}
        [Required(ErrorMessage = "Please enter your email address")]
        [EmailAddress]  
        public string Email {get; set;}
        [Required(ErrorMessage = "Please enter your phone number")]
        [Phone]
        public string Phone {get; set;}
        public string Grade {get; set;}
        public string BU {get; set;}
        [Required(ErrorMessage = "Please enter the date of hire")]
        [DataType(DataType.Date)]
        [DisplayName("Date Of Hire")]
         public DateTime DateOfHire {get; set;}

    }

}