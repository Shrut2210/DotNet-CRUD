using System.ComponentModel.DataAnnotations;

namespace AdminPanelCrud.Models
{
    public class User
    {
        [Required]
        public int UserID { get; set; }
        [Required]
        public string UserName { get; set; }
        [Required]
        public string Email { get; set; }
          [Required]
        public string Password { get; set; }
        [Required]
        public string MobileNo { get; set; }
        [Required]
        public string Address { get; set; }
            [Required]
        public bool IsActive { get; set; }

    }
}
