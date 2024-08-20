using System.ComponentModel.DataAnnotations;

namespace AdminPanelCrud.Models
{
    public class Product
    {
        [Required]
        public int ProductID { get; set; }

        [Required]
        public string ProductName { get; set; }

        [Required]
        public decimal ProductPrice { get; set; }

        [Required]
        public string ProductCode { get; set; }

        [Required]
        public string Description { get; set; }

        [Required]
        public int UserID { get; set; }
    }
}
