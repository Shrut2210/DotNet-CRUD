using System.ComponentModel.DataAnnotations;

namespace AdminPanelCrud.Models
{
    public class Order
    {
        [Required]
        public int OrderID { get; set; }
        [Required(ErrorMessage = " Enter field!!!")]
        public DateTime OrderDate { get; set; }
        [Required]
        public int CustomerID { get; set; }
            [Required]
        public string PaymentMode { get; set; }
            [Required(ErrorMessage = " Plz enter field!!!")]
        public decimal TotalAmount { get; set; }
        [Required]
        public string ShippingAddress { get; set; }
        [Required]
        public int UserID { get; set; }
    }
}
