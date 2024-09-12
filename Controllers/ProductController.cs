using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace AdminPanelCrud.Controllers
{
    public class ProductController : Controller
    {
        private IConfiguration configuration;

        public ProductController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }

        public IActionResult Index()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Product_Select_All";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return View(table);
        }

        public IActionResult ProductDelete(int ProductID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_Product_Delete";
                command.Parameters.Add("@ProductID", SqlDbType.Int).Value = ProductID;
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = ex.Message;
                Console.WriteLine(ex.ToString());
            }
            return RedirectToAction("Index");
        }

        //public void DropDown()
        //{
            
        //}


        public IActionResult ProductAddEdit(int ProductID)
        {

            string connectionString1 = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection1 = new SqlConnection(connectionString1);
            connection1.Open();
            SqlCommand command1 = connection1.CreateCommand();
            command1.CommandType = CommandType.StoredProcedure;
            command1.CommandText = "PR_User_DropDown";
            SqlDataReader reader1 = command1.ExecuteReader();
            DataTable dataTable1 = new DataTable();
            dataTable1.Load(reader1);
            connection1.Close();
            List<UserDropDownModel> userList = new List<UserDropDownModel>();

            foreach (DataRow data in dataTable1.Rows)
            {
                UserDropDownModel userDropDownModel = new UserDropDownModel();
                userDropDownModel.UserID = Convert.ToInt32(data["UserID"]);
                userDropDownModel.UserName = data["UserName"].ToString();
                userList.Add(userDropDownModel);
            }

            ViewBag.UserList = userList;

            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Product_Select_By_Primary_Key";
            command.Parameters.AddWithValue("@ProductID", ProductID);
            Console.WriteLine(ProductID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            connection.Close();
            Product productModel = new Product();

            foreach (DataRow dataRow in table.Rows)
            {
                productModel.ProductName = @dataRow["ProductName"].ToString();
                productModel.ProductCode = @dataRow["ProductCode"].ToString();
                productModel.ProductPrice = Convert.ToDecimal(@dataRow["ProductPrice"]);
                productModel.Description = @dataRow["Description"].ToString();
                productModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }

            return View("ProductAddEdit", productModel);
        }

        [HttpPost]
        public IActionResult ProductSave(Product productModel)
        {
            if (productModel.UserID <= 0)
            {
                ModelState.AddModelError("UserID", "A valid User is required.");
            }
            if (ModelState.IsValid)
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                Console.WriteLine(productModel.ProductID);
                if (productModel.ProductID == null )
                {
                    command.CommandText = "PR_Product_Insert";
                    
                }
                else
                {
                    command.CommandText = "PR_Product_Update";
                    command.Parameters.Add("@ProductID", SqlDbType.Int).Value = productModel.ProductID;
                }
                command.Parameters.Add("@ProductName", SqlDbType.VarChar).Value = productModel.ProductName;
                command.Parameters.Add("@ProductPrice", SqlDbType.Decimal).Value = productModel.ProductPrice;
                command.Parameters.Add("@ProductCode", SqlDbType.VarChar).Value = productModel.ProductCode;
                command.Parameters.Add("@Description", SqlDbType.VarChar).Value = productModel.Description;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = productModel.UserID;
                command.ExecuteNonQuery();
                return RedirectToAction("Index");
            }
            return View("ProductAddEdit", productModel);
        }
    }
}
