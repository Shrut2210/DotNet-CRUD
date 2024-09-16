using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;

namespace AdminPanelCrud.Controllers
{
    public class Ordercontroller : Controller
    {

        private IConfiguration configuration;

        public Ordercontroller(IConfiguration _configuration)
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
            command.CommandText = "PR_Order_Select_All";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return View(table);
        }

        public IActionResult OrderAddEdit(int OrderID)
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection1 = new SqlConnection(connectionString);
            connection1.Open();
            SqlCommand command1 = connection1.CreateCommand();
            command1.CommandType = System.Data.CommandType.StoredProcedure;
            command1.CommandText = "PR_User_DropDown";
            SqlDataReader reader1 = command1.ExecuteReader();
            DataTable dataTable1 = new DataTable();
            dataTable1.Load(reader1);
            List<UserDropDownModel> userList = new List<UserDropDownModel>();

            foreach (DataRow data in dataTable1.Rows)
            {
                UserDropDownModel userDropDownModel = new UserDropDownModel();
                userDropDownModel.UserID = Convert.ToInt32(data["UserID"]);
                userDropDownModel.UserName = data["UserName"].ToString();
                userList.Add(userDropDownModel);
            }
            ViewBag.UserList = userList;

            SqlConnection connection2 = new SqlConnection(connectionString);
            SqlCommand command2 = connection1.CreateCommand();
            command2.CommandText = "PR_Customer_DropDown";
            SqlDataReader reader2 = command2.ExecuteReader();
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            List<CustomerDropDownModel> customerList = new List<CustomerDropDownModel>();

            foreach (DataRow data in dataTable2.Rows)
            {
                CustomerDropDownModel customereDropDownModel = new CustomerDropDownModel();
                customereDropDownModel.CustomerId = Convert.ToInt32(data["CustomerID"]);
                customereDropDownModel.CustomerName = data["CustomerName"].ToString();
                customerList.Add(customereDropDownModel);
            }
            ViewBag.CustomerList = customerList;

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Order_Select_By_Primary_Key";
            command.Parameters.AddWithValue("@OrderID", OrderID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            connection.Close();
            Order orderModel = new Order();

            foreach (DataRow dataRow in table.Rows)
            {
                orderModel.OrderDate = Convert.ToDateTime(@dataRow["OrderDate"]);
                orderModel.CustomerID = Convert.ToInt32(@dataRow["CustomerID"]);
                orderModel.TotalAmount = Convert.ToDecimal(@dataRow["TotalAmount"]);
                orderModel.PaymentMode = Convert.ToString(@dataRow["PaymentMode"]);
                orderModel.ShippingAddress = Convert.ToString(dataRow["ShippingAddress"]);
                orderModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }
            return View("OrderAddEdit", orderModel);
        }

        [HttpPost]
        public ActionResult OrderSave(Order orderModel)
        {
            if (orderModel.UserID <= 0)
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
                if (orderModel.OrderID == null || orderModel.OrderID == 0)
                {
                    command.CommandText = "PR_Order_Insert";
                }
                else
                {
                    command.CommandText = "PR_Order_Update";
                    command.Parameters.Add("@OrderID", SqlDbType.Int).Value = orderModel.OrderID;

                }
                command.Parameters.Add("@OrderDate", SqlDbType.DateTime).Value = orderModel.OrderDate;
                command.Parameters.Add("@CustomerID", SqlDbType.Int).Value = orderModel.CustomerID;
                command.Parameters.Add("@PaymentMode", SqlDbType.VarChar).Value = orderModel.PaymentMode;
                command.Parameters.Add("@TotalAmount", SqlDbType.Decimal).Value = orderModel.TotalAmount;
                command.Parameters.Add("@ShippingAddress", SqlDbType.VarChar).Value = orderModel.ShippingAddress;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = orderModel.UserID;
                command.ExecuteNonQuery();
                return RedirectToAction("Index");
            }
            return View("OrderAddEdit", orderModel);
        }

        public IActionResult ExportToExcel()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Order_Select_All";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add the headers
                worksheet.Cells[1, 1].Value = "OrderID";
                worksheet.Cells[1, 2].Value = "OrderDate";
                worksheet.Cells[1, 3].Value = "CustomerName";
                worksheet.Cells[1, 4].Value = "PaymentMode";
                worksheet.Cells[1, 5].Value = "TotalAmount";
                worksheet.Cells[1, 6].Value = "ShippingAddress";
                worksheet.Cells[1, 7].Value = "UserName";

                // Add the data
                int rowNumber = 0;
                foreach (DataRow row in table.Rows)
                {
                    worksheet.Cells[rowNumber + 2, 1].Value = row["OrderID"];
                    worksheet.Cells[rowNumber + 2, 2].Value = row["OrderDate"];
                    worksheet.Cells[rowNumber + 2, 3].Value = row["CustomerName"];
                    worksheet.Cells[rowNumber + 2, 4].Value = row["PaymentMode"];
                    worksheet.Cells[rowNumber + 2, 5].Value = row["TotalAmount"];
                    worksheet.Cells[rowNumber + 2, 6].Value = row["ShippingAddress"];
                    worksheet.Cells[rowNumber + 2, 7].Value = row["UserName"];
                    rowNumber++;
                }
                var fileBytes = package.GetAsByteArray();
                var fileName = "OrdersData.xlsx";

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
    }
}
