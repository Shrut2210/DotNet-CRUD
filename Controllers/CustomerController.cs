using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;

namespace AdminPanelCrud.Controllers
{
    public class CustomerController : Controller
    {
        private IConfiguration configuration;

        public CustomerController(IConfiguration _configuration)
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
            command.CommandText = "PR_Customer_Select_All";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return View(table);
        }

        public void Drop_Down()
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
        }
        public IActionResult CustomerAddEdit(int CustomerID)
        {
            Drop_Down();
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Customer_Select_By_Primary_Key";
            command.Parameters.AddWithValue("@CustomerID", CustomerID);
            Console.WriteLine(CustomerID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            connection.Close();
            Customer customerModel = new Customer();

            foreach (DataRow dataRow in table.Rows)
            {
                customerModel.CustomerName = @dataRow["CustomerName"].ToString();
                customerModel.HomeAddress = @dataRow["HomeAddress"].ToString();
                customerModel.Email = @dataRow["Email"].ToString();
                customerModel.MobileNo = @dataRow["MobileNo"].ToString();
                customerModel.GSTNO = @dataRow["GSTNO"].ToString();
                customerModel.CityName = @dataRow["CityName"].ToString();
                customerModel.PinCode = @dataRow["PinCode"].ToString();
                customerModel.NetAmount = Convert.ToDecimal(@dataRow["NetAmount"]);
                customerModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }
            return View("CustomerAddEdit", customerModel);
        }

        public ActionResult CustomerSave(Customer customerModel)
        {
            if (customerModel.UserID <= 0)
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
                Console.WriteLine(customerModel.CustomerID);
                if (customerModel.CustomerID == null || customerModel.CustomerID == 0 )
                {
                    command.CommandText = "PR_Customer_Insert";

                }
                else
                {
                    command.CommandText = "PR_Customers_Update";
                    command.Parameters.Add("@CustomerID", SqlDbType.Int).Value = customerModel.CustomerID;
                }
                command.Parameters.Add("@CustomerName", SqlDbType.VarChar).Value = customerModel.CustomerName;
                command.Parameters.Add("@HomeAddress", SqlDbType.VarChar).Value = customerModel.HomeAddress;
                command.Parameters.Add("@Email", SqlDbType.VarChar).Value = customerModel.Email;
                command.Parameters.Add("@MobileNo", SqlDbType.VarChar).Value = customerModel.MobileNo;
                command.Parameters.Add("@GSTNO", SqlDbType.VarChar).Value = customerModel.GSTNO;
                command.Parameters.Add("@CityName", SqlDbType.VarChar).Value = customerModel.CityName;
                command.Parameters.Add("@PinCode", SqlDbType.VarChar).Value = customerModel.PinCode;
                command.Parameters.Add("@NetAmount", SqlDbType.Decimal).Value = customerModel.NetAmount;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = customerModel.UserID;
                command.ExecuteNonQuery();
                return RedirectToAction("Index");
            }
            return View( "CustomerAddEdit" ,customerModel);
        }

        public IActionResult ExportToExcel()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Customer_Select_All";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add the headers
                worksheet.Cells[1, 1].Value = "CustomerID";
                worksheet.Cells[1, 2].Value = "CustomerName";
                worksheet.Cells[1, 3].Value = "HomeAddress";
                worksheet.Cells[1, 4].Value = "Email";
                worksheet.Cells[1, 5].Value = "MobileNo";
                worksheet.Cells[1, 6].Value = "GSTNO";
                worksheet.Cells[1, 7].Value = "CityName";
                worksheet.Cells[1, 8].Value = "PinCode";
                worksheet.Cells[1, 9].Value = "NetAmount";

                // Add the data
                int rowNumber = 0;
                foreach (DataRow row in table.Rows)
                {
                    worksheet.Cells[rowNumber + 2, 1].Value = row["CustomerID"];
                    worksheet.Cells[rowNumber + 2, 2].Value = row["CustomerName"];
                    worksheet.Cells[rowNumber + 2, 3].Value = row["HomeAddress"];
                    worksheet.Cells[rowNumber + 2, 4].Value = row["Email"];
                    worksheet.Cells[rowNumber + 2, 5].Value = row["MobileNo"];
                    worksheet.Cells[rowNumber + 2, 6].Value = row["GSTNO"];
                    worksheet.Cells[rowNumber + 2, 7].Value = row["CityName"];
                    worksheet.Cells[rowNumber + 2, 8].Value = row["PinCode"];
                    worksheet.Cells[rowNumber + 2, 9].Value = row["NetAmount"];
                    rowNumber++;
                }
                var fileBytes = package.GetAsByteArray();
                var fileName = "CustomersData.xlsx";

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
    }
}
