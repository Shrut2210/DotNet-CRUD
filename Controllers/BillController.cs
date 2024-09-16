using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace AdminPanelCrud.Controllers
{
    public class BillController : Controller
    {
        private IConfiguration configuration;

        public BillController(IConfiguration _configuration)
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
            command.CommandText = "PR_Select_All_Bills";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return View(table);
        }

        public IActionResult BillDelete(int BillID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_Bills_Delete";
                command.Parameters.Add("@BillID", SqlDbType.Int).Value = BillID;
                command.ExecuteNonQuery();

            }

            catch (Exception ex)
            {
                TempData["ErrorMessage"] = ex.Message;
                Console.WriteLine(ex.ToString());
            }

            
            return RedirectToAction("Index");
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

            SqlConnection connection2 = new SqlConnection(connectionString);
            SqlCommand command2 = connection1.CreateCommand();
            command2.CommandText = "PR_Order_DropDown";
            SqlDataReader reader2 = command2.ExecuteReader();
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            List<OrderDropDownModel> orderList = new List<OrderDropDownModel>();

            foreach (DataRow data in dataTable2.Rows)
            {
                OrderDropDownModel orderDropDownModel = new OrderDropDownModel();
                orderDropDownModel.OrderID = Convert.ToInt32(data["OrderID"]);
                orderList.Add(orderDropDownModel);
            }
            ViewBag.OrderList = orderList;
        }

        public IActionResult BillAddEdit(int BillID)
        {
            Drop_Down();
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Bills_Select_By_Primary_Key";
            command.Parameters.AddWithValue("@BillID", BillID);
            Console.WriteLine(BillID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            connection.Close();
            Bill billModel = new Bill();

            foreach (DataRow dataRow in table.Rows)
            {
                billModel.BillNumber = @dataRow["BillNumber"].ToString();
                billModel.BillDate = Convert.ToDateTime(@dataRow["BillDate"]);
                billModel.OrderID = Convert.ToInt32(@dataRow["OrderID"]);
                billModel.TotalAmount = Convert.ToDecimal(@dataRow["TotalAmount"]);
                billModel.Discount = Convert.ToDecimal(@dataRow["Discount"]);
                billModel.NetAmount = Convert.ToDecimal(@dataRow["NetAmount"]);
                billModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }
            return View("BillAddEdit", billModel);
        }

        [HttpPost]
        public IActionResult BillSave(Bill billModel)
        {
            int flag = 0;

            if (billModel.UserID <= 0)
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
                Console.WriteLine(billModel.BillID);
                if (billModel.BillID == null || billModel.BillID == 0 )
                {
                    command.CommandText = "PR_Bills_Insert";

                }
                else
                {
                    flag = 1;
                    command.CommandText = "PR_Bills_Update";
                    command.Parameters.Add("@BillID", SqlDbType.Int).Value = billModel.BillID;
                }
                command.Parameters.Add("@BillNumber", SqlDbType.VarChar).Value = billModel.BillNumber;
                command.Parameters.Add("@BillDate", SqlDbType.DateTime).Value = billModel.BillDate;
                command.Parameters.Add("@OrderID", SqlDbType.Int).Value = billModel.OrderID;
                command.Parameters.Add("@TotalAmount", SqlDbType.Decimal).Value = billModel.TotalAmount;
                command.Parameters.Add("@Discount", SqlDbType.Decimal).Value = billModel.Discount;
                command.Parameters.Add("@NetAmount", SqlDbType.Decimal).Value = billModel.NetAmount;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = billModel.UserID;
                command.ExecuteNonQuery();
                return RedirectToAction("Index");
            }

            if (flag == 0)
            {
                
            }
            else
            {

            }

            return View("BillAddEdit", billModel);
        }

        public IActionResult ExportToExcel()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Select_All_Bills";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add the headers
                worksheet.Cells[1, 1].Value = "BillID";
                worksheet.Cells[1, 2].Value = "BillNumber";
                worksheet.Cells[1, 3].Value = "BillDate";
                worksheet.Cells[1, 4].Value = "TotalAmount";
                worksheet.Cells[1, 5].Value = "Discount";
                worksheet.Cells[1, 6].Value = "NetAmount";
                worksheet.Cells[1, 7].Value = "UserName";

                // Add the data
                int rowNumber = 0;
                foreach (DataRow row in table.Rows)
                {
                    worksheet.Cells[rowNumber + 2, 1].Value = row["BillID"];
                    worksheet.Cells[rowNumber + 2, 2].Value = row["BillNumber"];
                    worksheet.Cells[rowNumber + 2, 3].Value = row["BillDate"];
                    worksheet.Cells[rowNumber + 2, 4].Value = row["TotalAmount"];
                    worksheet.Cells[rowNumber + 2, 5].Value = row["Discount"];
                    worksheet.Cells[rowNumber + 2, 6].Value = row["NetAmount"];
                    worksheet.Cells[rowNumber + 2, 7].Value = row["UserName"];
                    rowNumber++;
                }
                var fileBytes = package.GetAsByteArray();
                var fileName = "BillsData.xlsx";

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
    }
}
