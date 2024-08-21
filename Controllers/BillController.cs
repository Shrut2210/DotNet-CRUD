using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Data.SqlClient;

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

        public IActionResult BillAddEdit()
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
            return View();
        }
    }
}
