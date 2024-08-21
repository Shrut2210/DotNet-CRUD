using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
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
        public IActionResult OrderAddEdit()
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
            return View();
        }
    }
}
