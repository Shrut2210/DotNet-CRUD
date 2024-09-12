using AdminPanelCrud.DAL;
using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Data.SqlClient;

namespace AdminPanelCrud.Controllers
{

    public class UserController : Controller
    {

        private IConfiguration _configuration;

        public UserController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            string connectionString = this._configuration.GetConnectionString("ConnectionString");
            SqlConnection sqlConnection = new SqlConnection(connectionString);
            sqlConnection.Open();
            SqlCommand cmd = sqlConnection.CreateCommand();
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandText = "PR_User_Select_All";
            SqlDataReader reader = cmd.ExecuteReader();
            DataTable dataTable = new DataTable();
            dataTable.Load(reader);
            return View(dataTable);
        }

        public IActionResult UserDelete(int UserID)
        {
            try
            {
                string connectionString = this._configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_User_Delete";
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = UserID;
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = ex.Message;
                Console.WriteLine(ex.ToString());
            }
            return RedirectToAction("Index");
        }

        public IActionResult UserAddEdit(int UserID)
        {
            string connectionString = this._configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_User_Select_By_Primary_Key";
            command.Parameters.AddWithValue("@UserID", UserID);
            Console.WriteLine(UserID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            connection.Close();
            User userModel = new User();

            foreach (DataRow dataRow in table.Rows)
            {
                userModel.UserName = @dataRow["UserName"].ToString();
                userModel.Email = @dataRow["Email"].ToString();
                userModel.Password = @dataRow["Password"].ToString();
                userModel.MobileNo = @dataRow["MobileNo"].ToString();
                userModel.Address = @dataRow["Address"].ToString();
                userModel.IsActive = Convert.ToBoolean(@dataRow["IsActive"]);
            }
            return View("UserAddEdit", userModel);
        }

        public IActionResult UserSave(User userModel)
        {
            if (userModel.UserID <= 0)
            {
                ModelState.AddModelError("UserID", "A valid User is required.");
            }
            if (ModelState.IsValid)
            {
                string connectionString = this._configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                Console.WriteLine(userModel.UserID);
                if (userModel.UserID == null || userModel.UserID == 0)
                {
                    command.CommandText = "PR_User_Insert";

                }
                else
                {
                    command.CommandText = "PR_User_Update";
                    command.Parameters.Add("@UserId", SqlDbType.Int).Value = userModel.UserID;
                }
                command.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userModel.UserName;
                command.Parameters.Add("@Email", SqlDbType.VarChar).Value = userModel.Email;
                command.Parameters.Add("@Password", SqlDbType.VarChar).Value = userModel.Password;
                command.Parameters.Add("@MobileNo", SqlDbType.VarChar).Value = userModel.MobileNo;
                command.Parameters.Add("@Address", SqlDbType.VarChar).Value = userModel.Address;
                command.Parameters.Add("@IsActive", SqlDbType.Bit).Value = userModel.IsActive;
                command.ExecuteNonQuery();
                return RedirectToAction("Index");
            }
            return View("UserAddEdit", userModel);
        }
    }
}
