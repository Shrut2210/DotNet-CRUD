using AdminPanelCrud.DAL;
using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
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
            //if (userModel.UserID <= 0)
            //{
            //    ModelState.AddModelError("UserID", "A valid User is required.");
            //}
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
                    Console.WriteLine("asbxhsbvxh");
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

        public IActionResult ExportToExcel()
        {
            string connectionString = this._configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_User_Select_All";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add the headers
                worksheet.Cells[1, 1].Value = "UserID";
                worksheet.Cells[1, 2].Value = "UserName";
                worksheet.Cells[1, 3].Value = "Email";
                worksheet.Cells[1, 4].Value = "Password";
                worksheet.Cells[1, 5].Value = "MobileNo";
                worksheet.Cells[1, 6].Value = "Address";

                // Add the data
                int rowNumber = 0;
                foreach (DataRow row in table.Rows)
                {
                    worksheet.Cells[rowNumber + 2, 1].Value = row["UserID"];
                    worksheet.Cells[rowNumber + 2, 2].Value = row["UserName"];
                    worksheet.Cells[rowNumber + 2, 3].Value = row["Email"];
                    worksheet.Cells[rowNumber + 2, 4].Value = row["Password"];
                    worksheet.Cells[rowNumber + 2, 5].Value = row["MobileNo"];
                    worksheet.Cells[rowNumber + 2, 6].Value = row["Address"];
                    rowNumber++;
                }
                var fileBytes = package.GetAsByteArray();
                var fileName = "UserData.xlsx";

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        [HttpPost]
        public IActionResult Login(UserLoginModel userLoginModel)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    string connectionString = this._configuration.GetConnectionString("ConnectionString");
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.CommandText = "PR_User_Login";
                    sqlCommand.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userLoginModel.UserName;
                    sqlCommand.Parameters.Add("@Password", SqlDbType.VarChar).Value = userLoginModel.Password;
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                    DataTable dataTable = new DataTable();
                    dataTable.Load(sqlDataReader);
                    if (dataTable.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            HttpContext.Session.SetString("UserID", dr["UserID"].ToString());
                            HttpContext.Session.SetString("UserName", dr["UserName"].ToString());
                        }
                        return RedirectToAction("Index", "User");

                    }

                }
            }
            catch (Exception e)
            {
                TempData["ErrorMessage"] = e.Message;
            }

            return RedirectToAction("Login");
        }

        public ActionResult Login()
        {
            return View();
        }
        public IActionResult UserRegister(UserRegisterModel userRegisterModel)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    string connectionString = this._configuration.GetConnectionString("ConnectionString");
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.CommandText = "PR_User_Register";
                    sqlCommand.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userRegisterModel.UserName;
                    sqlCommand.Parameters.Add("@Password", SqlDbType.VarChar).Value = userRegisterModel.Password;
                    sqlCommand.Parameters.Add("@Email", SqlDbType.VarChar).Value = userRegisterModel.Email;
                    sqlCommand.Parameters.Add("@MobileNo", SqlDbType.VarChar).Value = userRegisterModel.MobileNo;
                    sqlCommand.Parameters.Add("@Address", SqlDbType.VarChar).Value = userRegisterModel.Address;
                    sqlCommand.ExecuteNonQuery();
                    return RedirectToAction("Login", "User");
                }
            }
            catch (Exception e)
            {
                TempData["ErrorMessage"] = e.Message;
                return RedirectToAction("Register");
            }
            return RedirectToAction("Register");
        }

        [HttpPost]
        public IActionResult Logout()
        {
            HttpContext.Session.Clear();
            return RedirectToAction("Login", "User");
        }
    }
}
