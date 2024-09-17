using AdminPanelCrud.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace AdminPanelCrud.Controllers
{
    [CheckAccess]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult CreateCookie()
        {
            string key = "Zoro";
            string value = "You are dead!!!!!!!!!!!!!";

            CookieOptions options = new CookieOptions
            {
                Expires = DateTime.Now.AddDays(2)
            };

            Response.Cookies.Append(key, value, options);
            return View("Index");
        }

        public IActionResult ReadCookie()
        {
            string key = "Zoro";
            string cookieValue = Request.Cookies[key];
            return View("Index");
        }

        public IActionResult RemoveCookie()
        {
            string key = "Zoro";
            string value = string.Empty;

            CookieOptions options = new CookieOptions
            {
                Expires = DateTime.Now.AddDays(-1)
            };

            Response.Cookies.Append(key, value, options);
            return View("Index");
        }

    }
}
