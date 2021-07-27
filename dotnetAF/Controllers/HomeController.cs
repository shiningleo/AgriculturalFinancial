using Microsoft.AspNetCore.Mvc;

namespace dotnetHelloWorld.Controllers
{
    public class HomeController : Controller
    {
        // GET: /<controller>/
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult nongchang()
        {
             ViewData["tuiguang"] = "97";
             ViewData["licheng"] = "中级酒庄";
             ViewData["tongshu"] = "5000";
            return View();
        }
        
    }
}
