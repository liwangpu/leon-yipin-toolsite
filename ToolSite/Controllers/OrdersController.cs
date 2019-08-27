using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ToolSite.Controllers
{
    public class OrdersController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public ActionResult ExtractArea()
        {
            return View();
        }

        [HttpPost]
        public async Task<JsonResult> ExtractAreaHandle()
        {
            var tmpFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
            if (!Directory.Exists(tmpFolder))
                Directory.CreateDirectory(tmpFolder);
            var orderFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var extractArea = Request.Form["area"];
            var files = Request.Form.Files;

            if (files.Count > 0)
            {
                using (var targetStream = System.IO.File.Create(orderFilePath))
                    await files[0].CopyToAsync(targetStream);
            }



            return Json(new { Name = "test" });
        }
    }

}