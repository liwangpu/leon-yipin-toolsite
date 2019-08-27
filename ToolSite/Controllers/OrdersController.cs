using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using Microsoft.AspNetCore.Hosting;

namespace ToolSite.Controllers
{
    public class OrdersController : Controller
    {
        private readonly IHostingEnvironment env;

        public OrdersController(IHostingEnvironment env)
        {
            this.env = env;
        }

        public IActionResult Index()
        {
            return View();
        }

        public ActionResult ExtractArea()
        {
            return View();
        }

        [HttpPost]
        public async Task<PartialViewResult> ExtractAreaHandle()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var orderFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var extractArea = Request.Form["area"];
            var files = Request.Form.Files;

            if (files.Count > 0)
            {
                using (var targetStream = System.IO.File.Create(orderFilePath))
                    await files[0].CopyToAsync(targetStream);
            }

            ViewBag.DowloadFileName = "bob";
            return PartialView("_MetadataDowload");
        }
    }

}