using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace ToolSite.Controllers
{
    public class AccountantController : Controller
    {

        private readonly IHostingEnvironment env;

        public AccountantController(IHostingEnvironment env)
        {
            this.env = env;
        }

        public ActionResult Other()
        {
            return View();
        }

        [HttpPost]
        public async Task<PartialViewResult> ProcessOtherData()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var files = Request.Form.Files;
            var filePaths = files.Select(f =>
            {
                var aaa = 1;
                return Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            }).ToList();

            if (files.Count > 0)
            {
                for (var idx = 0; idx < filePaths.Count; idx++)
                {
                    var filePath = filePaths[idx];
                    using (var targetStream = System.IO.File.Create(filePath))
                    {
                        await files[idx].CopyToAsync(targetStream);
                    }
                }
            }



            //ViewBag.DowloadFileName = orderFileName;
            ViewBag.DowloadFileName = "";
            return PartialView("_MetadataDowload");
        }


    }

    public class _高新区企业从业人员情况
    {
        public string _年份 { get; set; }
        public int _年末就业 { get; set; }
        public int _留学归国 { get; set; }
        public int _外籍常驻 { get; set; }
        public int _大专以上 { get; set; }
        public int _中高级 { get; set; }
        //public int _年份 { get; set; }
    }
}
