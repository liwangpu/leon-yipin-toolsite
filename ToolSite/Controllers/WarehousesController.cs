using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace ToolSite.Controllers
{
    /// <summary>
    /// 仓库管理控制器
    /// </summary>
    public class WarehousesController : Controller
    {
        private readonly IHostingEnvironment env;

        public WarehousesController(IHostingEnvironment env)
        {
            this.env = env;
        }

        /// <summary>
        /// 库存盘点
        /// </summary>
        /// <returns></returns>
        public IActionResult StockTaking()
        {
            return View();
        }

        [HttpPost]
        public async Task<PartialViewResult> StockTakingHandler()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var files = Request.Form.Files;
            var stockoutFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");//缺货表
            var stockFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");//库存表

            if (files.Count > 0)
            {
                var fstockout = files.FirstOrDefault(x => x.Name == "stockout");
                if (fstockout != null)
                {
                    using (var targetStream = System.IO.File.Create(stockoutFilePath))
                        await fstockout.CopyToAsync(targetStream);
                }

                var fstock = files.FirstOrDefault(x => x.Name == "stock");
                if (fstock != null)
                {
                    using (var targetStream = System.IO.File.Create(stockFilePath))
                        await fstock.CopyToAsync(targetStream);
                }
            }







            ViewBag.DowloadFileName = "";
            return PartialView("_MetadataDowload");
        }

    }
}