using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace ToolSite.Controllers
{
    /// <summary>
    /// 考勤计算控制器
    /// </summary>
    public class SalariesController : Controller
    {
        private readonly IHostingEnvironment env;

        public SalariesController(IHostingEnvironment env)
        {
            this.env = env;
        }

        /// <summary>
        /// 仓库加班考勤
        /// </summary>
        /// <returns></returns>
        public ActionResult WarehouseOvertime()
        {
            return View();
        }
    }
}