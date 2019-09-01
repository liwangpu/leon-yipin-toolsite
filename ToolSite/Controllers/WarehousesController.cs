using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace ToolSite.Controllers
{
    /// <summary>
    /// 仓库管理控制器
    /// </summary>
    public class WarehousesController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 库存盘点
        /// </summary>
        /// <returns></returns>
        public IActionResult StockTaking()
        {
            return View();
        }

    }
}