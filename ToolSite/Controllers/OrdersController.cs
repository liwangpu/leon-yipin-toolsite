﻿using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

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
            var orderFileName = Guid.NewGuid().ToString() + ".xlsx";
            var orderFilePath = Path.Combine(tmpFolder, orderFileName);
            var files = Request.Form.Files;

            if (files.Count > 0)
            {
                using (var targetStream = System.IO.File.Create(orderFilePath))
                    await files[0].CopyToAsync(targetStream);
            }


            using (var package = new ExcelPackage(new FileInfo(orderFilePath)))
            {
                var mapping = new Dictionary<string, List<int>>();
                var dataSheet = package.Workbook.Worksheets[0];

                var endRow = dataSheet.Dimension.End.Row;
                var endColumn = dataSheet.Dimension.End.Column;
                var areaFlagColumn = endColumn + 1;

                for (int idx = endRow; idx >= 2; idx--)
                {
                    var areaObj = dataSheet.Cells[idx, 8].Value;
                    if (areaObj != null)
                    {
                        var areaStr = areaObj.ToString();
                        var areas = areaStr.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Substring(0, 1).ToUpper()).Distinct().ToArray();
                        if (areas.Count() == 1)
                        {
                            var _area = areas[0];
                            if (!mapping.ContainsKey(_area))
                                mapping[_area] = new List<int>();
                            mapping[_area].Add(idx);
                        }

                    }
                }

                var filterAreaStr = Request.Form["area"];
                var areaNames = mapping.Keys.OrderBy(x => x).ToList();
                if (!string.IsNullOrWhiteSpace(filterAreaStr))
                {
                    var s = filterAreaStr.ToString().Replace("，", ",").ToUpper();
                    var farr = s.Split(',', StringSplitOptions.RemoveEmptyEntries).ToList();
                    if (farr.Count > 0)
                    {
                        areaNames.Clear();
                        areaNames.AddRange(farr);
                    }
                }

                foreach (var areaName in areaNames)
                {
                    if (mapping.ContainsKey(areaName))
                    {
                        var sheet = package.Workbook.Worksheets.Add(areaName + "区域");
                        var rows = mapping[areaName];
                        //拷贝标题行
                        dataSheet.Cells[1, 1, 1, endColumn].Copy(sheet.Cells[1, 1, 1, endColumn]);
                        //数据行
                        for (int i = 0, len = rows.Count; i < len; i++)
                        {
                            var r = rows[i];
                            dataSheet.Cells[r, 1, r, endColumn].Copy(sheet.Cells[i + 2, 1, i + 2, endColumn]);
                        }
                    }
                }

                //删除原数据表
                package.Workbook.Worksheets.Delete(0);
                package.Save();
            }


            ViewBag.DowloadFileName = $"/tmp/{orderFileName}";
            return PartialView("_MetadataDowload");
        }

    }

}