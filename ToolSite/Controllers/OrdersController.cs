using Microsoft.AspNetCore.Hosting;
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
        public async Task<PartialViewResult> ExtractSingleAreaHandle()
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
                //默认的话,"库位号"在第8列,但是也有可能改变
                var areaColumn = 8;
                if (dataSheet.Cells[1, 8].Value == null || dataSheet.Cells[1, 8].Value.ToString().Trim() != "库位号")
                {
                    for (int i = 1; i <= endColumn; i++)
                    {
                        if (dataSheet.Cells[1, i].Value.ToString().Trim() == "库位号")
                        {
                            areaColumn = i;
                            break;
                        }
                    }
                }
                for (int idx = endRow; idx >= 2; idx--)
                {
                    var areaObj = dataSheet.Cells[idx, areaColumn].Value;
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
                    var s = filterAreaStr.ToString().Replace("，", ",").Replace(" ", string.Empty).ToUpper();
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

            ViewBag.DowloadFileName = orderFileName;
            return PartialView("_MetadataDowload");
        }


        [HttpPost]
        public async Task<PartialViewResult> ExtractMixtureAreaHandle()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var orderFileName = Guid.NewGuid().ToString() + ".xlsx";
            var orderFilePath = Path.Combine(tmpFolder, orderFileName);
            var files = Request.Form.Files;
            var filterAreaStr = Request.Form["area"].ToString().Replace("，", ",").Replace(" ", string.Empty).ToUpper();

            if (files.Count > 0)
            {
                using (var targetStream = System.IO.File.Create(orderFilePath))
                    await files[0].CopyToAsync(targetStream);
            }


            using (var package = new ExcelPackage(new FileInfo(orderFilePath)))
            {
                var filterAreas = filterAreaStr.Split(',', StringSplitOptions.RemoveEmptyEntries).ToList();
                var filterAreaCount = filterAreas.Count();
                var mapping = new List<int>();
                var dataSheet = package.Workbook.Worksheets[0];

                var endRow = dataSheet.Dimension.End.Row;
                var endColumn = dataSheet.Dimension.End.Column;
                var areaFlagColumn = endColumn + 1;
                //默认的话,"库位号"在第8列,但是也有可能改变
                var areaColumn = 8;
                if (dataSheet.Cells[1, 8].Value == null || dataSheet.Cells[1, 8].Value.ToString().Trim() != "库位号")
                {
                    for (int i = 1; i <= endColumn; i++)
                    {
                        if (dataSheet.Cells[1, i].Value.ToString().Trim() == "库位号")
                        {
                            areaColumn = i;
                            break;
                        }
                    }
                }
                for (int idx = endRow; idx >= 2; idx--)
                {
                    var areaObj = dataSheet.Cells[idx, areaColumn].Value;
                    if (areaObj != null)
                    {
                        var areaStr = areaObj.ToString();
                        var areas = areaStr.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Substring(0, 1).ToUpper()).Distinct().ToArray();

                        if (areas.Count() != filterAreaCount) continue;

                        var satisfied = true;
                        for (var k = areas.Count() - 1; k >= 0; k--)
                        {
                            if (!filterAreas.Contains(areas[k]))
                            {
                                satisfied = false;
                                break;
                            }
                        }

                        if (satisfied)
                            mapping.Add(idx);

                    }
                }


                var sheet = package.Workbook.Worksheets.Add(filterAreaStr + "区域");
                //拷贝标题行
                dataSheet.Cells[1, 1, 1, endColumn].Copy(sheet.Cells[1, 1, 1, endColumn]);
                //数据行
                for (int i = 0, len = mapping.Count; i < len; i++)
                {
                    var r = mapping[i];
                    dataSheet.Cells[r, 1, r, endColumn].Copy(sheet.Cells[i + 2, 1, i + 2, endColumn]);
                }

                //删除原数据表
                package.Workbook.Worksheets.Delete(0);
                package.Save();
            }

            ViewBag.DowloadFileName = orderFileName;
            return PartialView("_MetadataDowload");
        }

    }

}