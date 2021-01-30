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
    public class LaboratoryController : Controller
    {

        private readonly IHostingEnvironment env;

        public LaboratoryController(IHostingEnvironment env)
        {
            this.env = env;
        }

        public IActionResult InflorescenceSplit()
        {
            return View();
        }

        [HttpPost]
        public async Task<PartialViewResult> SplitInflorescenceData()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var fileName = Guid.NewGuid().ToString() + ".xlsx";
            var filePath = Path.Combine(tmpFolder, fileName);
            var charaterStr = Request.Form["charater"].ToString().Trim();
            var files = Request.Form.Files;

            if (files.Count > 0)
            {
                using (var targetStream = System.IO.File.Create(filePath))
                    await files[0].CopyToAsync(targetStream);
            }


            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var dataSheet = package.Workbook.Worksheets[0];

                var endRow = dataSheet.Dimension.End.Row;
                var endColumn = dataSheet.Dimension.End.Column;

                var inflorescenceColumns = new List<int>();
                // 查找"主花序角果数列"
                for (var i = 1; i <= endColumn; i++)
                {
                    if (dataSheet.Cells[1, i].Value != null && dataSheet.Cells[1, i].Value.ToString().Trim() == charaterStr)
                    {
                        inflorescenceColumns.Add(i);
                    }
                }

                // 获取实验角果编号
                var treeNumbers = new List<string>();
                for (var cdx = inflorescenceColumns.Count - 1; cdx >= 0; cdx--)
                {
                    var treeNumberColumn = inflorescenceColumns[cdx] - 1;
                    for (var i = 1; i <= endRow; i++)
                    {
                        if (dataSheet.Cells[i, treeNumberColumn].Value != null)
                        {
                            var name = dataSheet.Cells[i, treeNumberColumn].Value.ToString().Trim();
                            if (!treeNumbers.Contains(name))
                            {
                                treeNumbers.Add(name);
                            }
                        }
                    }
                }

                var result = new List<InflorescenceSplitData>();
                for (var i = 0; i < inflorescenceColumns.Count; i++)
                {
                    for (var row = 2; row < endRow; row++)
                    {
                        var treeNumberColumn = inflorescenceColumns[i] - 1;
                        var treeDataColumn = inflorescenceColumns[i];
                        var name = dataSheet.Cells[row, treeNumberColumn].Value != null ? dataSheet.Cells[row, treeNumberColumn].Value.ToString().Trim() : null;
                        if (name != null)
                        {
                            InflorescenceSplitData item;
                            var itemCreate = false;
                            if (result.Count(d => d.name == name) == 0)
                            {
                                item = new InflorescenceSplitData();
                                item.name = name;
                                itemCreate = true;
                            }
                            else
                            {
                                item = result.First(d => d.name == name);
                            }

                            var datas = item.datas.ElementAtOrDefault(i) == null ? new List<decimal>() : item.datas[i];
                            var countStr = dataSheet.Cells[row, treeDataColumn].Value != null ? dataSheet.Cells[row, treeDataColumn].Value.ToString().Trim() : null;
                            decimal count = 0;
                            decimal.TryParse(countStr, out count);
                            datas.Add(count);

                            if (item.datas.ElementAtOrDefault(i) == null)
                            {
                                item.datas.Add(datas);
                            }

                            if (itemCreate)
                            {
                                result.Add(item);
                            }
                        }
                    }
                }

                var sheet = package.Workbook.Worksheets.Add("处理结果");
                var r = 1;
                foreach (var item in result)
                {
                    foreach (var datas in item.datas)
                    {
                        sheet.Cells[r + 1, 1].Value = item.name;
                        var cell = 2;
                        foreach (var count in datas)
                        {
                            sheet.Cells[r + 1, cell].Value = count;
                            cell++;
                        }
                        r++;
                    }
                }
                package.Save();
            }

            ViewBag.DowloadFileName = fileName;
            return PartialView("_MetadataDowload");
        }

    }

    class InflorescenceSplitData
    {
        public string name { get; set; }
        public List<List<decimal>> datas = new List<List<decimal>>();

    }
}
