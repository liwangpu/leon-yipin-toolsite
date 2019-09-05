using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ToolSite.Models.Salary;

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

        /// <summary>
        /// 仓库加班考勤数据处理
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public async Task<PartialViewResult> WarehouseOvertimeHandle()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var resultFileName = Guid.NewGuid().ToString() + ".xlsx";
            var signFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var resultFilePath = Path.Combine(tmpFolder, resultFileName);
            var files = Request.Form.Files;
            var monthStr = Request.Form["month"];
            if (files.Count > 0)
            {
                using (var targetStream = System.IO.File.Create(signFilePath))
                    await files[0].CopyToAsync(targetStream);
            }

            var list考勤数据 = new List<_仓库加班考勤数据>();
            var list加班绩效 = new List<_仓库加班绩效>();

            #region 读取数据
            using (var package = new ExcelPackage(new FileInfo(signFilePath)))
            {
                var worksheet = package.Workbook.Worksheets["打卡时间"];
                var endRow = worksheet.Dimension.End.Row;
                var endColumn = worksheet.Dimension.End.Column;
                for (int idx = 4; idx <= endRow; idx++)
                {
                    var md = new _仓库加班考勤数据();
                    md._姓名 = worksheet.Cells[idx, 1].Value.ToString();
                    md._员工序号 = Convert.ToInt32(worksheet.Cells[idx, 3].Value);
                    var list = new List<string>();
                    for (int cll = 6; cll <= endColumn; cll++)
                    {
                        var vl = worksheet.Cells[idx, cll].Value != null ? worksheet.Cells[idx, cll].Value.ToString().Trim().Replace("\r\n", "").Replace(" ", "") : "";
                        list.Add(vl);
                    }
                    md._加班信息 = list;
                    list考勤数据.Add(md);
                }
            }
            System.IO.File.Delete(signFilePath);
            #endregion

            #region 处理数据
            {
                var _d包饭时间 = Convert.ToDateTime("2018-08-08 21:00:00");
                for (int idx = list考勤数据.Count - 1; idx >= 0; idx--)
                {
                    var cur = list考勤数据[idx];
                    var md = new _仓库加班绩效();
                    md._员工序号 = cur._员工序号;
                    md._姓名 = cur._姓名;

                    //if (md._姓名 == "曹雷")
                    //{

                    //}
                    var _list加班时长 = new List<double>();
                    var _list出勤时长 = new List<double>();
                    var _list打卡异常 = new List<bool>();
                    var _list原始打卡时间 = new List<string>();
                    for (int nnn = 0, count = cur._加班信息.Count; nnn < count; nnn++)
                    {
                        var timeStr = !string.IsNullOrWhiteSpace(cur._加班信息[nnn]) ? cur._加班信息[nnn].Trim() : "";
                        var errFlag = false;

                        _list原始打卡时间.Add(timeStr);
                        //一天打卡一次或没有打卡
                        if (string.IsNullOrWhiteSpace(timeStr) || timeStr.Length <= 5)
                        {
                            //只打了一次卡,标记异常
                            if (!string.IsNullOrWhiteSpace(timeStr))
                            {
                                errFlag = true;
                            }
                            _list出勤时长.Add(0);
                            _list加班时长.Add(0);
                        }
                        else
                        {
                            double _加班时长 = 0;
                            var d上班时间 = Convert.ToDateTime(string.Format("2018-08-08 {0}:00", timeStr.Substring(0, 5)));
                            var d下班时间 = Convert.ToDateTime(string.Format("2018-08-08 {0}:00", timeStr.Substring(timeStr.Length - 5, 5)));

                            var timespan = (d下班时间 - d上班时间).TotalMinutes;
                            var remain = timespan % 30;
                            var halfHours = (timespan - remain) / 30;
                            //误差六分钟
                            if (remain >= 24)
                                halfHours += 1;

                            var hours = halfHours / 2;
                            if (hours >= 8.5)
                            {
                                _加班时长 = hours - 8.5;
                                //超过饭点,减去半个小时吃饭时间
                                if (d下班时间 >= _d包饭时间)
                                    _加班时长 -= 0.5;

                                _list出勤时长.Add(8.5);
                            }
                            else
                            {
                                _list出勤时长.Add(hours);
                            }

                            _list加班时长.Add(_加班时长);
                        }
                        _list打卡异常.Add(errFlag);

                    }
                    md._加班时长 = _list加班时长;
                    md._打卡出现异常 = _list打卡异常;
                    md._出勤时长 = _list出勤时长;
                    md._原始打卡时间 = _list原始打卡时间;
                    list加班绩效.Add(md);

                }
            }
            #endregion

            #region 生成表格
            {
                #region 订单分配
                using (ExcelPackage package = new ExcelPackage(new FileInfo(resultFilePath)))
                {
                    var workbox = package.Workbook;
                    var sheet1 = workbox.Worksheets.Add("Sheet1");

                    #region 标题行
                    using (var rng = sheet1.Cells[1, 1, 3, 1])
                    {
                        rng.Value = "姓名";
                        rng.Merge = true;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }
                    sheet1.Cells[1, 2].Value = "星期";
                    using (var rng = sheet1.Cells[2, 2, 3, 2])
                    {
                        rng.Value = "项目";
                        rng.Merge = true;
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                    }

                    if (list加班绩效[0] != null)
                    {
                        string[] Day = new string[] { "周日", "周一", "周二", "周三", "周四", "周五", "周六" };
                        var days = list加班绩效[0]._加班时长.Count + 3;
                        for (int column = 3, idx = 1; column < days; column++, idx++)
                        {
                            var ct = DateTime.Now;
                            var month = Convert.ToInt32(monthStr);
                            var dateStr = string.Format("{0}-{1}-{2}", ct.Year, month > 9 ? "" + month : "0" + month, idx > 9 ? "" + idx : "0" + idx);
                            var date = DateTime.MinValue;
                            var isValid = DateTime.TryParse(dateStr, out date);
                            if (isValid)
                            {
                                sheet1.Column(column).Width = 5;//设置列宽
                                using (var rng = sheet1.Cells[2, column, 3, column])
                                {
                                    rng.Value = idx;
                                    rng.Merge = true;
                                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                                }
                                //背景色标记周末
                                var ddd = Day[Convert.ToInt32(Convert.ToDateTime(dateStr).DayOfWeek.ToString("d"))].ToString();
                                if (ddd == "周日" || ddd == "周六")
                                {
                                    sheet1.Column(column).Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    sheet1.Column(column).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(14277081));
                                }
                                sheet1.Cells[1, column].Value = ddd;
                            }
                        }

                        using (var rng = sheet1.Cells[1, days, 2, days + 2])
                        {
                            rng.Value = "合计加班";
                            rng.Merge = true;
                            sheet1.Column(days).Width = 5;//设置列宽
                            sheet1.Column(days + 1).Width = 5;//设置列宽
                            sheet1.Column(days + 2).Width = 5;//设置列宽
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }

                        using (var rng = sheet1.Cells[3, days])
                        {
                            rng.Value = "平时（H|日）";
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }

                        using (var rng = sheet1.Cells[3, days + 1])
                        {
                            rng.Value = "周末（日）";
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }

                        using (var rng = sheet1.Cells[3, days + 2])
                        {
                            rng.Value = "节日（日）";
                            rng.Style.WrapText = true;//自动换行
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                    }
                    sheet1.Row(3).Height = 79;//设置行高
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 4, len = list加班绩效.Count; idx < len; idx++)
                    {
                        var curOrder = list加班绩效[idx];
                        using (var rng = sheet1.Cells[rowIdx, 1, rowIdx + 2, 1])
                        {
                            rng.Value = curOrder._姓名;
                            rng.Merge = true;
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        }
                        sheet1.Cells[rowIdx, 2].Value = "出勤";
                        sheet1.Cells[rowIdx + 1, 2].Value = "请假";
                        sheet1.Cells[rowIdx + 2, 2].Value = "加班";
                        sheet1.Cells[rowIdx + 3, 2].Value = "打卡情况";
                        var _i请假总计 = 0;
                        double _i出勤合计 = 0;
                        for (int nnn = 0, nlen = curOrder._加班时长.Count; nnn < nlen; nnn++)
                        {
                            //比如这一天是第31号,但是这个月没有
                            if (sheet1.Cells[1, 3 + nnn].Value == null)
                                continue;


                            _i出勤合计 += curOrder._出勤时长[nnn];
                            sheet1.Cells[rowIdx, 3 + nnn].Value = curOrder._出勤时长[nnn];

                            //首先请假默认写一个0
                            if (sheet1.Cells[1, 3 + nnn].Value != null && sheet1.Cells[1, 3 + nnn].Value.ToString().IndexOf("周") > -1)
                            {
                                sheet1.Cells[rowIdx + 1, 3 + nnn].Value = 0;
                            }
                            //判断是否请假
                            if (curOrder._出勤时长[nnn] == 0 && curOrder._打卡出现异常[nnn] != true && sheet1.Cells[1, 3 + nnn].Value != null && sheet1.Cells[1, 3 + nnn].Value.ToString() != "周日")
                            {
                                _i请假总计++;
                                sheet1.Cells[rowIdx + 1, 3 + nnn].Value = 1;

                            }

                            if (curOrder._打卡出现异常[nnn])
                            {
                                using (var rng = sheet1.Cells[rowIdx, 3 + nnn, rowIdx + 3, 3 + nnn])
                                {
                                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    rng.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                }
                                sheet1.Cells[rowIdx + 3, 3 + nnn].Value = curOrder._原始打卡时间[nnn];
                            }
                            sheet1.Cells[rowIdx + 2, 3 + nnn].Value = curOrder._加班时长[nnn];
                        }

                        sheet1.Cells[rowIdx, 3 + curOrder._加班时长.Count].Value = _i出勤合计;

                        //请假合计因为单位是天,字体颜色另类一些
                        using (var rng = sheet1.Cells[rowIdx + 1, 3 + curOrder._加班时长.Count])
                        {
                            rng.Value = _i请假总计;
                            rng.Style.Font.Color.SetColor(Color.FromArgb(15773696));//字体颜色

                        }

                        sheet1.Cells[rowIdx + 2, 3 + curOrder._加班时长.Count].Value = curOrder._加班时长.Sum();
                        rowIdx += 4;
                    }
                    #endregion

                    #region 全部边框
                    {
                        var endRow = sheet1.Dimension.End.Row;
                        var endColumn = sheet1.Dimension.End.Column;
                        using (var rng = sheet1.Cells[1, 1, endRow, endColumn])
                        {
                            rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                    }
                    #endregion

                    package.Save();
                }
                #endregion
            }
            #endregion

            ViewBag.DowloadFileName = resultFileName;
            return PartialView("_MetadataDowload");
        }


        public ActionResult PickingPerf()
        {
            return View();
        }

        public async Task<IActionResult> MonthlyWorkingHours()
        {



            return Ok();
        }


    }
}