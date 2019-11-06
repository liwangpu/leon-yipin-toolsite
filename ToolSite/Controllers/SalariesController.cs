using EpplusHelper;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
        private const string PickingPerfMonthlyWorkingHoursCacheFolder = "配货绩效_月上班时间";
        private const string HistoryPickingPerfMonthlyHoursCacheFolder = "历史配货绩效";

        #region ctor
        public SalariesController(IHostingEnvironment env)
        {
            this.env = env;
        }
        #endregion

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

        /// <summary>
        /// 配货绩效
        /// </summary>
        /// <returns></returns>
        public ActionResult PickingPerf()
        {
            return View();
        }

        /// <summary>
        /// 配货绩效-月上班时间数据处理
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public async Task<PartialViewResult> MonthlyWorkingHoursHandle()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var workingHoursCacheFolder = Path.Combine(env.WebRootPath, "cache", PickingPerfMonthlyWorkingHoursCacheFolder);
            if (!Directory.Exists(workingHoursCacheFolder)) Directory.CreateDirectory(workingHoursCacheFolder);
            var workHoursFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var files = Request.Form.Files;
            var monthStr = Request.Form["month"].ToString();
            var workHours = new List<_配货绩效_员工月上班时间>();
            if (files.Count > 0)
            {
                var monthlyWorkingHoursFile = files.FirstOrDefault(x => x.Name == "monthlyWorkingHoursFile");
                if (monthlyWorkingHoursFile != null)
                {
                    using (var targetStream = System.IO.File.Create(workHoursFilePath))
                        await monthlyWorkingHoursFile.CopyToAsync(targetStream);
                    using (var package = new ExcelPackage(new FileInfo(workHoursFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        workHours = SheetReader<_配货绩效_员工月上班时间>.From(worksheet);
                        for (int idx = workHours.Count - 1; idx >= 0; idx--)
                        {
                            if (string.IsNullOrWhiteSpace(workHours[idx]._姓名))
                            {
                                workHours.RemoveAt(idx);
                                continue;
                            }
                            workHours[idx].GenerateWorkingTime();
                        }
                        var json = JsonConvert.SerializeObject(workHours);
                        var workHoursCacheFilePath = Path.Combine(workingHoursCacheFolder, monthStr + ".json");
                        using (var fs = new StreamWriter(workHoursCacheFilePath, false, Encoding.UTF8))
                            fs.Write(json);
                    }
                    System.IO.File.Delete(workHoursFilePath);
                }
            }
            ViewBag.DowloadFileName = "";
            return PartialView("_MetadataDowload");
        }

        /// <summary>
        /// 配货绩效-每日绩效计算
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public async Task<PartialViewResult> DailyWorkingHoursHandler()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var workingHoursCacheFolder = Path.Combine(env.WebRootPath, "cache", PickingPerfMonthlyWorkingHoursCacheFolder);
            var historyPickingPerfCacheFolder = Path.Combine(env.WebRootPath, "cache", HistoryPickingPerfMonthlyHoursCacheFolder);//历史配货绩效存储
            if (!Directory.Exists(historyPickingPerfCacheFolder)) Directory.CreateDirectory(historyPickingPerfCacheFolder);
            var pickingFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var randomFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var flowFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var areaRepFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var helpHoursFilePath = Path.Combine(tmpFolder, Guid.NewGuid().ToString() + ".xlsx");
            var paperAmount = Convert.ToDouble(Request.Form["paperAmount"]);//张数定值
            var paperRate = Convert.ToDouble(Request.Form["paperRate"]);//张数占比
            var pickingAmount = Convert.ToDouble(Request.Form["pickingAmount"]);//数量定值
            var pickingRate = Convert.ToDouble(Request.Form["pickingRate"]);//数量占比
            var perfDate = Convert.ToDateTime(Request.Form["pickingDate"]);
            var resultFileName = $"{perfDate.ToString("yyyy-MM-dd")}[当天配货绩效].xlsx";
            var dailyPerfFilePath = Path.Combine(tmpFolder, resultFileName);
            var list拣货单 = new List<_配货绩效_拣货单>();
            var list乱单 = new List<_配货绩效_乱单>();
            var list本楼层乱单原始数据 = new List<_配货绩效_本楼层乱单>();
            var list人员负责库位信息 = new List<_配货绩效_拣货人员配置信息>();
            var list最终绩效 = new List<_配货绩效_配货绩效结果>();
            var list本月上班时间 = new List<_配货绩效_员工月上班时间>();
            var list当天帮忙时间 = new List<_配货绩效_帮忙点货时间>();

            var files = Request.Form.Files;

            #region 读取表格信息
            if (files.Count > 0)
            {
                var pickingFile = files.FirstOrDefault(x => x.Name == "pickingFile");
                if (pickingFile != null)
                {
                    using (var targetStream = System.IO.File.Create(pickingFilePath))
                        await pickingFile.CopyToAsync(targetStream);
                    using (var package = new ExcelPackage(new FileInfo(pickingFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        list拣货单 = SheetReader<_配货绩效_拣货单>.From(worksheet);
                    }
                    System.IO.File.Delete(pickingFilePath);
                }

                var randomFile = files.FirstOrDefault(x => x.Name == "randomFile");
                if (randomFile != null)
                {
                    using (var targetStream = System.IO.File.Create(randomFilePath))
                        await randomFile.CopyToAsync(targetStream);
                    using (var package = new ExcelPackage(new FileInfo(randomFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        list乱单 = SheetReader<_配货绩效_乱单>.From(worksheet);
                        //将乱单转换正常拣货单
                        foreach (var item乱单 in list乱单)
                        {
                            if (item乱单 == null || string.IsNullOrWhiteSpace(item乱单._商品明细) || string.IsNullOrWhiteSpace(item乱单._完整库位号))
                                continue;
                            list拣货单.AddRange(item乱单.ToData());
                        }
                    }
                    System.IO.File.Delete(randomFilePath);
                }

                //本楼层格式就是乱单,不一定有
                var flowFile = files.FirstOrDefault(x => x.Name == "flowFile");
                if (flowFile != null)
                {
                    using (var targetStream = System.IO.File.Create(flowFilePath))
                        await flowFile.CopyToAsync(targetStream);
                    using (var package = new ExcelPackage(new FileInfo(flowFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        list本楼层乱单原始数据 = SheetReader<_配货绩效_本楼层乱单>.From(worksheet);
                        ////将乱单转换正常拣货单
                        //foreach (var item乱单 in list本楼层乱单)
                        //{
                        //    list拣货单.AddRange(item乱单.ToData());
                        //}
                    }
                    System.IO.File.Delete(flowFilePath);
                }


                var areaRepFile = files.FirstOrDefault(x => x.Name == "areaRepFile");
                if (areaRepFile != null)
                {
                    using (var targetStream = System.IO.File.Create(areaRepFilePath))
                        await areaRepFile.CopyToAsync(targetStream);
                    using (var package = new ExcelPackage(new FileInfo(areaRepFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        list人员负责库位信息 = SheetReader<_配货绩效_拣货人员配置信息>.From(worksheet);
                    }
                    System.IO.File.Delete(areaRepFilePath);
                }

                var helpingHoursFile = files.FirstOrDefault(x => x.Name == "helpingHoursFile");
                if (helpingHoursFile != null)
                {
                    using (var targetStream = System.IO.File.Create(helpHoursFilePath))
                        await helpingHoursFile.CopyToAsync(targetStream);
                    using (var package = new ExcelPackage(new FileInfo(helpHoursFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        list当天帮忙时间 = SheetReader<_配货绩效_帮忙点货时间>.From(worksheet);
                    }
                    System.IO.File.Delete(helpHoursFilePath);
                }
            }
            #endregion

            #region 加载月缓存上班时间信息
            {
                var workHoursCacheFilePath = Path.Combine(workingHoursCacheFolder, perfDate.Month + ".json");
                if (System.IO.File.Exists(workHoursCacheFilePath))
                {
                    using (var fs = new StreamReader(workHoursCacheFilePath, Encoding.UTF8))
                    {
                        var json = fs.ReadToEnd();
                        list本月上班时间 = JsonConvert.DeserializeObject<List<_配货绩效_员工月上班时间>>(json);
                    }
                }
            }
            #endregion

            #region 处理数据
            if (list拣货单.Count > 0)
            {
                for (int idx = list拣货单.Count - 1; idx >= 0; idx--)
                {
                    if (string.IsNullOrWhiteSpace(list拣货单[idx]._库位号))
                        list拣货单.RemoveAt(idx);
                }



                var allEmpNames = list人员负责库位信息.Select(x => x._姓名).Distinct().ToList();
                allEmpNames.ForEach(name =>
                {

                    if (!string.IsNullOrEmpty(name))
                    {

                        //if (name.Trim() == "欧于书")
                        //{

                        //}
                        var md = new _配货绩效_配货绩效结果();
                        md._d张数占比 = paperRate;
                        md._d张数定值 = paperAmount;
                        md._d数量定值 = pickingAmount;
                        md._d数量占比 = pickingRate;


                        md._业绩归属人 = name;
                        var _订单详情数据 = new List<_配货绩效_订单详情数据>();

                        #region 抽取详细信息
                        {
                            var query = from it in list拣货单
                                        join s in list人员负责库位信息 on it._库位号 equals s._管理库位 into joined
                                        from j in joined.DefaultIfEmpty()
                                        where j != null && j._姓名 == name
                                        select it;
                            var refLh = query.ToList();


                            //从本楼层原始数据抽出本楼层
                            //因为本楼层是帮忙的数据,所以不按照拣货单的区域算绩效,用原始数据里面的配货人
                            {
                                var ds = list本楼层乱单原始数据.Where(x => x._拣货人 == name).ToList();
                                ds.ForEach(it =>
                                {
                                    refLh.AddRange(it.ToData());
                                });

                            }

                            foreach (var deitem in refLh)
                            {
                                var item = deitem._拣货明细;
                                foreach (var it in item)
                                {
                                    var arr = it.Replace(".", string.Empty).Split(new string[] { "*" }, StringSplitOptions.RemoveEmptyEntries);
                                    if (arr.Length >= 2)
                                    {
                                        var detail = new _配货绩效_订单详情数据();
                                        detail.SKU = arr[0].Trim();
                                        detail.Amount = Convert.ToDouble(arr[1]);
                                        detail._乱单 = deitem._乱单;
                                        detail._本楼层 = deitem._本楼层;
                                        _订单详情数据.Add(detail);
                                    }

                                }
                            }

                        }
                        #endregion

                        var list订单详情数据_拣货单 = _订单详情数据.Where(x => x._乱单 == false).ToList();
                        var list订单详情数据_乱单 = _订单详情数据.Where(x => x._乱单 == true).ToList();
                        var list订单详情数据_本楼层乱单 = _订单详情数据.Where(x => x._本楼层 == true).ToList();

                        var str_帮忙总时长 = "";
                        decimal refTime = 0;

                        #region 计算上班时间和帮忙时间
                        {
                            var d绩效日期 = perfDate.Date;
                            if (list本月上班时间 != null && list本月上班时间.Count > 0)
                            {
                                var refer工作时间 = list本月上班时间.Where(x => x._姓名 == name).FirstOrDefault();
                                var refer帮忙时间 = list当天帮忙时间.Where(x => x._姓名 == name).FirstOrDefault();
                                if (refer工作时间 != null)
                                {
                                    decimal d上班时间 = 0;
                                    decimal d帮忙时间 = 0;
                                    if (refer帮忙时间 != null && refer帮忙时间._帮忙总时间 != null)
                                    {
                                        var h = (refer帮忙时间._帮忙总时间).Hours;
                                        var mh = Math.Round((refer帮忙时间._帮忙总时间).Minutes / 60m, 1);
                                        d帮忙时间 = h + mh;
                                    }
                                    str_帮忙总时长 = d帮忙时间.ToString();
                                    d上班时间 = refer工作时间._工作时间[d绩效日期.Day - 1];
                                    if (d上班时间 > 0)
                                        refTime = d上班时间 - d帮忙时间;
                                }
                            }
                        }
                        #endregion


                        md._拣货单张数_正常 = list订单详情数据_拣货单.Select(x => x.SKU).Distinct().Count();
                        md._购买总数量_正常 = list订单详情数据_拣货单.Select(x => x.Amount).Sum();
                        md._拣货单张数_乱单 = list订单详情数据_乱单.Select(x => x.SKU).Distinct().Count();
                        md._购买总数量_乱单 = list订单详情数据_乱单.Select(x => x.Amount).Sum();
                        md._拣货单张数_本楼层乱单 = list订单详情数据_本楼层乱单.Select(x => x.SKU).Distinct().Count();
                        md._购买总数量_本楼层乱单 = list订单详情数据_本楼层乱单.Select(x => x.Amount).Sum();

                        md._总时长 = refTime.ToString();
                        md._帮忙总时长 = str_帮忙总时长;
                        md._分钟 = Convert.ToDouble(refTime * 60);


                        list最终绩效.Add(md);
                    }
                });

                if (list最终绩效.Count > 0)
                {
                    #region 生成绩效表格
                    GenerateDailyPerfExcelTable(dailyPerfFilePath, list最终绩效);
                    #endregion

                    #region 缓存数据
                    {
                        var historyPickingPerfCacheFilePath = Path.Combine(historyPickingPerfCacheFolder, $"{perfDate.ToString("yyyy-MM")}.json");

                        _配货绩效_全月绩效结果 historyPickingPerf;

                        if (System.IO.File.Exists(historyPickingPerfCacheFilePath))
                            using (var fs = new StreamReader(historyPickingPerfCacheFilePath, Encoding.UTF8))
                            {
                                var json = fs.ReadToEnd();
                                historyPickingPerf = JsonConvert.DeserializeObject<_配货绩效_全月绩效结果>(json);
                            }
                        else
                            historyPickingPerf = new _配货绩效_全月绩效结果();


                        using (var fs = new StreamWriter(historyPickingPerfCacheFilePath, false, Encoding.UTF8))
                        {
                            historyPickingPerf.Perf[perfDate.Day] = list最终绩效;
                            fs.Write(JsonConvert.SerializeObject(historyPickingPerf));
                        }
                    }
                    #endregion
                }
            }
            #endregion

            ViewBag.DowloadFileName = resultFileName;
            return PartialView("_MetadataDowload");
        }

        /// <summary>
        /// 下载指定日期的绩效
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public string DownloadSpecifyDatePerf()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var historyPickingPerfCacheFolder = Path.Combine(env.WebRootPath, "cache", HistoryPickingPerfMonthlyHoursCacheFolder);//历史配货绩效存储
            var dowloadPickingDate = Convert.ToDateTime(Request.Form["dowloadPickingDate"]);
            var resultFileName = $"{dowloadPickingDate.ToString("yyyy-MM-dd")}[当天配货绩效].xlsx";
            var resultFilePath = Path.Combine(tmpFolder, resultFileName);
            var historyPickingPerfCacheFilePath = Path.Combine(historyPickingPerfCacheFolder, $"{dowloadPickingDate.ToString("yyyy-MM")}.json");
            var historyPickingPerf = new _配货绩效_全月绩效结果();

            #region 加载缓存结果
            if (System.IO.File.Exists(historyPickingPerfCacheFilePath))
            {
                using (var fs = new StreamReader(historyPickingPerfCacheFilePath, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    historyPickingPerf = JsonConvert.DeserializeObject<_配货绩效_全月绩效结果>(json);
                }
            }
            #endregion

            if (historyPickingPerf.Perf.ContainsKey(dowloadPickingDate.Day))
            {
                GenerateDailyPerfExcelTable(resultFilePath, historyPickingPerf.Perf[dowloadPickingDate.Day]);
                return resultFileName;
            }

            return string.Empty;
        }

        /// <summary>
        /// 下载指定日期的全月绩效
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public string DownloadMonthPerf()
        {
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var historyPickingPerfCacheFolder = Path.Combine(env.WebRootPath, "cache", HistoryPickingPerfMonthlyHoursCacheFolder);//历史配货绩效存储
            var dowloadPickingDate = Convert.ToDateTime(Request.Form["dowloadPickingDate"]);
            var resultFileName = $"{dowloadPickingDate.ToString("yyyy-MM")}[全月配货绩效].xlsx";
            var resultFilePath = Path.Combine(tmpFolder, resultFileName);
            var historyPickingPerfCacheFilePath = Path.Combine(historyPickingPerfCacheFolder, $"{dowloadPickingDate.ToString("yyyy-MM")}.json");
            var historyPickingPerf = new _配货绩效_全月绩效结果();

            #region 加载缓存结果
            if (System.IO.File.Exists(historyPickingPerfCacheFilePath))
            {
                using (var fs = new StreamReader(historyPickingPerfCacheFilePath, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    historyPickingPerf = JsonConvert.DeserializeObject<_配货绩效_全月绩效结果>(json);
                }
            }
            #endregion

            var source = new List<_配货绩效_配货绩效结果>();
            foreach (var item in historyPickingPerf.Perf)
            {
                for (int i = 0, len = item.Value.Count; i < len; i++)
                {
                    var it = item.Value[i];
                    it._绩效日期 = item.Key;
                    source.Add(it);
                }
            }

            var datas = new List<_配货绩效_配货绩效结果>();
            var expNames = source.Select(x => x._业绩归属人).Distinct().ToList();
            expNames.ForEach(n =>
            {
                if (!string.IsNullOrWhiteSpace(n))
                {
                    var md = new _配货绩效_配货绩效结果();
                    md._业绩归属人 = n;

                    var defaultItem = source[0];
                    md._d张数占比 = defaultItem._d张数占比;
                    md._d张数定值 = defaultItem._d张数定值;
                    md._d数量占比 = defaultItem._d数量占比;
                    md._d数量定值 = defaultItem._d数量定值;


                    md._购买总数量_正常 = source.Where(x => x._业绩归属人 == n).Select(x => x._购买总数量_正常).Sum();
                    md._拣货单张数_正常 = source.Where(x => x._业绩归属人 == n).Select(x => x._拣货单张数_正常).Sum();
                    md._购买总数量_乱单 = source.Where(x => x._业绩归属人 == n).Select(x => x._购买总数量_乱单).Sum();
                    md._拣货单张数_乱单 = source.Where(x => x._业绩归属人 == n).Select(x => x._拣货单张数_乱单).Sum();

                    md._分钟 = source.Where(x => x._业绩归属人 == n).Select(x => x._分钟).Sum();
                    md._总时长 = source.Where(x => x._业绩归属人 == n).Select(x => x._小时).Sum().ToString();
                    md._帮忙总时长 = source.Where(x => x._业绩归属人 == n).Select(x => string.IsNullOrWhiteSpace(x._帮忙总时长) ? 0 : decimal.Parse(x._帮忙总时长)).Sum().ToString();
                    datas.Add(md);
                }
            });

            if (datas.Count > 0)
            {
                GenerateMonthlyPerfExcelTable(resultFilePath, datas, source.OrderBy(x => x._绩效日期).ToList(), historyPickingPerf);
                return resultFileName;
            }

            return string.Empty;
        }

        /// <summary>
        /// 查看某个月份的所有绩效信息
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IEnumerable<string> WatchWholeMonthPerfMessage()
        {
            var perfs = new List<string>();
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            var historyPickingPerfCacheFolder = Path.Combine(env.WebRootPath, "cache", HistoryPickingPerfMonthlyHoursCacheFolder);//历史配货绩效存储
            var dowloadPickingDate = Convert.ToDateTime(Request.Form["dowloadPickingDate"]);
            var resultFileName = $"{dowloadPickingDate.ToString("yyyy-MM")}[全月配货绩效].xlsx";
            var resultFilePath = Path.Combine(tmpFolder, resultFileName);
            var historyPickingPerfCacheFilePath = Path.Combine(historyPickingPerfCacheFolder, $"{dowloadPickingDate.ToString("yyyy-MM")}.json");
            var historyPickingPerf = new _配货绩效_全月绩效结果();

            #region 加载缓存结果
            if (System.IO.File.Exists(historyPickingPerfCacheFilePath))
            {
                using (var fs = new StreamReader(historyPickingPerfCacheFilePath, Encoding.UTF8))
                {
                    var json = fs.ReadToEnd();
                    historyPickingPerf = JsonConvert.DeserializeObject<_配货绩效_全月绩效结果>(json);
                }
            }
            else
                return new List<string>();
            #endregion

            var prefx = dowloadPickingDate.ToString("yyyy-MM");
            return historyPickingPerf.Perf.Keys.ToList().OrderBy(x => x).Select(x => prefx + "-" + x.ToString().PadLeft(2, '0'));
        }

        /// <summary>
        /// 根据绩效结果生成表格
        /// </summary>
        /// <param name="path"></param>
        /// <param name="datas"></param>
        private void GenerateDailyPerfExcelTable(string path, List<_配货绩效_配货绩效结果> datas)
        {
            if (System.IO.File.Exists(path)) System.IO.File.Delete(path);

            #region 生成绩效表格
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var workbox = package.Workbook;
                var sheet1 = workbox.Worksheets.Add("Sheet1");


                #region 标题行
                sheet1.Cells[1, 1].Value = "姓名";
                sheet1.Cells[1, 2].Value = "拣货单数量";
                sheet1.Cells[1, 3].Value = "乱单数量";
                sheet1.Cells[1, 4].Value = "本楼层数量";
                sheet1.Cells[1, 5].Value = "总数量";
                sheet1.Cells[1, 6].Value = "拣货单张数";
                sheet1.Cells[1, 7].Value = "乱单张数";
                sheet1.Cells[1, 8].Value = "本楼层张数";
                sheet1.Cells[1, 9].Value = "总张数";
                sheet1.Cells[1, 10].Value = "帮忙总时长";
                sheet1.Cells[1, 11].Value = "工作总时长";
                sheet1.Cells[1, 12].Value = "分钟";
                sheet1.Cells[1, 13].Value = "拣货单效率";
                sheet1.Cells[1, 14].Value = "购买数量效率";
                sheet1.Cells[1, 15].Value = "小时";
                sheet1.Cells[1, 16].Value = "拣货单每小时";
                sheet1.Cells[1, 17].Value = "个数每小时";
                sheet1.Cells[1, 18].Value = "定值倍数";
                sheet1.Cells[1, 19].Value = "工资";

                #endregion

                #region 数据行
                for (int idx = 0, rowIdx = 2, len = datas.Count; idx < len; idx++)
                {
                    var curOrder = datas[idx];
                    sheet1.Cells[rowIdx, 1].Value = curOrder._业绩归属人;
                    sheet1.Cells[rowIdx, 2].Value = curOrder._购买总数量_正常;
                    sheet1.Cells[rowIdx, 3].Value = curOrder._购买总数量_乱单;
                    sheet1.Cells[rowIdx, 4].Value = curOrder._购买总数量_本楼层乱单;
                    sheet1.Cells[rowIdx, 5].Value = curOrder._购买总数量;
                    sheet1.Cells[rowIdx, 6].Value = curOrder._拣货单张数_正常;
                    sheet1.Cells[rowIdx, 7].Value = curOrder._拣货单张数_乱单;
                    sheet1.Cells[rowIdx, 8].Value = curOrder._拣货单张数_本楼层乱单;
                    sheet1.Cells[rowIdx, 9].Value = curOrder._拣货单张数;
                    sheet1.Cells[rowIdx, 10].Value = string.IsNullOrWhiteSpace(curOrder._帮忙总时长) ? 0 : decimal.Parse(curOrder._帮忙总时长);
                    sheet1.Cells[rowIdx, 11].Value = string.IsNullOrWhiteSpace(curOrder._总时长) ? 0 : decimal.Parse(curOrder._总时长);
                    sheet1.Cells[rowIdx, 12].Value = curOrder._分钟;
                    sheet1.Cells[rowIdx, 13].Value = curOrder._拣货单效率;
                    sheet1.Cells[rowIdx, 14].Value = curOrder._购买数量效率;
                    sheet1.Cells[rowIdx, 15].Value = curOrder._小时;
                    sheet1.Cells[rowIdx, 16].Value = curOrder._拣货单每小时;
                    sheet1.Cells[rowIdx, 17].Value = curOrder._个数每小时;
                    sheet1.Cells[rowIdx, 18].Value = curOrder._定值倍数;
                    sheet1.Cells[rowIdx, 19].Value = curOrder._工资;
                    rowIdx++;
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

                ////////linux系统里面自动适应列宽有bug
                //////sheet1.Cells[sheet1.Dimension.Address].AutoFitColumns();

                package.Save();
            }
            #endregion
        }

        /// <summary>
        /// 根据绩效结果生成表格
        /// </summary>
        /// <param name="path"></param>
        /// <param name="summary"></param>
        /// <param name="details"></param>
        private void GenerateMonthlyPerfExcelTable(string path, List<_配货绩效_配货绩效结果> summary, List<_配货绩效_配货绩效结果> details, _配货绩效_全月绩效结果 cache)
        {
            if (System.IO.File.Exists(path)) System.IO.File.Delete(path);

            #region 生成绩效表格
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var workbox = package.Workbook;

                #region 绩效汇总表
                {
                    var sheet1 = workbox.Worksheets.Add("当月汇总");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "姓名";
                    sheet1.Cells[1, 2].Value = "拣货单数量";
                    sheet1.Cells[1, 3].Value = "乱单数量";
                    sheet1.Cells[1, 4].Value = "本楼层数量";
                    sheet1.Cells[1, 5].Value = "总数量";
                    sheet1.Cells[1, 6].Value = "拣货单张数";
                    sheet1.Cells[1, 7].Value = "乱单张数";
                    sheet1.Cells[1, 8].Value = "本楼层张数";
                    sheet1.Cells[1, 9].Value = "总张数";
                    sheet1.Cells[1, 10].Value = "帮忙总时长";
                    sheet1.Cells[1, 11].Value = "工作总时长";
                    sheet1.Cells[1, 12].Value = "分钟";
                    sheet1.Cells[1, 13].Value = "拣货单效率";
                    sheet1.Cells[1, 14].Value = "购买数量效率";
                    sheet1.Cells[1, 15].Value = "小时";
                    sheet1.Cells[1, 16].Value = "拣货单每小时";
                    sheet1.Cells[1, 17].Value = "个数每小时";
                    sheet1.Cells[1, 18].Value = "定值倍数";
                    sheet1.Cells[1, 19].Value = "工资";

                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = summary.Count; idx < len; idx++)
                    {
                        var curOrder = summary[idx];
                        sheet1.Cells[rowIdx, 1].Value = curOrder._业绩归属人;
                        sheet1.Cells[rowIdx, 2].Value = curOrder._购买总数量_正常;
                        sheet1.Cells[rowIdx, 3].Value = curOrder._购买总数量_乱单;
                        sheet1.Cells[rowIdx, 4].Value = curOrder._购买总数量_本楼层乱单;
                        sheet1.Cells[rowIdx, 5].Value = curOrder._购买总数量;
                        sheet1.Cells[rowIdx, 6].Value = curOrder._拣货单张数_正常;
                        sheet1.Cells[rowIdx, 7].Value = curOrder._拣货单张数_乱单;
                        sheet1.Cells[rowIdx, 8].Value = curOrder._拣货单张数_本楼层乱单;
                        sheet1.Cells[rowIdx, 9].Value = curOrder._拣货单张数;
                        sheet1.Cells[rowIdx, 10].Value = string.IsNullOrWhiteSpace(curOrder._帮忙总时长) ? 0 : decimal.Parse(curOrder._帮忙总时长);
                        sheet1.Cells[rowIdx, 11].Value = string.IsNullOrWhiteSpace(curOrder._总时长) ? 0 : decimal.Parse(curOrder._总时长);
                        sheet1.Cells[rowIdx, 12].Value = curOrder._分钟;
                        sheet1.Cells[rowIdx, 13].Value = curOrder._拣货单效率;
                        sheet1.Cells[rowIdx, 14].Value = curOrder._购买数量效率;
                        sheet1.Cells[rowIdx, 15].Value = curOrder._小时;
                        sheet1.Cells[rowIdx, 16].Value = curOrder._拣货单每小时;
                        sheet1.Cells[rowIdx, 17].Value = curOrder._个数每小时;
                        sheet1.Cells[rowIdx, 18].Value = curOrder._定值倍数;
                        sheet1.Cells[rowIdx, 19].Value = curOrder._工资;
                        rowIdx++;
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

                        //添加几个汇总计算
                        sheet1.Cells[endRow + 1, 1].Value = "合计";
                        sheet1.Cells[endRow + 1, 2].Formula = $"SUM(B2:B{endRow})";
                        sheet1.Cells[endRow + 1, 3].Formula = $"SUM(C2:C{endRow})";
                        sheet1.Cells[endRow + 1, 4].Formula = $"SUM(D2:D{endRow})";
                        sheet1.Cells[endRow + 1, 5].Formula = $"SUM(E2:E{endRow})";
                        sheet1.Cells[endRow + 1, 6].Formula = $"SUM(F2:F{endRow})";
                        sheet1.Cells[endRow + 1, 7].Formula = $"SUM(G2:G{endRow})";
                        sheet1.Cells[endRow + 1, 7].Formula = $"SUM(G2:G{endRow})";
                        sheet1.Cells[endRow + 1, 17].Formula = $"SUM(Q2:Q{endRow})";
                        using (var rng = sheet1.Cells[endRow + 1, 1, endRow + 1, 17])
                        {
                            rng.Style.Font.Color.SetColor(Color.Blue);//字体颜色
                        }
                    }
                    #endregion
                }
                #endregion

                #region 绩效详情表
                {
                    var sheet1 = workbox.Worksheets.Add("详情");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "姓名";
                    sheet1.Cells[1, 2].Value = "拣货单数量";
                    sheet1.Cells[1, 3].Value = "乱单数量";
                    sheet1.Cells[1, 4].Value = "本楼层数量";
                    sheet1.Cells[1, 5].Value = "总数量";
                    sheet1.Cells[1, 6].Value = "拣货单张数";
                    sheet1.Cells[1, 7].Value = "乱单张数";
                    sheet1.Cells[1, 8].Value = "本楼层张数";
                    sheet1.Cells[1, 9].Value = "总张数";
                    sheet1.Cells[1, 10].Value = "帮忙总时长";
                    sheet1.Cells[1, 11].Value = "工作总时长";
                    sheet1.Cells[1, 12].Value = "分钟";
                    sheet1.Cells[1, 13].Value = "拣货单效率";
                    sheet1.Cells[1, 14].Value = "购买数量效率";
                    sheet1.Cells[1, 15].Value = "小时";
                    sheet1.Cells[1, 16].Value = "拣货单每小时";
                    sheet1.Cells[1, 17].Value = "个数每小时";
                    sheet1.Cells[1, 18].Value = "定值倍数";
                    sheet1.Cells[1, 19].Value = "工资";
                    sheet1.Cells[1, 20].Value = "日期";
                    #endregion

                    #region 数据行
                    for (int idx = 0, rowIdx = 2, len = details.Count; idx < len; idx++)
                    {
                        var curOrder = details[idx];
                        sheet1.Cells[rowIdx, 1].Value = curOrder._业绩归属人;
                        sheet1.Cells[rowIdx, 2].Value = curOrder._购买总数量_正常;
                        sheet1.Cells[rowIdx, 3].Value = curOrder._购买总数量_乱单;
                        sheet1.Cells[rowIdx, 4].Value = curOrder._购买总数量_本楼层乱单;
                        sheet1.Cells[rowIdx, 5].Value = curOrder._购买总数量;
                        sheet1.Cells[rowIdx, 6].Value = curOrder._拣货单张数_正常;
                        sheet1.Cells[rowIdx, 7].Value = curOrder._拣货单张数_乱单;
                        sheet1.Cells[rowIdx, 8].Value = curOrder._拣货单张数_本楼层乱单;
                        sheet1.Cells[rowIdx, 9].Value = curOrder._拣货单张数;
                        sheet1.Cells[rowIdx, 10].Value = string.IsNullOrWhiteSpace(curOrder._帮忙总时长) ? 0 : decimal.Parse(curOrder._帮忙总时长);
                        sheet1.Cells[rowIdx, 11].Value = string.IsNullOrWhiteSpace(curOrder._总时长) ? 0 : decimal.Parse(curOrder._总时长);
                        sheet1.Cells[rowIdx, 12].Value = curOrder._分钟;
                        sheet1.Cells[rowIdx, 13].Value = curOrder._拣货单效率;
                        sheet1.Cells[rowIdx, 14].Value = curOrder._购买数量效率;
                        sheet1.Cells[rowIdx, 15].Value = curOrder._小时;
                        sheet1.Cells[rowIdx, 16].Value = curOrder._拣货单每小时;
                        sheet1.Cells[rowIdx, 17].Value = curOrder._个数每小时;
                        sheet1.Cells[rowIdx, 18].Value = curOrder._定值倍数;
                        sheet1.Cells[rowIdx, 19].Value = curOrder._工资;
                        sheet1.Cells[rowIdx, 20].Value = curOrder._绩效日期;
                        rowIdx++;
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
                }
                #endregion

                #region 每日小计
                {
                    var sheet1 = workbox.Worksheets.Add("每日汇总");

                    #region 标题行
                    sheet1.Cells[1, 1].Value = "日期";
                    sheet1.Cells[1, 2].Value = "拣货单数量";
                    sheet1.Cells[1, 3].Value = "乱单数量";
                    sheet1.Cells[1, 4].Value = "本楼层数量";
                    sheet1.Cells[1, 5].Value = "总数量";
                    sheet1.Cells[1, 6].Value = "拣货单张数";
                    sheet1.Cells[1, 7].Value = "乱单张数";
                    sheet1.Cells[1, 8].Value = "本楼层张数";
                    sheet1.Cells[1, 9].Value = "总张数";
                    #endregion

                    #region 数据行
                    var dates = cache.Perf.Keys.OrderBy(x => x).ToList();
                    for (int idx = 0, rowIdx = 2, len = dates.Count; idx < len; idx++)
                    {
                        var it = cache.Perf[dates[idx]];
                        sheet1.Cells[rowIdx, 1].Value = dates[idx];
                        sheet1.Cells[rowIdx, 2].Value = it.Sum(x => x._购买总数量_正常);
                        sheet1.Cells[rowIdx, 3].Value = it.Sum(x => x._购买总数量_乱单);
                        sheet1.Cells[rowIdx, 4].Value = it.Sum(x => x._购买总数量_本楼层乱单);
                        sheet1.Cells[rowIdx, 5].Value = it.Sum(x => x._购买总数量);
                        sheet1.Cells[rowIdx, 6].Value = it.Sum(x => x._拣货单张数_正常);
                        sheet1.Cells[rowIdx, 7].Value = it.Sum(x => x._拣货单张数_乱单);
                        sheet1.Cells[rowIdx, 8].Value = it.Sum(x => x._拣货单张数_本楼层乱单);
                        sheet1.Cells[rowIdx, 9].Value = it.Sum(x => x._拣货单张数);
                        rowIdx++;
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
                }
                #endregion

                package.Save();
            }
            #endregion
        }
    }
}