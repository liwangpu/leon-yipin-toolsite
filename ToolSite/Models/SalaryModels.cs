using EpplusHelper;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ToolSite.Models.Salary
{
    public class _仓库加班考勤数据
    {
        public int _员工序号 { get; set; }
        public string _姓名 { get; set; }
        public List<string> _加班信息 { get; set; }
    }

    public class _仓库加班绩效
    {
        public int _员工序号 { get; set; }
        public string _姓名 { get; set; }
        public List<double> _出勤时长 { get; set; }
        public List<double> _加班时长 { get; set; }
        public List<bool> _打卡出现异常 { get; set; }
        public List<string> _原始打卡时间 { get; set; }
    }

    public class _配货绩效_员工月上班时间
    {
        [ExcelColumn("姓名", 1)]
        public string _姓名 { get; set; }

        [ExcelColumn("1号", 2)]
        public decimal _1号 { get; set; }

        [ExcelColumn("2号", 3)]
        public decimal _2号 { get; set; }

        [ExcelColumn("3号", 4)]
        public decimal _3号 { get; set; }

        [ExcelColumn("4号", 5)]
        public decimal _4号 { get; set; }

        [ExcelColumn("5号", 6)]
        public decimal _5号 { get; set; }

        [ExcelColumn("6号", 7)]
        public decimal _6号 { get; set; }

        [ExcelColumn("7号", 8)]
        public decimal _7号 { get; set; }

        [ExcelColumn("8号", 9)]
        public decimal _8号 { get; set; }

        [ExcelColumn("9号", 10)]
        public decimal _9号 { get; set; }

        [ExcelColumn("10号", 11)]
        public decimal _10号 { get; set; }

        [ExcelColumn("11号", 12)]
        public decimal _11号 { get; set; }

        [ExcelColumn("12号", 13)]
        public decimal _12号 { get; set; }

        [ExcelColumn("13号", 14)]
        public decimal _13号 { get; set; }

        [ExcelColumn("14号", 15)]
        public decimal _14号 { get; set; }

        [ExcelColumn("15号", 16)]
        public decimal _15号 { get; set; }

        [ExcelColumn("16号", 17)]
        public decimal _16号 { get; set; }

        [ExcelColumn("17号", 18)]
        public decimal _17号 { get; set; }

        [ExcelColumn("18号", 19)]
        public decimal _18号 { get; set; }

        [ExcelColumn("19号", 20)]
        public decimal _19号 { get; set; }

        [ExcelColumn("20号", 21)]
        public decimal _20号 { get; set; }

        [ExcelColumn("21号", 22)]
        public decimal _21号 { get; set; }

        [ExcelColumn("22号", 23)]
        public decimal _22号 { get; set; }

        [ExcelColumn("23号", 24)]
        public decimal _23号 { get; set; }

        [ExcelColumn("24号", 25)]
        public decimal _24号 { get; set; }

        [ExcelColumn("25号", 26)]
        public decimal _25号 { get; set; }

        [ExcelColumn("26号", 27)]
        public decimal _26号 { get; set; }

        [ExcelColumn("27号", 28)]
        public decimal _27号 { get; set; }

        [ExcelColumn("28号", 29)]
        public decimal _28号 { get; set; }

        [ExcelColumn("29号", 30)]
        public decimal _29号 { get; set; }

        [ExcelColumn("30号", 31)]
        public decimal _30号 { get; set; }

        [ExcelColumn("31号", 32)]
        public decimal _31号 { get; set; }
        public List<decimal> _工作时间 { get; set; }
    }

    public class _配货绩效_拣货单
    {

        [ExcelColumn("商品明细", 5)]
        public string _商品明细 { get; set; }

        [ExcelColumn("库位号", 7)]
        public string _完整库位号 { get; set; }

        public string _库位号
        {
            get
            {
                var arr = _完整库位号.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                if (arr.Count() > 0)
                {
                    var first = arr[0];
                    if (first.Length < 3)
                        return string.Empty;
                    return first.Substring(0, 3).ToUpper();
                }
                return "";
            }
        }

        public List<string> _拣货明细
        {
            get
            {
                var list = new List<string>();
                if (!string.IsNullOrEmpty(_商品明细))
                {
                    var arr = _商品明细.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Count() > 0)
                        list.AddRange(arr.ToList());
                }
                return list;
            }
        }

        public bool _乱单 { get; set; }
    }

    public class _配货绩效_乱单
    {
        [ExcelColumn("商品明细")]
        public string _商品明细 { get; set; }

        [ExcelColumn("库位号")]
        public string _完整库位号 { get; set; }

        public List<_配货绩效_拣货单> ToData()
        {
            var list = new List<_配货绩效_拣货单>();
            var _明细Arr = _商品明细.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            var _库位号Arr = _完整库位号.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (_明细Arr.Count == _库位号Arr.Count)
            {
                for (int idx = 0, len = _明细Arr.Count; idx < len; idx++)
                {
                    var model = new _配货绩效_拣货单();
                    model._商品明细 = _明细Arr[idx] + ";";
                    model._完整库位号 = _库位号Arr[idx] + ";";
                    model._乱单 = true;
                    list.Add(model);
                }
            }
            return list;
        }
    }

    //public class _配货绩效_拣货人员配置
    //{
    //    [ExcelColumn("库位")]
    //    public string _库位 { get; set; }

    //    [ExcelColumn("配货人员")]
    //    public string _配货人员 { get; set; }
    //}

    public class _配货绩效_订单详情数据
    {
        public string SKU { get; set; }
        public double Amount { get; set; }
        public bool _乱单 { get; set; }
    }

    public class _配货绩效_帮忙点货时间
    {
        [ExcelColumn("姓名")]
        public string _姓名 { get; set; }

        [ExcelColumn("工作时间")]
        public DateTime _工作时间 { get; set; }

        public TimeSpan _帮忙总时间
        {
            get
            {
                if (_工作时间 != null)
                    return new TimeSpan(_工作时间.Hour, _工作时间.Minute, 0);

                return new TimeSpan(0, 0, 0);
            }
        }
    }

    public class _配货绩效_拣货人员配置信息
    {
        [ExcelColumn("配货人员")]
        public string _姓名 { get; set; }
        [ExcelColumn("库位")]
        public string 管理库位 { get; set; }
    }

    public class _配货绩效_配货绩效结果
    {
        public string _业绩归属人 { get; set; }
        public double _购买总数量
        {
            get
            {
                return _购买总数量_正常 + _购买总数量_乱单;
            }
        }
        public double _拣货单张数
        {
            get
            {
                return _拣货单张数_正常 + _拣货单张数_乱单;
            }
        }

        public double _购买总数量_正常 { get; set; }
        public double _购买总数量_乱单 { get; set; }
        public double _拣货单张数_正常 { get; set; }
        public double _拣货单张数_乱单 { get; set; }

        public string _总时长 { get; set; }
        public string _帮忙总时长 { get; set; }
        public double _分钟 { get; set; }
        public double _d张数定值 { get; set; }
        public double _d张数占比 { get; set; }
        public double _d数量定值 { get; set; }
        public double _d数量占比 { get; set; }

        public double _拣货单效率
        {
            get
            {
                if (_分钟 <= 0)
                    return 0;
                return Math.Round(_拣货单张数 / _分钟, 4);
            }
        }

        public double _购买数量效率
        {
            get
            {
                if (_分钟 <= 0)
                    return 0;
                return Math.Round(_购买总数量 / _分钟, 4);
            }
        }

        public double _小时
        {
            get
            {
                if (_分钟 <= 0)
                    return 0;
                var mm = _分钟 % 60;
                var hh = (_分钟 - mm) / 60;
                return hh + Math.Round(mm / 60, 4);
            }
        }

        public double _拣货单每小时
        {
            get
            {
                if (_小时 <= 0)
                    return 0;
                return Math.Round(_拣货单张数 / _小时, 4);
            }
        }

        public double _个数每小时
        {
            get
            {
                if (_小时 <= 0)
                    return 0;
                return Math.Round(_购买总数量 / _小时, 4);
            }
        }

        public double _定值倍数
        {
            get
            {
                //= 拣货单每小时 / 208 * 0.75 + 个数每小时 / 1186 * 0.25

                return Math.Round(_拣货单每小时 / _d张数定值 * _d张数占比 + _个数每小时 / _d数量定值 * _d数量占比, 4);
            }
        }

        public double _工资
        {
            get
            {
                if (_小时 <= 0)
                    return 0;
                //=IF(定值倍数>1,(定值倍数-1)*3000,0)
                if (_定值倍数 > 1)
                    return Math.Round((_定值倍数 - 1) * 3000, 2);
                else
                    return 0;
            }
        }
    }
}
