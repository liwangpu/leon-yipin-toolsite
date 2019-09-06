using EpplusHelper;
using System.Collections.Generic;

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
}
