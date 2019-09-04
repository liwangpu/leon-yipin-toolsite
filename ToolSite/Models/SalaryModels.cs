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
}
