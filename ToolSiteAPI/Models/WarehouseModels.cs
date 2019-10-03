using EpplusHelper;
using System;

namespace ToolSiteAPI.Models
{
    public class _库存明细
    {
        [ExcelColumn("SKU码")]
        public string SKU { get; set; }
        [ExcelColumn("商品名称")]
        public string _商品名称 { get; set; }
        [ExcelColumn("单位")]
        public string _单位 { get; set; }
        [ExcelColumn("库位")]
        public string _库位 { get; set; }
        [ExcelColumn("可用数量")]
        public decimal _可用数量 { get; set; }
        [ExcelColumn("占用数量")]
        public decimal _占用数量 { get; set; }
        [ExcelColumn("库存数量")]
        public decimal _库存数量 { get; set; }
        [ExcelColumn("30天销量")]
        public decimal _30天销量 { get; set; }
        [ExcelColumn("15天销量")]
        public decimal _15天销量 { get; set; }
        [ExcelColumn("5天销量")]
        public decimal _5天销量 { get; set; }

        public decimal _平均日销量
        {
            get
            {
                return Math.Round((_30天销量 / 30 + _15天销量 / 15 + _5天销量 / 5) / 3, 0);
            }
        }
    }




}
