﻿using OfficeOpenXml;
using System;
using System.IO;

namespace EpplusHelper.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var tablePath = @"C:\Users\Leon\Desktop\昆山全部库存.xlsx";

                using (var package = new ExcelPackage(new FileInfo(tablePath)))
                {
                    var dataSheet = package.Workbook.Worksheets[0];

                    var list = SheetReader<_库存明细>.From(dataSheet);




                }














            }
            catch (Exception ex)
            {
                Console.WriteLine("error:{0}", ex.Message);
            }
            Console.WriteLine("++++++++++++++++++++");
            Console.ReadKey();
        }
    }




    public class _库存明细
    {
        [ExcelColumn("SKU码", 5)]
        public string Sku { get; set; }

        [ExcelColumn("商品重量(克)", 11)]
        public decimal Weight { get; set; }

        [ExcelColumn("商品创建时间", 38)]
        public DateTime Test { get; set; }

        [ExcelColumn("商品成本单价", 33)]
        public decimal Price { get; set; }

    }


}