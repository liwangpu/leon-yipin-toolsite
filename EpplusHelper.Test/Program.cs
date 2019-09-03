using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace EpplusHelper.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var tablePath = @"C:\Users\Leon\Desktop\昆山全部库存 - 副本.xlsx";

                using (var package = new ExcelPackage(new FileInfo(tablePath)))
                {
                    var dataSheet = package.Workbook.Worksheets[0];

                    SheetReader<_库存明细>.From(dataSheet);

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

    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public string Tile { get; set; }
        public int Column { get; set; }
        public ExcelColumnAttribute(string title, int column = 1)
        {
            Tile = title;
            Column = column;
        }
    }

    public class SheetReader<T>
        where T : class, new()
    {
        public static List<T> From(ExcelWorksheet sheet, int headerRow = 1, int dataRow = 2)
        {
            var mappingType = typeof(T);

            var endColumn = sheet.Dimension.End.Column;
            var endRow = sheet.Dimension.End.Row;

            var mapping = new Dictionary<string, Tuple<string, int>>();
            var list = new List<T>();

            #region 根据标注,获取表格匹配信息
            {
                var exAttrType = typeof(ExcelColumnAttribute);
                var mappingTypeProps = mappingType.GetProperties();

                foreach (var prop in mappingTypeProps)
                {



                    var attrs = prop.GetCustomAttributes(exAttrType, true);
                    if (attrs.Count() > 0)
                    {
                        var attr = attrs[0] as ExcelColumnAttribute;

                        var distType = "string";
                        var ptName = prop.PropertyType.Name.ToLower();
                        if (ptName.Contains("int"))
                            distType = "int";
                        else if (ptName.Contains("decimal"))
                            distType = "decimal";
                        else if (ptName.Contains("double"))
                            distType = "double";
                        else if (ptName.Contains("datetime"))
                            distType = "datetime";
                        else { }


                        var distColumn = attr.Column;
                        //验证一下列数对不对,不对需要遍历纠正
                        if (sheet.Cells[headerRow, distColumn].Value == null || sheet.Cells[headerRow, distColumn].Value.ToString().Trim() != attr.Tile)
                        {
                            for (int i = 1; i <= endColumn; i++)
                            {
                                if (sheet.Cells[1, i].Value.ToString().Trim() == attr.Tile)
                                {
                                    distColumn = i;
                                    break;
                                }
                            }
                        }



                        if (!mapping.ContainsKey(prop.Name))
                            mapping[prop.Name] = new Tuple<string, int>(distType, distColumn);

                    }
                }
                #endregion

                for (int idx = endRow; idx >= dataRow; idx--)
                {
                    var instance = new T();
                    foreach (var item in mapping)
                    {
                        var obj = sheet.Cells[idx, item.Value.Item2].Value;
                        if (obj == null) continue;
                        var value = obj.ToString().Trim();

                        if (item.Value.Item1 == "int")
                            mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { Convert.ToInt32(value) });
                        else if (item.Value.Item1 == "decimal")
                            mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { Convert.ToDecimal(value) });
                        else if (item.Value.Item1 == "double")
                            mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { Convert.ToDouble(value) });
                        else if (item.Value.Item1 == "datetime")
                            mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { Convert.ToDateTime(value) });
                        else
                            mappingType.InvokeMember(item.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty, Type.DefaultBinder, instance, new object[] { value });

                 }

                    list.Add(instance);
                }





                var aaa = 1;


            }


            //var attrs = t.GetCustomAttributes(typeof(ExcelColumnAttribute), true);
            //var aaa = attrs.Count();
            //foreach (var item in collection)
            //{

            //}



            return list;
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
