using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.IO;
using System.Configuration;
using System.Data;
using System.Diagnostics;

namespace AsposeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var appStartDir = AppDomain.CurrentDomain.BaseDirectory;

            // 1. 设置Aspose的License
            //var licensePath = appStartDir + ConfigurationManager.AppSettings["LicensePath"];
            //new License().SetLicense(licensePath);

            // 2. 加载模板
            var templatePath = appStartDir + ConfigurationManager.AppSettings["TemplatePath"];
            var designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(templatePath);

            // 3. 获取数据源
            var dt = new DataTable();
            dt.Columns.Add("DeptNO", typeof(string));
            dt.Columns.Add("DeptName", typeof(string));
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Field_1", typeof(string));
            dt.Columns.Add("Field_2", typeof(string));
            dt.Columns.Add("Field_3", typeof(string));
            dt.Columns.Add("Field_4", typeof(string));
            dt.Columns.Add("Field_5", typeof(string));

            for (int i = 1; i < 4; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    var dr = dt.NewRow();
                    dr["DeptNO"] = "D00" + i;
                    dr["DeptName"] = "部门-D00" + i;
                    dr["ID"] = "U00" + j;
                    dr["Name"] = $"D00{i}-User-{j}";
                    dr["Age"] = 20 + j;
                    dr["Field_1"] = $"Field-1";
                    dr["Field_2"] = $"Field-2";
                    dr["Field_3"] = $"Field-3";
                    dr["Field_4"] = $"Field-4";
                    dr["Field_5"] = $"Field-5";
                    dt.Rows.Add(dr);
                }
            }

            // 4. 绑定数据源
            var deptGroup = dt.AsEnumerable().GroupBy(e => new { deptNO = e.Field<string>("DeptNO"), deptName = e.Field<string>("DeptName") }).Select(e => new { DeptNo = e.Key.deptNO, DeptName = e.Key.deptName });
            foreach (var dept in deptGroup)
            {
                var newSheet = designer.Workbook.Worksheets[designer.Workbook.Worksheets.AddCopy(0)];
                newSheet.Name = $"Sheet-{dept.DeptNo}";
                newSheet.Replace("[A]", $"[A-{dept.DeptNo}]");

                // 模板Sheet添加列
                newSheet.Cells.InsertColumns(4, 1);
                var newColTitle = newSheet.Cells.GetCell(3, 4);
                newColTitle.PutValue("Field_5");
                var newColValue = newSheet.Cells.GetCell(4, 4);
                newColValue.PutValue($"&=[A-{dept.DeptNo}].Field_5");

                var newDt = dt.Select($"DeptNO='{dept.DeptNo}'").CopyToDataTable();
                newDt.TableName = $"A-{dept.DeptNo}";

                designer.SetDataSource(newDt);
                designer.SetDataSource($"[A-{dept.DeptNo}]DeptNO", dept.DeptNo);
                designer.SetDataSource($"[A-{dept.DeptNo}]DeptName", dept.DeptName);
                designer.SetDataSource($"[A-{dept.DeptNo}]Comments", "This is a comment");
            }

            designer.Workbook.Worksheets.RemoveAt(0);
            designer.Workbook.Worksheets.ActiveSheetIndex = 0;

            // 5. 生成Excel
            designer.Process();
            var reportFolder = appStartDir + ConfigurationManager.AppSettings["ReportFolder"];
            if (!Directory.Exists(reportFolder)) Directory.CreateDirectory(reportFolder);
            var reportPath = reportFolder + DateTime.Now.ToString("HHmmss")+".xlsx";
            designer.Workbook.Save(reportPath, SaveFormat.Xlsx);

            // 6. 打开生成好的文件
            Process.Start(reportPath);
        }
    }
}
