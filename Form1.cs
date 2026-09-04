using DevExpress.Xpo.DB.Helpers;
using MySql.Data.MySqlClient;
using MySql.Data.MySqlClient;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using out_execl;
using out_exel;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace @out
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private const string conn_mes = "Server=10.10.210.130;user id=sa;password=Aa123456;initial catalog=UFTData545228_900000;Connect Timeout=30;Persist Security Info=True;Current Language=Simplified Chinese;pooling=false";
        private const string constring = "data source=172.17.41.196;port=3306;database=yongyou_zjedu_st_2025;user id=st;password=123456;pooling=true;charset=utf8;SslMode=None;";
        //private const string constring = "data source=localhost;port=3306;database=test;user id=root;password=123456;pooling=true;charset=utf8;";
        private const string constring2 = "Server=127.0.0.1;user id=sa;password=123456;initial catalog=text;Connect Timeout=30;Persist Security Info=True;Current Language=Simplified Chinese;pooling=false";

        //execl导入到MySQL
        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置对话框属性
            openFileDialog.InitialDirectory = "C:\\"; // 默认目录
            //选择路径
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "Excel 2007+ (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*"; // 文件过滤器
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox2.Text = openFileDialog.FileName;
                    //Process.Start(openFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("无法打开文件: " + ex.Message);
                }
            }
            Thread.Sleep(5000);
            try
            {
                MySqlConnection msc = new MySqlConnection(constring);
                msc.Open();
                //读取execl文件
                Workbook workbook = new Workbook();
                // 加载Excel文件
                workbook.LoadFromFile(textBox2.Text);
                //获取表的数量
                int sheetCount = workbook.Worksheets.Count;

                foreach (Worksheet worksheet in workbook.Worksheets)
                // 获取第一个工作表
                //Worksheet worksheet = workbook.Worksheets[0];
                // 将工作表中的数据导出到一个DataTable对象
                {
                    DataTable dataTable = worksheet.ExportDataTable();
                    // 构建 插入或更新 语句
                    

                    string columns = string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                    string parameters = string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => "@" + c.ColumnName));
                    var updateColumns = dataTable.Columns.Cast<DataColumn>().Where(c => !dataTable.Columns.Contains(c.ColumnName))
            .Select(c => $"{c.ColumnName} = VALUES({c.ColumnName})").ToList();
                    // 如果没有可更新的列，添加一个虚拟更新
                    if (updateColumns.Count == 0)
                    {
                        // 假设表中有一个LastUpdated字段
                        updateColumns.Add("GXSJ = NOW()");
                        // 或者使用虚拟更新（某些MySQL版本支持）
                        // updateColumns.Add("Id = Id");
                    }
                    //string sql = $"INSERT INTO {worksheet.Name} ({columns}) VALUES ({parameters}) ON DUPLICATE KEY UPDATE {string.Join(", ", updateColumns)}";
                    string sql = $@"
            INSERT INTO {worksheet.Name} ({columns}) 
            VALUES ({parameters})
            ON DUPLICATE KEY UPDATE {string.Join(", ", updateColumns)}";
                    using (var command = new MySqlCommand(sql, msc))
                    {
                        // 添加参数
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            command.Parameters.Add("@" + column.ColumnName, GetMySqlDbType(column.DataType));
                        }

                        // 插入每一行
                        foreach (DataRow row in dataTable.Rows)
                        {
                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                command.Parameters[i].Value = row[i] == DBNull.Value ? "" : row[i];
                            }
                            command.ExecuteNonQuery();
                        }
                    }
                    
                    
                    
                }
                MessageBox.Show("导入成功！");
            }
            catch (Exception ex)
            {
                //执行update语句
                MessageBox.Show("导入失败！"+ex.Message );
            }
        }
        private static MySqlDbType GetMySqlDbType(Type type)
        {
            if (type == typeof(int)) return MySqlDbType.Int32;
            if (type == typeof(string)) return MySqlDbType.VarChar;
            if (type == typeof(DateTime)) return MySqlDbType.DateTime;
            if (type == typeof(decimal)) return MySqlDbType.Decimal;
            if (type == typeof(bool)) return MySqlDbType.Bit;
            if (type == typeof(long)) return MySqlDbType.Int64;
            if (type == typeof(float)) return MySqlDbType.Float;
            if (type == typeof(double)) return MySqlDbType.Double;
            if (type == typeof(byte[])) return MySqlDbType.Blob;
            // 添加更多类型映射...
            return MySqlDbType.VarChar;
        }

        //mssql导出成execl
        private void button1_Click(object sender, EventArgs e)
        {
            int count = 0;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = "D:\\";
            saveFileDialog1.Title = "保存路径";
            saveFileDialog1.Filter = "Excel 2007+ (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif|文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";
            //saveFileDialog1.Title = "Save an Image File";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox1.Text = saveFileDialog1.FileName;
                    //Process.Start(openFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            DateTime stratTime = Convert.ToDateTime(this.dateTimePicker1.Text);
            DateTime endTime =Convert.ToDateTime( this.dateTimePicker2.Text);
            string filePath = textBox1.Text;
            try
            {

                // 创建工作簿（.xlsx 格式）
                IWorkbook workbook = new XSSFWorkbook();
                // 如果是 .xls（旧版 Excel），用：
                // IWorkbook workbook = new HSSFWorkbook();
                #region 采购汇总单订单信息
                string sql = string.Format(@"SELECT a.code AS HZDDH,'4133013854' as  XXID,'JY33301090174786' as STID,
1 as CGLX,a.voucherdate AS XDRQ,a.auditeddate AS SDRQ,a.createdtime AS CJSJ,a.updated AS GXSJ,
a.totalAmount AS DDZJ,CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
FROM PU_PurchaseOrder a 
LEFT JOIN PU_PurchaseOrder_b b on a.ID = b.idPurchaseOrderDTO
left join AA_Inventory c on c.id = b.idinventory 
left join AA_Partner d on d.id = a.idpartner
left join AA_Unit e on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass = f.id
where a.createdtime between '{0}' and '{1}' and auditeddate is not null and c.Disabled =0 and (Left(f.code,2))<>'14' and (Left(f.code,2))<>'99'and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00' 
and a.auditeddate is not null GROUP BY a.voucherdate,a.auditeddate,a.createdtime,a.updated,a.totalAmount,a.code order by a.code asc", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));//SQL语句
                DataTable dt = DBHelper.ExecuterQuery(conn_mes, sql);
                if (dt != null && dt.Rows.Count > 0)
                {                   
                    // 添加工作表
                    ISheet sheet = workbook.CreateSheet("gxcgzz_cghzdddxx"); //采购汇总单订单信息
                    // 创建行（第 0 行）
                    IRow row = sheet.CreateRow(0);
                    // 写入数据
                    //row.CreateCell(0).SetCellValue(dt.Columns[0].ColumnName);
                    //row.CreateCell(1).SetCellValue("CreatedDate");
                    foreach (DataColumn dr in dt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < dt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        //赋值XXID                        
                        //row.CreateCell(0).SetCellValue(dt.Rows[i - 1]["QName"].ToString());
                        //row.CreateCell(1).SetCellValue(dt.Rows[i - 1]["CreatedDate"].ToString());
                        foreach (DataColumn dr in dt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(dt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                    // 保存文件
                    //using (var fileStream = new FileStream(filePath, FileMode.Create))
                    //{
                    //    workbook.Write(fileStream);
                    //}

                    //Console.WriteLine($"Excel 文件已创建：{filePath}");

                }
                else
                {
                    count = count + 1;
                    //MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 采购汇总单订单详情信息
                string sql1 = string.Format(@"SELECT b.id AS HZDXQID,'4133013854' as  XXID,a.code AS HZDDH,
d.code AS GYSID,c.code AS SCID,b.OrigDiscountPrice AS DQJG,b.taxamount AS DXZJ,e.name AS SCDW,b.Quantity AS CGSJ,
a.createdtime AS CJSJ,a.updated AS GXSJ,CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
from  PU_PurchaseOrder a LEFT JOIN PU_PurchaseOrder_b b on a.ID = b.idPurchaseOrderDTO
left join AA_Inventory c on c.id = b.idinventory 
left join AA_Partner d  on d.id =  a.idpartner
left  join  AA_Unit  e  on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass  =  f.id
where c.Disabled  =0 and (Left(f.code,2))<>'14' and a.auditeddate is not null and (Left(f.code,2))<>'99'and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00' 
AND a.createdtime  between '{0}'  and '{1}'", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable dt1 = DBHelper.ExecuterQuery(conn_mes, sql1);
                if (dt1 != null && dt1.Rows.Count > 0)
                {
                    
                    // 添加工作表
                    ISheet sheet = workbook.CreateSheet("gxcgzz_cghzdddxxxx");//采购汇总单订单详情信息                                                                             // 创建行（第 0 行）
                    IRow row = sheet.CreateRow(0);
                    // 写入数据
                    //row.CreateCell(0).SetCellValue(dt.Columns[0].ColumnName);
                    //row.CreateCell(1).SetCellValue("CreatedDate");
                    foreach (DataColumn dr in dt1.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < dt1.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);                        
                        //row.CreateCell(0).SetCellValue(dt.Rows[i - 1]["QName"].ToString());
                        //row.CreateCell(1).SetCellValue(dt.Rows[i - 1]["CreatedDate"].ToString());
                        foreach (DataColumn dr in dt1.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(dt1.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    //MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 入库信息
                string rkxx = string.Format(@"SELECT DISTINCT  a.code AS RKDH,'4133013854' as  XXID,1 as RKFS,
a.Createdtime AS CJSJ,a.updated AS GXSJ,a.totalAmount AS RKZJE,
CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
from PU_PurchaseArrival a 
left join PU_PurchaseArrival_b b on a.id = b.idPurchaseArrivalDTO
left join AA_Inventory c on c.id = b.idinventory
left  join  AA_Unit  e  on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass  =  f.id
where c.Disabled  =0 and (Left(f.code,2))<>'14' and (Left(f.code,2))<>'99'and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00'
and a.createdtime  between '{0}'  and '{1}'", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable rkdt = DBHelper.ExecuterQuery(conn_mes, rkxx);
                if (rkdt != null && rkdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_rkxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in rkdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < rkdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        //row.CreateCell(0).SetCellValue(dt.Rows[i - 1]["QName"].ToString());
                        //row.CreateCell(1).SetCellValue(dt.Rows[i - 1]["CreatedDate"].ToString());
                        foreach (DataColumn dr in rkdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(rkdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    //MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 入库详情信息
                string rkxqxx = string.Format(@"SELECT b.ID AS RKXQID,'4133013854' as  XXID,a.code AS RKDH,
c.code AS SCID,b.OrigDiscountPrice AS SCDJ,b.quantity AS SCSL,b.taxamount AS SCZJ,a.createdtime AS CJSJ,
a.updated AS GXSJ,'1' as SFZJZC,CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT  
FROM PU_PurchaseArrival   a
left join PU_PurchaseArrival_b b on a.id = b.idPurchaseArrivalDTO
left join AA_Inventory c on c.id = b.idinventory
left  join  AA_Unit  e  on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass  =  f.id
where c.Disabled  =0 and (Left(f.code,2))<>'14' and (Left(f.code,2))<>'99'and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00'
AND a.createdtime  between '{0}'  and '{1}'", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable rkxqdt = DBHelper.ExecuterQuery(conn_mes, rkxqxx);
                if (rkxqdt != null && rkxqdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_rkxqxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in rkxqdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < rkxqdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        foreach (DataColumn dr in rkxqdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(rkxqdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    //MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 出库信息
                string ckxx = string.Format(@"SELECT DISTINCT a.code AS CKDH,'4133013854' as  XXID,a.Createdtime AS CJSJ,
a.updated AS GXSJ,a.auditeddate AS SHRQ,isnull(a.amount,0) AS CKZJE,
CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
from ST_RDRecord a 
left join ST_RDRecord_b b on a.id = b.idRDRecordDTO 
left join AA_Inventory c on c.id = b.idinventory  
left  join  AA_Unit  e  on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass  =  f.id
where c.Disabled  =0 and (Left(f.code,2))<>'14' and (Left(f.code,2))<>'99' 
and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00' 
and a.Createdtime between '{0}'  and '{1}'", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable ckdt = DBHelper.ExecuterQuery(conn_mes, ckxx);
                if (ckdt != null && ckdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_ckxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in ckdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < ckdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        foreach (DataColumn dr in ckdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(ckdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    // MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 出库详情信息
                string ckxqxx = string.Format(@"SELECT b.ID AS CKXQID,'4133013854' as  XXID,a.code AS CKDH,c.code AS SCID,
b.Price AS SCDJ,b.amount AS SCZJ,b.quantity AS SCSL,a.createdtime AS CJSJ,a.updated AS GXSJ,'1' as SFZJZC,
CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
from ST_RDRecord a 
left join ST_RDRecord_b b on a.id = b.idRDRecordDTO 
left join AA_Inventory c on c.id = b.idinventory  
left  join  AA_Unit  e  on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass  =  f.id
where c.Disabled  =0 and (Left(f.code,2))<>'14' and (Left(f.code,2))<>'99'and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00'
AND a.createdtime  between '{0}'  and '{1}'", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable ckxqdt = DBHelper.ExecuterQuery(conn_mes, ckxqxx);
                if (ckxqdt != null && ckxqdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_ckxqxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in ckxqdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < ckxqdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        foreach (DataColumn dr in ckxqdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(ckxqdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    // MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 验收信息
                string ysxx = string.Format(@"select a.code as YSDDH,'4133013854' as  XXID,a.code  AS HZDDH,
a.voucherdate AS JDRQ,a.voucherdate AS SDRQ,'JY33301090174786' as STID, d.code AS GYSID,a.totalAmount AS DDZJE,
sum(b.Quantity) as DDZSL,3 as DDZT,4 AS PSZT,a.voucherdate AS YSRQ ,a.totalAmount AS YSZJE,0 AS SFTH,
a.createdtime AS CJSJ,a.updated AS GXSJ,CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
FROM PU_PurchaseOrder a 
LEFT JOIN PU_PurchaseOrder_b b on a.ID = b.idPurchaseOrderDTO
left join AA_Inventory c on c.id = b.idinventory 
left join AA_Partner d  on d.id =  a.idpartner
left  join  AA_Unit  e  on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass  =  f.id
where c.Disabled  =0 and (Left(f.code,2))<>'14' and (Left(f.code,2))<>'99'and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00' and a.auditeddate is not null AND 
a.createdtime  between '{0}'  and '{1}' GROUP BY a.code,a.code,a.voucherdate,a.voucherdate,d.code,a.voucherdate,a.voucherdate,a.updated,a.createdtime,a.totalAmount", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable ysdt = DBHelper.ExecuterQuery(conn_mes, ysxx);
                if (ysdt != null && ysdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_cgddysxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in ysdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < ysdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        foreach (DataColumn dr in ysdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(ysdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    //MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 验收详细信息
                string ysxqxx = string.Format(@"select b.id as CGDDHXQID,'4133013854' as XXID,a.code as YSDDH,
c.code AS SCID,d.code AS GYSID,b.Quantity AS CGSL,c.name as SCGG,e.name as SCDW,b.OrigDiscountPrice AS CGDJ,
b.taxamount AS CGZJ,b.Quantity AS QRSL,0 AS JSSL,b.taxamount AS SJYSJE,a.createdtime AS CJSJ,a.updated AS GXSJ,
CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
from  PU_PurchaseOrder a LEFT JOIN PU_PurchaseOrder_b b on a.ID = b.idPurchaseOrderDTO
left join AA_Inventory c on c.id = b.idinventory 
left join AA_Partner d  on d.id =  a.idpartner
left join AA_Unit  e  on e.id = c.idunit
left join AA_InventoryClass f on c.idinventoryclass  =  f.id
where c.Disabled  =0 and (Left(f.code,2))<>'14' and (Left(f.code,2))<>'99'and (Left(f.code,2))<>'98' and (Left(f.code,2))<>'00' and a.auditeddate is not null
AND a.createdtime  between '{0}'  and '{1}'", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));//,b.taxamount AS ZZJG
                DataTable ysxqdt = DBHelper.ExecuterQuery(conn_mes, ysxqxx);
                if (ysxqdt != null && ysxqdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_cgddysxqxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in ysxqdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < ysxqdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        foreach (DataColumn dr in ysxqdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(ysxqdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    // MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 食材信息
                string scxx = string.Format(@"SELECT a.code as SCID,'4133013854' as XXID,a.name AS MC,
a.createdtime AS CJSJ,a.updated AS GXSJ,b.name AS DW,
CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT
FROM AA_Inventory a
left join  AA_Unit b on b.id = a.idunit 
left join AA_InventoryClass c on a.idinventoryclass  =  c.id
where a.Disabled  =0 and (Left(c.code,2))<>'14' and (Left(c.code,2))<>'99'and (Left(c.code,2))<>'98' and (Left(c.code,2))<>'00' 
and a.createdTime  between '{0}'  and '{1}' ", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable scdt = DBHelper.ExecuterQuery(conn_mes, scxx);
                if (scdt != null && scdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_scxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in scdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < scdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        foreach (DataColumn dr in scdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(scdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    //MessageBox.Show("没有查询到数据！");
                }
                #endregion

                #region 食材价格信息
                string scjgxx = string.Format(@"SELECT A.code AS SCJDID,'4133013854' AS XXID,
a.code as SCID,ISNULL(d.latestCost,1) AS DQJG,a.createdtime AS KSRQ,a.updated AS JSRQ,a.createdtime AS CJSJ,
a.updated AS GXSJ,CONVERT(VARCHAR(32), CONVERT(VARBINARY(16), NEWID()), 2) AS UUID,'0' as SJZT 
FROM AA_Inventory a
left join  AA_Unit b on b.id = a.idunit 
left join AA_InventoryClass c on a.idinventoryclass  =  c.id
LEFT JOIN AA_Inventory_Cost D ON A.ID = D.idInventory 
where a.Disabled  =0 and (Left(c.code,2))<>'14' and (Left(c.code,2))<>'99'and (Left(c.code,2))<>'98' and (Left(c.code,2))<>'00' 
and a.createdTime  between '{0}'  and '{1}' 
GROUP BY d.latestCost,a.createdtime,a.updated,A.ID,A.CODE ORDER BY a.Code ", stratTime.ToString("yyyy-MM-dd 00:00:00"), endTime.ToString("yyyy-MM-dd 23:59:59"));
                DataTable scjgdt = DBHelper.ExecuterQuery(conn_mes, scjgxx);
                if (scjgdt != null && scjgdt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet("gxcgzz_scjgxx");
                    IRow row = sheet.CreateRow(0);
                    foreach (DataColumn dr in scjgdt.Columns)
                    {
                        row.CreateCell(dr.Ordinal).SetCellValue(dr.ColumnName);
                    }
                    for (int i = 1; i < scjgdt.Rows.Count + 1; i++)
                    {
                        row = sheet.CreateRow(i);
                        foreach (DataColumn dr in scjgdt.Columns)
                        {
                            row.CreateCell(dr.Ordinal).SetCellValue(scjgdt.Rows[i - 1][dr.ColumnName].ToString());
                        }
                    }
                }
                else
                {
                    count = count + 1;
                    //MessageBox.Show("没有查询到数据！");
                }
                #endregion

                

                //保存文件
                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    workbook.Write(fileStream);
                }
                MessageBox.Show("有" + count + "个表格没有查询到数据！");
                MessageBox.Show("Excel 文件已创建：" + filePath);

                //System.Diagnostics.Process.Start("explorer.exe", "/select," + filePath);
                //关闭应用窗口
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
    }
}
