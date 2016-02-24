using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace DBConvertToExcel
{
    public partial class Form1 : Form
    {
        string saveFilePath;
        string openFilePath;
        //string foldPath;
        public Form1()
        {
            InitializeComponent();
            
        }
        //属性数据转换为excel表格
        private void ConverToExcel_Click(object sender, EventArgs e)
        {
            CreateExcel(saveFilePath);
            string sqlstr = @"select * from [JMDData]";
            DataSet jmdds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateJMDExcel(jmdds.Tables[0], saveFilePath);
            }
            sqlstr = @"select * from [DLData]";
            DataSet dlds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateDLExcel(dlds.Tables[0], saveFilePath);
            }

            sqlstr = @"select * from [SXData]";
            DataSet sxds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateSXExcel(sxds.Tables[0], saveFilePath);
            }

            sqlstr = @"select * from [GXData]";
            DataSet gxds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateGXExcel(gxds.Tables[0], saveFilePath);
            }

            sqlstr = @"select * from [JJXData]";
            DataSet jjxds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateJJXExcel(jjxds.Tables[0], saveFilePath);
            }

            sqlstr = @"select * from [DMData]";
            DataSet dmds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateTZDBExcel(dmds.Tables[0], saveFilePath);
            }

            sqlstr = @"select * from [PZJData]";
            DataSet pzjds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateDLZJExcel(pzjds.Tables[0], saveFilePath);
            }
            sqlstr = @"select * from [WZBZData]";
            DataSet wzbzds = SqliteHelper.ExcelDataSet(sqlstr, openFilePath);
            if (saveFilePath != null)
            {
                CreateWZZJExcel(wzbzds.Tables[0], saveFilePath);
            }
            MessageBox.Show("导出成功！");
        }
        //选择db文件
        private void ChoosePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "打开文件";
            openFileDialog.Filter = "DB文件（*.db)|*db";
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                openFilePath = openFileDialog.FileName;
            }
            AttributePath.Text = openFilePath;
        }
        //打开存在的excel表（刚刚创建的）
        public HSSFWorkbook OpenExistExcel(string filePath)
        {
            HSSFWorkbook MyHSSFWorkBook;
            Stream MyExcelStream = OpenClasspathResource(filePath);
            MyHSSFWorkBook = new HSSFWorkbook(MyExcelStream);
            return MyHSSFWorkBook;
        }
        //读入流打开文件（excel）
        private Stream OpenClasspathResource(String fileName)
        {
            FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            return file;
        }
        #region 属性数据导出excel
        //居民地要素写入excel
        public void CreateJMDExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("居民地") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(0);
            IRow jmdrow = sheet.CreateRow(0);
            ICell jmdcell = jmdrow.CreateCell(0);
            jmdcell.SetCellValue("ID");
            jmdcell = jmdrow.CreateCell(1);
            jmdcell.SetCellValue("LinkID");
            jmdcell = jmdrow.CreateCell(2);
            jmdcell.SetCellValue("要素名称");
            jmdcell = jmdrow.CreateCell(3);
            jmdcell.SetCellValue("房屋层数");
            jmdcell = jmdrow.CreateCell(4);
            jmdcell.SetCellValue("房屋材质");
            jmdcell = jmdrow.CreateCell(5);
            jmdcell.SetCellValue("房檐改正");
            jmdcell = jmdrow.CreateCell(6);
            jmdcell.SetCellValue("时间");
            jmdcell = jmdrow.CreateCell(7);
            jmdcell.SetCellValue("备注");
            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    jmdrow = sheet.CreateRow(j);
                    j++;

                    jmdcell = jmdrow.CreateCell(0);
                    jmdcell.SetCellValue(dataRow["ID"].ToString());
                    jmdcell = jmdrow.CreateCell(1);
                    jmdcell.SetCellValue(dataRow["LinkID"].ToString());
                    jmdcell = jmdrow.CreateCell(2);
                    jmdcell.SetCellValue(dataRow["FTName"].ToString());
                    jmdcell = jmdrow.CreateCell(3);
                    jmdcell.SetCellValue(dataRow["FWCS"].ToString());
                    jmdcell = jmdrow.CreateCell(4);
                    jmdcell.SetCellValue(dataRow["FWCZ"].ToString());
                    jmdcell = jmdrow.CreateCell(5);
                    jmdcell.SetCellValue(dataRow["FYGZ"].ToString());
                    jmdcell = jmdrow.CreateCell(6);
                    jmdcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    jmdcell = jmdrow.CreateCell(7);
                    jmdcell.SetCellValue(dataRow["BZ"].ToString());
                    MemoryStream ms = new MemoryStream();
                    workbook.Write(ms);

                    using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                    {
                        byte[] bArr = ms.ToArray();
                        fs.Write(bArr, 0, bArr.Length);
                        fs.Flush();

                    }
                }
            }
        }
        //道路要素写入excel
        public void CreateDLExcel(System.Data.DataTable dt, string path)
        {
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(1);
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("道路") : workbook.CreateSheet(dt.TableName);
            IRow dlrow = sheet.CreateRow(0);
            ICell dlidcell = dlrow.CreateCell(0);
            dlidcell.SetCellValue("ID");
            ICell dllinkidcell = dlrow.CreateCell(1);
            dllinkidcell.SetCellValue("LinkID");
            ICell dlmccell = dlrow.CreateCell(2);
            dlmccell.SetCellValue("道路名称");
            ICell dlxhcell = dlrow.CreateCell(3);
            dlxhcell.SetCellValue("道路线号");
            ICell dldmcell = dlrow.CreateCell(4);
            dldmcell.SetCellValue("道路代码");
            ICell dlsjcell = dlrow.CreateCell(5);
            dlsjcell.SetCellValue("时间");
            ICell dlbzcell = dlrow.CreateCell(6);
            dlbzcell.SetCellValue("备注");

            int j = 1;
            for (int i = 1; i < dt.Rows.Count + 1; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    dlrow = sheet.CreateRow(j);
                    j++;

                    dlidcell = dlrow.CreateCell(0);
                    dlidcell.SetCellValue(dataRow["ID"].ToString());
                    dllinkidcell = dlrow.CreateCell(1);
                    dllinkidcell.SetCellValue(dataRow["LinkID"].ToString());
                    dlmccell = dlrow.CreateCell(2);
                    dlmccell.SetCellValue(dataRow["DLMC"].ToString());
                    dlxhcell = dlrow.CreateCell(3);
                    dlxhcell.SetCellValue(dataRow["DLXH"].ToString());
                    dldmcell = dlrow.CreateCell(4);
                    dldmcell.SetCellValue(dataRow["DJDM"].ToString());
                    dlsjcell = dlrow.CreateCell(5);
                    dlsjcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    dlbzcell = dlrow.CreateCell(6);
                    dlbzcell.SetCellValue(dataRow["BZ"].ToString());
                }
            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }
        }
        //水系要素写入excel
        public void CreateSXExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("水系") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(2);
            IRow sxrow = sheet.CreateRow(0);
            ICell sxcell = sxrow.CreateCell(0);
            sxcell.SetCellValue("ID");
            sxcell = sxrow.CreateCell(1);
            sxcell.SetCellValue("LinkID");
            sxcell = sxrow.CreateCell(2);
            sxcell.SetCellValue("要素名称");
            sxcell = sxrow.CreateCell(3);
            sxcell.SetCellValue("附属设施");
            sxcell = sxrow.CreateCell(4);
            sxcell.SetCellValue("时间");
            sxcell = sxrow.CreateCell(5);
            sxcell.SetCellValue("备注");
            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    sxrow = sheet.CreateRow(j);
                    j++;

                    sxcell = sxrow.CreateCell(0);
                    sxcell.SetCellValue(dataRow["ID"].ToString());
                    sxcell = sxrow.CreateCell(1);
                    sxcell.SetCellValue(dataRow["LinkID"].ToString());
                    sxcell = sxrow.CreateCell(2);
                    sxcell.SetCellValue(dataRow["YSMC"].ToString());
                    sxcell = sxrow.CreateCell(3);
                    sxcell.SetCellValue(dataRow["FSSS"].ToString());
                    sxcell = sxrow.CreateCell(4);
                    sxcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    sxcell = sxrow.CreateCell(5);
                    sxcell.SetCellValue(dataRow["BZ"].ToString());

                }
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }


        }
        //管线电力线写入excel
        public void CreateGXExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("管线") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(3);
            IRow gxrow = sheet.CreateRow(0);
            ICell gxcell = gxrow.CreateCell(0);
            gxcell.SetCellValue("ID");
            gxcell = gxrow.CreateCell(1);
            gxcell.SetCellValue("LinkID");
            gxcell = gxrow.CreateCell(2);
            gxcell.SetCellValue("要素名称");
            gxcell = gxrow.CreateCell(3);
            gxcell.SetCellValue("电力线走向");
            gxcell = gxrow.CreateCell(4);
            gxcell.SetCellValue("电力线伏数");
            gxcell = gxrow.CreateCell(5);
            gxcell.SetCellValue("时间");
            gxcell = gxrow.CreateCell(6);
            gxcell.SetCellValue("备注");
            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    gxrow = sheet.CreateRow(j);
                    j++;

                    gxrow = sheet.CreateRow(i);
                    gxcell = gxrow.CreateCell(0);
                    gxcell.SetCellValue(dataRow["ID"].ToString());
                    gxcell = gxrow.CreateCell(1);
                    gxcell.SetCellValue(dataRow["LinkID"].ToString());
                    gxcell = gxrow.CreateCell(2);
                    gxcell.SetCellValue(dataRow["FTName"].ToString());
                    gxcell = gxrow.CreateCell(3);
                    gxcell.SetCellValue(dataRow["DLXZX"].ToString());
                    gxcell = gxrow.CreateCell(4);
                    gxcell.SetCellValue(dataRow["DLXFS"].ToString());
                    gxcell = gxrow.CreateCell(5);
                    gxcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    gxcell = gxrow.CreateCell(6);
                    gxcell.SetCellValue(dataRow["BZ"].ToString());
                }

            }
        }
        //植被要素写入excel
        public void CreateZBExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("植被") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(4);
            IRow zbrow = sheet.CreateRow(0);
            ICell zbcell = zbrow.CreateCell(0);
            zbcell.SetCellValue("ID");
            zbcell = zbrow.CreateCell(1);
            zbcell.SetCellValue("LinkID");
            zbcell = zbrow.CreateCell(2);
            zbcell.SetCellValue("要素名称");
            zbcell = zbrow.CreateCell(3);
            zbcell.SetCellValue("要素种类");
            zbcell = zbrow.CreateCell(4);
            zbcell.SetCellValue("所属林场");
            zbcell = zbrow.CreateCell(5);
            zbcell.SetCellValue("时间");
            zbcell = zbrow.CreateCell(6);
            zbcell.SetCellValue("备注");

            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    zbrow = sheet.CreateRow(j);
                    j++;
                    zbcell = zbrow.CreateCell(0);
                    zbcell.SetCellValue(dataRow["ID"].ToString());
                    zbcell = zbrow.CreateCell(1);
                    zbcell.SetCellValue(dataRow["LinkID"].ToString());
                    zbcell = zbrow.CreateCell(2);
                    zbcell.SetCellValue(dataRow["YSMC"].ToString());
                    zbcell = zbrow.CreateCell(3);
                    zbcell.SetCellValue(dataRow["YSZL"].ToString());
                    zbcell = zbrow.CreateCell(4);
                    zbcell.SetCellValue(dataRow["SSLC"].ToString());
                    zbcell = zbrow.CreateCell(5);
                    zbcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    zbcell = zbrow.CreateCell(6);
                    zbcell.SetCellValue(dataRow["BZ"].ToString());

                }
            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }
        }
        //土质地貌要素写入excel
        public void CreateTZDBExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("土质地貌") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(5);
            IRow dmrow = sheet.CreateRow(0);
            ICell dmcell = dmrow.CreateCell(0);
            dmcell.SetCellValue("ID");
            dmcell = dmrow.CreateCell(1);
            dmcell.SetCellValue("LinkID");
            dmcell = dmrow.CreateCell(2);
            dmcell.SetCellValue("地貌名称");
            dmcell = dmrow.CreateCell(3);
            dmcell.SetCellValue("时间");
            dmcell = dmrow.CreateCell(4);
            dmcell.SetCellValue("备注");
            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    dmrow = sheet.CreateRow(j);
                    j++;
                    dmcell = dmrow.CreateCell(0);
                    dmcell.SetCellValue(dataRow["ID"].ToString());
                    dmcell = dmrow.CreateCell(1);
                    dmcell.SetCellValue(dataRow["LinkID"].ToString());
                    dmcell = dmrow.CreateCell(2);
                    dmcell.SetCellValue(dataRow["DMMC"].ToString());
                    dmcell = dmrow.CreateCell(3);
                    dmcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    dmcell = dmrow.CreateCell(4);
                    dmcell.SetCellValue(dataRow["BZ"].ToString());
                }
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }
        }
        //地理注记写入excel
        public void CreateDLZJExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("地理注记") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(6);
            IRow dlzjrow = sheet.CreateRow(0);
            ICell dlzjcell = dlzjrow.CreateCell(0);
            dlzjcell.SetCellValue("ID");
            dlzjcell = dlzjrow.CreateCell(1);
            dlzjcell.SetCellValue("LinkID");
            dlzjcell = dlzjrow.CreateCell(2);
            dlzjcell.SetCellValue("要素名称");
            dlzjcell = dlzjrow.CreateCell(3);
            dlzjcell.SetCellValue("要素类型");
            dlzjcell = dlzjrow.CreateCell(4);
            dlzjcell.SetCellValue("时间");
            dlzjcell = dlzjrow.CreateCell(5);
            dlzjcell.SetCellValue("备注");
            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    dlzjrow = sheet.CreateRow(j);
                    j++;

                    dlzjcell = dlzjrow.CreateCell(0);
                    dlzjcell.SetCellValue(dataRow["ID"].ToString());
                    dlzjcell = dlzjrow.CreateCell(1);
                    dlzjcell.SetCellValue(dataRow["LinkID"].ToString());
                    dlzjcell = dlzjrow.CreateCell(2);
                    dlzjcell.SetCellValue(dataRow["YSName"].ToString());
                    dlzjcell = dlzjrow.CreateCell(3);
                    dlzjcell.SetCellValue(dataRow["YSType"].ToString());
                    dlzjcell = dlzjrow.CreateCell(4);
                    dlzjcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    dlzjcell = dlzjrow.CreateCell(5);
                    dlzjcell.SetCellValue(dataRow["BZ"].ToString());
                }
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }
        }
        //文字注记写入excel
        public void CreateWZZJExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("文字注记") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(7);
            IRow wzzjrow = sheet.CreateRow(0);
            ICell wzzjcell = wzzjrow.CreateCell(0);
            wzzjcell.SetCellValue("ID");
            wzzjcell = wzzjrow.CreateCell(1);
            wzzjcell.SetCellValue("LinkID");
            wzzjcell = wzzjrow.CreateCell(2);
            wzzjcell.SetCellValue("要素代码");
            wzzjcell = wzzjrow.CreateCell(3);
            wzzjcell.SetCellValue("要素名称");
            wzzjcell = wzzjrow.CreateCell(4);
            wzzjcell.SetCellValue("要素类型");
            wzzjcell = wzzjrow.CreateCell(5);
            wzzjcell.SetCellValue("时间");
            wzzjcell = wzzjrow.CreateCell(6);
            wzzjcell.SetCellValue("备注");

            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    wzzjrow = sheet.CreateRow(j);
                    j++;

                    wzzjcell = wzzjrow.CreateCell(0);
                    wzzjcell.SetCellValue(dataRow["ID"].ToString());
                    wzzjcell = wzzjrow.CreateCell(1);
                    wzzjcell.SetCellValue(dataRow["LinkID"].ToString());
                    wzzjcell = wzzjrow.CreateCell(2);
                    wzzjcell.SetCellValue(dataRow["YSDM"].ToString());
                    wzzjcell = wzzjrow.CreateCell(3);
                    wzzjcell.SetCellValue(dataRow["YSName"].ToString());
                    wzzjcell = wzzjrow.CreateCell(4);
                    wzzjcell.SetCellValue(dataRow["YSType"].ToString());
                    wzzjcell = wzzjrow.CreateCell(5);
                    wzzjcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    wzzjcell = wzzjrow.CreateCell(6);
                    wzzjcell.SetCellValue(dataRow["BZ"].ToString());
                }
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }
        }
        //境界线要素写入excel
        public void CreateJJXExcel(System.Data.DataTable dt, string path)
        {
            //HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("境界线") : workbook.CreateSheet(dt.TableName);
            HSSFWorkbook workbook = OpenExistExcel(path);
            ISheet sheet = workbook.GetSheetAt(8);
            IRow jjxrow = sheet.CreateRow(0);
            ICell jjxcell = jjxrow.CreateCell(0);
            jjxcell.SetCellValue("ID");
            jjxcell = jjxrow.CreateCell(1);
            jjxcell.SetCellValue("LinkID");
            jjxcell = jjxrow.CreateCell(2);
            jjxcell.SetCellValue("要素名称");
            jjxcell = jjxrow.CreateCell(3);
            jjxcell.SetCellValue("国界");
            jjxcell = jjxrow.CreateCell(4);
            jjxcell.SetCellValue("国内境界线");
            jjxcell = jjxrow.CreateCell(5);
            jjxcell.SetCellValue("时间");
            jjxcell = jjxrow.CreateCell(6);
            jjxcell.SetCellValue("备注");

            int j = 1;
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i - 1];
                if (dataRow["LinkID"].ToString() != "")
                {
                    jjxrow = sheet.CreateRow(j);
                    j++;

                    jjxcell = jjxrow.CreateCell(0);
                    jjxcell.SetCellValue(dataRow["ID"].ToString());
                    jjxcell = jjxrow.CreateCell(1);
                    jjxcell.SetCellValue(dataRow["LinkID"].ToString());
                    jjxcell = jjxrow.CreateCell(2);
                    jjxcell.SetCellValue(dataRow["FTName"].ToString());
                    jjxcell = jjxrow.CreateCell(3);
                    jjxcell.SetCellValue(dataRow["GJ"].ToString());
                    jjxcell = jjxrow.CreateCell(4);
                    jjxcell.SetCellValue(dataRow["NBJJX"].ToString());
                    jjxcell = jjxrow.CreateCell(5);
                    jjxcell.SetCellValue(dataRow["ZJTIME"].ToString());
                    jjxcell = jjxrow.CreateCell(6);
                    jjxcell.SetCellValue(dataRow["BZ"].ToString());
                }
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }
        }
        //创建总表
        public void CreateExcel(string path)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("") : workbook.CreateSheet(dt.TableName);//创建属性表       
            string[] featuresname = { "居民地", "道路", "水系", "管线", "植被", "土质地貌", "地理注记", "文字注记", "境界线" };


            for (int i = 0; i < featuresname.Length; i++)
            {
                string name = featuresname[i];
                ISheet sheet = workbook.CreateSheet();
                workbook.SetSheetName(i, name);
            }

            HSSFSheet JMDworksheet = (HSSFSheet)workbook.GetSheet("居民地");
            HSSFSheet DLworksheet = (HSSFSheet)workbook.GetSheet("道路");
            HSSFSheet SXworksheet = (HSSFSheet)workbook.GetSheet("水系");
            HSSFSheet GXworksheet = (HSSFSheet)workbook.GetSheet("管线");
            HSSFSheet ZBworksheet = (HSSFSheet)workbook.GetSheet("植被");
            HSSFSheet TZDMworksheet = (HSSFSheet)workbook.GetSheet("土质地貌");
            HSSFSheet DLZJworksheet = (HSSFSheet)workbook.GetSheet("地理注记");
            HSSFSheet WZZJworksheet = (HSSFSheet)workbook.GetSheet("文字注记");
            HSSFSheet JJXworksheet = (HSSFSheet)workbook.GetSheet("境界线");

            IRow jmdrow = JMDworksheet.CreateRow(0);
            ICell jmdcell = jmdrow.CreateCell(0);
            jmdcell.SetCellValue("ID");
            jmdcell = jmdrow.CreateCell(1);
            jmdcell.SetCellValue("LinkID");
            jmdcell = jmdrow.CreateCell(2);
            jmdcell.SetCellValue("要素名称");
            jmdcell = jmdrow.CreateCell(3);
            jmdcell.SetCellValue("房屋层数");
            jmdcell = jmdrow.CreateCell(4);
            jmdcell.SetCellValue("房屋材质");
            jmdcell = jmdrow.CreateCell(5);
            jmdcell.SetCellValue("房檐改正");
            jmdcell = jmdrow.CreateCell(6);
            jmdcell.SetCellValue("时间");
            jmdcell = jmdrow.CreateCell(7);
            jmdcell.SetCellValue("备注");

            IRow dlrow = DLworksheet.CreateRow(0);
            ICell dlcell = dlrow.CreateCell(0);
            dlcell.SetCellValue("ID");
            dlcell = dlrow.CreateCell(1);
            dlcell.SetCellValue("LinkID");
            dlcell = dlrow.CreateCell(2);
            dlcell.SetCellValue("道路名称");
            dlcell = dlrow.CreateCell(3);
            dlcell.SetCellValue("道路线号");
            dlcell = dlrow.CreateCell(4);
            dlcell.SetCellValue("道路等级");
            dlcell = dlrow.CreateCell(5);
            dlcell.SetCellValue("时间");
            dlcell = dlrow.CreateCell(6);
            dlcell.SetCellValue("备注");

            IRow sxrow = SXworksheet.CreateRow(0);
            ICell sxcell = sxrow.CreateCell(0);
            sxcell.SetCellValue("ID");
            sxcell = sxrow.CreateCell(1);
            sxcell.SetCellValue("LinkID");
            sxcell = sxrow.CreateCell(2);
            sxcell.SetCellValue("要素名称");
            sxcell = sxrow.CreateCell(3);
            sxcell.SetCellValue("附属设施");
            sxcell = sxrow.CreateCell(4);
            sxcell.SetCellValue("时间");
            sxcell = sxrow.CreateCell(5);
            sxcell.SetCellValue("备注");

            IRow zbrow = ZBworksheet.CreateRow(0);
            ICell zbcell = zbrow.CreateCell(0);
            zbcell.SetCellValue("ID");
            zbcell = zbrow.CreateCell(1);
            zbcell.SetCellValue("LinkID");
            zbcell = zbrow.CreateCell(2);
            zbcell.SetCellValue("要素名称");
            zbcell = zbrow.CreateCell(3);
            zbcell.SetCellValue("要素种类");
            zbcell = zbrow.CreateCell(4);
            zbcell.SetCellValue("所属林场");
            zbcell = zbrow.CreateCell(5);
            zbcell.SetCellValue("时间");
            zbcell = zbrow.CreateCell(6);
            zbcell.SetCellValue("备注");

            IRow gxrow = GXworksheet.CreateRow(0);
            ICell gxcell = gxrow.CreateCell(0);
            gxcell.SetCellValue("ID");
            gxcell = gxrow.CreateCell(1);
            gxcell.SetCellValue("LinkID");
            gxcell = gxrow.CreateCell(2);
            gxcell.SetCellValue("要素名称");
            gxcell = gxrow.CreateCell(3);
            gxcell.SetCellValue("电力线走向");
            gxcell = gxrow.CreateCell(4);
            gxcell.SetCellValue("电力线伏数");
            gxcell = gxrow.CreateCell(5);
            gxcell.SetCellValue("时间");
            gxcell = gxrow.CreateCell(6);
            gxcell.SetCellValue("备注");

            IRow jjxrow = JJXworksheet.CreateRow(0);
            ICell jjxcell = jjxrow.CreateCell(0);
            jjxcell.SetCellValue("ID");
            jjxcell = jjxrow.CreateCell(1);
            jjxcell.SetCellValue("LinkID");
            jjxcell = jjxrow.CreateCell(2);
            jjxcell.SetCellValue("要素名称");
            jjxcell = jjxrow.CreateCell(3);
            jjxcell.SetCellValue("国界");
            jjxcell = jjxrow.CreateCell(4);
            jjxcell.SetCellValue("国内境界线");
            jjxcell = jjxrow.CreateCell(5);
            jjxcell.SetCellValue("时间");
            jjxcell = jjxrow.CreateCell(6);
            jjxcell.SetCellValue("备注");

            IRow dmrow = TZDMworksheet.CreateRow(0);
            ICell dmcell = dmrow.CreateCell(0);
            dmcell.SetCellValue("ID");
            dmcell = dmrow.CreateCell(1);
            dmcell.SetCellValue("LinkID");
            dmcell = dmrow.CreateCell(2);
            dmcell.SetCellValue("地貌名称");
            dmcell = dmrow.CreateCell(3);
            dmcell.SetCellValue("时间");
            dmcell = dmrow.CreateCell(4);
            dmcell.SetCellValue("备注");
            dmcell = dmrow.CreateCell(5);

            IRow dlzjrow = DLZJworksheet.CreateRow(0);
            ICell dlzjcell = dlzjrow.CreateCell(0);
            dlzjcell.SetCellValue("ID");
            dlzjcell = dlzjrow.CreateCell(1);
            dlzjcell.SetCellValue("LinkID");
            dlzjcell = dlzjrow.CreateCell(2);
            dlzjcell.SetCellValue("要素名称");
            dlzjcell = dlzjrow.CreateCell(3);
            dlzjcell.SetCellValue("要素类型");
            dlzjcell = dlzjrow.CreateCell(4);
            dlzjcell.SetCellValue("时间");
            dlzjcell = dlzjrow.CreateCell(5);
            dlzjcell.SetCellValue("备注");

            IRow wzzjrow = WZZJworksheet.CreateRow(0);
            ICell wzzjcell = wzzjrow.CreateCell(0);
            wzzjcell.SetCellValue("ID");
            wzzjcell = wzzjrow.CreateCell(1);
            wzzjcell.SetCellValue("LinkID");
            wzzjcell = wzzjrow.CreateCell(2);
            wzzjcell.SetCellValue("要素代码");
            wzzjcell = wzzjrow.CreateCell(3);
            wzzjcell.SetCellValue("要素名称");
            wzzjcell = wzzjrow.CreateCell(4);
            wzzjcell.SetCellValue("要素类型");
            wzzjcell = wzzjrow.CreateCell(5);
            wzzjcell.SetCellValue("时间");
            wzzjcell = wzzjrow.CreateCell(6);
            wzzjcell.SetCellValue("备注");

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] bArray = ms.ToArray();
                fs.Write(bArray, 0, bArray.Length);
                fs.Flush();
            }

        }
        #endregion
        

        #region  没有用到导出excel
        //居民地导出excel
        public static bool SaveJJMToExcel(System.Data.DataTable excelTable, string filePath)
        {

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "", str7 = "", str8 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "FTName")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "FWCS")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "FWCZ")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "FYGZ")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str7 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str8 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "要素名称";
                worksheet.Cells[1, 4] = "房屋层数";
                worksheet.Cells[1, 5] = "房屋材质";
                worksheet.Cells[1, 6] = "房檐改正";
                worksheet.Cells[1, 7] = "时间";
                worksheet.Cells[1, 8] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        //道路
        public static bool SaveDLToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "", str7 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "DLMC")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "DLXH")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "DJDM")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str7 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                        if (str2 != "")
                        {
                            worksheet.Cells[i + 2, 1] = str1;
                            worksheet.Cells[i + 2, 2] = str2;
                            worksheet.Cells[i + 2, 3] = str3;
                            worksheet.Cells[i + 2, 4] = str4;
                            worksheet.Cells[i + 2, 5] = str5;
                            worksheet.Cells[i + 2, 6] = str6;
                            worksheet.Cells[i + 2, 7] = str7;
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "道路名称";
                worksheet.Cells[1, 4] = "道路线号";
                worksheet.Cells[1, 5] = "道路代码";
                worksheet.Cells[1, 6] = "时间";
                worksheet.Cells[1, 7] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }

        }
        //水系
        public static bool SaveSXToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSMC")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "FSSS")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "要素名称";
                worksheet.Cells[1, 4] = "附属设施";
                worksheet.Cells[1, 5] = "时间";
                worksheet.Cells[1, 6] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        //植被
        public static bool SaveZBToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "", str7 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSMC")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSZL")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "SSLC")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str7 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "要素代码";
                worksheet.Cells[1, 4] = "要素种类";
                worksheet.Cells[1, 5] = "所属林场";
                worksheet.Cells[1, 6] = "时间";
                worksheet.Cells[1, 7] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        //管线
        public static bool SaveGXToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "", str7 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "FTName")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "DLXZX")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "DLXFS")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str7 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "要素名称";
                worksheet.Cells[1, 4] = "电力线走向";
                worksheet.Cells[1, 5] = "电力线伏数";
                worksheet.Cells[1, 6] = "时间";
                worksheet.Cells[1, 7] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        //土质地貌
        public static bool SaveDMToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "DMMC")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "ZJTIME")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "地貌名称";
                worksheet.Cells[1, 4] = "时间";
                worksheet.Cells[1, 5] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        //地理注记
        public static bool SaveDLZJToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSName")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSType")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "要素名称";
                worksheet.Cells[1, 4] = "要素类型";
                worksheet.Cells[1, 5] = "时间";
                worksheet.Cells[1, 6] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        //文字注记
        public static bool SaveWZZJToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "", str7 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSDM")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSName")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "YSType")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str7 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "要素代码";
                worksheet.Cells[1, 4] = "要素名称";
                worksheet.Cells[1, 5] = "要素类型";
                worksheet.Cells[1, 6] = "时间";
                worksheet.Cells[1, 7] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        //境界线
        public static bool SaveJJXToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook workbook = app.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int colums = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        string str1 = "", str2 = "", str3 = "", str4 = "", str5 = "", str6 = "", str7 = "";
                        for (int j = 0; j < colums; j++)
                        {
                            string colName = excelTable.Columns[j].ColumnName;
                            if (excelTable.Rows[i][j] != null)
                            {
                                if (colName == "ID")
                                {
                                    str1 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "LinkID")
                                {
                                    str2 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "FTName")
                                {
                                    str3 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "GJ")
                                {
                                    str4 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "NBJJX")
                                {
                                    str5 = excelTable.Rows[i][j].ToString();
                                }

                                if (colName == "ZJTIME")
                                {
                                    str6 = excelTable.Rows[i][j].ToString();
                                }
                                if (colName == "BZ")
                                {
                                    str7 = excelTable.Rows[i][j].ToString();
                                }
                            }
                        }
                    }
                }
                worksheet.Cells[1, 1] = "ID";
                worksheet.Cells[1, 2] = "LinkID";
                worksheet.Cells[1, 3] = "要素名称";
                worksheet.Cells[1, 4] = "国界";
                worksheet.Cells[1, 5] = "国内界线";
                worksheet.Cells[1, 6] = "时间";
                worksheet.Cells[1, 7] = "备注";
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                workbook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示消息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
                MessageBox.Show("导出成功！");
            }
        }
        #endregion

        #region 空间数据导出为excel

        public void CreatreSHP(System.Data.DataTable dt,string filePath)
        {

        }
        #endregion


        //选择保存excel路径
        private void ChooseExcelPath_Click(object sender, EventArgs e)
        {
            SaveFileDialog savefileDialog = new SaveFileDialog();
            savefileDialog.Title = "保存文件";
            savefileDialog.Filter = "Excel文件（*.xls;*.xlsx)|*.xls;*xlsx|所有文件（*.*)|*.*";
            savefileDialog.RestoreDirectory = true;
            if (savefileDialog.ShowDialog() == DialogResult.OK)
            {
                saveFilePath = savefileDialog.FileName;
            }
            ExcelPath.Text = saveFilePath;
        }
    }
}
