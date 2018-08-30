using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Net.Mail;

namespace BLF.ActivationCodeCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        XSSFWorkbook xssfworkbook;
        public MainWindow()
        {
            InitializeComponent();
        }

        void InitializeWorkbook(string path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                xssfworkbook = new XSSFWorkbook(file);
            }
        }

        System.Data.DataTable InitializedtSalesEngineerDataTable()
        {

            System.Data.DataTable dtSalesEngineer = new System.Data.DataTable();
            //dtAccount.TableName = Constants.AccountTable;
            dtSalesEngineer.Columns.Add("SalesEngineer");
            dtSalesEngineer.Columns.Add("Account");
            dtSalesEngineer.Columns.Add("Password");
            dtSalesEngineer.Columns.Add("ActivationCode");
            return dtSalesEngineer;
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = this.filePath.Text.Trim();
                if (string.IsNullOrEmpty(filePath))
                {
                    this.Error.Text = "Please input excel file path...";
                }
                else
                {
                    string sheetName = Constants.Activation;
                    string xmlPath = CreateXml(sheetName, filePath);

                    //CreateActivationCode(sheetName, xmlPath);
                    MessageBox.Show("create successfully..", "Create", MessageBoxButton.OK);
                }
            }
            catch (Exception ex)
            {
                this.Error.Text = ex.Message;
            }
            
        }

        /**   
       * 字符串转换成十六进制字符串  
       * @param String str 待转换的ASCII字符串  
       * @return String 每个Byte之间空格分隔，如: [61 6C 6B]  
       */
        private String str2HexStr(String str)
        {

            char[] chars = "0123456789ABCDEF".ToCharArray();
            StringBuilder sb = new StringBuilder("");
            byte[] bs = Encoding.ASCII.GetBytes(str);
            int bit;

            for (int i = 0; i < bs.Length; i++)
            {
                bit = (bs[i] & 0x0f0) >> 4;
                sb.Append(chars[bit]);
                bit = bs[i] & 0x0f;
                sb.Append(chars[bit]);
            }
            return sb.ToString().Trim().Replace('0', 'G').Replace('1', 'H').Replace('2', 'I').Replace('3', 'J').Replace('4', 'K').Replace('5', 'L').Replace('6', 'M').Replace('7', 'N').Replace('8', 'O').Replace('9', 'P');
        }

        //private void CreateActivationCode(string tableName, string filePath)
        //{
        //    XmlDataDocument xmlDoc = new XmlDataDocument();
        //    SqlCommand cmd = new SqlCommand();

        //    try
        //    {
        //        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        //        {
        //            if (fs != null)
        //            {
        //                xmlDoc.Load(fs);
        //                XmlNodeList NodeList = xmlDoc.GetElementsByTagName("Item");
        //                for (int i = 0; i < NodeList.Count; i++)
        //                {
        //                    //string ID = NodeList[i];
        //                    cmd = SetParameters(tableName, NodeList[i]);
        //                    ExecuteInsertSQL(cmd);
        //                    if (!string.IsNullOrEmpty(this.Error.Text.ToString()))
        //                    {
        //                        continue;
        //                    }
        //                }
        //                if (string.IsNullOrEmpty(this.Error.Text.ToString()))
        //                {
        //                    MessageBox.Show("upload successfully..", "Upload Data", MessageBoxButton.OK);
        //                }
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        this.Error.Text = ex.Message;
        //    }
        //    finally
        //    {
        //        cmd.Dispose();
        //    }
        //}
       

        private string CreateXml(string tableName, string path)
        {
            System.Data.DataTable dt = null;
            string sheetName = Constants.Activation;
            dt = InitializedtSalesEngineerDataTable();           
            if (dt != null)
            {
                //'Item' is tagName in xml
                dt.TableName = "Item";
            }

            InitializeWorkbook(path);
            ISheet sheet = xssfworkbook.GetSheet(sheetName);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();



            if (dt == null)
            {
                return null;
            }
            //start at line 2
            rows.MoveNext();

            while (rows.MoveNext())
            {
                IRow row = (XSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }
                if (dr[0].ToString() != "")
                {
                    dr[0] = dr[0].ToString().Trim().ToLower();
                    dr[3] = str2HexStr(dr[0].ToString().Trim().ToLower() + " " + dr[1].ToString().Trim() + " " + dr[2].ToString().Trim());
                    dt.Rows.Add(dr);
                }

            }

            string xmlPath = @"E:\ActivationCodeCreator\" + tableName + DateTime.Now.ToString("yyyyMMdd") + ".xml";
            //create xml file to local
            dt.WriteXml(xmlPath);
            dt.Dispose();
            return xmlPath;

        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "excel files (*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                this.filePath.Text = dlg.FileName;
            }
        }
    }

     
}
