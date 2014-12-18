using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;


namespace ConvertPinYin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            FileStream fs = new FileStream ( "F:\\1.txt" , FileMode.Open , FileAccess.Read ) ;
            StreamReader strR = new StreamReader(fs, Encoding.GetEncoding("gb2312")); //使用StreamReader类来读取文件 
            strR.BaseStream.Seek ( 0 , SeekOrigin.Begin ) ; // 从数据流中读取每一行，直到文件的最后一行，并在richTextBox1中显示出内容             
            string strLine = strR.ReadLine ( ) ;  

            while ( strLine != null ) 
            {
                string strPinYin = null;
                string strIniPinyin = null;
                char[] arr = strLine.ToCharArray();
                string[] strWordArr = new string[arr.Length];
                for (int i = 0; i < arr.Length; i++) strWordArr[i] = arr[i].ToString();

                foreach (string strSingle in strWordArr)
                {                    
                    strPinYin += Pinyin.Convert(strSingle);
                    strIniPinyin += IniPinYin.getSpell(strSingle);
                }

                string[] InsertItem = { strLine, strPinYin, strIniPinyin };
                ListViewItem lvi = new ListViewItem( InsertItem );
                listView1.Items.Insert(listView1.Items.Count,lvi);
                strLine = strR.ReadLine ( ) ; 
            } 
            strR.Close();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                object m_objOpt = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Workbooks ExcelBooks = (Microsoft.Office.Interop.Excel.Workbooks)ExcelApp.Workbooks;
                Microsoft.Office.Interop.Excel._Workbook ExcelBook = (Microsoft.Office.Interop.Excel._Workbook)(ExcelBooks.Add(m_objOpt));
                Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel._Worksheet)ExcelBook.ActiveSheet;

                ExcelApp.Caption = "Test";                //设置标题
                for (int i = 1; i <= listView1.Columns.Count; i++) ExcelSheet.Cells[2, i] = listView1.Columns[i - 1].Text;
                for (int i = 3; i < listView1.Items.Count + 3; i++)
                {
                    ExcelSheet.Cells[i, 1] = listView1.Items[i - 3].Text;
                    for (int j = 2; j <= listView1.Columns.Count; j++) ExcelSheet.Cells[i, j] = listView1.Items[i - 3].SubItems[j - 1].Text;
                }
                ExcelApp.Visible = true;
            }
            catch (SystemException sysE) { MessageBox.Show(sysE.ToString()); }          
        }

       
    }
}
