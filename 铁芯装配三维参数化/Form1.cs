using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;

namespace 铁芯装配三维参数化
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "请选择计算单文件";
            openFileDialog1.Filter = "*.xlsx|*.xlsx";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string str = openFileDialog1.FileName;
                MessageBox.Show(str);
                ExcelUtility excelUtility = new ExcelUtility();
                excelUtility.LoadData(str);
  
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            {
                ISldWorks swApp = SwUtility.ConnectToSolidWorks();

                if (swApp != null)
                {
                    string msg = "This message from C#. solidworks version is " + swApp.RevisionNumber();
                    swApp.SendMsgToUser(msg);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }
    }
}
