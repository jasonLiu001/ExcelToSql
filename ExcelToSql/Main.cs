using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelToJson;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Formatting = System.Xml.Formatting;

namespace ExcelToSql
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void btn_broswerFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Title = "选择Excel文件";
            fileDialog.Filter = "All Files(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string file = fileDialog.FileName;
                this.txt_filepath.Text = file.Trim();
            }
        }

        private void btn_convert_Click(object sender, EventArgs e)
        {
            var excelFilePath = this.txt_filepath.Text.Trim();
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("请选择Excel文件路径");
                return;
            }

            var excelDataTable = NPOIUtility.GetDataFromExcel(excelFilePath);
            var sqlData = new StringBuilder();
            for (var i = 0; i < excelDataTable.Rows.Count; i++)
            {
                var userId = int.Parse(excelDataTable.Rows[i]["UserId"].ToString());
                var publicTestId = int.Parse(excelDataTable.Rows[i]["PublicTestId"].ToString());
                sqlData.Append("insert into [PublicTestUser]([UserId],[PublicTestId],[CityId]) values(");
                sqlData.Append(userId);
                sqlData.Append(",");
                sqlData.Append(publicTestId);
                sqlData.Append(",");
                sqlData.Append("201);");
            }
                

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.CreatePrompt = true;
            saveDialog.OverwritePrompt = true;

            saveDialog.DefaultExt = "sql";
            saveDialog.Filter = "Text Files(*.sql)|*.sql";

            if (DialogResult.OK == saveDialog.ShowDialog())
            {
                if (File.Exists(saveDialog.FileName))
                {
                    DialogResult result = MessageBox.Show("是否覆盖？", "确定", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        StreamWriter swLog = new StreamWriter(saveDialog.FileName);
                        swLog.WriteLine(sqlData.ToString());
                        swLog.WriteLine();
                        swLog.Close();
                    }
                    if (result == DialogResult.No)
                    {
                        return;
                    }
                }
                else
                {
                    StreamWriter swLog = new StreamWriter(saveDialog.FileName);
                    swLog.WriteLine(sqlData.ToString());
                    swLog.WriteLine();
                    swLog.Close();
                }
            }
        }
    }
}
