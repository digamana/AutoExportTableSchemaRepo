using AutoExportTableSchema.Domain;
using AutoExportTableSchemaDll.Domain;
using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

namespace AutoExportTableSchema
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadConfig();
            Center center = new Center();
        }
        public void LoadConfig()
        {
            txtbServeName.Text = Properties.Settings.Default["ServerName"].ToString();
            txtbDbName.Text = Properties.Settings.Default["DBName"].ToString();
            txtbAccount.Text = Properties.Settings.Default["Account"].ToString();
            txtbPwd.Password = Properties.Settings.Default["Pwd"].ToString();
            txtbConnectString.Text = Properties.Settings.Default["ConnectString"].ToString();
            var temp = Properties.Settings.Default["ExistColumnDescription"].ToString();
            chkDescription.IsEnabled = Convert.ToBoolean(temp);

        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Center.CreatTempelteExcel();
        }
 

        private void Window_Closed(object sender, EventArgs e)
        {
            Properties.Settings.Default["ServerName"] = txtbServeName.Text;
            Properties.Settings.Default["DBName"] = txtbDbName.Text;
            Properties.Settings.Default["Account"] =txtbAccount.Text ;
            Properties.Settings.Default["Pwd"] = txtbPwd.Password;
            Properties.Settings.Default["ConnectString"] = txtbConnectString.Text;
            Properties.Settings.Default["ExistColumnDescription"] = chkDescription.IsEnabled;

            Properties.Settings.Default.Save(); // Saves settings in application configuration file
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (rdoBasic.IsChecked == false && rdoConnectString.IsChecked == false && rdoFile.IsChecked == false)
            {
                MessageBox.Show("請選擇項目");
                return;
            }

            if (rdoBasic.IsChecked == true)
            {
                if (string.IsNullOrEmpty(txtbAccount.Text) || string.IsNullOrEmpty(txtbDbName.Text) || string.IsNullOrEmpty(txtbServeName.Text) || string.IsNullOrEmpty(txtbPwd.Password))
                {
                    MessageBox.Show("伺服器名稱、資料庫名稱、帳號、密碼請勿空白");
                    return;
                }
                Center center = new Center();
                string ServerName = txtbServeName.Text;
                string DbName = txtbDbName.Text;
                string AccName = txtbAccount.Text;
                string Pwd = txtbPwd.Password;
                string ConnectString = $"data source={ServerName};initial catalog={DbName};persist security info=True;user id={AccName};password={Pwd};MultipleActiveResultSets=True;App=EntityFramework providerName = System.Data.SqlClient";
                if (!Center.SqlConnectionTest(ConnectString))
                {
                    MessageBox.Show("資料庫連線失敗,請確認輸入資訊");
                    return;
                }
                center.Run(ConnectString, txtbDbName.Text,(bool)chkDescription.IsChecked);

            }
            else if (rdoConnectString.IsChecked == true)
            {
                if (string.IsNullOrEmpty(txtbConnectString.Text))
                {
                    MessageBox.Show("請輸入ConnectString");
                    return;
                }
                Center center = new Center();
                if (!Center.SqlConnectionTest(txtbConnectString.Text))
                {
                    MessageBox.Show("資料庫連線失敗,請確認輸入資訊");
                    return;
                }
                center.Run(txtbConnectString.Text, "MIDDB", (bool)chkDescription.IsChecked);
            }
            else if(rdoFile.IsChecked==true)
            {
                if (string.IsNullOrEmpty(txtbFilePath.Text) || !txtbFilePath.Text.Contains("xlsx"))
                {
                    MessageBox.Show("請輸入xlsx的檔案路徑");
                    return;
                }
                Center center = new Center();
                var lst = center.ReadTempleteExecl(txtbFilePath.Text);
 
                foreach (var item in lst)
                {
                    string ServerName = $"{item.ServerName}";
                    string DbName = $"{item.DbName}";
                    string AccName = $"{item.Account}";
                    string Pwd = $"{item.Password}";
                    string ConnectString = $"data source={ServerName};initial catalog={DbName};persist security info=True;user id={AccName};password={Pwd};MultipleActiveResultSets=True;App=EntityFramework providerName = System.Data.SqlClient";
                    if (Center.SqlConnectionTest(ConnectString))
                    {
                        center.Run(ConnectString, DbName, (bool)chkDescription.IsChecked);
                    } 
                }
            }

        }

        private void btnImportTemplete_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = openFileDlg.ShowDialog();
            if (result == true) // Test result.
            {
                txtbFilePath.Text = openFileDlg.FileName;
            }
        }

        private void btnSourceFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = openFileDlg.ShowDialog();
            if (result == true) // Test result.
            {
                txtbSourceFilePath.Text = openFileDlg.FileName;
            }
        }

        private void btnTargetFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = openFileDlg.ShowDialog();
            if (result == true) // Test result.
            {
                txtbTargetFilePath.Text = openFileDlg.FileName;
            }
        }

        private void btnMapping_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtbTargetFilePath.Text) || string.IsNullOrEmpty(txtbSourceFilePath.Text))
            {
                MessageBox.Show("請輸入xlsx的檔案路徑");
                return;
            }
            Mapping mapping = new Mapping(txtbSourceFilePath.Text, txtbTargetFilePath.Text);
            mapping.Run();

            Center.openFile(txtbTargetFilePath.Text);
        }

        private void btnTargetFile_CreatCommand_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            Nullable<bool> result = openFileDlg.ShowDialog();
            if (result == true) // Test result.
            {
                txtbTargetFilePath_CreatCommand.Text = openFileDlg.FileName;
            }
        }

        private void btnMapping_CreatCommand_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtbTargetFilePath_CreatCommand.Text) || string.IsNullOrEmpty(txtbTargetFilePath_CreatCommand.Text))
            {
                MessageBox.Show("請輸入xlsx的檔案路徑");
                return;
            }
            /**
             * 下面開始編寫讀取Excel裡面的描述欄位,並組成SQL寫入欄位指令的方法
             */
            Reading reading = new Reading(txtbTargetFilePath_CreatCommand.Text);
            reading.Run();
            //Mapping mapping = new Mapping(txtbSourceFilePath.Text, txtbTargetFilePath.Text);
            //mapping.Run();

            //Center.openFile(txtbTargetFilePath.Text);
        }
    }
}
