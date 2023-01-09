using AutoExportTableSchema.Domain;
using AutoExportTableSchemaDll.Domain;
using System;
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
            Center center = new Center();
            Mapping mapping = new Mapping("", "");
            mapping.Run();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Center.CreatTempelteExcel();
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
                center.Run(ConnectString, txtbDbName.Text);

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
                center.Run(txtbConnectString.Text, "MIDDB");
            }
            else if(rdoFile.IsChecked==true)
            {
                if (string.IsNullOrEmpty(txtbFilePath.Text) || !txtbFilePath.Text.Contains("xlsx"))
                {
                    MessageBox.Show("請輸入xlsm的檔案路徑");
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
                        center.Run(ConnectString, DbName);
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
    }
}
