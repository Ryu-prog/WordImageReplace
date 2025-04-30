using Microsoft.Win32;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;
using System.ComponentModel;
using System.IO;
using System.Linq.Expressions;
using System.Runtime.InteropServices;

namespace WordImageReplace
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();

            if (!String.IsNullOrEmpty(Properties.Settings.Default.TemplatePath)){
                this.tbWordFilePath.Text = Properties.Settings.Default.TemplatePath;
            }

            if (!String.IsNullOrEmpty(Properties.Settings.Default.FrontFilePath))
            {
                this.tbFrontFilePath.Text = Properties.Settings.Default.FrontFilePath;
            }

            if (!String.IsNullOrEmpty(Properties.Settings.Default.BackFilePath))
            {
                this.tbBackFilePath.Text = Properties.Settings.Default.BackFilePath;
            }

            if (!String.IsNullOrEmpty(Properties.Settings.Default.Password))
            {
                this.pbWordFile.Password = Properties.Settings.Default.Password;
            }

            Closing += FormMain_FormClosing;
        }

        /// <summary>
        /// アプリ終了時のプロパティ保存
        /// </summary>
        private void FormMain_FormClosing(object sender, CancelEventArgs e)
        {
            Properties.Settings.Default.TemplatePath = this.tbWordFilePath.Text;
            Properties.Settings.Default.FrontFilePath = this.tbFrontFilePath.Text;
            Properties.Settings.Default.BackFilePath = this.tbBackFilePath.Text;
            Properties.Settings.Default.Password = this.pbWordFile.Password;
            Properties.Settings.Default.Save();

        }


        private void Replace_CommandButton_Click(object sender, RoutedEventArgs e)
        {
            bool isChangeFront = Convert.ToBoolean(this.cbFrontSide.IsChecked);
            bool isChangeBack = Convert.ToBoolean(this.cbBackSide.IsChecked);

            string templatePath = this.tbWordFilePath.Text;
            string frontFilePath = this.tbFrontFilePath.Text;
            string backFilePath = this.tbBackFilePath.Text;
            string wdPassword = this.pbWordFile.Password;

            //入力欄チェック
            if (File.Exists(templatePath) == false) {
                System.Windows.MessageBox.Show(templatePath + "が見つかりませんでした。\n対象のファイルが存在するか確認してください。"
                    , "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (File.Exists(frontFilePath) == false) {
                System.Windows.MessageBox.Show(frontFilePath + "が見つかりませんでした。\n対象のファイルが存在するか確認してください。"
                    , "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (isChangeBack && File.Exists(backFilePath) == false)
            {
                System.Windows.MessageBox.Show(backFilePath + "が見つかりませんでした。\n対象のファイルが存在するか確認してください。"
                    , "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }



            WordController wd = new WordController();
            try
            {
                wd.OpenDocument(templatePath,true, wdPassword);

                //画像の左端位置を設定
                float left = 0;
                //画像の上端位置を設定
                float top = 0;
                //画像の横幅を設定
                float width = wd.wdPageSetup.PageWidth;
                //画像の縦幅を設定
                float height = wd.wdPageSetup.PageHeight;


                //全セクションに対して画像を差し替え
                //（セクションが分けられている場合を考慮しない）
                for (int secNum=1; secNum <= wd.wdSecCnt; secNum++) {
                    if (isChangeFront)
                    {
                        wd.HeaderPictureChange(secNum, frontFilePath, left, top, width, height, WdHeaderFooterIndex.wdHeaderFooterPrimary);
                        wd.SetSharpRange(secNum, WdHeaderFooterIndex.wdHeaderFooterPrimary);
                    }

                    if (isChangeBack)
                    {
                        wd.HeaderPictureChange(secNum, backFilePath, left, top, width, height, WdHeaderFooterIndex.wdHeaderFooterEvenPages);
                        wd.SetSharpRange(secNum, WdHeaderFooterIndex.wdHeaderFooterEvenPages);
                    }
                }
                //保存先選択ダイアログを表示
                OpenFileDialog opd = new OpenFileDialog();

                opd.Filter = "Wordファイル(*.docx)|*.docx";

                opd.FileName = templatePath;

                if ((bool)opd.ShowDialog())
                {
                    wd.SaveDoc(opd.FileName);
                }


                //保存完了メッセージ
                System.Windows.MessageBox.Show("Wordファイルのヘッダー画像を差し替えました。"
                        , "正常終了", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("エラーが発生しました" + Environment.NewLine + ex.ToString(), "エラー");
            }



        }

        private string FilePathSelect(string FilePath, string ExtensionLabel, string Extension)
        {


            return null;
        }

        private void WordFileSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opd = new OpenFileDialog();

            opd.Filter = "Wordファイル(*.docx)|*.docx";

            if ((bool)opd.ShowDialog())
            {
                this.tbWordFilePath.Text = opd.FileName;
            }
        }

        private void FrontSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opd = new OpenFileDialog();

            opd.Filter = "画像ファイル|*.png;*.jp*g";

            if ((bool)opd.ShowDialog())
            {
                this.tbFrontFilePath.Text = opd.FileName;
            }
        }

        private void BackSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opd = new OpenFileDialog();

            opd.Filter = "画像ファイル|*.png;*.jp*g";

            if ((bool)opd.ShowDialog())
            {
                this.tbBackFilePath.Text = opd.FileName;
            }
        }
    }
}