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
            Properties.Settings.Default.Save();

        }


        private void Replace_CommandButton_Click(object sender, RoutedEventArgs e)
        {
            bool isChangeFront = Convert.ToBoolean(this.cbFrontSide.IsChecked);
            bool isChangeBack = Convert.ToBoolean(this.cbBackSide.IsChecked);

            string templatePath = this.tbWordFilePath.Text;
            string frontFilePath = this.tbFrontFilePath.Text;
            string backFilePath = this.tbBackFilePath.Text;

            // 入力欄チェック（既存のチェックはそのまま）
            if (!File.Exists(templatePath)) {
                System.Windows.MessageBox.Show(templatePath + "が見つかりませんでした。\n対象のファイルが存在するか確認してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (!File.Exists(frontFilePath)) {
                System.Windows.MessageBox.Show(frontFilePath + "が見つかりませんでした。\n対象のファイルが存在するか確認してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (isChangeBack && !File.Exists(backFilePath)) {
                System.Windows.MessageBox.Show(backFilePath + "が見つかりませんでした。\n対象のファイルが存在するか確認してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // 保存先選択ダイアログ
                var sfd = new Microsoft.Win32.SaveFileDialog();
                sfd.Filter = "Wordファイル(*.docx)|*.docx";
                sfd.FileName = System.IO.Path.GetFileName(templatePath);

                if (sfd.ShowDialog() == true)
                {
                    string savePath = sfd.FileName;
                    if (System.IO.Path.GetFullPath(savePath).Equals(System.IO.Path.GetFullPath(templatePath), StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show("保存先に元ファイルを選択できません。別名を選択してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    try
                    {
                        OpenXmlWordHeaderReplacer.ReplaceHeaderImages(templatePath, savePath, frontFilePath, backFilePath, isChangeFront, isChangeBack);
                        MessageBox.Show("Wordファイルのヘッダー画像を差し替えました。", "正常終了", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (IOException ioEx)
                    {
                        MessageBox.Show("ファイルにアクセスできません。他のアプリで開かれている可能性があります。\n" + ioEx.Message, "ファイル使用中", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("エラーが発生しました" + Environment.NewLine + ex.ToString(), "エラー");
            }
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

                // 選択されたファイルのパスからURIを作成
                Uri fileUri = new Uri(opd.FileName);

                // BitmapImageを作成してImageコントロールにセット
                FrontImage.Source = new BitmapImage(fileUri);
            }
        }

        private void BackSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opd = new OpenFileDialog();

            opd.Filter = "画像ファイル|*.png;*.jp*g";

            if ((bool)opd.ShowDialog())
            {
                this.tbBackFilePath.Text = opd.FileName;

                // 選択されたファイルのパスからURIを作成
                Uri fileUri = new Uri(opd.FileName);

                // BitmapImageを作成してImageコントロールにセット
                BackImage.Source = new BitmapImage(fileUri);
            }
        }
    }
}