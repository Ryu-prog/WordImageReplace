using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using Microsoft.Office;

namespace WordImageReplace
{
    public class WordController
    {

        private Microsoft.Office.Interop.Word.Application wordApplication = null;
        private Microsoft.Office.Interop.Word.Documents wordDocuments = null;
        private Microsoft.Office.Interop.Word.Document wordDocument = null;

        public Microsoft.Office.Interop.Word.Sections wdSections { get { return this.wordDocument.Sections; } }
        public Microsoft.Office.Interop.Word.PageSetup wdPageSetup { get { return this.wordDocument.PageSetup; } }    

        public void OpenDocument(string filePath, bool isVisible = false, string passWord = "")
        {
            this.wordApplication = new Application();
            this.wordDocuments = this.wordApplication.Documents;

            object fPathObj = filePath;

            object passWordObj = passWord;

            // Word 文書を開く
            if (passWord == "")
            {
                this.wordDocument = this.wordApplication.Documents.Open(ref fPathObj);
            }
            else
            {
                this.wordDocument = this.wordApplication.Documents.Open(ref fPathObj, PasswordDocument: ref passWordObj);
            }
            // Word アプリケーションを画面に表示するか設定
            this.wordApplication.Visible = isVisible;
        }

        /// <summary>
        /// 編集の制限パスワードを解除
        /// </summary>
        /// <returns></returns>
        public void UnprotectPassword(string password = "")
        {

            if (this.wordDocument.ProtectionType != WdProtectionType.wdNoProtection)
            {

                this.wordDocument.Unprotect(password);
            }
        }

        /// <summary>
        /// 編集の制限を有効化
        /// </summary>
        /// <param name="password"></param>
        public void ProtectPassword(string password = "")
        {

            if (password != "")
            {
                this.wordDocument.Protect(WdProtectionType.wdAllowOnlyReading, false, password);
            }
            else
            {
                this.wordDocument.Protect(WdProtectionType.wdAllowOnlyReading, false);
            }
        }


        /// <summary>
        /// ヘッダーの画像差し替えを行う
        /// </summary>
        /// <param name="sec"></param>
        /// <param name="imagePath">画像ファイルの位置</param>
        /// <param name="left">画像の左端位置</param>
        /// <param name="top">画像の上端位置</param>
        /// <param name="width">画像の横幅</param>
        /// <param name="height">画像の縦幅</param>
        /// <param name="headerIndex">
        /// wdHeaderFooterPrimary:奇数ページの画像差し替え（奇数ページと偶数ページのヘッダーが分かれていない場合全ページ）
        /// wdHeaderFooterEvenPages:偶数ページの画像差し替え
        /// </param>
        public void HeaderPictureChange(Section sec, string imagePath, double left, double top, double width, double height, WdHeaderFooterIndex headerIndex = WdHeaderFooterIndex.wdHeaderFooterPrimary)
        {
            //奇数ページの場合、奇数ページと偶数ページのヘッダーを分けるをON
            if (headerIndex == WdHeaderFooterIndex.wdHeaderFooterEvenPages)
            {
                //ヘッダーの奇数ページと偶数ページを分ける
                //-1でTrue,0でFalse
                this.wordDocument.PageSetup.OddAndEvenPagesHeaderFooter = -1;
            }

            //差し替え後の画像位置の設定を行う。
            object leftObj = left;
            object topObj = top;
            object widthObj = width;
            object heightObj = height;

            //全セクションに対して画像を差し替え
            //（セクションが分けられている場合を考慮しない）
            HeaderFooter header = sec.Headers[headerIndex];

            //既に存在する画像を削除
            header.Range.Delete();

            //画像を追加
            header.Shapes.AddPicture(@imagePath, Left: leftObj, Top: topObj, Width: widthObj, Height: heightObj);

            //図形を背景に設定
            header.Range.ShapeRange.WrapFormat.Type = WdWrapType.wdWrapBehind;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sec"></param>
        /// <param name="headerIndex"></param>
        /// <param name="state"></param>
        /// <param name="wdRVP">文書内の図形範囲の相対位置を指定（垂直方向）</param>
        /// <param name="wdRHP">図形範囲の相対位置を指定（平行方向）</param>
        /// <param name="left">/図形範囲の相対位置（左）</param>
        public void SetSharpRange(
            Section sec,
            WdHeaderFooterIndex headerIndex,
            Microsoft.Office.Core.MsoTriState state = Microsoft.Office.Core.MsoTriState.msoTrue,
            WdRelativeVerticalPosition wdRVP = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage,
            WdRelativeHorizontalPosition wdRHP = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage,
            float left = 0
            )
        {
            ShapeRange shapeRange = sec.Headers[headerIndex].Range.ShapeRange;

            //図形のロックを設定
            shapeRange.LockAspectRatio = state;

            ///図形範囲の設定
            //文書内の図形範囲の相対位置を指定（垂直方向）
            shapeRange.RelativeVerticalPosition = wdRVP;

            //図形範囲の相対位置を指定（平行方向）
            shapeRange.RelativeHorizontalPosition = wdRHP;

            //図形範囲を左詰め
            shapeRange.Left = left;

        }

        ///// <summary>
        ///// ヘッダーの画像差し替えを行う
        ///// </summary>
        ///// <param name="headerIndex">
        ///// wdHeaderFooterPrimary:奇数ページの画像差し替え（奇数ページと偶数ページのヘッダーが分かれていない場合全ページ）
        ///// wdHeaderFooterEvenPages:偶数ページの画像差し替え
        ///// </param>
        ///// <param name="imagePath">差し替え後の画像ファイルパス</param>
        //public void HeaderPictureChange(WdHeaderFooterIndex headerIndex, string imagePath)
        //{
        //    //裏面画像が存在する場合、奇数ページと偶数ページのヘッダーを分けるをON
        //    if (headerIndex == WdHeaderFooterIndex.wdHeaderFooterEvenPages)
        //    {
        //        //ヘッダーの奇数ページと偶数ページを分ける
        //        //-1でTrue,0でFalse
        //        this.wordDocument.PageSetup.OddAndEvenPagesHeaderFooter = -1;
        //    }

        //    //差し替え後の画像位置の設定を行う。

        //    //画像の左端位置を設定
        //    double left = 0;
        //    //画像の上端位置を設定
        //    double top = 0;

        //    PageSetup pageSetup = this.wordDocument.PageSetup;

        //    //画像の横幅を設定
        //    double width = pageSetup.PageWidth;

        //    //画像の縦幅を設定
        //    double height = pageSetup.PageHeight;

        //    object leftObj = left;
        //    object topObj = top;
        //    object widthObj = width;
        //    object heightObj = height;

        //    //全セクションに対して画像を差し替え
        //    //（セクションが分けられている場合を考慮しない）
        //    foreach (Section sec in this.wordDocument.Sections)
        //    {
        //        HeaderFooter header = sec.Headers[headerIndex];

        //        //既に存在する画像を削除
        //        header.Range.Delete();

        //        //画像を追加
        //        header.Shapes.AddPicture(@imagePath, Left: leftObj, Top: topObj, Width: widthObj, Height: heightObj);

        //        //図形を背景に設定
        //        header.Range.ShapeRange.WrapFormat.Type = WdWrapType.wdWrapBehind;

        //        ShapeRange shapeRange = sec.Headers[headerIndex].Range.ShapeRange;

        //        //図形のロックを設定
        //        shapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

        //        ///図形範囲の設定
        //        //文書内の図形範囲の相対位置を指定（垂直方向）
        //        shapeRange.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;

        //        //図形範囲の相対位置を指定（平行方向）
        //        shapeRange.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;

        //        //図形範囲を左詰め
        //        shapeRange.Left = 0;

        //    }

        //}

        public void SaveDoc(string saveFilePath)
        {
            this.wordDocument.SaveAs(FileName: saveFilePath);
        }

    }
}
