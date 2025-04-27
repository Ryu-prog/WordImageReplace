using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;

namespace WordImageReplace
{
        public class WordController
        {

            private Microsoft.Office.Interop.Word.Application wordApplication = null;
            private Microsoft.Office.Interop.Word.Documents wordDocuments = null;
            private Microsoft.Office.Interop.Word.Document wordDocument = null;


            public void OpenDocument(string filePath)
            {
                Application wordApp = new Application();
                this.wordDocuments = wordApp.Documents;

                object fPathObj = filePath;

                // Word 文書を開く
                this.wordDocument = wordApp.Documents.Open(ref fPathObj);

                // Word アプリケーションを画面に表示する
                wordApp.Visible = true;
            }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
            public void UnprotectPassword(string password = "")
            {

                if (this.wordDocument.ProtectionType != WdProtectionType.wdNoProtection)
                {

                    this.wordDocument.Unprotect(password);
                }
            }

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
            /// 
            /// </summary>
            /// <param name="headerIndex"></param>
            /// <param name="imagePath"></param>
            public void HeaderPictureChange(WdHeaderFooterIndex headerIndex, string imagePath)
            {

            ////裏面画像が存在する場合、奇数ページと偶数ページのヘッダーを分けるをON
            //if (headerIndex == WdHeaderFooterIndex.wdHeaderFooterEvenPages)
            //{

            //    foreach (Section section in this.wordDocument.Sections)
            //    {
            //        section.PageSetup.OddAndEvenPagesHeaderFooter = 1;
            //        //this.wordDocument.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = (int)WdHeaderFooterIndex.wdHeaderFooterEvenPages;


            //    }
            //}
                PageSetup pageSetup = this.wordDocument.PageSetup;

                double Left = 0;
                double Top = 0;
                double Width = pageSetup.PageWidth;
                double Height = pageSetup.PageHeight;

                foreach (Section sec in this.wordDocument.Sections)
                {

                    HeaderFooter header = sec.Headers[headerIndex];

                    header.Range.Delete();
                    header.Shapes.AddPicture(@imagePath, false, true, Left, Top, Width, Height);
                    header.Range.ShapeRange.WrapFormat.Type = WdWrapType.wdWrapBehind;

                    ShapeRange shapeRange = sec.Headers[headerIndex].Range.ShapeRange;

                    shapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                    shapeRange.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;

                    shapeRange.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;

                    shapeRange.Left = 0;

                }

            }

            public void SaveDoc(string saveFilePath)
            {
                this.wordDocument.SaveAs(FileName: saveFilePath);
            }

        }
}
