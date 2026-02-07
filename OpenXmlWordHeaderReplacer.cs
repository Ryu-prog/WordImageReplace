using System;
using System.IO;
using System.Linq;
using System.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordImageReplace
{
    public static class OpenXmlWordHeaderReplacer
    {
        const long EMU_PER_PIXEL = 9525;

        public static void ReplaceHeaderImages(
            string templatePath,
            string savePath,
            string frontImagePath,
            string backImagePath,
            bool changeFront,
            bool changeBack)
        {
            if (string.IsNullOrWhiteSpace(savePath))
            {
                throw new ArgumentException("The save path must not be null or empty.", nameof(savePath));
            }

            var saveDirectory = Path.GetDirectoryName(savePath);
            if (string.IsNullOrWhiteSpace(saveDirectory))
            {
                throw new ArgumentException("The save path must include a directory.", nameof(savePath));
            }

            if (!Directory.Exists(saveDirectory))
            {
                throw new DirectoryNotFoundException($"The destination directory does not exist: '{saveDirectory}'.");
            }

            // Check that we can write to the destination directory so we can provide a clear error message.
            try
            {
                var testFilePath = Path.Combine(saveDirectory, Path.GetRandomFileName());
                using (File.Create(testFilePath))
                {
                    // If this succeeds, we have write access; the file will be deleted immediately after.
                }
                File.Delete(testFilePath);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new UnauthorizedAccessException($"Cannot write to the destination directory: '{saveDirectory}'.", ex);
            }
            // 元ファイルを上書きしないようにコピーして編集します。
            // ファイルをUTF-8で保存し、コメントの文字化けを防止してください。
            File.Copy(templatePath, savePath, overwrite: true);

            using (var doc = WordprocessingDocument.Open(savePath, true))
            {
                var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart がありません。");

                // 全 SectionProperties を取得（各セクションのヘッダ参照を調べる）
                var sectionProps = mainPart.Document.Descendants<SectionProperties>().ToList();

                foreach (var sec in sectionProps)
                {
                    // HeaderReference を列挙
                    foreach (var headerRef in sec.Elements<HeaderReference>())
                    {
                        var type = headerRef.Type?.Value ?? HeaderFooterValues.Default;
                        var headerPart = (HeaderPart)mainPart.GetPartById(headerRef.Id.Value);

                        if ((type == HeaderFooterValues.Default && changeFront))
                        {
                            ReplaceOrAddImageInHeader(headerPart, frontImagePath);
                        }
                        else if ((type == HeaderFooterValues.Even && changeBack))
                        {
                            ReplaceOrAddImageInHeader(headerPart, backImagePath);
                        }
                        // First (first page) を扱いたい場合は HeaderFooterValues.First を追加してください
                    }
                }

                // 保存は using とともに行われる
            }
        }

        private static void ReplaceOrAddImageInHeader(HeaderPart headerPart, string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath)) return;

            // 画像バイト
            using var imgStream = File.OpenRead(imagePath);

            // 既存の image part があれば最初のものを上書きする
            var existingImagePart = headerPart.ImageParts.FirstOrDefault();
            if (existingImagePart != null)
            {
                using var partStream = existingImagePart.GetStream(FileMode.Create, FileAccess.Write);
                imgStream.Position = 0;
                imgStream.CopyTo(partStream);
                return;
            }

            // 既存画像がない場合は新しい image part を作り、ヘッダー本文に Drawing を挿入する
            var ext = Path.GetExtension(imagePath).ToLowerInvariant();
            var ipt = ext switch
            {
                ".png" => ImagePartType.Png,
                ".jpg" => ImagePartType.Jpeg,
                ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                _ => ImagePartType.Png
            };

            var imagePart = headerPart.AddImagePart(ipt);
            imgStream.Position = 0;
            imagePart.FeedData(imgStream);

            string rId = headerPart.GetIdOfPart(imagePart);

            // 画像のピクセルサイズを取得して EMU に変換
            long cx = 0, cy = 0;
            using (var img = Image.FromFile(imagePath))
            {
                cx = (long)(img.Width * EMU_PER_PIXEL);
                cy = (long)(img.Height * EMU_PER_PIXEL);
            }

            // Drawing 要素を作成（Microsoft のサンプルを簡略化）
            var element = CreateImageDrawing(rId, cx, cy);

            // ヘッダーのルート要素を用意
            if (headerPart.Header == null)
            {
                headerPart.Header = new Header();
            }

            // 既存内容を削除して画像だけにする場合は下の行をアンコメントする
            // headerPart.Header.RemoveAllChildren();

            // 段落->実行->drawing で追加
            var paragraph = new Paragraph(new Run(element));
            headerPart.Header.Append(paragraph);
            headerPart.Header.Save();
        }

        private static Drawing CreateImageDrawing(string relationshipId, long cx, long cy)
        {
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = cx, Cy = cy },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)BitConverter.ToUInt32(Guid.NewGuid().ToByteArray(), 0),
                            Name = "Picture " + relationshipId
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)BitConverter.ToUInt32(Guid.NewGuid().ToByteArray(), 0),
                                            Name = "Image " + relationshipId
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip() { Embed = relationshipId },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = cx, Cy = cy }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                                )
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        // behind の配置（必要ならさらに wrap を調整）
                    });

            return element;
        }
    }
}