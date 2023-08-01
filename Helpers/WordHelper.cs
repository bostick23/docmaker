using System;
using System.IO;
using System.IO.Packaging;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocMaker.Helpers
{
    public class WordHelper
    {
        private readonly string FILE_PATH;
        public WordHelper(string _filePath)
        {
            FILE_PATH = _filePath;
        }
        public string CreateAndInsertAllPictures()
        {
            string wordFileName = System.IO.Path.Combine(FILE_PATH, "note.docx");
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(wordFileName, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                string[] files = Directory.GetFiles(FILE_PATH, $"ScreenCapture*.png");
                if (files != null && files.Length > 0)
                {
                    for (int i = 0; i < files.Length; i++)
                    {
                        var img = new BitmapImage();
                        using (var fs = new FileStream(files[i], FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            img.BeginInit();
                            img.StreamSource = fs;
                            img.EndInit();
                        }
                        int maxWidthCm = 15;
                        var widthPx = img.PixelWidth;
                        var heightPx = img.PixelHeight;
                        var horzRezDpi = img.DpiX;
                        var vertRezDpi = img.DpiY;
                        const int emusPerInch = 914400;
                        const int emusPerCm = 360000;
                        var widthEmus = (long)(widthPx / horzRezDpi * emusPerInch);
                        long heightEmus = (long)(heightPx / vertRezDpi * emusPerInch);
                        long maxWidthEmus = (long)(maxWidthCm * emusPerCm);
                        if (widthEmus > maxWidthEmus)
                        {
                            var ratio = (heightEmus * 1.0m) / widthEmus;
                            widthEmus = maxWidthEmus;
                            heightEmus = (long)(widthEmus * ratio);
                        }
                        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                        using (FileStream stream = new FileStream(files[i], FileMode.Open))
                        {
                            imagePart.FeedData(stream);
                        }
                        wordDocument.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text("Step n. " + (i + 1)))));

                        AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart), widthEmus, heightEmus);
                        File.Delete(files[i]);
                    }
                }
            }
            return wordFileName;
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, long widthEmus, long heightEmus)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         {
                                             Preset = A.ShapeTypeValues.Rectangle
                                         }))
                             )
                             {
                                 Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                             })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}
