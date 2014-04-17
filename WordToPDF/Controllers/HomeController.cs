using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WordToPDF.Models;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.IO;
//using Aspose.Pdf.Generator;
using Aspose.Words;
using Aspose.Pdf;
using Aspose.Pdf.Generator;

namespace WordToPDF.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Message = "Modify this template to jump-start your ASP.NET MVC application.";


            //Aspose.Words.Document doc = new Aspose.Words.Document(@"C:\Users\芸芸\Desktop\菠萝咕咾肉的做法.docx");
            //Word word = new Word();
            //word.WordInsertWatermark(ref doc, new Word.WordWatermark("内部资料"));
            //doc.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".pdf", Aspose.Words.SaveFormat.Pdf);

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your app description page.";
            // 从图片到PDF
            //Pdf pdf = new Pdf();
            //Aspose.Pdf.Generator.Section section = pdf.Sections.Add();

            //System.Drawing.Image imge = System.Drawing.Image.FromFile(@"C:\Users\芸芸\Desktop\cb71b07fc15cbcc3e587e1e83bc751c7.jpg");
            //ImgAddWaterMark(ref imge);

            //Aspose.Pdf.Generator.Image img = Aspose.Pdf.Generator.Image.FromSystemImage(imge);
            //img.ImageInfo.ImageFileType = ImageFileType.Jpeg;
            //section.PageInfo.PageWidth = imge.Width / imge.HorizontalResolution * 72;
            //section.PageInfo.PageHeight = imge.Height / imge.HorizontalResolution * 72;
            //section.PageInfo.Margin = new MarginInfo();

            //pdf.Watermarks.Add(fb);


            //for (int i = 0; i < 10; i++)
            //{
            //    section.Paragraphs.Add(img);
            //}
            //pdf.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".pdf");
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            try
            {
                //Pdf pdf1 = new Pdf();
                //Aspose.Pdf.Generator.Section sec1 = pdf1.Sections.Add();

                //Aspose.Pdf.Generator.FloatingBox wk = new Aspose.Pdf.Generator.FloatingBox(500, 500);
                //wk.BoxHorizontalAlignment = BoxHorizontalAlignmentType.Center;
                //wk.BoxVerticalAlignment = BoxVerticalAlignmentType.Center;
                //wk.BoxHorizontalPositioning = BoxHorizontalPositioningType.Page;
                //wk.BoxVerticalPositioning = BoxVerticalPositioningType.Page;

                //Aspose.Pdf.Generator.Text tt = new Aspose.Pdf.Generator.Text("这里是水印");
                //tt.TextInfo.Color = new Aspose.Pdf.Generator.Color("gray");
                //tt.TextInfo.Alignment = AlignmentType.Center;
                //tt.TextInfo.FontSize = 40.0f;
                //tt.Opacity = 0.5f;
                //wk.Paragraphs.Add(tt);

                //pdf1.Watermarks.Add(wk);
                //pdf1.IsWatermarkOnTop = true;
                //pdf1.Sections[0].Paragraphs.Add(new Text(" Aspose.Pdf is a .NET Component built to ease the job of developers to create PDF documents ranging from simple to complex on the fly programmatically. Aspose.Pdf allows developers to insert Tables, Graphs, Images, Hyperlinks and Custom Fonts etc. in the PDF documents. Moreover, it is also possible to compress PDF documents. Aspose.Pdf provides excellent security features to develop secure PDF documents. And the most distinct feature of Aspose.Pdf is that it supports the creation of PDF documents through both an API and from XML templates Aspose.Pdf is a .NET Component built to ease the job of developers to create PDF documents ranging from simple to complex on the fly programmatically. Aspose.Pdf allows developers to insert Tables, Graphs, Images, Hyperlinks and Custom Fonts etc. in the PDF documents. Moreover, it is also possible to compress PDF documents. Aspose.Pdf provides excellent security features to develop secure PDF documents. And the most distinct feature of Aspose.Pdf is that it supports the creation of PDF documents through both an API and from XML templates Aspose.Pdf is a .NET Component built to ease the job of developers to create PDF documents ranging from simple to complex on the fly programmatically. Aspose.Pdf allows developers to insert Tables, Graphs, Images, Hyperlinks and Custom Fonts etc. in the PDF documents. Moreover, it is also possible to compress PDF documents. Aspose.Pdf provides excellent security features to develop secure PDF documents. And the most distinct feature of Aspose.Pdf is that it supports the creation of PDF documents through both an API and from XML templates "));
                //pdf1.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".pdf");

                Pdf pdf1 = new Pdf();
                Aspose.Pdf.Generator.Section sec1 = pdf1.Sections.Add();

                //Text text = new Text("这里是水印");
                //text.TextInfo.Alignment = AlignmentType.Center;
                //text.TextInfo.FontSize = 40.0f;
                //text.Opacity = 0.3f;
                //text.RotatingAngle = 45.0f;
                //text.TextInfo.Color = new Aspose.Pdf.Generator.Color("gray");

                //Aspose.Pdf.Generator.FloatingBox watermark = new Aspose.Pdf.Generator.FloatingBox(500, 200);
                //watermark.BoxVerticalPositioning = BoxVerticalPositioningType.Margin;
                //watermark.BoxVerticalAlignment = BoxVerticalAlignmentType.Top;
                //watermark.ZIndex = -1;
                //watermark.BoxHorizontalPositioning = BoxHorizontalPositioningType.Margin;
                //watermark.Paragraphs.Add(text);
                //pdf1.Watermarks.Add(watermark);

                //pdf1.IsWatermarkOnTop = true;

                PDF.PDFWatermark pwm = new PDF.PDFWatermark("这里是水印");
                PDF pdf = new PDF();
                pdf.InsertWatermark(ref pdf1, pwm);

                var img = new Aspose.Pdf.Generator.Image();
                img.ImageInfo.File = @"C:\Users\芸芸\Desktop\cb71b07fc15cbcc3e587e1e83bc751c7.jpg";
                pdf1.Sections[0].Paragraphs.Add(img);
                //pdf1.Sections[0].Paragraphs.Add(new Text("111"));
                pdf1.Save(@"E:\2\" + Guid.NewGuid().ToString() + ".pdf");
            }
            catch (Exception e)
            {
                throw;
            }
            return View();
        }

        public void ImgAddWaterMark(System.Drawing.Image waterImg, ImageAttributes imgAttr, ref System.Drawing.Image image)
        {
            System.Drawing.Graphics graph = System.Drawing.Graphics.FromImage(image);
            graph.DrawImage(image, 0, 0, image.Width, image.Height);

            graph.SmoothingMode = SmoothingMode.HighQuality;
            graph.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graph.CompositingQuality = CompositingQuality.HighQuality;

            using (MemoryStream ms = new MemoryStream())
            {
                //WordToPDF.Models.Watermark.WaterImage wt = new WordToPDF.Models.Watermark.WaterImage(image.Width, image.Height);
                //wt.FontSize = 200;
                //wt.Create();
                graph.DrawImage(waterImg, new System.Drawing.Rectangle(0, 0, waterImg.Width, waterImg.Height)
                    , 0, 0, waterImg.Width, waterImg.Height, GraphicsUnit.Pixel, imgAttr);
            }

            graph.Dispose();
        }
    }
}
