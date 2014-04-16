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
using Aspose.Pdf.Devices;
using Aspose.Pdf.Generator;

namespace WordToPDF.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Message = "Modify this template to jump-start your ASP.NET MVC application.";


            Aspose.Words.Document doc = new Aspose.Words.Document(@"C:\Users\芸芸\Desktop\2013年工作总结.doc");

            Pdf pdf = new Pdf();
            Aspose.Pdf.Generator.Section section = pdf.Sections.Add();

            using (FileStream fs = new FileStream(@"E:\2\" + Guid.NewGuid().ToString() + "___1.pdf", FileMode.Create))
            {
                doc.Save(fs, Aspose.Words.SaveFormat.Pdf);

                // 把pdf转换为图片,并加水印
                Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(fs);
                var docPageInfo = doc.GetPageInfo(1);
                section.PageInfo.PageWidth = docPageInfo.WidthInPoints;
                section.PageInfo.PageHeight = docPageInfo.HeightInPoints;
                section.PageInfo.Margin = new Aspose.Pdf.Generator.MarginInfo();

                int imgWith = (int)docPageInfo.WidthInPoints / 72 * 300;
                int imgHeight = (int)docPageInfo.HeightInPoints / 72 * 300;

                for (int pageCount = 1; pageCount <= pdfDoc.Pages.Count; pageCount++)
                {
                    Resolution resolution = new Resolution(300);

                    JpegDevice jpgBuilder = new JpegDevice(imgWith, imgHeight, resolution, 100);
                    using (MemoryStream jpgMs = new MemoryStream())
                    {
                        jpgBuilder.Process(pdfDoc.Pages[pageCount], jpgMs);

                        System.Drawing.Image img = System.Drawing.Image.FromStream(jpgMs);
                        ImgAddWaterMark(ref img);

                        Aspose.Pdf.Generator.Image pdfImg = Aspose.Pdf.Generator.Image.FromSystemImage(img);

                        section.Paragraphs.Add(pdfImg);
                    }

                }
            }
            pdf.Save(@"E:\2\" + Guid.NewGuid().ToString() + ".pdf");
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

            return View();
        }

        public void ImgAddWaterMark(ref System.Drawing.Image image)
        {
            System.Drawing.Graphics graph = System.Drawing.Graphics.FromImage(image);
            graph.DrawImage(image, 0, 0, image.Width, image.Height);

            graph.SmoothingMode = SmoothingMode.HighQuality;
            graph.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graph.CompositingQuality = CompositingQuality.HighQuality;

            using (MemoryStream ms = new MemoryStream())
            {
                WordToPDF.Models.Watermark.WaterImage wt = new WordToPDF.Models.Watermark.WaterImage(image.Width, image.Height);
                wt.FontSize = 200;
                wt.Create();
                graph.DrawImage(wt.ResultImage, new System.Drawing.Rectangle(0, 0, wt.ResultImage.Width, wt.ResultImage.Height)
                    , 0, 0, wt.ResultImage.Width, wt.ResultImage.Height, GraphicsUnit.Pixel, wt.SetTransparency(wt.Transparency));
            }

            graph.Dispose();
        }
    }
}
