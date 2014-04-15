using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Aspose.Pdf.Generator;
using WordToPDF.Models;
using System.Drawing;
using System.Drawing.Imaging;
//using Aspose.Words;
//using Aspose.Pdf;

namespace WordToPDF.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Message = "Modify this template to jump-start your ASP.NET MVC application.";
            //Document doc = new Document(@"E:\工作文档\模板属性.docx");
            //doc.Save(@"E:\工作文档\1\" + Guid.NewGuid().ToString() + ".pdf", SaveFormat.Pdf);

            Pdf pdf = new Pdf();
            Aspose.Pdf.Generator.Section section = pdf.Sections.Add();

            System.Drawing.Image imge = System.Drawing.Image.FromFile(@"C:\Users\芸芸\Desktop\cb71b07fc15cbcc3e587e1e83bc751c7.jpg");
            AddWaterMark(ref imge);

            Aspose.Pdf.Generator.Image img = Aspose.Pdf.Generator.Image.FromSystemImage(imge);
            img.ImageInfo.ImageFileType = ImageFileType.Jpeg;
            section.PageInfo.PageWidth = imge.Width / imge.HorizontalResolution * 72;
            section.PageInfo.PageHeight = imge.Height / imge.HorizontalResolution * 72;
            section.PageInfo.Margin = new MarginInfo();

            //FloatingBox fb = new FloatingBox(700, 800);
            //TextInfo ti = new TextInfo();
            //ti.Color = new Color("#0000ff");
            //ti.FontSize = 24;
            //ti.IsTrueTypeFontBold = true;
            //fb.Paragraphs.Add(new Text("这里是水印", ti));

            //pdf.Watermarks.Add(fb);


            for (int i = 0; i < 10; i++)
            {
                section.Paragraphs.Add(img);
            }
            pdf.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".pdf");
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your app description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public void AddWaterMark(ref System.Drawing.Image image) {
            System.Drawing.Graphics graph = System.Drawing.Graphics.FromImage(image);
            graph.DrawImage(image, 0,0, image.Width, image.Height);
            Watermark.WaterImage wt = new Watermark.WaterImage();
            wt.Create();

            wt.ResultImage.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".png", ImageFormat.Png);
            graph.DrawImage(wt.ResultImage, new PointF(0, 0));
            image.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".png", ImageFormat.Png);
        }
    }
}
