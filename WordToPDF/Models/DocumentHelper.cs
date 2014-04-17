using Aspose.Pdf.Generator;
using Aspose.Words;
using Aspose.Words.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;

namespace WordToPDF.Models
{
    public class DocumentHelper
    {
        public class Watermark
        {
            #region 属性 及 变量
            public string Text { get; set; }
            /// <summary>
            /// 字体
            /// </summary>
            private string _fontFamily = "宋体";
            public string FontFamily
            {
                get { return _fontFamily; }
                set { _fontFamily = value; }
            }
            /// <summary>
            /// 填充颜色
            /// </summary>
            private System.Drawing.Color _fillColor = System.Drawing.Color.LightGray;
            public System.Drawing.Color FillColor
            {
                get { return _fillColor; }
                set { _fillColor = value; }
            }
            /// <summary>
            /// 线条颜色
            /// </summary>
            private System.Drawing.Color _lineColor = System.Drawing.Color.Empty;
            public System.Drawing.Color LineColor
            {
                get { return _lineColor; }
                set { _lineColor = value; }
            }
            /// <summary>
            /// 加粗
            /// </summary>
            public bool FontBold { get; set; }
            /// <summary>
            /// 倾斜
            /// </summary>
            public bool FontItalic { get; set; }

            private float _width = 500;
            public float Width
            {
                get { return _width; }
                set { _width = value; }
            }

            private float _height = 100;
            public float Height
            {
                get { return _height; }
                set { _height = value; }
            }

            /// <summary>
            /// 旋转角度
            /// </summary>
            private float _rotation = 315;
            public float Rotation
            {
                get { return _rotation; }
                set { _rotation = value; }
            }
            /// <summary>
            /// 填充透明度
            /// </summary>
            private float _fillOpacity = 0.5f;
            public float FillOpacity
            {
                get { return _fillOpacity; }
                set { _fillOpacity = value; }
            }
            /// <summary>
            /// 线条透明度
            /// </summary>
            public float LineOpacity { get; set; }
            /// <summary>
            /// 垂直对齐
            /// </summary>
            private VerticalAlignEnum _verticalAlign = VerticalAlignEnum.Center;
            public VerticalAlignEnum VerticalAlign
            {
                get { return _verticalAlign; }
                set { _verticalAlign = value; }
            }
            /// <summary>
            /// 水平对齐
            /// </summary>
            private AlignEnum _align = AlignEnum.Center;
            public AlignEnum Align
            {
                get { return _align; }
                set { _align = value; }
            }
            #endregion

            public Watermark()
            {
            }
            public Watermark(string text)
            {
                this.Text = text;
            }

            /// <summary>
            /// 垂直对齐
            /// </summary>
            public enum VerticalAlignEnum
            {
                None = 0,
                Top = 1,
                Center = 2,
                Bottom = 3
            }

            /// <summary>
            /// 水平对齐
            /// </summary>
            public enum AlignEnum
            {
                None = 0,
                Left = 1,
                Center = 2,
                Right = 3
            }
        }
    }
    public class Word
    {
        public class WordWatermark : DocumentHelper.Watermark
        {
            #region 属性 及 变量
            #           endregion

            public WordWatermark() : base() { }
            public WordWatermark(string text) : base(text) { }

            public VerticalAlignment GetVerticalAlignment()
            {
                return (VerticalAlignment)Enum.Parse(typeof(VerticalAlignment), VerticalAlign.ToString());
            }
            public HorizontalAlignment GetHorizontalAlignment()
            {
                return (HorizontalAlignment)Enum.Parse(typeof(AlignEnum), Align.ToString());
            }
        }

        /// <summary>
        /// Word转PDF
        /// </summary>
        /// <param name="st">word数据流</param>
        /// <param name="pdfPath">pdf数据流</param>
        public void WordToPDF(Stream st, ref FileStream pdfSt)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(st);
            doc.Save(pdfSt, Aspose.Words.SaveFormat.Pdf);
        }

        /// <summary>
        /// Word中添加水印
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="watermark"></param>
        public void InsertWatermark(ref Aspose.Words.Document doc, WordWatermark wwm)
        {
            var a = wwm.GetVerticalAlignment();
            InsertTextWatermark(doc, wwm);
        }

        #region 水印操作

        /// <summary>
        /// 插入水印
        /// </summary>
        /// <param name="doc">Word Document</param>
        /// <param name="watermarkText">水印文字</param>
        private void InsertTextWatermark(Aspose.Words.Document doc, WordWatermark wwm)
        {
            // 创建word艺术字
            Aspose.Words.Drawing.Shape watermark = new Aspose.Words.Drawing.Shape(doc, ShapeType.TextPlainText);

            watermark.TextPath.Text = wwm.Text;
            watermark.TextPath.FontFamily = wwm.FontFamily;
            watermark.Width = wwm.Width;
            watermark.Height = wwm.Height;

            //旋转角度
            watermark.Rotation = wwm.Rotation;

            watermark.Font.Bold = wwm.FontBold;     //文字加粗
            watermark.Font.Italic = wwm.FontItalic;   //文字倾斜

            // 设置填充颜色,透明度
            watermark.Fill.Color = wwm.FillColor;
            watermark.Fill.Opacity = wwm.FillOpacity;

            // 线条颜色,透明度
            watermark.StrokeColor = wwm.LineColor;
            watermark.Stroke.Opacity = wwm.LineOpacity;

            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;     //环绕方式 浮于文字上方
            watermark.VerticalAlignment = VerticalAlignment.Center;     // 垂直对齐方式
            watermark.HorizontalAlignment = HorizontalAlignment.Center;     //水平对齐方式

            // 把水印图片插入到一个新段落中(word中任何东西都是以段落方式存在的)
            Aspose.Words.Paragraph watermarkPara = new Aspose.Words.Paragraph(doc);
            watermarkPara.AppendChild(watermark);

            // 在页眉中插入水印(word中的水印是在页眉中的)
            foreach (Aspose.Words.Section sect in doc.Sections)
            {
                //在各种页眉中插入水印
                InsertWatermarkIntoHeader(watermarkPara, sect, Aspose.Words.HeaderFooterType.HeaderPrimary);
                InsertWatermarkIntoHeader(watermarkPara, sect, Aspose.Words.HeaderFooterType.HeaderFirst);
                InsertWatermarkIntoHeader(watermarkPara, sect, Aspose.Words.HeaderFooterType.HeaderEven);
            }
        }

        private void InsertWatermarkIntoHeader(Aspose.Words.Paragraph watermarkPara, Aspose.Words.Section sect, Aspose.Words.HeaderFooterType headerType)
        {
            Aspose.Words.HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // 如果word文档没有页眉,创建一个
                header = new Aspose.Words.HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // 插入水印
            header.AppendChild(watermarkPara.Clone(true));
        }
        #endregion

    }

    public class PDF
    {
        public class PDFWatermark : DocumentHelper.Watermark
        {
            #region 属性 及 变量
            private float _fontSize = 50.0f;
            public float FontSize
            {
                get { return _fontSize; }
                set { _fontSize = value; }
            }
            /// <summary>
            /// 文本对齐方式
            /// </summary>
            private AlignmentType _fontAlign = AlignmentType.Center;
            public AlignmentType FontAlign
            {
                get { return _fontAlign; }
                set { _fontAlign = value; }
            }
            /// <summary>
            /// 水印是在最顶层
            /// </summary>
            public bool _watermarkOnTop = true;
            public bool WatermarkOnTop
            {
                get { return _watermarkOnTop; }
                set { _watermarkOnTop = value; }
            }
            #endregion

            public PDFWatermark() : base() { }
            public PDFWatermark(string text) : base(text) { }
            public BoxVerticalAlignmentType GetBoxVerticalAlignment()
            {
                return (BoxVerticalAlignmentType)Enum.Parse(typeof(BoxVerticalAlignmentType), VerticalAlign.ToString());
            }
            public BoxHorizontalAlignmentType GetBoxHorizontalAlignment()
            {
                return (BoxHorizontalAlignmentType)Enum.Parse(typeof(BoxHorizontalAlignmentType), Align.ToString());
            }
        }
        public void InsertWatermark(ref Aspose.Pdf.Generator.Pdf pdf, PDFWatermark pwm)
        {
            InsertTextWatermark(pdf, pwm);
        }

        private void InsertTextWatermark(Aspose.Pdf.Generator.Pdf pdf, PDFWatermark pwm)
        {

            pdf.IsWatermarkOnTop = pwm.WatermarkOnTop;

            Image.ImageBuilderFromText imageBuilder = new Image.ImageBuilderFromText((int)pwm.Width, (int)pwm.Height);
            imageBuilder.Text = pwm.Text;
            imageBuilder.FontSize = pwm.FontSize;
            imageBuilder.FontColor = pwm.FillColor;
            //imageBuilder.Transparency = pwm.FillOpacity;
            imageBuilder.Rotation = pwm.Rotation;
            imageBuilder.Adaptable = false;
            imageBuilder.Tiling = false;

            imageBuilder.Create();
            imageBuilder.ResultImage.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".png", System.Drawing.Imaging.ImageFormat.Png);
            Aspose.Pdf.Generator.Image image = Aspose.Pdf.Generator.Image.FromSystemImage(imageBuilder.ResultImage);
            //image.Opacity = pwm.FillOpacity;
            image.ImageInfo.ImageFileType = ImageFileType.Png;

            // 创建水印盒子
            FloatingBox watermark = new FloatingBox(pwm.Width, pwm.Height);
            watermark.BoxVerticalPositioning = BoxVerticalPositioningType.Margin;
            watermark.BoxVerticalAlignment = pwm.GetBoxVerticalAlignment();         //垂直对齐
            watermark.BoxHorizontalPositioning = BoxHorizontalPositioningType.Margin;
            watermark.BoxHorizontalAlignment = pwm.GetBoxHorizontalAlignment();     //水平对齐
            watermark.ZIndex = -1;

            watermark.Paragraphs.Add(image);
            pdf.Watermarks.Add(watermark);
        }
    }
}