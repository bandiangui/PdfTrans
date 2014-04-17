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
    public class Word
    {
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
        public void WordInsertWatermark(ref Aspose.Words.Document doc, WordWatermark wwm)
        {
            var a = wwm.GetVerticalAlignment();
            InsertWatermarkText(doc, wwm);
        }

        public class WordWatermark
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
            /// 加粗
            /// </summary>
            public bool FontBold { get; set; }
            /// <summary>
            /// 倾斜
            /// </summary>
            public bool FontItalic { get; set; }
            /// <summary>
            /// 旋转角度
            /// </summary>
            private double _rotation = 315;
            public double Rotation
            {
                get { return _rotation; }
                set { _rotation = value; }
            }
            /// <summary>
            /// 填充颜色
            /// </summary>
            private Color _fillColor = Color.LightGray;
            public Color FillColor
            {
                get { return _fillColor; }
                set { _fillColor = value; }
            }
            /// <summary>
            /// 填充透明度
            /// </summary>
            private double _fillOpacity = 0.5f;
            public double FillOpacity
            {
                get { return _fillOpacity; }
                set { _fillOpacity = value; }
            }
            /// <summary>
            /// 线条颜色
            /// </summary>
            public Color _lineColor = Color.Empty;
            public Color LineColor
            {
                get { return _lineColor; }
                set { _lineColor = value; }
            }
            /// <summary>
            /// 线条透明度
            /// </summary>
            public double LineOpacity { get; set; }
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

            public WordWatermark()
            {
            }
            public WordWatermark(string text)
            {
                this.Text = text;
            }

            public VerticalAlignment GetVerticalAlignment()
            {
                return (VerticalAlignment)Enum.Parse(typeof(VerticalAlignment), this._verticalAlign.ToString());
            }
            public HorizontalAlignment GetHorizontalAlignment()
            {
                return (HorizontalAlignment)Enum.Parse(typeof(AlignEnum), this._align.ToString());
            }

            /// <summary>
            /// 垂直对齐
            /// </summary>
            public enum VerticalAlignEnum
            {
                Default = 0,
                Top = 1,
                Center = 2,
                Bottom = 3
            }

            /// <summary>
            /// 水平对齐
            /// </summary>
            public enum AlignEnum
            {
                Default = 0,
                Left = 1,
                Center = 2,
                Right = 3
            }
        }


        /// <summary>
        /// 插入水印
        /// </summary>
        /// <param name="doc">Word Document</param>
        /// <param name="watermarkText">水印文字</param>
        private void InsertWatermarkText(Aspose.Words.Document doc, WordWatermark wwm)
        {
            // 创建word艺术字
            Shape watermark = new Shape(doc, ShapeType.TextPlainText);

            watermark.TextPath.Text = wwm.Text;
            watermark.TextPath.FontFamily = wwm.FontFamily;
            watermark.Width = 500;
            watermark.Height = 100;

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
            Paragraph watermarkPara = new Paragraph(doc);
            watermarkPara.AppendChild(watermark);

            // 在页眉中插入水印(word中的水印是在页眉中的)
            foreach (Section sect in doc.Sections)
            {
                //在各种页眉中插入水印
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
            }
        }

        private void InsertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, HeaderFooterType headerType)
        {
            HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // 如果word文档没有页眉,创建一个
                header = new HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // 插入水印
            header.AppendChild(watermarkPara.Clone(true));
        }
    }
}