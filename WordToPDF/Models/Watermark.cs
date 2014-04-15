using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Web;

namespace WordToPDF.Models
{
    public class Watermark
    {
        /**/
        /// <summary>
        /// WaterMark
        /// </summary>
        public class WaterImage
        {
            private int _width;
            private int _height;
            private string _fontFamily;
            private int _fontSize;
            private bool _adaptable;
            private FontStyle _fontStyle;
            private bool _shadow;
            private Stream _backgroundImage;
            private Color _bgColor;
            private bool _noBg;
            private int _left;
            private Image _resultImage;
            private string _text;
            private int _top;
            private int _BgAlpha;
            private int _red;
            private int _green;
            private int _blue;
            private int _incline;

            public WaterImage()
            {
                //
                // TODO: Add constructor logic here
                //
                _width = 460;
                _height = 30;
                _fontFamily = "华文行楷";
                _fontSize = 20;
                _fontStyle = FontStyle.Regular;
                _adaptable = true;
                _shadow = false;
                _left = 0;
                _top = 0;
                _BgAlpha = 100;
                _red = 0;
                _green = 0;
                _blue = 0;
                _incline = 45;
                _text = "这里是水印";
                _backgroundImage = null;
                _noBg = true;
                _bgColor = Color.FromArgb(_BgAlpha, 255, 255, 255);
            }

            /**/
            /// <summary>
            /// 字体
            /// </summary>
            public string FontFamily
            {
                set { this._fontFamily = value; }
            }

            /**/
            /// <summary>
            /// 文字大小
            /// </summary>
            public int FontSize
            {
                set { this._fontSize = value; }
            }

            /**/
            /// <summary>
            /// 文字风格
            /// </summary>
            public FontStyle FontStyle
            {
                get { return _fontStyle; }
                set { _fontStyle = value; }
            }

            /**/
            /// <summary>
            /// 透明度0-255,255表示不透明
            /// </summary>
            public int BgAlpha
            {
                get { return _BgAlpha; }
                set { _BgAlpha = value; }
            }

            /**/
            /// <summary>
            /// 水印文字是否使用阴影
            /// </summary>
            public bool Shadow
            {
                get { return _shadow; }
                set { _shadow = value; }
            }

            public int Red
            {
                get { return _red; }
                set { _red = value; }
            }

            public int Green
            {
                get { return _green; }
                set { _green = value; }
            }

            public int Blue
            {
                get { return _blue; }
                set { _blue = value; }
            }

            /**/
            /// <summary>
            /// 底图
            /// </summary>
            public Stream BackgroundImage
            {
                set { this._backgroundImage = value; }
            }

            /**/
            /// <summary>
            /// 水印文字的左边距
            /// </summary>
            public int Left
            {
                set { this._left = value; }
            }


            /**/
            /// <summary>
            /// 水印文字的顶边距
            /// </summary>
            public int Top
            {
                set { this._top = value; }
            }

            /**/
            /// <summary>
            /// 生成后的图片
            /// </summary>
            public Image ResultImage
            {
                set { this._resultImage = value; }
                get { return _resultImage; }
            }

            /**/
            /// <summary>
            /// 水印文本
            /// </summary>
            public string Text
            {
                set { this._text = value; }
            }

            /**/
            /// <summary>
            /// 生成图片的宽度
            /// </summary>
            public int Width
            {
                get { return _width; }
                set { _width = value; }
            }

            /**/
            /// <summary>
            /// 生成图片的高度
            /// </summary>
            public int Height
            {
                get { return _height; }
                set { _height = value; }
            }

            /**/
            /// <summary>
            /// 根据字体调整背景
            /// </summary>
            public bool Adaptable
            {
                get { return _adaptable; }
                set { _adaptable = value; }
            }

            public Color BgColor
            {
                get { return _bgColor; }
                set { _bgColor = value; }
            }
            /// <summary>
            /// 无背景
            /// </summary>
            public bool NoBg
            {
                get { return _noBg; }
                set { _noBg = value; }
            }

            /// <summary>
            /// 倾斜角度
            /// </summary>
            public int Incline
            {
                get { return _incline; }
                set { _incline = value; }
            }

            public void Create()
            {

                try
                {
                    Bitmap bitmap;
                    Graphics graph;

                    if (this._backgroundImage == null)
                    {
                        bitmap = new Bitmap(this._width, this._height, PixelFormat.Format64bppArgb);
                        graph = Graphics.FromImage(bitmap);
                        if (!_noBg)
                        {
                            graph.Clear(this._bgColor);
                        }
                    }
                    else
                    {
                        bitmap = new Bitmap(Image.FromStream(_backgroundImage));
                        graph = Graphics.FromImage(bitmap);
                    }

                    graph.SmoothingMode = SmoothingMode.HighQuality;
                    graph.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graph.CompositingQuality = CompositingQuality.HighQuality;

                    Font font = new Font(_fontFamily, _fontSize, _fontStyle);
                    SizeF size = graph.MeasureString(_text, font);

                    if (_adaptable && this._backgroundImage == null)
                    {
                        // 自适应大小
                        _width = (int)size.Width;
                        _height = (int)size.Height;
                        _adaptable = false;
                        Create();
                        return;
                    }

                    Brush brush = new SolidBrush(Color.FromArgb(_BgAlpha, _red, _green, _blue));
                    StringFormat strFormat = new StringFormat();
                    strFormat.Alignment = StringAlignment.Near;

                    if (_shadow)
                    {
                        // 加阴影
                        Brush brushShadow = new SolidBrush(Color.FromArgb(90, 0, 0, 0));
                        graph.DrawString(_text, font, brushShadow, _left + 2, _top + 1);
                    }
                    if (_incline > 0)
                    {
                        graph.RotateTransform(30.0f);
                        _width = Convert.ToInt32(graph.VisibleClipBounds.Width);
                        _height = Convert.ToInt32(graph.VisibleClipBounds.Height);
                        _incline = 0;
                        Create();
                        return;
                    }
                    graph.DrawString(_text, font, brush, new PointF(_left, _top), strFormat);
                    //bitmap.Save(@"E:\1\" + Guid.NewGuid().ToString() + ".png", ImageFormat.Png);
                    _resultImage = (Image)bitmap.Clone();
                    bitmap.Dispose();
                    graph.Dispose();
                }
                catch (Exception)
                {

                    throw;
                }
            }
        }
    }
}