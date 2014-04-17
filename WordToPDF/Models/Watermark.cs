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
            private int _width = 460;
            private int _height = 30;
            private int _left = 0;
            private int _top = 0;

            private string _text = "内部资料";
            private string _fontFamily = "华文行楷";
            private int _fontSize = 140;
            private Color _fontColor = Color.Black;
            private bool _bold = true;

            /// <summary>
            /// 根据背景调整字体大小
            /// </summary>
            private bool _adaptable = true;
            /// <summary>
            /// 水印文字是否使用阴影
            /// </summary>
            private bool _shadow = false;
            /// <summary>
            /// 水印平铺
            /// </summary>
            private bool _tiling = true;
            /// <summary>
            /// 无背景色
            /// </summary>
            private bool _noBg = true;
            /// <summary>
            /// 背景颜色
            /// </summary>
            private Color _bgColor = Color.White;
            /// <summary>
            /// 水印得透明度
            /// </summary>
            private float _transparency = 0.5f;
            /// <summary>
            /// 倾斜角度
            /// </summary>
            private int _rotation = 45;

            /// <summary>
            /// 底图
            /// </summary>
            private Stream _backgroundImage;
            /// <summary>
            /// 生成后的图片
            /// </summary>
            private Image _resultImage;

            #region 构造函数
            public WaterImage() { }
            public WaterImage(int width, int height)
            {
                _width = width;
                _height = height;
            }
            #endregion

            #region 属性
            /// <summary>
            /// 字体
            /// </summary>
            public string FontFamily
            {
                set { this._fontFamily = value; }
            }

            /// <summary>
            /// 文字大小
            /// </summary>
            public int FontSize
            {
                set { this._fontSize = value; }
            }

            /// <summary>
            /// 水印得透明度
            /// </summary>
            public float Transparency
            {
                get
                {
                    if (_transparency > 1.0f)
                    {
                        _transparency = 1.0f;
                    }
                    return _transparency;
                }
                set { _transparency = value; }
            }

            /// <summary>
            /// 水印文字是否使用阴影
            /// </summary>
            public bool Shadow
            {
                get { return _shadow; }
                set { _shadow = value; }
            }
            /// <summary>
            /// 水印平铺
            /// </summary>
            public bool Tiling
            {
                get { return _tiling; }
                set { _tiling = value; }
            }

            public Color FontColor
            {
                get { return _fontColor; }
                set { _fontColor = value; }
            }

            /// <summary>
            /// 底图
            /// </summary>
            public Stream BackgroundImage
            {
                set { this._backgroundImage = value; }
            }

            /// <summary>
            /// 水印文字的左边距
            /// </summary>
            public int Left
            {
                set { this._left = value; }
            }

            /// <summary>
            /// 水印文字的顶边距
            /// </summary>
            public int Top
            {
                set { this._top = value; }
            }

            /// <summary>
            /// 生成后的图片
            /// </summary>
            public Image ResultImage
            {
                set { this._resultImage = value; }
                get { return _resultImage; }
            }

            /// <summary>
            /// 水印文本
            /// </summary>
            public string Text
            {
                set { this._text = value; }
            }

            /// <summary>
            /// 生成图片的宽度
            /// </summary>
            public int Width
            {
                get { return _width; }
                set { _width = value; }
            }

            /// <summary>
            /// 生成图片的高度
            /// </summary>
            public int Height
            {
                get { return _height; }
                set { _height = value; }
            }

            /// <summary>
            /// 根据背景调整字体大小
            /// </summary>
            public bool Adaptable
            {
                get { return _adaptable; }
                set { _adaptable = value; }
            }

            /// <summary>
            /// 背景颜色
            /// </summary>
            public Color BgColor
            {
                get { return _bgColor; }
                set { _bgColor = value; }
            }

            /// <summary>
            /// 无背景色
            /// </summary>
            public bool NoBg
            {
                get { return _noBg; }
                set { _noBg = value; }
            }

            /// <summary>
            /// 倾斜角度
            /// </summary>
            public int Rotation
            {
                get { return _rotation; }
                set { _rotation = value; }
            }

            /// <summary>
            /// 水印文字是否加粗
            /// </summary>
            public bool Bold
            {
                get { return _bold; }
                set { _bold = value; }
            }
            #endregion

            public void Create()
            {

                try
                {
                    Bitmap bitmap;
                    Graphics graph;

                    bitmap = new Bitmap(this._width, this._height, PixelFormat.Format32bppArgb);
                    graph = Graphics.FromImage(bitmap);

                    graph.SmoothingMode = SmoothingMode.AntiAlias;
                    graph.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graph.CompositingQuality = CompositingQuality.HighQuality;

                    Font font = _bold ? new Font(_fontFamily, _fontSize, FontStyle.Bold) : new Font(_fontFamily, _fontSize, FontStyle.Regular);
                    SizeF size = graph.MeasureString(_text, font);

                    if (_adaptable)
                    {
                        while (size.Width > _width || size.Height > _height)
                        {
                            _fontSize--;
                            font = _bold ? new Font(_fontFamily, _fontSize, FontStyle.Bold) : new Font(_fontFamily, _fontSize, FontStyle.Regular);
                            size = graph.MeasureString(_text, font);
                        }
                    }

                    // 根据字体大小裁剪画布
                    bitmap = bitmap.Clone(new Rectangle(0, 0, (int)size.Width, (int)size.Height), bitmap.PixelFormat);
                    graph = Graphics.FromImage(bitmap);
                    if (!_noBg)
                    {
                        graph.Clear(this._bgColor);
                    }


                    Brush brush = new SolidBrush(_fontColor);
                    StringFormat strFormat = new StringFormat();
                    strFormat.Alignment = StringAlignment.Near;

                    if (_shadow)
                    {
                        // 加阴影
                        Brush brushShadow = new SolidBrush(Color.Gray);
                        graph.DrawString(_text, font, brushShadow, _left + 2, _top + 1);
                    }

                    graph.DrawString(_text, font, brush, new PointF(_left, _top), strFormat);
                    if (_rotation != 0)
                    {
                        bitmap = KiRotate(bitmap);
                    }
                    if (_tiling)
                    {
                        // 创建水印笔刷
                        TextureBrush wtbrush = new TextureBrush(bitmap);
                        // 按照原画布大小,创建新的画布
                        Bitmap newBitmap = new Bitmap(this._width, this._height, PixelFormat.Format32bppArgb);
                        Graphics newGraph = Graphics.FromImage(newBitmap);
                        // 平铺水印笔刷
                        newGraph.FillRectangle(wtbrush, new System.Drawing.Rectangle(0, 0, _width, _height));
                        bitmap = (Bitmap)newBitmap.Clone();
                    }

                    _resultImage = (Image)bitmap.Clone();
                    bitmap.Dispose();
                    graph.Dispose();
                }
                catch (Exception)
                {
                    throw;
                }
            }

            /// <summary>
            /// 任意角度旋转
            /// </summary>
            /// <param name="bmp">原始图Bitmap</param>
            /// <param name="angle">旋转角度</param>
            /// <param name="bkColor">背景色</param>
            /// <returns>输出Bitmap</returns>
            private Bitmap KiRotate(Bitmap bmp)
            {
                int w = bmp.Width + 2;
                int h = bmp.Height + 2;
                Bitmap tmp = new Bitmap(w, h, bmp.PixelFormat);
                Graphics graph = Graphics.FromImage(tmp);
                if (!_noBg)
                {
                    graph.Clear(_bgColor);
                }

                graph.SmoothingMode = SmoothingMode.HighQuality;
                graph.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graph.CompositingQuality = CompositingQuality.HighQuality;

                graph.DrawImageUnscaled(bmp, 1, 1);

                graph.Dispose();

                GraphicsPath path = new GraphicsPath();
                path.AddRectangle(new RectangleF(0f, 0f, w, h));
                Matrix mtrx = new Matrix();
                mtrx.Rotate(_rotation);
                RectangleF rct = path.GetBounds(mtrx);
                Bitmap dst = new Bitmap((int)rct.Width, (int)rct.Height, bmp.PixelFormat);
                graph = Graphics.FromImage(dst);

                if (!_noBg)
                {
                    graph.Clear(_bgColor);
                }

                graph.SmoothingMode = SmoothingMode.HighQuality;
                graph.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graph.CompositingQuality = CompositingQuality.HighQuality;

                graph.TranslateTransform(-rct.X, -rct.Y);
                graph.RotateTransform(_rotation);

                graph.DrawImageUnscaled(tmp, 0, 0);
                graph.Dispose();
                tmp.Dispose();
                return dst;
            }

            public ImageAttributes SetTransparency(float transparency)
            {

                //创建颜色矩阵
                float[][] ptsArray ={ 
                      new float[] {1, 0, 0, 0, 0},
                      new float[] {0, 1, 0, 0, 0},
                      new float[] {0, 0, 1, 0, 0},
                      new float[] {0, 0, 0, transparency, 0}, //注意：此处为0.0f为完全透明，1.0f为完全不透明
                      new float[] {0, 0, 0, 0, 1}};
                ColorMatrix colorMatrix = new ColorMatrix(ptsArray);
                //新建一个Image属性
                ImageAttributes imageAttributes = new ImageAttributes();
                //将颜色矩阵添加到属性
                imageAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default,
                 ColorAdjustType.Default);

                return imageAttributes;
            }
        }
    }
}