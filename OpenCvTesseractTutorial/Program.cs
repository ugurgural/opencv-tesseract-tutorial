using ImageMagick;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using OpenQA.Selenium.Chrome;
using System.Drawing;
using System.IO;
using System.Reflection;
using Tesseract;

namespace OpenCvTesseractTutorial
{
    class Program
    {
        static void Main(string[] args)
        {
            Bitmap bmpCaptcha = GetCaptchaFromWebPage("https://www.bonus.com.tr/kredi-karti-basvurusu", true);
            Bitmap blackAndWhiteCaptcha = ConvertCaptchaToBlackAndWhite(bmpCaptcha, true);
            Bitmap linelessCaptcha = RemoveLinesFromCaptcha(blackAndWhiteCaptcha, true);
            string OcrResult = OcrCaptcha(linelessCaptcha);
        }

        static Bitmap GetCaptchaFromWebPage(string pageUrl, bool isGenerateLocalCopies = false)
        {
            Bitmap bmpCaptcha = null;

            using (var driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)))
            {
                driver.Navigate().GoToUrl(pageUrl);
                driver.Manage().Window.Maximize();
                driver.ExecuteScript("document.body.style.zoom = '200%'");
                driver.ExecuteScript("window.scrollTo(0, 250);");
                var arrScreen = driver.GetScreenshot().AsByteArray;

                using (var msScreen = new MemoryStream(arrScreen))
                {
                    var bmpScreen = new Bitmap(msScreen);
                    var captchaCropCoordinates = new System.Drawing.Rectangle(new System.Drawing.Point(960, 670), new System.Drawing.Size(355, 120));
                    bmpCaptcha = bmpScreen.Clone(captchaCropCoordinates, bmpScreen.PixelFormat);

                    if (isGenerateLocalCopies)
                    {
                        bmpScreen.Save(@"webpage.png", System.Drawing.Imaging.ImageFormat.Png);
                        bmpCaptcha.Save(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\captcha.png");
                    }
                }
            }

            return bmpCaptcha;
        }

        static Bitmap RemoveLinesFromCaptcha(Bitmap captchaBmp, bool isGenerateLocalCopies = false)
        {
            // load the file
            Mat captchaMat = BitmapConverter.ToMat(captchaBmp);

            using (var src = captchaMat)
            {
                // Transform source image to gray if it is not
                Mat gray = new Mat();
                if (src.Channels() == 3)
                {
                    Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
                }
                else
                {
                    gray = src;
                }

                // Apply adaptiveThreshold at the bitwise_not of gray, notice the ~ symbol
                Mat bw = new Mat();
                Cv2.AdaptiveThreshold(~gray, bw, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 15, -2);

                // Create the images that will use to extract the horizontal and vertical lines
                Mat horizontal = bw.Clone();
                Mat vertical = bw.Clone();

                // Specify size on horizontal axis
                int horizontalsize = horizontal.Cols / 8;
                // Create structure element for extracting horizontal lines through morphology operations
                Mat horizontalStructure = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(horizontalsize, 1));
                // Apply morphology operations
                Cv2.Erode(horizontal, horizontal, horizontalStructure, new OpenCvSharp.Point(-1, -1));
                Cv2.Dilate(horizontal, horizontal, horizontalStructure, new OpenCvSharp.Point(-1, -1));

                // Specify size on vertical axis
                int verticalsize = vertical.Rows / 30;
                // Create structure element for extracting vertical lines through morphology operations           
                Mat verticalStructure = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(1, verticalsize));
                // Apply morphology operations
                Cv2.Erode(vertical, vertical, verticalStructure, new OpenCvSharp.Point(-1, -1));
                Cv2.Dilate(vertical, vertical, verticalStructure, new OpenCvSharp.Point(-1, -1));

                // Inverse vertical image
                Cv2.BitwiseNot(vertical, vertical);

                // Extract edges and smooth image according to the logic
                // 1. extract edges
                // 2. dilate(edges)
                // 3. src.copyTo(smooth)
                // 4. blur smooth img
                // 5. smooth.copyTo(src, edges)
                // Step 1
                Mat edges = new Mat();
                Cv2.AdaptiveThreshold(vertical, edges, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 3, -2);
                // Step 2
                Mat kernel = Mat.Ones(2, 2, MatType.CV_8UC1);
                Cv2.Dilate(edges, edges, kernel);
                // Step 3
                Mat smooth = new Mat();
                vertical.CopyTo(smooth);
                // Step 4
                Cv2.Blur(smooth, smooth, new OpenCvSharp.Size(2, 2));
                // Step 5
                smooth.CopyTo(vertical, edges);

                MagickImage image = new MagickImage(vertical.ToBitmap());
                image.Morphology(MorphologyMethod.Close, "1x4: 0,1,1,0", 7);
                Bitmap bmpCaptchaLineRemoved = image.ToBitmap();

                if (isGenerateLocalCopies)
                {
                    Bitmap bmpBlurred = BitmapConverter.ToBitmap(vertical);
                    bmpBlurred.Save(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\blurred.png");
                    bmpCaptchaLineRemoved.Save(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\lineRemoved.png");
                }

                
                image.Dispose();
                return bmpCaptchaLineRemoved;
            }
        }

        static Bitmap ConvertCaptchaToBlackAndWhite(Bitmap captchaBmp, bool isGenerateLocalCopies = false)
        {
            // load the file
            Mat captchaMat = BitmapConverter.ToMat(captchaBmp);
            using (var src = captchaMat)
            {
                using (var binaryMask = new Mat())
                {

                    // lines color is different than text
                    var linesColor = Scalar.FromRgb(0x70, 0x70, 0x70);

                    // build a mask of lines
                    Cv2.InRange(src, linesColor, linesColor, binaryMask);
                    using (var masked = new Mat())
                    {
                        // build the corresponding image
                        // dilate lines a bit because aliasing may have filtered borders too much during masking
                        src.CopyTo(masked, binaryMask);
                        int linesDilate = 3;
                        using (var element = Cv2.GetStructuringElement(MorphShapes.Ellipse, new OpenCvSharp.Size(linesDilate, linesDilate)))
                        {
                            Cv2.Dilate(masked, masked, element);
                        }

                        // convert mask to grayscale
                        Cv2.CvtColor(masked, masked, ColorConversionCodes.BGR2GRAY);
                        using (var dst = src.EmptyClone())
                        {
                            // repaint big lines
                            Cv2.Inpaint(src, masked, dst, 3, InpaintMethod.NS);

                            // destroy small lines
                            linesDilate = 2;
                            using (var element = Cv2.GetStructuringElement(MorphShapes.Ellipse, new OpenCvSharp.Size(linesDilate, linesDilate)))
                            {
                                Cv2.Dilate(dst, dst, element);
                            }

                            Cv2.GaussianBlur(dst, dst, new OpenCvSharp.Size(5, 5), 0);
                            using (var dst2 = dst.BilateralFilter(5, 75, 75))
                            {
                                // basically make it B&W
                                Cv2.CvtColor(dst2, dst2, ColorConversionCodes.BGR2GRAY);
                                Cv2.Threshold(dst2, dst2, 255, 255, ThresholdTypes.Otsu);

                                if (isGenerateLocalCopies)
                                {
                                    dst2.SaveImage(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\blackandwhite.png");
                                }

                                // save the file
                                return dst2.ToBitmap();
                            }
                        }
                    }
                }
            }
        }

        public static string OcrCaptcha(Bitmap b)
        {
            string res = "";
            using (var engine = new TesseractEngine(@"tessdata", "eng", EngineMode.Default))
            {
                engine.SetVariable("tessedit_char_whitelist", "1234567890");
                engine.SetVariable("tessedit_unrej_any_wd", true);
                engine.SetVariable("classify_bln_numeric_mode", true);
                using (var page = engine.Process(b, PageSegMode.SingleWord))
                    res = page.GetText();
            }
            return res;
        }
    }
}