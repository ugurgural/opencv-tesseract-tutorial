using ImageMagick;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using OpenCvTesseractTutorial.Dto;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using OpenCvTesseractTutorial.Interfaces;
using Tesseract;

namespace OpenCvTesseractTutorial.Implementations
{
    public class MilesSmilesCCService : ICaptchaImplementation
    {
        public bool ProcessApplication(CCApplicant applicantInfo)
        {
            bool result = false;

            using (var driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)))
            {
                //Specifically scaled for onboard graphic cards, if client resolution is different change the resolution and captcha coordinates.
                driver.Navigate().GoToUrl("https://www.bonus.com.tr/kredi-karti-basvurusu");
                driver.Manage().Window.Size = new System.Drawing.Size(1920, 1080);
                driver.Manage().Window.Maximize();
                driver.ExecuteScript("document.body.style.zoom = '200%'");
                driver.ExecuteScript("window.scrollTo(0, 250);");

                Bitmap bmpCaptcha = GetCaptchaFromWebPage(driver, true);
                Bitmap blackAndWhiteCaptcha = ConvertCaptchaToBlackAndWhite(bmpCaptcha, true);
                Bitmap linelessCaptcha = RemoveLinesFromCaptcha(blackAndWhiteCaptcha, true);

                string decodedCaptchaText = OcrCaptcha(linelessCaptcha);

                if (!string.IsNullOrWhiteSpace(decodedCaptchaText))
                {
                    // For selecting the card type in page, we have to resize to window to mobile size to choose cardtype from dropdown. Otherwise there is a gallery showing in desktop mode for card type selection, which is harder to manipulate the dom.
                    driver.Manage().Window.Size = new System.Drawing.Size(480, 1080);
                    driver.ExecuteScript("document.body.style.zoom = '100%'");

                    driver.FindElement(By.CssSelector(".input.text.names")).SendKeys(applicantInfo.FirstName);
                    driver.FindElement(By.CssSelector(".input.text.surnames")).SendKeys(applicantInfo.LastName);
                    driver.FindElement(By.CssSelector(".input.phone-number")).SendKeys(applicantInfo.PhoneNumber);
                    driver.FindElement(By.CssSelector(".input.tc-number")).SendKeys(applicantInfo.TcIdentityNo);

                    var cardTypeSelect = new SelectElement(driver.FindElement(By.CssSelector(".hidden-select")).FindElement(By.TagName("select")));
                    cardTypeSelect.SelectByText(applicantInfo.AppliedCCName);

                    driver.FindElement(By.CssSelector(".input.captcha")).SendKeys(decodedCaptchaText);
                    driver.FindElement(By.Id("yasalsozlesme")).Click();

                    driver.FindElement(By.CssSelector(".button.large.green.transition.apply-submit ")).Click();
                }

            }

            return result;
        }


        /// <summary>
        /// Selenium operation for extracting captcha from the desired webpage. Note that the screen resolution and captcha coordinates are machine-specific and needs to be reviewed if resolution is different than 1920x1080.
        /// </summary>
        /// 
        /// <param name="pageUrl">Url of the web page</param>
        /// <param name="isGenerateLocalCopies">Turn this on if you want to save captcha files that processed to see the steps</param>
        /// <returns></returns>
        private Bitmap GetCaptchaFromWebPage(ChromeDriver driver, bool isGenerateLocalCopies = false)
        {
            Bitmap bmpCaptcha = null;
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

            return bmpCaptcha;
        }

        /// <summary>
        /// Second and last operation for handling captcha, sharpen the image and remove the line that crossing the captcha to make it readable for ocr operation.
        /// </summary>
        /// <param name="captchaBmp">Bitmap of captcha</param>
        /// <param name="isGenerateLocalCopies">Turn this on if you want to save captcha files that processed to see the steps</param>
        /// <returns>Bitmap of processed captcha</returns>
        private Bitmap RemoveLinesFromCaptcha(Bitmap captchaBmp, bool isGenerateLocalCopies = false)
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

                // Inverse vertical image
                Cv2.BitwiseNot(bw, bw);

                // Extract edges and smooth image according to the logic
                // 1. extract edges
                // 2. dilate(edges)
                // 3. src.copyTo(smooth)
                // 4. blur smooth img
                // 5. smooth.copyTo(src, edges)
                // Step 1
                Mat edges = new Mat();
                Cv2.AdaptiveThreshold(bw, edges, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 3, -2);
                // Step 2
                Mat kernel = Mat.Ones(2, 2, MatType.CV_8UC1);
                Cv2.Dilate(edges, edges, kernel);
                // Step 3
                Mat smooth = new Mat();
                bw.CopyTo(smooth);
                // Step 4
                Cv2.Blur(smooth, smooth, new OpenCvSharp.Size(2, 2));
                // Step 5
                smooth.CopyTo(bw, edges);

                //Remove lines using imagemagick's morphology tool
                MagickImage image = new MagickImage(bw.ToBitmap());
                image.Morphology(MorphologyMethod.Close, "1x4: 0,1,1,0", 7);
                Bitmap bmpCaptchaLineRemoved = image.ToBitmap();

                if (isGenerateLocalCopies)
                {
                    Bitmap bmpBlurred = BitmapConverter.ToBitmap(bw);
                    bmpBlurred.Save(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\blurred.png");
                    bmpCaptchaLineRemoved.Save(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\lineRemoved.png");
                }

                image.Dispose();
                return bmpCaptchaLineRemoved;
            }
        }

        /// <summary>
        /// First operation handling captcha, convert image to grayscale, remove all little noises, make it b&w.
        /// </summary>
        /// <param name="captchaBmp">Bitmap of captcha</param>
        /// <param name="isGenerateLocalCopies">Turn this on if you want to save captcha files that processed to see the steps</param>
        /// <returns>Bitmap of processed captcha</returns>
        private Bitmap ConvertCaptchaToBlackAndWhite(Bitmap captchaBmp, bool isGenerateLocalCopies = false)
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

        /// <summary>
        /// Doing the final ocr process with tesseract.
        /// </summary>
        /// <param name="captchaBmp">Bitmap of captcha</param>
        /// <returns>String of captcha</returns>
        private string OcrCaptcha(Bitmap captchaBmp)
        {
            string captchaText = string.Empty;

            using (var engine = new TesseractEngine(@"tessdata", "eng", EngineMode.Default))
            {
                //Set engine to process only numeric characters
                engine.DefaultPageSegMode = PageSegMode.SingleLine;
                engine.SetVariable("tessedit_char_blacklist", "!?@#$%&*()<>_-+=/:;'\"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz");
                engine.SetVariable("tessedit_char_whitelist", "1234567890");
                engine.SetVariable("tessedit_unrej_any_wd", true);
                engine.SetVariable("classify_bln_numeric_mode", "1");

                using (var page = engine.Process(captchaBmp, PageSegMode.SingleWord))
                {
                    captchaText = Regex.Replace(page.GetText(), "[ \n\r\t]", "");
                }
            }

            return captchaText;
        }
    }
}
