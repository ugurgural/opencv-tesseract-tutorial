using ImageMagick;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using OpenQA.Selenium.Chrome;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using OpenCvTesseractTutorial.Dto;
using OpenCvTesseractTutorial.Implementations;
using Tesseract;

namespace OpenCvTesseractTutorial
{
    class Program
    {
        static void Main(string[] args)
        {
            BonusCCService bonusCCService = new BonusCCService();
            bonusCCService.ProcessApplication(new CCApplicant() { AppliedCCName = "BONUS", TcIdentityNo = "93280992608", FirstName = "Ali", LastName = "Veli", PhoneNumber = "337202680"});
        }
    }
}