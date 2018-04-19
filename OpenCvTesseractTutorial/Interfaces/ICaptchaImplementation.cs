using OpenCvTesseractTutorial.Dto;

namespace OpenCvTesseractTutorial.Interfaces
{
    interface ICaptchaImplementation
    {
        bool ProcessApplication(CCApplicant applicantInfo);
    }
}