using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Freeware;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace SlideCat
{
    internal class PPTPDF : PPT
    {
        private readonly string _mPptPDFPath = string.Empty;

        public PPTPDF()
        {
            _mPptPDFPath = mSlideCatPath + "PPTPDF\\";

            if (Directory.Exists(_mPptPDFPath)) Directory.Delete(_mPptPDFPath, true);
            Directory.CreateDirectory(_mPptPDFPath);
        }

        public override int nrSlides => 1;

        public override void Load(string file = null)
        {
            mSrc = file;
        }

        public override void CreatePresentation()
        {
            mApplication = new Application();
            mPresentation = mApplication.Presentations.Add(MsoTriState.msoFalse);
            CustomLayout customLayout = mPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];


            Stream pdfFileStream = new FileStream(mSrc, FileMode.Open, FileAccess.Read);
            List<byte[]> imagesBytes = Pdf2Png.ConvertAllPages(pdfFileStream);


            int i = 0;
            foreach (byte[] imageBytes in imagesBytes)
            {
                i++;

                Image image = Image.FromStream(new MemoryStream(imageBytes));
                image.Save(_mPptPDFPath + "PPTPDF_" + i + ".png", ImageFormat.Png);

                Slide slide = mPresentation.Slides.AddSlide(i, customLayout);
                Shape shape = slide.Shapes.AddPicture(_mPptPDFPath + "PPTPDF_" + i + ".png", MsoTriState.msoFalse,
                    MsoTriState.msoTrue, 0, 0);
                shape.Left = mPresentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
                shape.Top = mPresentation.PageSetup.SlideHeight * .5f - shape.Height / 2;
                Color slideBGColor = Color.Black;
                slide.Design.SlideMaster.Background.Fill.ForeColor.RGB = slideBGColor.ToArgb();
                slide.FollowMasterBackground = MsoTriState.msoFalse;
                slide.Background.Fill.ForeColor.RGB = 0;

                File.Delete(_mPptPDFPath + "PPTPDF_" + i + ".png");
            }
        }

        public override Presentation GetPresentation()
        {
            return mPresentation;
        }
    }
}