using Microsoft.Office.Core;
using System;
using System.Drawing;
using System.IO;
using WebSupergoo.ABCpdf12;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    internal class PPTPDF : PPT
    {

        private string _pptPDFPath = "";

        public PPTPDF() : base()
        {
            this._pptPDFPath = this._pptPath + "PPTPDF\\";

            if (Directory.Exists(this._pptPDFPath))
            {
                Directory.Delete(this._pptPDFPath, true);
            }
            Directory.CreateDirectory(this._pptPDFPath);
        }

        public override int nrSlides { get { return 1; } }
        public override void Load(String file = null)
        {
            this._src = file;
        }

        public override void createPresentation()
        {
            this._application = new PowerPoint.Application();
            this._presentation = this._application.Presentations.Add(MsoTriState.msoFalse);
            PowerPoint.CustomLayout customLayout = this._presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutTitle];
            

            Doc PDFDoc = new Doc();
            PDFDoc.Read(this._src);
            int pageCount = PDFDoc.PageCount;
            for(int i = 1; i <= pageCount; i++)
            {
                PDFDoc.PageNumber = i;
                XRendering xRendering = PDFDoc.Rendering;
                xRendering.Save(this._pptPDFPath + "PPTPDF_" + i + ".emf");

                PowerPoint.Slide slide = this._presentation.Slides.AddSlide(i, customLayout);
                PowerPoint.Shape shape = slide.Shapes.AddPicture(this._pptPDFPath + "PPTPDF_" + i + ".emf", MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
                //shape.Width = (float)(0.9 * shape.Width);
                //shape.Height = (float)(0.9 * shape.Height);
                shape.Left = this._presentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
                shape.Top = this._presentation.PageSetup.SlideHeight * .5f - shape.Height / 2;
                Color slideBGColor = Color.Black;
                slide.Design.SlideMaster.Background.Fill.ForeColor.RGB = slideBGColor.ToArgb();
                slide.FollowMasterBackground = MsoTriState.msoFalse;
                slide.Background.Fill.ForeColor.RGB = 0;

                File.Delete(this._pptPDFPath + "PPTPDF_" + i + ".emf");
            }


            
            /*
            this._application = new PowerPoint.Application();
            this._presentation = new PowerPoint.Presentation();
            PowerPoint.CustomLayout customLayout = this._presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutBlank];
            PowerPoint.Slide slide = this._presentation.Slides.AddSlide(1, customLayout);
            PowerPoint.Shape shape = slide.Shapes.AddPicture(this._src, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 500, 500);
            slide.Design.SlideMaster.Background.Fill.ForeColor.RGB = 0;
            */
        }

        public override PowerPoint.Presentation getPresentation()
        {
            return this._presentation;
        }
    }
}
