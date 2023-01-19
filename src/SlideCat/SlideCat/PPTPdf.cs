using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Freeware;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    internal class PPTPDF : PPT
    {
        private readonly string _pptPDFPath = "";

        public PPTPDF()
        {
            _pptPDFPath = _slideCatPath + "PPTPDF\\";

            if (Directory.Exists(_pptPDFPath)) Directory.Delete(_pptPDFPath, true);
            Directory.CreateDirectory(_pptPDFPath);
        }

        public override int nrSlides => 1;

        public override void Load(string file = null)
        {
            _src = file;
        }

        public override void createPresentation()
        {
            _application = new Application();
            _presentation = _application.Presentations.Add(MsoTriState.msoFalse);
            var customLayout = _presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];

            /*org.pdfclown.files.File file = new org.pdfclown.files.File(this._src);
            org.pdfclown.documents.Document document = file.Document;
            org.pdfclown.documents.Pages pages = document.Pages;*/

            Stream pdfFileStream = new FileStream(_src, FileMode.Open, FileAccess.Read);
            var imagesBytes = Pdf2Png.ConvertAllPages(pdfFileStream);


            var i = 0;
            /*foreach(org.pdfclown.documents.Page page in pages)*/
            foreach (var imageBytes in imagesBytes)
            {
                i++;

                /*SizeF imageSize = page.Size;
                Console.WriteLine(imageSize);
                Renderer renderer = new Renderer();
                Image image = renderer.Render(page, imageSize);

                image.Save(this._pptPDFPath + "PPTPDF_" + i + ".png", ImageFormat.Png);*/
                var image = Image.FromStream(new MemoryStream(imageBytes));
                image.Save(_pptPDFPath + "PPTPDF_" + i + ".png", ImageFormat.Png);

                var slide = _presentation.Slides.AddSlide(i, customLayout);
                var shape = slide.Shapes.AddPicture(_pptPDFPath + "PPTPDF_" + i + ".png", MsoTriState.msoFalse,
                    MsoTriState.msoTrue, 0, 0);
                shape.Left = _presentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
                shape.Top = _presentation.PageSetup.SlideHeight * .5f - shape.Height / 2;
                var slideBGColor = Color.Black;
                slide.Design.SlideMaster.Background.Fill.ForeColor.RGB = slideBGColor.ToArgb();
                slide.FollowMasterBackground = MsoTriState.msoFalse;
                slide.Background.Fill.ForeColor.RGB = 0;

                File.Delete(_pptPDFPath + "PPTPDF_" + i + ".png");
            }


            /* Doc PDFDoc = new Doc();
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
             }*/


            /*
            this._application = new PowerPoint.Application();
            this._presentation = new PowerPoint.Presentation();
            PowerPoint.CustomLayout customLayout = this._presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutBlank];
            PowerPoint.Slide slide = this._presentation.Slides.AddSlide(1, customLayout);
            PowerPoint.Shape shape = slide.Shapes.AddPicture(this._src, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 500, 500);
            slide.Design.SlideMaster.Background.Fill.ForeColor.RGB = 0;
            */
        }

        public override Presentation getPresentation()
        {
            return _presentation;
        }
    }
}