using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace SlideCat
{
    public class PPTImage : PPT
    {
        public override void Load(string file = null)
        {
            _src = file;
        }

        public override void createPresentation()
        {
            _application = new Application();
            _presentation = _application.Presentations.Add(MsoTriState.msoFalse);
            CustomLayout customLayout = _presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];

            Slide slide = _presentation.Slides.AddSlide(1, customLayout);
            Shape shape = slide.Shapes.AddPicture(_src, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);


            shape.Left = _presentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
            shape.Top = _presentation.PageSetup.SlideHeight * .5f - shape.Height / 2;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = 0;
        }

        public override Presentation getPresentation()
        {
            return _presentation;
        }
    }
}