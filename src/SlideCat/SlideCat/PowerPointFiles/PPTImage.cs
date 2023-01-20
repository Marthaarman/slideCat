using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace SlideCat
{
    public class PPTImage : PPT
    {
        public override void Load(string file = null)
        {
            mSrc = file;
        }

        public override void CreatePresentation()
        {
            mApplication = new Application();
            mPresentation = mApplication.Presentations.Add(MsoTriState.msoFalse);
            CustomLayout customLayout = mPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];

            Slide slide = mPresentation.Slides.AddSlide(1, customLayout);
            Shape shape = slide.Shapes.AddPicture(mSrc, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);


            shape.Left = mPresentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
            shape.Top = mPresentation.PageSetup.SlideHeight * .5f - shape.Height / 2;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = 0;
        }

        public override Presentation GetPresentation()
        {
            return mPresentation;
        }
    }
}