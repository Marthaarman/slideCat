using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace SlideCat
{
    public class PPTVideo : PPT
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

            Shape shape = slide.Shapes.AddMediaObject2(mSrc);
            shape.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;
            shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;

            shape.Left = mPresentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
            shape.Top = mPresentation.PageSetup.SlideHeight * .5f - shape.Height / 2;

            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = 0;
            

            /*SlideShowSettings settings = mPresentation.SlideShowSettings;
            settings.ShowType = PpSlideShowType.ppShowTypeSpeaker;
            settings.ShowPresenterView = MsoTriState.msoFalse;
            SlideShowWindow sw = settings.Run();
            mPresentation.SlideShowWindow.View.FirstAnimationIsAutomatic();*/
        }

        public override Presentation GetPresentation()
        {
            return mPresentation;
        }
    }
}