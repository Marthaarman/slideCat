using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class PPTVideo : PPT
    {
        public override void Load(string file = null)
        {
            _src = file;
        }

        public override void createPresentation()
        {
            _application = new Application();
            _presentation = _application.Presentations.Add(MsoTriState.msoFalse);
            var customLayout = _presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];

            var slide = _presentation.Slides.AddSlide(1, customLayout);

            var shape = slide.Shapes.AddMediaObject2(_src);
            shape.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;

            shape.Left = _presentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
            shape.Top = _presentation.PageSetup.SlideHeight * .5f - shape.Height / 2;

            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = 0;

            var settings = _presentation.SlideShowSettings;
            settings.ShowType = PpSlideShowType.ppShowTypeSpeaker;
            settings.ShowPresenterView = MsoTriState.msoFalse;
            var sw = settings.Run();
            _presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();
        }

        public override Presentation getPresentation()
        {
            return _presentation;
        }
    }
}