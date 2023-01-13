using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class PPTVideo : PPT
    {
        public override void Load(String file = null)
        {

            this._src = file;
        }

        public override void createPresentation()
        {
            this._application = new PowerPoint.Application();
            this._presentation = this._application.Presentations.Add(MsoTriState.msoFalse);
            PowerPoint.CustomLayout customLayout = this._presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutTitle];

            PowerPoint.Slide slide = this._presentation.Slides.AddSlide(1, customLayout);
            
            PowerPoint.Shape shape = slide.Shapes.AddMediaObject2(this._src, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
            shape.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;

            shape.Left = this._presentation.PageSetup.SlideWidth * .5f - shape.Width / 2;
            shape.Top = this._presentation.PageSetup.SlideHeight * .5f - shape.Height / 2;
 
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = 0;

            PowerPoint.SlideShowSettings settings = this._presentation.SlideShowSettings;
            settings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeSpeaker;
            settings.ShowPresenterView = MsoTriState.msoFalse;
            PowerPoint.SlideShowWindow sw = settings.Run();
            this._presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();
        }

        public override PowerPoint.Presentation getPresentation()
        {
            return this._presentation;
        }
    }
}
