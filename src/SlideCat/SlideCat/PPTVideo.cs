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
            this._presentation = new PowerPoint.Presentation();
            PowerPoint.CustomLayout customLayout = this._presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutBlank];
            PowerPoint.Slide slide = this._presentation.Slides.AddSlide(1, customLayout);
            PowerPoint.Shape shape = slide.Shapes.AddMediaObject2(this._src, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 500, 500);
            slide.Design.SlideMaster.Background.Fill.ForeColor.RGB = 0;
        }

        public override PowerPoint.Presentation getPresentation()
        {
            return this._presentation;
        }
    }
}
