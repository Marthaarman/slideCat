using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class PPT
    {
        protected PowerPoint.Slides _slides;
        protected PowerPoint.Application _application;
        protected PowerPoint.Presentation _presentation;
        protected String _src = "";
        protected String _name = "";
        
        public int nrSlides {  get { return _presentation.Slides.Count; } }

        virtual public void Load(String file = null)
        {
            if (file == null)
            {
                return;
            }
            this._src = file;
            this._application = new PowerPoint.Application();
            this._presentation = _application.Presentations.Open2007(this._src, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);

            //  used to add one safetyslide into the back
            //PowerPoint.SlideRange dupliSlide = this._presentation.Slides[this._presentation.Slides.Count].Duplicate();
            //dupliSlide.MoveTo(this._presentation.Slides.Count);
        }

        virtual public void createPresentation() {}

        virtual public PowerPoint.Presentation getPresentation()
        {
            return this._presentation;
        }

    }
}
