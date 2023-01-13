using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace SlideCat
{
    public class PPT
    {
        protected PowerPoint.Slides _slides;
        protected PowerPoint.Application _application;
        protected PowerPoint.Presentation _presentation;
        
        private bool _pptValid = false;
        protected string _pptPath = System.IO.Path.GetTempPath() + "slidecat\\";
        
        protected String _src = "";
        protected String _name = "";

        public PPT()
        {
            if (!Directory.Exists(_pptPath))
            {
                Directory.CreateDirectory(this._pptPath);
            }
        }
        
        virtual public int nrSlides {  get { return this._pptValid ? _presentation.Slides.Count : 0; } }

        virtual public void Load(String file = null)
        {
            if (file == null)
            {
                return;
            }
            this._src = file;
            this._application = new PowerPoint.Application();
            this._presentation = _application.Presentations.Open2007(this._src, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
            this._pptValid = true;
            //  used to add one safetyslide into the back
            //PowerPoint.SlideRange dupliSlide = this._presentation.Slides[this._presentation.Slides.Count].Duplicate();
            //dupliSlide.MoveTo(this._presentation.Slides.Count);
        }

        virtual public void createPresentation() {}

        virtual public PowerPoint.Presentation getPresentation()
        {
            return this._presentation;
        }

        virtual public void exitPresentation()
        {

        }

    }
}
