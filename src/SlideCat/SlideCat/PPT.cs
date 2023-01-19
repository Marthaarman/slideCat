using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class PPT
    {
        protected Application _application;
        protected string _name = "";
        protected string _powerPointPath = "";

        private bool _pptValid;
        protected Presentation _presentation;
        protected string _slideCatPath = Path.GetTempPath() + "slidecat\\";
        protected Slides _slides;

        protected string _src = "";

        public PPT()
        {
            if (!Directory.Exists(_slideCatPath)) Directory.CreateDirectory(_slideCatPath);
        }

        public virtual int nrSlides => _pptValid ? _presentation.Slides.Count : 0;

        public virtual void Load(string file = null)
        {
            if (file == null) return;
            _src = file;
            _application = new Application();
            _presentation = _application.Presentations.Open2007(_src, MsoTriState.msoTrue, MsoTriState.msoFalse,
                MsoTriState.msoFalse, MsoTriState.msoTrue);
            _pptValid = true;
            //  used to add one safetyslide into the back
            //PowerPoint.SlideRange dupliSlide = this._presentation.Slides[this._presentation.Slides.Count].Duplicate();
            //dupliSlide.MoveTo(this._presentation.Slides.Count);
        }

        public virtual void createPresentation()
        {
        }

        public virtual Presentation getPresentation()
        {
            return _presentation;
        }

        public virtual void exitPresentation()
        {
        }
    }
}