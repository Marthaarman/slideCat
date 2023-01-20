using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class PPT
    {
        protected Application mApplication;

        private bool _mPptValid;
        protected Presentation mPresentation;
        protected string mSlideCatPath = Path.GetTempPath() + "slidecat\\";

        protected string mSrc = "";

        public PPT()
        {
            if (!Directory.Exists(mSlideCatPath)) Directory.CreateDirectory(mSlideCatPath);
        }

        public virtual int nrSlides => _mPptValid ? mPresentation.Slides.Count : 0;

        public virtual void Load(string file = null)
        {
            if (file == null) return;
            mSrc = file;
            mApplication = new Application();
            mPresentation = mApplication.Presentations.Open2007(mSrc, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            _mPptValid = true;
        }

        public virtual void CreatePresentation()
        {
        }

        public virtual Presentation GetPresentation()
        {
            return mPresentation;
        }
    }
}