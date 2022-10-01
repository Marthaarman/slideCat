using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace SlideCat
{
    public class Presentation
    {
        private int _currentSlideIndex;

        private bool _isPlaying = false;

        private PowerPoint.Application _application;
        private PowerPoint.Presentation _presentation;

        private int _intervalCounter = 0;
        private String _slideNotes = String.Empty;
        private String _slideNotesNext = String.Empty;

        public String slideNotes {  get { return this._slideNotes; } }
        public String slideNotesNext { get { return this._slideNotesNext; } }

        public bool IsPlaying { get { return _isPlaying; } }

        private string _pptxPath = System.IO.Path.GetTempPath() + "/slidecat/";

        public Presentation()
        {
            if(!Directory.Exists(this._pptxPath))
            {
                Directory.CreateDirectory(this._pptxPath);
            }
            this._pptxPath += new Random().Next() + "/";

            if(!Directory.Exists(_pptxPath))
            {
                Directory.CreateDirectory(this._pptxPath);
            }
        }

        private void _emptyPresentationDirectory()
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(this._pptxPath);

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        public void createPresentation(MediaItems mediaItems)
        {

            //  initiate new application and main presentation
            this._application = new PowerPoint.Application();
            this._presentation = this._application.Presentations.Add(MsoTriState.msoFalse);

            //  clear the folder to which temporary files are stored
            this._emptyPresentationDirectory();
            
            //  save each presentation as powerpoint presentation into the tmp folder
            //  add each temporary powerpoint into the main powerpoint
            mediaItems.sort();
            int i = 0;
            foreach (MediaItem mediaItem in mediaItems.mediaItems)
            {
                i++;
                PowerPoint.Presentation pres = mediaItem.presentation;
                pres.SaveCopyAs(this._pptxPath + "pptx_"+i, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                this._presentation.Slides.InsertFromFile(this._pptxPath + "pptx_" + i + ".pptx", this._presentation.Slides.Count);
                pres.Close();
            }
            
            //  store the main presentation as file
            this._presentation.SaveCopyAs(this._pptxPath + "pptxfinal", PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

            //  store each slide as image for thumbs
            Directory.CreateDirectory(this._pptxPath + "thumbs/");
            i = 0;
            foreach(PowerPoint.Slide slide in this._presentation.Slides)
            {
                i++;
                slide.Export(this._pptxPath + "thumbs/" + i + ".png", "PNG");
            }
        }

        public void playPresentation()
        {
            this._isPlaying = false;

            if(!(this._presentation.Slides.Count > 0))
            {
                return;
            }

            this._currentSlideIndex = 0;

            PowerPoint.SlideShowSettings settings = this._presentation.SlideShowSettings;

            settings.ShowType = (PowerPoint.PpSlideShowType)1;
            settings.ShowPresenterView = MsoTriState.msoFalse;
            PowerPoint.SlideShowWindow sw = settings.Run();

            this._presentation.SlideShowWindow.View.GotoSlide(this._currentSlideIndex + 1);
            this._presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();


            this._isPlaying = true;
        }

        public void stopPresentation()
        {
            //String file = _presentation.Path + "\\" + _presentation.Name;
            this._stopPresentation(this._presentation);
            this._emptyPresentationDirectory();
            this._isPlaying = false;
        }

        private void _stopPresentation(PowerPoint.Presentation _presentation)
        {
            try
            {
                _presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("LOG - Presentation.cs - _stopPresentation() - catch");
                Console.WriteLine(ex.Message);
            }
        }

        
        public void nextSlide()
        {
            this._presentation.SlideShowWindow.View.Next();
            this._focus();
        }

        public void prevSlide()
        {
            this._presentation.SlideShowWindow.View.Previous();
            this._focus();
        }

        public void goToSlideIndex(int _index)
        {
            this._presentation.SlideShowWindow.View.GotoSlide(_index + 1);
        }

        public bool exitSlide()
        {
            if(this._isPlaying && this._validPresentation())
            {
                try
                {
                    int tmp = this._presentation.SlideShowWindow.View.Slide.SlideIndex;
                }catch
                {
                    return true;
                }
            }
            return false;
        }

        

        public bool firstSlide()
        {
            return false;
        }

        public int getSlideIndex()
        {
            return _currentSlideIndex;
        }        

   

        

        public void focus()
        {
            this._focus();
        }

        public bool runInterval()
        {
            this._intervalCounter++;
            this._setSlideIndex();
            if(this._intervalCounter >= 5)
            {
                this._intervalCounter = 0;
                this._focus();
                this._obtainSlideNotes();
                return true;
            }
            return false;
        }

        public String getThumb()
        {
            return this._pptxPath + "thumbs/" + (this._currentSlideIndex + 1) + ".png";
        }

        public String getNextThumb()
        {
            if (_currentSlideIndex != (this._presentation.Slides.Count - 1))
            {
                return this._pptxPath + "thumbs/" + (this._currentSlideIndex + 2) + ".png";
            }
            
            return String.Empty;
        }
        
        private void _focus()
        {
            if(this._isPlaying && this._validPresentation())
            {
                this._presentation.SlideShowWindow.Activate();
            }
        }

        private void _obtainSlideNotes()
        {
            this._slideNotes = this._getSlideNotes(_currentSlideIndex + 1);
            this._slideNotesNext = this._getSlideNotes(_currentSlideIndex + 2);
        }

        private string _getSlideNotes(int slideNR)
        {
            String notes = String.Empty;
            if (this.IsPlaying && !this.exitSlide() && this._validPresentation() && this._presentation.Slides.Count >= slideNR)
            {

                PowerPoint.Slide slide = this._presentation.Slides[slideNR];
                if (slide.HasNotesPage == MsoTriState.msoTrue)
                {
                    int length = 0;
                    foreach (PowerPoint.Shape shape in slide.NotesPage.Shapes)
                    {
                        if (shape.Type == MsoShapeType.msoPlaceholder)
                        {
                            var tf = shape.TextFrame;
                            try
                            {
                                var range = tf.TextRange;
                                if (range.Length > length)
                                {
                                    length = range.Length;
                                    notes = range.Text;
                                }
                            }
                            catch (Exception ex)
                            { }
                        }
                    }
                }
            }
            return notes;
        }
        private void _setSlideIndex()
        {
            this._currentSlideIndex = 0;
            if (this._isPlaying && this._validPresentation())
            {
                
                try
                {
                    this._currentSlideIndex = this._presentation.SlideShowWindow.View.Slide.SlideIndex - 1;
                } catch
                { }
                    
                
            }
        }

        private bool _validPresentation()
        {
            try
            {
                int tmp = this._presentation.Slides.Count;
            }
            catch
            {
                return false;
            }
            return true;
        }


        public bool validPresentation
        {
            get { return this._validPresentation(); }
        }
    }
}
