using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace SlideCat
{
    internal class Presentation
    {
        private int _currentMediaItemIndex;
        private int _currentSlideIndex;

        private MediaItem _currentMediaItem;
        private MediaItem _nextMediaItem;
        private bool _isPlaying = false;

        private PowerPoint.Application _current_application;
        private PowerPoint.Presentation _current_presentation;
        private PowerPoint.SlideShowWindow _current_slideShowWindow;

        private PowerPoint.Application _next_application;
        private PowerPoint.Presentation _next_presentation;

        private PowerPoint.Application _prev_application;
        private PowerPoint.Presentation _prev_presentation;

        private int _intervalCounter = 0;
        private String _slideNotes = String.Empty;


        public bool IsPlaying { get { return _isPlaying; } }

        public void loadNextPresentationItem(MediaItem _item)
        {
            _nextMediaItem = _item;
            if(_item.type == MediaType.powerpoint)
            {
                _next_application = _item.application;
                _next_presentation = _item.presentation;
            }
        }

        public int presentationItemOrder()
        {
            return _currentMediaItem.order;
        }

        public void playPresentation()
        {
            this._isPlaying = false;
            
            _prev_application = _current_application;
            _prev_presentation = _current_presentation;

            this._currentSlideIndex = 0;

            _currentMediaItem = _nextMediaItem;
            _currentMediaItem.reload();
            _current_application = (PowerPoint.Application)_next_application;
            _current_presentation = (PowerPoint.Presentation)_next_presentation;
            PowerPoint.SlideShowSettings settings = this._current_presentation.SlideShowSettings;

            settings.ShowType = (PowerPoint.PpSlideShowType)1;
            settings.ShowPresenterView = MsoTriState.msoFalse;
            PowerPoint.SlideShowWindow sw = settings.Run();

            this._current_presentation.SlideShowWindow.View.GotoSlide(this._currentSlideIndex + 1);
            this._current_presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();

            if (this._prev_presentation != null)
            {
                this._stopPresentation(_prev_presentation);
            }


            this._isPlaying = true;
        }

        public void stopPresentation()
        {
            //String file = _presentation.Path + "\\" + _presentation.Name;
            this._stopPresentation(this._current_presentation);
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
            this._current_presentation.SlideShowWindow.View.Next();
            this._focus();
        }

        public void prevSlide()
        {
            this._current_presentation.SlideShowWindow.View.Previous();
            this._focus();
        }

        public void goToSlideIndex(int _index)
        {
            this._current_presentation.SlideShowWindow.View.GotoSlide(_index + 1);
        }

        public bool exitSlide()
        {
            if(this._isPlaying && this._validPresentation())
            {
                try
                {
                    int tmp = this._current_presentation.SlideShowWindow.View.Slide.SlideIndex;
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

        public String getSlideNotes()
        {
            return this._slideNotes;
        }

        public void focus()
        {
            this._focus();
        }

        public bool runInterval()
        {
            this._intervalCounter++;
            this._setSlideIndex();
            if(this._intervalCounter >= 50)
            {
                this._intervalCounter = 0;
                this._focus();
                this._getSlideNotes();
                return true;
            }
            return false;
        }

        private void _focus()
        {
            if(this._isPlaying && this._validPresentation())
            {
                this._current_presentation.SlideShowWindow.Activate();
            }
        }

        private void _getSlideNotes()
        {
            String notes = String.Empty;
            if (this.IsPlaying && !this.exitSlide() && this._validPresentation())
            {
                PowerPoint.Slide slide = this._current_presentation.Slides[_currentSlideIndex + 1];
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
                                Console.WriteLine(range.Text);
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
            this._slideNotes = notes;
        }
        private void _setSlideIndex()
        {
            this._currentSlideIndex = 0;
            if (this._isPlaying && this._validPresentation())
            {
                if (this._currentMediaItem.type == MediaType.powerpoint)
                {
                    try
                    {
                        this._currentSlideIndex = this._current_presentation.SlideShowWindow.View.Slide.SlideIndex - 1;
                    } catch
                    { }
                    
                }
            }
        }

        private bool _validPresentation()
        {
            if (this._currentMediaItem.type != MediaType.powerpoint)
            {
                return true;
            }
            try
            {
                int tmp = this._current_presentation.Slides.Count;
            }
            catch
            {
                return false;
            }
            return true;
        }

    }
}
