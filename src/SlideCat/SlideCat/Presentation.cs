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

        private int _focusCounter = 0;


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

            _prev_application = _current_application;
            _prev_presentation = _current_presentation;

            _currentMediaItem = _nextMediaItem;
            _currentMediaItem.reload();
            _current_application = (PowerPoint.Application)_next_application;
            _current_presentation = (PowerPoint.Presentation)_next_presentation;
            PowerPoint.SlideShowSettings settings = this._current_presentation.SlideShowSettings;

            settings.ShowType = (PowerPoint.PpSlideShowType)1;
            settings.ShowPresenterView = MsoTriState.msoFalse;
            PowerPoint.SlideShowWindow sw = settings.Run();
            //sw.View.AcceleratorsEnabled = MsoTriState.msoFalse;

            this._current_presentation.SlideShowWindow.View.GotoSlide(this._currentSlideIndex + 1);
            this._current_presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();


            if (this._isPlaying)
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

        public bool lastSlide()
        {
            return this._currentMediaItem.nrSlides == this.slide() ? true : false;
        }

        public bool firstSlide()
        {
            return false;
        }

        public int slide()
        {
            if (this._isPlaying) {
                if (this._currentMediaItem.type == MediaType.powerpoint)
                {
                    try
                    {
                        return this._current_presentation.SlideShowWindow.View.Slide.SlideIndex - 1;
                    }
                    catch (Exception ex1)
                    {
                        Console.WriteLine("LOG - Presentation.cs - slide() - catch ex1");
                        try
                        {
                            return this._current_presentation.Slides.Count;
                        } catch (Exception ex2)
                        {
                            Console.WriteLine("LOG - Presentation.cs - slide() - catch ex2");
                            Console.WriteLine(ex2.Message);
                            return 0;
                        }
                    }
                }
            }
            return 0;
        }

        public void focus()
        {
            this._focusCounter++;
            if(this._focusCounter >= 50)
            {
                this._focus();
            }
        }

        private void _focus()
        {
            if(this._isPlaying)
            {
                this._focusCounter = 0;
                this._current_presentation.SlideShowWindow.Activate();
            }
            
        }
    }
}
