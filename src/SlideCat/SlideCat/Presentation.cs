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
        private bool _isPlaying = false;

        private PowerPoint.Application _current_application;
        private PowerPoint.Presentation _current_presentation;
        private PowerPoint.SlideShowWindow _current_slideShowWindow;


        public bool IsPlaying { get { return _isPlaying; } }

        public void loadPresentationItem(MediaItem _item)
        {
            _currentMediaItem = _item;
            if(_item.type == MediaType.powerpoint)
            {
                _current_application = _item.application;
                _current_presentation = _item.presentation;
                
            }
        }

        public void playPresentation()
        {
            PowerPoint.SlideShowSettings settings = this._current_presentation.SlideShowSettings;

            settings.ShowType = (PowerPoint.PpSlideShowType)1;
            settings.ShowPresenterView = MsoTriState.msoFalse;
            PowerPoint.SlideShowWindow sw = settings.Run();
            sw.View.AcceleratorsEnabled = MsoTriState.msoFalse;

            this._current_presentation.SlideShowWindow.View.GotoSlide(this._currentSlideIndex + 1);
            this._current_presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();

            this._isPlaying = true;
            
        }

        public void stopPresentation()
        {
            try
            {
                String file = this._current_presentation.Path + "\\" + this._current_presentation.Name;
                this._current_presentation.Close();
                this._current_presentation = this._currentMediaItem.presentation;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            this._isPlaying = false;
        }

        public void nextSlide()
        {
            this._current_presentation.SlideShowWindow.View.Next();
        }

        public void prevSlide()
        {
            this._current_presentation.SlideShowWindow.View.Previous();
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
            if(this._currentMediaItem.type == MediaType.powerpoint && this._isPlaying)
            {
                try
                {
                    return this._current_presentation.SlideShowWindow.View.Slide.SlideIndex - 1;
                }
                catch (Exception ex)
                {
                    return this._current_presentation.Slides.Count;
                }
                
            }

            if(this._currentMediaItem.type == MediaType.powerpoint)
            {
                return 1;
            }else
            {
                return 2;
            }

            return 0;
        }
    }
}
