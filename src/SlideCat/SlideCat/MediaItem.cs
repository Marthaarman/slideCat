using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class MediaItem
    {

        private String _name = null;
        private String _src = null;
        private int _nrSlides = 1;
        private int _order = 0;
        private int _id = 0;
        private MediaType _type = MediaType.unsupported;

        private PowerPoint.Application _application;
        private PowerPoint.Presentation _presentation;

        public PowerPoint.Application application { get { return _application; } }
        public PowerPoint.Presentation presentation { get { return _presentation; } }

        public String name
        {
            get { return this._name; }
        }
            
        public string src
        {
            get { return this._src; }
        }

        public int order
        {
            get { return this._order; }
            set { this._order = value; }
        }
        
        public MediaType type
        {
            get { return this._type; }
        }

        public bool valid
        {
            get {  return (this._type != MediaType.unsupported); }
        }

        public int nrSlides
        {
            get { return this._nrSlides; }
        }

        public MediaItem(String src, int order)
        {
            if(File.Exists(src))
            {
                this._src = src;
                this._order = order;
                this._name = Path.GetFileName(src);
                this._id = new Random().Next(1, 1000);
                switch(Path.GetExtension(src))
                {
                    case ".pptx":
                    case ".ppt":
                        this._type = MediaType.powerpoint;
                        this._application = new PowerPoint.Application();
                        this._presentation = _application.Presentations.Open2007(this._src, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
                        PowerPoint.SlideRange dupliSlide = this._presentation.Slides[this._presentation.Slides.Count].Duplicate();
                        dupliSlide.MoveTo(this._presentation.Slides.Count);
                        this._nrSlides = this._presentation.Slides.Count;
                        this.setThumbs();
                        break;
                    case ".mov":
                    case ".mp4":
                    case ".mp3":
                    case ".avi":
                        this._type = MediaType.video;
                        break;
                    case ".jpg":
                    case ".png":
                    case ".gif":
                        this._type = MediaType.image;
                        break;
                    case ".pdf":
                        this._type = MediaType.pdf;
                        break;
                    default:
                        this._type = MediaType.unsupported;
                        break;
                } 
                
            }
        }

        public void reload()
        {
            if(this._presentation != null && this._type == MediaType.powerpoint)
            {
                this._application = new PowerPoint.Application();
                this._presentation = this._application.Presentations.Open2007(this._src, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
            }
        }

        public void setThumbs()
        {
            
            if(this._type == MediaType.powerpoint)
            {
                String _path = Path.GetTempPath() + "SlideCat\\";
                if (!Directory.Exists(_path))
                {
                    Directory.CreateDirectory(_path);
                }

                _path += this._id + "\\";
                if(Directory.Exists(_path))
                {
                    Directory.Delete(_path, true);
                }
                Directory.CreateDirectory(_path);

                foreach (PowerPoint.Slide slide in this._presentation.Slides)
                {
                    String src = _path + slide.SlideIndex + ".jpg";
                    slide.Export(src, "jpg", 1080, 960);
                }    
            }
        }

        public String getThumb(int _index)
        {
            return Path.GetTempPath() + "SlideCat\\" + this._id + "\\" + _index + ".jpg";
        }

    }

    public enum MediaType
    {
        powerpoint,
        image,
        video,
        pdf,
        unsupported
    }

    

    public class MediaItemComparer : IComparer
    {
        int IComparer.Compare(Object _x, Object _y)
        {
            MediaItem _item_x = (MediaItem)_x;
            MediaItem _item_y = (MediaItem)_y;
            return _item_x.order.CompareTo(_item_y.order);
        }
    }
}

