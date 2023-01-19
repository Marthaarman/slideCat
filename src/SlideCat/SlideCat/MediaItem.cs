using System;
using System.Collections;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class MediaItem
    {
        private readonly int _id;

        public MediaItem(string src, int order)
        {
            if (File.Exists(src))
            {
                this.src = src;
                this.order = order;
                name = Path.GetFileName(src);
                _id = new Random().Next(1, 1000);

                PPT ppt = new PPT();

                switch (Path.GetExtension(src))
                {
                    case ".pptx":
                    case ".ppt":
                        type = MediaType.powerpoint;
                        ppt = new PPT();
                        break;
                    case ".mov":
                    case ".mp4":
                    case ".mp3":
                    case ".avi":
                        type = MediaType.video;
                        ppt = new PPTVideo();

                        break;
                    case ".jpg":
                    case ".jpeg":
                    case ".JPEG":
                    case ".JPG":
                    case ".png":
                    case ".gif":
                        type = MediaType.image;
                        ppt = new PPTImage();
                        break;
                    case ".pdf":
                        //currently unsupported
                        type = MediaType.pdf;
                        ppt = new PPTPDF();
                        break;
                    default:
                        type = MediaType.unsupported;
                        break;
                }

                if (type != MediaType.unsupported)
                {
                    ppt.Load(src);
                    ppt.createPresentation();
                    presentation = ppt.getPresentation();
                    nrSlides = ppt.nrSlides;
                }
            }
        }

        public Presentation presentation { get; }

        public Slides slides => presentation.Slides;

        public string name { get; }

        public string src { get; }

        public int order { get; set; }

        public MediaType type { get; } = MediaType.unsupported;

        public bool valid => type != MediaType.unsupported;

        public int nrSlides { get; } = 1;


        public void setThumbs()
        {
            if (type == MediaType.powerpoint)
            {
                string _path = Path.GetTempPath() + "SlideCat\\";
                if (!Directory.Exists(_path)) Directory.CreateDirectory(_path);

                _path += _id + "\\";
                if (Directory.Exists(_path)) Directory.Delete(_path, true);
                Directory.CreateDirectory(_path);

                foreach (Slide slide in presentation.Slides)
                {
                    string src = _path + slide.SlideIndex + ".jpg";
                    slide.Export(src, "jpg", 1080, 960);
                }
            }
        }

        public string getThumb(int _index)
        {
            return Path.GetTempPath() + "SlideCat\\" + _id + "\\" + _index + ".jpg";
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
        int IComparer.Compare(object _x, object _y)
        {
            MediaItem _item_x = (MediaItem)_x;
            MediaItem _item_y = (MediaItem)_y;
            return _item_x.order.CompareTo(_item_y.order);
        }
    }
}