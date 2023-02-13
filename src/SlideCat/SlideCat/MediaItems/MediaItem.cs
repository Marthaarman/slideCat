using System;
using System.Collections;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    public class MediaItem
    {
        private readonly int _id;

        private readonly string _mSrc = string.Empty;
        private readonly string _mName = string.Empty;

        private readonly MediaType _mMediaType = MediaType.unsupported;

        private PPT _mPpt;
        private int _mOrder = 0;

        public MediaItem(string src, int order)
        {
            if (!File.Exists(src))
            {
                return;
            }
            
            _mSrc = src;
            _mOrder = order; 
            _mName = Path.GetFileName(src);
            _id = new Random().Next(1, 1000);

            switch (Path.GetExtension(src))
            {
                case ".pptx":
                case ".ppt":
                    _mMediaType = MediaType.powerpoint;
                    break;
                case ".mov":
                case ".mp4":
                case ".mp3":
                case ".avi":
                    _mMediaType = MediaType.video;
                    break;
                case ".jpg":
                case ".jpeg":
                case ".JPEG":
                case ".JPG":
                case ".png":
                case ".gif":
                    _mMediaType = MediaType.image;
                    break;
                case ".pdf":
                    _mMediaType = MediaType.pdf;
                    break;
                default:
                    _mMediaType = MediaType.unsupported;
                    break;
            }
            
        }

        public void Load()
        {
            switch (_mMediaType)
            {
                case MediaType.video:
                    _mPpt = new PPTVideo();
                    break;
                case MediaType.image:
                    _mPpt = new PPTImage();
                    break;
                case MediaType.pdf:
                    _mPpt = new PPTPDF();
                    break;
                case MediaType.powerpoint:
                    _mPpt = new PPT();
                    break;
                case MediaType.unsupported:
                default:
                    _mPpt = null;
                    break;
            }

            if (_mMediaType == MediaType.unsupported || _mPpt == null)
            {
                return;
            }
            
            _mPpt.Load(_mSrc);
            _mPpt.CreatePresentation();
            
        }

        public Presentation presentation => _mPpt.GetPresentation();

        public string name => _mName;


        public int order
        {
            get => _mOrder;
            set => _mOrder = value;
        }

        public bool valid => _mMediaType != MediaType.unsupported;

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