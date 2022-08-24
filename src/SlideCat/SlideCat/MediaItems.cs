using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace SlideCat
{
    public class MediaItems
    {
        private ArrayList _mediaItems = new ArrayList();

        public ArrayList mediaItems { get { return _mediaItems; } }

        public void add(MediaItem _mediaItem)
        {
            this._mediaItems.Add(_mediaItem);
        }

        public void addFile(String file)
        {
            MediaItem mediaItem = new MediaItem(file, this._mediaItems.Count);
            if (mediaItem.valid)
            {
                this.add(mediaItem);
            }
        }

        public void sort()
        {
            this._mediaItems.Sort(new MediaItemComparer());
        }

        public void updateMediaItems()
        {
            this.sort();
        }

        public void moveMediaItem(int _positionA, int _delta)
        {
            int _next_index = _positionA + _delta;
            if (_next_index >= 0 && _next_index < this.mediaItems.Count && _positionA >= 0)
            {
                swapMediaItems(_positionA, _next_index);
            }
        }

        public void swapMediaItems(int _indexA, int _indexB)
        {
            if (_indexA >= 0 && _indexB >= 0)
            {
                MediaItem _itemA = (MediaItem)mediaItems[_indexA];
                MediaItem _itemB = (MediaItem)mediaItems[_indexB];
                _itemA.order = _indexB;
                _itemB.order = _indexA;
                mediaItems[_indexA] = _itemA;
                mediaItems[_indexB] = _itemB;
                this.updateMediaItems();
            }
        }

        public void removeMediaItem(int _index)
        {
            if (_index >= 0)
            {
                this.mediaItems.RemoveAt(_index);

                for (int i = _index + 1; i < this.mediaItems.Count; i++)
                {
                    MediaItem _item = (MediaItem)mediaItems[i];
                    _item.order -= 1;
                    mediaItems[i] = _item;
                }
                this.updateMediaItems();
            }
        }

        public MediaItem get(int _index)
        {
            return (MediaItem)this._mediaItems[_index];
        }
    }
}
