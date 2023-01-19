using System.Collections;

namespace SlideCat
{
    public class MediaItems
    {
        public ArrayList mediaItems { get; } = new ArrayList();

        public void add(MediaItem _mediaItem)
        {
            mediaItems.Add(_mediaItem);
        }

        public void addFile(string file)
        {
            MediaItem mediaItem = new MediaItem(file, mediaItems.Count);
            if (mediaItem.valid) add(mediaItem);
        }

        public void sort()
        {
            mediaItems.Sort(new MediaItemComparer());
        }

        public void updateMediaItems()
        {
            sort();
        }

        public void moveMediaItem(int _positionA, int _delta)
        {
            int _next_index = _positionA + _delta;
            if (_next_index >= 0 && _next_index < mediaItems.Count && _positionA >= 0)
                swapMediaItems(_positionA, _next_index);
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
                updateMediaItems();
            }
        }

        public void removeMediaItem(int _index)
        {
            if (_index >= 0)
            {
                mediaItems.RemoveAt(_index);

                for (int i = _index + 1; i < mediaItems.Count; i++)
                {
                    MediaItem _item = (MediaItem)mediaItems[i];
                    _item.order -= 1;
                    mediaItems[i] = _item;
                }

                updateMediaItems();
            }
        }

        public MediaItem get(int _index)
        {
            if (_index >= 0 && _index < mediaItems.Count)
                return (MediaItem)mediaItems[_index];
            return (MediaItem)mediaItems[0];
        }
    }
}