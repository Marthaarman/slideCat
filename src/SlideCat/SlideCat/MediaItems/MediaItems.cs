using System.Collections;

namespace SlideCat
{
    public class ItemManager
    {
        public ArrayList mMediaItems { get; } = new ArrayList();

        public void Add(MediaItem mediaItem)
        {
            mMediaItems.Add(mediaItem);
        }

        public void AddFile(string file)
        {
            MediaItem mediaItem = new MediaItem(file, mMediaItems.Count);
            if (mediaItem.valid) Add(mediaItem);
        }

        public void Sort()
        {
            mMediaItems.Sort(new MediaItemComparer());
        }

        public void UpdateMediaItems()
        {
            Sort();
        }

        public void MoveMediaItem(int positionA, int delta)
        {
            int nextIndex = positionA + delta;
            if (nextIndex >= 0 && nextIndex < mMediaItems.Count && positionA >= 0)
                SwapMediaItems(positionA, nextIndex);
        }

        public void SwapMediaItems(int indexA, int indexB)
        {
            if (indexA < 0 || indexB < 0) return;
            MediaItem itemA = (MediaItem)mMediaItems[indexA];
            MediaItem itemB = (MediaItem)mMediaItems[indexB];
            itemA.order = indexB;
            itemB.order = indexA;
            mMediaItems[indexA] = itemA;
            mMediaItems[indexB] = itemB;
            UpdateMediaItems();
        }

        public void RemoveMediaItem(int index)
        {
            if (index < 0) return;
            mMediaItems.RemoveAt(index);
            for (int i = index + 1; i < mMediaItems.Count; i++)
            {
                MediaItem item = (MediaItem)mMediaItems[i];
                item.order -= 1;
                mMediaItems[i] = item;
            }
            UpdateMediaItems();
        }
    }
}