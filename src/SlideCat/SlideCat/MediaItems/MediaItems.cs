using System.Collections;

namespace SlideCat
{
    public class ItemManager
    {
        ///<summary>
        /// returns a list of media items selected for the tool
        /// </summary>
        public ArrayList mMediaItems { get; } = new ArrayList();

        /// <summary>adds a mediaItem to the list of media items</summary>
        /// <param name="mediaItem">The mediaItem to add to the list</param>
        public void Add(MediaItem mediaItem)
        {
            mMediaItems.Add(mediaItem);
        }

        /// <summary>takes a file and converts it into a mediaItem, then adds it to the list</summary>
        /// <param name="file">a file to be added to the list of mediaItems</param>
        public void AddFile(string file)
        {
            MediaItem mediaItem = new MediaItem(file, mMediaItems.Count);
            if (mediaItem.valid) Add(mediaItem);
        }

        /// <summary>Sorts the list of mediaItems by the given order</summary> 
        public void Sort()
        {
            mMediaItems.Sort(new MediaItemComparer());
        }

        /// <summary>Moves a mediaItem in the order list (both up and down)</summary>
        /// <param name="positionA">positionA is the current position of the media item</param>
        /// <param name="delta">
        /// delta is the difference with which it has to move. positive value for down and negative for up
        /// </param>
        public void MoveMediaItem(int positionA, int delta)
        {
            int nextIndex = positionA + delta;
            if (nextIndex >= 0 && nextIndex < mMediaItems.Count && positionA >= 0)
                SwapMediaItems(positionA, nextIndex);
        }

        /// <summary> swaps two items of position, takes items A and B</summary>
        /// <param name="indexA">indexA is the current position of item A</param>
        /// <param name="indexB">indexB is the current position of item B</param>
        public void SwapMediaItems(int indexA, int indexB)
        {
            if (indexA < 0 || indexB < 0) return;
            MediaItem itemA = (MediaItem)mMediaItems[indexA];
            MediaItem itemB = (MediaItem)mMediaItems[indexB];
            itemA.order = indexB;
            itemB.order = indexA;
            mMediaItems[indexA] = itemA;
            mMediaItems[indexB] = itemB;
            Sort();
        }

        /// <summary>removes an item from the list of items</summary>
        /// <param name="index">takes the current index of the item to be removed from the list</param>
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
            Sort();
        }
    }
}