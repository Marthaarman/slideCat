using System.Windows.Forms;

namespace SlideCat
{
    internal class SlideNotes
    {
        private Label _labelCurrent;
        private Label _labelNext;
        private PictureBox _pictureBoxCurrent;
        private PictureBox _pictureBoxNext;


        public SlideNotes(PictureBox pictureBoxCurrent, PictureBox pictureBoxNext, Label labelCurrent, Label labelNext)
        {
            _pictureBoxCurrent = pictureBoxCurrent;
            _pictureBoxNext = pictureBoxNext;
            _labelCurrent = labelCurrent;
            _labelNext = labelNext;
        }

        public void trigger()
        {
        }
    }
}