using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideCat
{
    internal class SlideNotes
    {
        private PictureBox _pictureBoxCurrent;
        private PictureBox _pictureBoxNext;

        private Label _labelCurrent;
        private Label _labelNext;


        public SlideNotes(PictureBox pictureBoxCurrent, PictureBox pictureBoxNext, Label labelCurrent, Label labelNext)
        {
            this._pictureBoxCurrent = pictureBoxCurrent;
            this._pictureBoxNext = pictureBoxNext;
            this._labelCurrent = labelCurrent;
            this._labelNext = labelNext;
        }

        public void trigger()
        {

        }
    }
}
