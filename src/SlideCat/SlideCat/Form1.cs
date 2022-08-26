using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SlideCat
{
    public partial class form_main : Form
    {

        //private ArrayList mediaItems = new ArrayList();
        private MediaItems _mediaItems = new MediaItems();
        private Presentation _presentation = new Presentation();

        private System.ComponentModel.BackgroundWorker backgroundWorker1;

        public form_main()
        {
            InitializeComponent();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            InitializeBackgroundWorker();

        }

        private void InitializeBackgroundWorker()
        {
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            while (true)
            {
                worker.ReportProgress(_presentation.getSlideIndex());
                Thread.Sleep(100);
                if (worker.CancellationPending)
                {
                    break;
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            this.label_slideNotes.Text = "";
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if(_presentation.turnOverSlide())
            {
                if(_presentation.presentationItemOrder() == (_mediaItems.mediaItems.Count - 1)) {
                    _presentation.nextSlide();
                }else
                {
                    _presentation.loadNextPresentationItem(_mediaItems.get(_presentation.presentationItemOrder() + 1));
                    _presentation.playPresentation();
                }
            } else
            {
                if(!_presentation.validPresentation)
                {
                    backgroundWorker1.CancelAsync();
                    this._presentation.stopPresentation();
                }
            }

            if(_presentation.runInterval())
            {
                this.label_slideNotes.Text = _presentation.getSlideNotes();
                String thumbURL = this._presentation.getThumb();
                if (thumbURL != String.Empty)
                {
                    this.pictureBox_currentSlideThumb.Image = new Bitmap(thumbURL);
                }else
                {
                    this.pictureBox_currentSlideThumb.Image=null;
                }

                thumbURL = this._presentation.getNextThumb();
                if(thumbURL != String.Empty)
                {
                    this.pictureBox_nextSlideThumb.Image = new Bitmap(thumbURL);
                }else
                {
                    this.pictureBox_nextSlideThumb.Image = null;
                }
                _presentation.focus();
            }
        }

        private void button_mediaItem_add_click(object sender, EventArgs e)
        {
            /*SOURCE: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.openfiledialog?view=windowsdesktop-6.0*/
            //  Variables the file info is stored to
            var fileContent = string.Empty;
            var filePath = string.Empty;

            //  set image thumb sizes
            //setImageListSize();

            //  setup dialog
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //  File type filter
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = false;

                //  open the dialog and if result found
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
                }
            }
            /* END SOURCE */

            //  process file
            if (filePath != String.Empty)
            {
                _mediaItems.addFile(filePath);
                this.reloadMediaItems();
            }
        }

       

        private void reloadMediaItems()
        {
            this.comboBox_mediaItems.Items.Clear();
            foreach (MediaItem item in _mediaItems.mediaItems)
            {
                this.comboBox_mediaItems.Items.Add(item.name);
            }
        }

        private void button_mediaItem_moveUp_Click(object sender, EventArgs e)
        {
            int _selected_index = this.comboBox_mediaItems.SelectedIndex;
            _mediaItems.moveMediaItem(_selected_index, -1);
            this.reloadMediaItems();
            this.comboBox_mediaItems.SelectedIndex = _selected_index -1;
        }

        private void button_mediaItem_moveDown_Click(object sender, EventArgs e)
        {
            int _selected_index = this.comboBox_mediaItems.SelectedIndex;
            _mediaItems.moveMediaItem(_selected_index, 1);
            this.reloadMediaItems();
            this.comboBox_mediaItems.SelectedIndex = _selected_index + 1;
        }     

        private void button_mediaItem_remove_Click(object sender, EventArgs e)
        {
            int _selected_index = this.comboBox_mediaItems.SelectedIndex;
            _mediaItems.removeMediaItem(_selected_index);
            this.reloadMediaItems();
            if(this.comboBox_mediaItems.Items.Count > 0)
            {
                this.comboBox_mediaItems.SelectedIndex = 0;
            }
        }

        private void button_control_start_Click(object sender, EventArgs e)
        {
            if (_presentation.IsPlaying == false)
            {
                int _selected_index = this.comboBox_mediaItems.SelectedIndex;
                if (_selected_index < 0)
                {
                    _selected_index = 0;
                }
                if (this.comboBox_mediaItems.Items.Count > 0) 
                {
                    _presentation.loadNextPresentationItem((MediaItem)this._mediaItems.get(_selected_index));
                    _presentation.playPresentation();
                    this.backgroundWorker1.RunWorkerAsync();
                }
                
            }
        }

        private void button_control_stop_Click(object sender, EventArgs e)
        {
            if(_presentation.IsPlaying)
            {
                this.backgroundWorker1.CancelAsync();
                _presentation.stopPresentation();
                this.label_slideNotes.Text = String.Empty;
                this.label_slideNotesNext.Text = String.Empty;
                this.pictureBox_currentSlideThumb.Image = null;
                this.pictureBox_nextSlideThumb.Image = null;
            }
        }

        private void button_control_next_Click(object sender, EventArgs e)
        {
            if(_presentation.IsPlaying)
            {
                _presentation.nextSlide();
            }
        }

        private void button_control_previous_Click(object sender, EventArgs e)
        {
            if(_presentation.IsPlaying)
            {
                if(_presentation.getSlideIndex() == 0)
                {
                    int _next_index = _presentation.presentationItemOrder() - 1;
                    if(_next_index >= 0)
                    {
                        _presentation.loadNextPresentationItem(_mediaItems.get(_next_index));
                        _presentation.playPresentation();
                        _presentation.goToSlideIndex(_mediaItems.get(_next_index).nrSlides - 1);
                    }
                } else
                {
                    _presentation.prevSlide();
                }
                
            }
        }
    }
}
