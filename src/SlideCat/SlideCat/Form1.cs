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
                worker.ReportProgress(_presentation.slide());
                Thread.Sleep(10);
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
            else if (e.Cancelled)
            {
                
            }
            else
            {
                //Console.WriteLine(e.Result.ToString());
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if(_presentation.lastSlide())
            {
                _presentation.prevSlide();
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
            if(_presentation.IsPlaying == false)
            {
                this.backgroundWorker1.RunWorkerAsync();
                _presentation.loadPresentationItem((MediaItem)this._mediaItems.get(0));
                _presentation.playPresentation();
            }
        }

        private void button_control_stop_Click(object sender, EventArgs e)
        {
            if(_presentation.IsPlaying)
            {
                this.backgroundWorker1.CancelAsync();
                _presentation.stopPresentation();
            }
        }
    }
}
