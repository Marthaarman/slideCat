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
        private int _percentage = 0;
        

        private BackgroundWorker backgroundWorker_createPresentation;
        private BackgroundWorker backgroundWorker_statusPresentation;
        public form_main()
        {
            InitializeComponent();
            InitializeBackgroundWorkers();
            this.FormClosing += _formClosing;
            this.progressBar.Visible = false;
        }

        private void _formClosing(object sender, FormClosingEventArgs e)
        {
            this._presentation.stop();
        }

        private void InitializeBackgroundWorkers()
        {
            this.backgroundWorker_createPresentation = new BackgroundWorker();
            this.backgroundWorker_createPresentation.WorkerReportsProgress = true;
            this.backgroundWorker_createPresentation.DoWork += new DoWorkEventHandler(backgroundWorker_createPresentation_DoWork);
            this.backgroundWorker_createPresentation.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_createPresentation_RunWorkerCompleted);
            this.backgroundWorker_createPresentation.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker_createPresentation_ProgressChanged);


            this.backgroundWorker_statusPresentation = new BackgroundWorker();
            this.backgroundWorker_statusPresentation.DoWork += new DoWorkEventHandler(backgroundWorker_statusPresentation_DoWork);
            this.backgroundWorker_statusPresentation.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_statusPresentation_RunWorkerCompleted);
        }

        private void backgroundWorker_createPresentation_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            _presentation.createPresentation(this._mediaItems, ref worker);
        }

        private void backgroundWorker_createPresentation_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            this.progressBar.Value = 100;
            _presentation.playPresentation();
            this.backgroundWorker_statusPresentation.RunWorkerAsync();
            this.progressBar.Visible = false;
        }

        
        private void backgroundWorker_createPresentation_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            int percentage = e.ProgressPercentage;
            this.progressBar.Value = percentage;
        }

        private void backgroundWorker_statusPresentation_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true) { 
                if (!this._presentation.IsPlaying)
                {
                    break;
                }
            }
        }

        private void backgroundWorker_statusPresentation_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            this._presentationStopped();
        }
        



        private void button_mediaItem_add_click(object sender, EventArgs e)
        {
            /*SOURCE: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.openfiledialog?view=windowsdesktop-6.0*/
            //  Variables the file info is stored to
            var fileContent = string.Empty;
            var filePath = string.Empty;

            //  set image thumb sizes
            //setImageListSize();
            this.Cursor = Cursors.WaitCursor;

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

            this.Cursor=Cursors.Default;
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
                    this.Cursor = Cursors.WaitCursor;

                    this.button_control_start.Text = "Processing, please wait.";
                    this.button_control_start.Enabled = false;
                    this.progressBar.Visible = true;
                    this.progressBar.Value = 0;
                    this.backgroundWorker_createPresentation.RunWorkerAsync();
                    this._setControlsStateStart();

                }
                this.Cursor=Cursors.Default;
            }
        }

        private void _presentationStopped()
        {
            _setControlsStateSstop();
        }

        

        private void _setControlsStateStart()
        {
            this.button_control_start.Text = "Started, press [esc] to stop";
            this.button_control_start.Enabled = false;
            this.button_mediaItem_add.Enabled = false;
            this.button_mediaItem_remove.Enabled = false;
            this.button_mediaItem_moveDown.Enabled = false;
            this.button_mediaItem_moveUp.Enabled = false;
        }

        private void _setControlsStateSstop()
        {
            this.button_control_start.Text = "Start";
            this.button_control_start.Enabled = true;
            this.button_mediaItem_add.Enabled = true;
            this.button_mediaItem_remove.Enabled = true;
            this.button_mediaItem_moveDown.Enabled = true;
            this.button_mediaItem_moveUp.Enabled = true;
        }
    }
}
