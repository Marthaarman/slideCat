using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace SlideCat
{
    public partial class form_main : Form
    {
        //private ArrayList mediaItems = new ArrayList();
        private readonly MediaItems _mediaItems = new MediaItems();
        private int _percentage = 0;
        private readonly SlideCatPresentation _presentation = new SlideCatPresentation();


        private BackgroundWorker backgroundWorker_createPresentation;
        private BackgroundWorker backgroundWorker_statusPresentation;

        public form_main()
        {
            InitializeComponent();
            InitializeBackgroundWorkers();
            FormClosing += _formClosing;
            progressBar.Visible = false;
        }

        private void _formClosing(object sender, FormClosingEventArgs e)
        {
            _presentation.stop();
        }

        private void InitializeBackgroundWorkers()
        {
            backgroundWorker_createPresentation = new BackgroundWorker();
            backgroundWorker_createPresentation.WorkerReportsProgress = true;
            backgroundWorker_createPresentation.DoWork += backgroundWorker_createPresentation_DoWork;
            backgroundWorker_createPresentation.RunWorkerCompleted +=
                backgroundWorker_createPresentation_RunWorkerCompleted;
            backgroundWorker_createPresentation.ProgressChanged += backgroundWorker_createPresentation_ProgressChanged;


            backgroundWorker_statusPresentation = new BackgroundWorker();
            backgroundWorker_statusPresentation.DoWork += backgroundWorker_statusPresentation_DoWork;
            backgroundWorker_statusPresentation.RunWorkerCompleted +=
                backgroundWorker_statusPresentation_RunWorkerCompleted;
        }

        private void backgroundWorker_createPresentation_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            _presentation.createPresentation(_mediaItems, ref worker);
        }

        private void backgroundWorker_createPresentation_RunWorkerCompleted(object sender,
            RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null) MessageBox.Show(e.Error.Message);
            progressBar.Value = 100;
            _presentation.playPresentation();
            backgroundWorker_statusPresentation.RunWorkerAsync();
            progressBar.Visible = false;
        }


        private void backgroundWorker_createPresentation_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var percentage = e.ProgressPercentage;
            progressBar.Value = percentage;
        }

        private void backgroundWorker_statusPresentation_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true)
                if (!_presentation.IsPlaying)
                    break;
        }

        private void backgroundWorker_statusPresentation_RunWorkerCompleted(object sender,
            RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null) MessageBox.Show(e.Error.Message);
            _presentationStopped();
        }

        private void button_mediaItem_add_click(object sender, EventArgs e)
        {
            /*SOURCE: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.openfiledialog?view=windowsdesktop-6.0*/
            //  Variables the file info is stored to
            var fileContent = string.Empty;
            var filePath = string.Empty;

            //  set image thumb sizes
            //setImageListSize();
            Cursor = Cursors.WaitCursor;

            //  setup dialog
            using (var openFileDialog = new OpenFileDialog())
            {
                //  File type filter
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = false;

                //  open the dialog and if result found
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
            }
            /* END SOURCE */

            //  process file
            if (filePath != string.Empty)
            {
                _mediaItems.addFile(filePath);
                reloadMediaItems();
            }

            Cursor = Cursors.Default;
        }


        private void reloadMediaItems()
        {
            comboBox_mediaItems.Items.Clear();
            foreach (MediaItem item in _mediaItems.mediaItems) comboBox_mediaItems.Items.Add(item.name);
        }

        private void button_mediaItem_moveUp_Click(object sender, EventArgs e)
        {
            var _selected_index = comboBox_mediaItems.SelectedIndex;
            _mediaItems.moveMediaItem(_selected_index, -1);
            reloadMediaItems();
            comboBox_mediaItems.SelectedIndex = _selected_index - 1;
        }

        private void button_mediaItem_moveDown_Click(object sender, EventArgs e)
        {
            var _selected_index = comboBox_mediaItems.SelectedIndex;
            _mediaItems.moveMediaItem(_selected_index, 1);
            reloadMediaItems();
            comboBox_mediaItems.SelectedIndex = _selected_index + 1;
        }

        private void button_mediaItem_remove_Click(object sender, EventArgs e)
        {
            var _selected_index = comboBox_mediaItems.SelectedIndex;
            _mediaItems.removeMediaItem(_selected_index);
            reloadMediaItems();
            if (comboBox_mediaItems.Items.Count > 0) comboBox_mediaItems.SelectedIndex = 0;
        }

        private void button_control_start_Click(object sender, EventArgs e)
        {
            if (_presentation.IsPlaying == false)
            {
                var _selected_index = comboBox_mediaItems.SelectedIndex;
                if (_selected_index < 0) _selected_index = 0;

                if (comboBox_mediaItems.Items.Count > 0)
                {
                    Cursor = Cursors.WaitCursor;

                    button_control_start.Text = "Processing, please wait.";
                    button_control_start.Enabled = false;
                    progressBar.Visible = true;
                    progressBar.Value = 0;
                    backgroundWorker_createPresentation.RunWorkerAsync();
                    _setControlsStateStart();
                }

                Cursor = Cursors.Default;
            }
        }

        private void _presentationStopped()
        {
            _setControlsStateSstop();
        }


        private void _setControlsStateStart()
        {
            button_control_start.Text = "Started, press [esc] to stop";
            button_control_start.Enabled = false;
            button_mediaItem_add.Enabled = false;
            button_mediaItem_remove.Enabled = false;
            button_mediaItem_moveDown.Enabled = false;
            button_mediaItem_moveUp.Enabled = false;
        }

        private void _setControlsStateSstop()
        {
            button_control_start.Text = "Start";
            button_control_start.Enabled = true;
            button_mediaItem_add.Enabled = true;
            button_mediaItem_remove.Enabled = true;
            button_mediaItem_moveDown.Enabled = true;
            button_mediaItem_moveUp.Enabled = true;
        }
    }
}