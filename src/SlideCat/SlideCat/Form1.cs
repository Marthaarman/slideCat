using System;
using System.ComponentModel;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Color = System.Drawing.Color;

namespace SlideCat
{
    public partial class form_main : Form
    {
        //private ArrayList ItemManager = new ArrayList();
        private readonly ItemManager _mItemManager = new ItemManager();
        private readonly SlideCatPresentation _mPresentation = new SlideCatPresentation();

        private BackgroundWorker _mBackgroundWorkerCreatePresentation;
        private BackgroundWorker _mBackgroundWorkerStatusPresentation;

        

        public form_main()
        {
            InitializeComponent();
            InitializeBackgroundWorkers();

            FormClosing += _formClosing;
            progressBar.Visible = false;

            
            _CheckPresenterView();
            checkBox_presenterview.CheckStateChanged += _ChangePresenterView;
        }

        private void _CheckPresenterView()
        {
            Screen[] screens = Screen.AllScreens;
            if (screens.Length <= 1)
            {
                checkBox_presenterview.Checked = false;
                checkBox_presenterview.Enabled = false;
                _mPresentation.PresenterView(false);
            }
        }

        private void _ChangePresenterView(object sender, EventArgs e)
        {
            _mPresentation.PresenterView(checkBox_presenterview.Checked);
        }

        private void _formClosing(object sender, FormClosingEventArgs e)
        {
            _mPresentation.Stop();
        }

        private void InitializeBackgroundWorkers()
        {
            _mBackgroundWorkerCreatePresentation = new BackgroundWorker();
            _mBackgroundWorkerCreatePresentation.WorkerReportsProgress = true;
            _mBackgroundWorkerCreatePresentation.DoWork += BackgroundWorkerCreatePresentationDoWork;
            _mBackgroundWorkerCreatePresentation.RunWorkerCompleted +=
                BackgroundWorkerCreatePresentationRunWorkerCompleted;
            _mBackgroundWorkerCreatePresentation.ProgressChanged += BackgroundWorkerCreatePresentationProgressChanged;


            _mBackgroundWorkerStatusPresentation = new BackgroundWorker();
            _mBackgroundWorkerStatusPresentation.DoWork += BackgroundWorkerStatusPresentationDoWork;
            _mBackgroundWorkerStatusPresentation.RunWorkerCompleted +=
                BackgroundWorkerStatusPresentationRunWorkerCompleted;
        }

        private void BackgroundWorkerCreatePresentationDoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            _mPresentation.CreatePresentation(_mItemManager, ref worker);
        }

        private void BackgroundWorkerCreatePresentationRunWorkerCompleted(object sender,
            RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null) MessageBox.Show(e.Error.Message);
            progressBar.Value = 100;
            _mPresentation.PlayPresentation();
            _mBackgroundWorkerStatusPresentation.RunWorkerAsync();
            progressBar.Visible = false;
        }


        private void BackgroundWorkerCreatePresentationProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int percentage = e.ProgressPercentage;
            progressBar.Value = percentage;
        }

        private void BackgroundWorkerStatusPresentationDoWork(object sender, DoWorkEventArgs e)
        {
            while (!_mPresentation.mIsPlaying)
            {
            }
        }

        private void BackgroundWorkerStatusPresentationRunWorkerCompleted(object sender,
            RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null) MessageBox.Show(e.Error.Message);
            _PresentationStopped();
        }

        private void button_mediaItem_add_click(object sender, EventArgs e)
        {
            /*SOURCE: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.openfiledialog?view=windowsdesktop-6.0*/
            //  Variables the file info is stored to
            string filePath = string.Empty;

            //  set image thumb sizes
            //setImageListSize();
            Cursor = Cursors.WaitCursor;

            //  setup dialog
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
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
                _mItemManager.AddFile(filePath);
                ReloadMediaItems();
            }

            Cursor = Cursors.Default;
        }


        private void ReloadMediaItems()
        {
            comboBox_mediaItems.Items.Clear();
            foreach (MediaItem item in _mItemManager.mMediaItems) comboBox_mediaItems.Items.Add(item.name);
            if (_mItemManager.mMediaItems.Count > 0)
            {
                button_control_start.BackColor = Color.ForestGreen;
                button_control_start.Enabled = true;
            }
            else
            {
                button_control_start.BackColor = Color.Salmon;
                button_control_start.Enabled = false;
            }
        }

        private void _PresentationStopped()
        {
            _SetControlsStateStop();
        }


        private void _SetControlsStateStart()
        {
            button_control_start.Text = "Started, press [esc] to stop";
            button_control_start.BackColor = Color.Salmon;
            button_control_start.Enabled = false;
            button_mediaItem_add.Enabled = false;
            button_mediaItem_remove.Enabled = false;
            button_mediaItem_moveDown.Enabled = false;
            button_mediaItem_moveUp.Enabled = false;
        }

        private void _SetControlsStateStop()
        {
            button_control_start.Text = "Start";
            button_control_start.Enabled = true;
            button_control_start.BackColor = Color.ForestGreen;
            button_mediaItem_add.Enabled = true;
            button_mediaItem_remove.Enabled = true;
            button_mediaItem_moveDown.Enabled = true;
            button_mediaItem_moveUp.Enabled = true;
        }
        private void button_mediaItem_moveUp_Click(object sender, EventArgs e)
        {
            int selectedIndex = comboBox_mediaItems.SelectedIndex;
            _mItemManager.MoveMediaItem(selectedIndex, -1);
            ReloadMediaItems();
            comboBox_mediaItems.SelectedIndex = selectedIndex - 1;
        }

        private void button_mediaItem_moveDown_Click(object sender, EventArgs e)
        {
            int selectedIndex = comboBox_mediaItems.SelectedIndex;
            _mItemManager.MoveMediaItem(selectedIndex, 1);
            ReloadMediaItems();
            comboBox_mediaItems.SelectedIndex = selectedIndex + 1;
        }

        private void button_mediaItem_remove_Click(object sender, EventArgs e)
        {
            int selectedIndex = comboBox_mediaItems.SelectedIndex;
            _mItemManager.RemoveMediaItem(selectedIndex);
            ReloadMediaItems();
            if (comboBox_mediaItems.Items.Count > 0) comboBox_mediaItems.SelectedIndex = 0;
        }

        private void button_control_start_Click(object sender, EventArgs e)
        {
            if (_mPresentation.mIsPlaying) return;
            if (comboBox_mediaItems.Items.Count > 0)
            {
                Cursor = Cursors.WaitCursor;

                button_control_start.Text = "Processing, please wait.";
                button_control_start.Enabled = false;
                button_control_start.BackColor = Color.Salmon;
                progressBar.Visible = true;
                progressBar.Value = 0;
                _mBackgroundWorkerCreatePresentation.RunWorkerAsync();
                _SetControlsStateStart();
            }

            Cursor = Cursors.Default;
            
        }

    }
}