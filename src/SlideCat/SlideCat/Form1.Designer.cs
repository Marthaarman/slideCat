namespace SlideCat
{
    partial class form_main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel_main = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox_media_items = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.button_mediaItem_moveUp = new System.Windows.Forms.Button();
            this.button_mediaItem_moveDown = new System.Windows.Forms.Button();
            this.button_mediaItem_add = new System.Windows.Forms.Button();
            this.button_mediaItem_remove = new System.Windows.Forms.Button();
            this.comboBox_mediaItems = new System.Windows.Forms.ComboBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_control_start = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.label_version = new System.Windows.Forms.Label();
            this.tableLayoutPanel_main.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.groupBox_media_items.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel_main
            // 
            this.tableLayoutPanel_main.ColumnCount = 2;
            this.tableLayoutPanel_main.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 80F));
            this.tableLayoutPanel_main.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel_main.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel_main.Controls.Add(this.tableLayoutPanel1, 1, 0);
            this.tableLayoutPanel_main.Controls.Add(this.progressBar, 0, 1);
            this.tableLayoutPanel_main.Controls.Add(this.label_version, 1, 1);
            this.tableLayoutPanel_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel_main.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel_main.Name = "tableLayoutPanel_main";
            this.tableLayoutPanel_main.RowCount = 2;
            this.tableLayoutPanel_main.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel_main.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel_main.Size = new System.Drawing.Size(1099, 620);
            this.tableLayoutPanel_main.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Controls.Add(this.groupBox_media_items, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(873, 564);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // groupBox_media_items
            // 
            this.groupBox_media_items.Controls.Add(this.tableLayoutPanel3);
            this.groupBox_media_items.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox_media_items.Location = new System.Drawing.Point(3, 3);
            this.groupBox_media_items.Name = "groupBox_media_items";
            this.groupBox_media_items.Size = new System.Drawing.Size(867, 558);
            this.groupBox_media_items.TabIndex = 0;
            this.groupBox_media_items.TabStop = false;
            this.groupBox_media_items.Text = "Media items";
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 2;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 85F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel3.Controls.Add(this.tableLayoutPanel5, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.comboBox_mediaItems, 0, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 18);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(861, 537);
            this.tableLayoutPanel3.TabIndex = 0;
            // 
            // tableLayoutPanel5
            // 
            this.tableLayoutPanel5.ColumnCount = 1;
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel5.Controls.Add(this.button_mediaItem_moveUp, 0, 0);
            this.tableLayoutPanel5.Controls.Add(this.button_mediaItem_moveDown, 0, 1);
            this.tableLayoutPanel5.Controls.Add(this.button_mediaItem_add, 0, 3);
            this.tableLayoutPanel5.Controls.Add(this.button_mediaItem_remove, 0, 4);
            this.tableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel5.Location = new System.Drawing.Point(734, 3);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.RowCount = 5;
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel5.Size = new System.Drawing.Size(124, 531);
            this.tableLayoutPanel5.TabIndex = 0;
            // 
            // button_mediaItem_moveUp
            // 
            this.button_mediaItem_moveUp.BackColor = System.Drawing.Color.Silver;
            this.button_mediaItem_moveUp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_mediaItem_moveUp.FlatAppearance.BorderSize = 0;
            this.button_mediaItem_moveUp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_mediaItem_moveUp.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.button_mediaItem_moveUp.Location = new System.Drawing.Point(3, 3);
            this.button_mediaItem_moveUp.Name = "button_mediaItem_moveUp";
            this.button_mediaItem_moveUp.Size = new System.Drawing.Size(118, 100);
            this.button_mediaItem_moveUp.TabIndex = 0;
            this.button_mediaItem_moveUp.Text = "Move up";
            this.button_mediaItem_moveUp.UseVisualStyleBackColor = false;
            this.button_mediaItem_moveUp.Click += new System.EventHandler(this.button_mediaItem_moveUp_Click);
            // 
            // button_mediaItem_moveDown
            // 
            this.button_mediaItem_moveDown.BackColor = System.Drawing.Color.Silver;
            this.button_mediaItem_moveDown.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_mediaItem_moveDown.FlatAppearance.BorderSize = 0;
            this.button_mediaItem_moveDown.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_mediaItem_moveDown.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.button_mediaItem_moveDown.Location = new System.Drawing.Point(3, 109);
            this.button_mediaItem_moveDown.Name = "button_mediaItem_moveDown";
            this.button_mediaItem_moveDown.Size = new System.Drawing.Size(118, 100);
            this.button_mediaItem_moveDown.TabIndex = 1;
            this.button_mediaItem_moveDown.Text = "Move down";
            this.button_mediaItem_moveDown.UseVisualStyleBackColor = false;
            this.button_mediaItem_moveDown.Click += new System.EventHandler(this.button_mediaItem_moveDown_Click);
            // 
            // button_mediaItem_add
            // 
            this.button_mediaItem_add.BackColor = System.Drawing.Color.Silver;
            this.button_mediaItem_add.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_mediaItem_add.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_mediaItem_add.FlatAppearance.BorderSize = 0;
            this.button_mediaItem_add.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_mediaItem_add.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.button_mediaItem_add.Location = new System.Drawing.Point(3, 321);
            this.button_mediaItem_add.Name = "button_mediaItem_add";
            this.button_mediaItem_add.Size = new System.Drawing.Size(118, 100);
            this.button_mediaItem_add.TabIndex = 2;
            this.button_mediaItem_add.Text = "Add media";
            this.button_mediaItem_add.UseVisualStyleBackColor = false;
            this.button_mediaItem_add.Click += new System.EventHandler(this.button_mediaItem_add_click);
            // 
            // button_mediaItem_remove
            // 
            this.button_mediaItem_remove.BackColor = System.Drawing.Color.Silver;
            this.button_mediaItem_remove.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_mediaItem_remove.FlatAppearance.BorderSize = 0;
            this.button_mediaItem_remove.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_mediaItem_remove.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.button_mediaItem_remove.Location = new System.Drawing.Point(3, 427);
            this.button_mediaItem_remove.Name = "button_mediaItem_remove";
            this.button_mediaItem_remove.Size = new System.Drawing.Size(118, 101);
            this.button_mediaItem_remove.TabIndex = 3;
            this.button_mediaItem_remove.Text = "Remove";
            this.button_mediaItem_remove.UseVisualStyleBackColor = false;
            this.button_mediaItem_remove.Click += new System.EventHandler(this.button_mediaItem_remove_Click);
            // 
            // comboBox_mediaItems
            // 
            this.comboBox_mediaItems.BackColor = System.Drawing.Color.White;
            this.comboBox_mediaItems.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox_mediaItems.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple;
            this.comboBox_mediaItems.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_mediaItems.FormattingEnabled = true;
            this.comboBox_mediaItems.Location = new System.Drawing.Point(3, 3);
            this.comboBox_mediaItems.Name = "comboBox_mediaItems";
            this.comboBox_mediaItems.Size = new System.Drawing.Size(725, 531);
            this.comboBox_mediaItems.TabIndex = 1;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.button_control_start, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(882, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(214, 564);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // button_control_start
            // 
            this.button_control_start.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.button_control_start.Enabled = false;
            this.button_control_start.FlatAppearance.BorderSize = 0;
            this.button_control_start.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_control_start.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.button_control_start.Location = new System.Drawing.Point(3, 3);
            this.button_control_start.Name = "button_control_start";
            this.button_control_start.Size = new System.Drawing.Size(194, 179);
            this.button_control_start.TabIndex = 0;
            this.button_control_start.Text = "Start";
            this.button_control_start.UseVisualStyleBackColor = false;
            this.button_control_start.Click += new System.EventHandler(this.button_control_start_Click);
            // 
            // progressBar
            // 
            this.progressBar.BackColor = System.Drawing.Color.DarkGray;
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar.Location = new System.Drawing.Point(3, 573);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(873, 44);
            this.progressBar.TabIndex = 2;
            // 
            // label_version
            // 
            this.label_version.AutoSize = true;
            this.label_version.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_version.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_version.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label_version.Location = new System.Drawing.Point(882, 570);
            this.label_version.Name = "label_version";
            this.label_version.Size = new System.Drawing.Size(214, 50);
            this.label_version.TabIndex = 3;
            this.label_version.Text = "github.com/Marthaarman/slidecat\r\nVersion: 1.0.0";
            this.label_version.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // form_main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1099, 620);
            this.Controls.Add(this.tableLayoutPanel_main);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.MaximizeBox = false;
            this.Name = "form_main";
            this.ShowIcon = false;
            this.Text = "SlideCat";
            this.tableLayoutPanel_main.ResumeLayout(false);
            this.tableLayoutPanel_main.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.groupBox_media_items.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel5.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel_main;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.GroupBox groupBox_media_items;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private System.Windows.Forms.Button button_mediaItem_moveUp;
        private System.Windows.Forms.Button button_mediaItem_moveDown;
        private System.Windows.Forms.Button button_mediaItem_add;
        private System.Windows.Forms.Button button_mediaItem_remove;
        private System.Windows.Forms.ComboBox comboBox_mediaItems;
        private System.Windows.Forms.Button button_control_start;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label label_version;
    }
}

