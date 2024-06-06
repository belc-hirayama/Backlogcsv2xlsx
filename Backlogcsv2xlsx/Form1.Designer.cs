namespace Backlogcsv2xlsx
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this._Button_OpenFile = new System.Windows.Forms.Button();
            this._TextBox_FileName = new System.Windows.Forms.TextBox();
            this.startCalendar = new System.Windows.Forms.MonthCalendar();
            this.endCalendar = new System.Windows.Forms.MonthCalendar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this._textBoxStartDate = new System.Windows.Forms.TextBox();
            this._textBoxEndDate = new System.Windows.Forms.TextBox();
            this._textBoxDiff = new System.Windows.Forms.TextBox();
            this._radioButtonWeek = new System.Windows.Forms.RadioButton();
            this._radioButtonDay = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // _Button_OpenFile
            // 
            this._Button_OpenFile.Location = new System.Drawing.Point(695, 9);
            this._Button_OpenFile.Name = "_Button_OpenFile";
            this._Button_OpenFile.Size = new System.Drawing.Size(70, 48);
            this._Button_OpenFile.TabIndex = 0;
            this._Button_OpenFile.Text = "参照";
            this._Button_OpenFile.UseVisualStyleBackColor = true;
            this._Button_OpenFile.Click += new System.EventHandler(this._Button_OpenFile_Click);
            // 
            // _TextBox_FileName
            // 
            this._TextBox_FileName.AllowDrop = true;
            this._TextBox_FileName.Font = new System.Drawing.Font("MS UI Gothic", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this._TextBox_FileName.Location = new System.Drawing.Point(78, 12);
            this._TextBox_FileName.Name = "_TextBox_FileName";
            this._TextBox_FileName.Size = new System.Drawing.Size(601, 37);
            this._TextBox_FileName.TabIndex = 1;
            this._TextBox_FileName.DragDrop += new System.Windows.Forms.DragEventHandler(this._TextBox_FileName_DragDrop);
            this._TextBox_FileName.DragEnter += new System.Windows.Forms.DragEventHandler(this._TextBox_FileName_DragEnter);
            // 
            // startCalendar
            // 
            this.startCalendar.Location = new System.Drawing.Point(22, 209);
            this.startCalendar.Name = "startCalendar";
            this.startCalendar.TabIndex = 3;
            this.startCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.startCalendar_DateChanged);
            // 
            // endCalendar
            // 
            this.endCalendar.Location = new System.Drawing.Point(317, 209);
            this.endCalendar.Name = "endCalendar";
            this.endCalendar.TabIndex = 4;
            this.endCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.endCalendar_DateChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(18, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "File : ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("MS UI Gothic", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(73, 129);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 25);
            this.label2.TabIndex = 6;
            this.label2.Text = "集計開始日";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("MS UI Gothic", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(384, 129);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(137, 25);
            this.label3.TabIndex = 7;
            this.label3.Text = "集計終了日";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("MS UI Gothic", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(625, 325);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(140, 82);
            this.button1.TabIndex = 8;
            this.button1.Text = "実行";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // _textBoxStartDate
            // 
            this._textBoxStartDate.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this._textBoxStartDate.Location = new System.Drawing.Point(88, 170);
            this._textBoxStartDate.Name = "_textBoxStartDate";
            this._textBoxStartDate.Size = new System.Drawing.Size(122, 27);
            this._textBoxStartDate.TabIndex = 9;
            this._textBoxStartDate.LostFocus += new System.EventHandler(this._textBoxStartDate_LostFocus);
            // 
            // _textBoxEndDate
            // 
            this._textBoxEndDate.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this._textBoxEndDate.Location = new System.Drawing.Point(389, 170);
            this._textBoxEndDate.Name = "_textBoxEndDate";
            this._textBoxEndDate.Size = new System.Drawing.Size(117, 27);
            this._textBoxEndDate.TabIndex = 10;
            this._textBoxEndDate.LostFocus += new System.EventHandler(this._textBoxEndDate_LostFocus);
            // 
            // _textBoxDiff
            // 
            this._textBoxDiff.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this._textBoxDiff.Location = new System.Drawing.Point(22, 70);
            this._textBoxDiff.MaxLength = 3;
            this._textBoxDiff.Name = "_textBoxDiff";
            this._textBoxDiff.Size = new System.Drawing.Size(84, 27);
            this._textBoxDiff.TabIndex = 11;
            this._textBoxDiff.Text = "1";
            this._textBoxDiff.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this._textBoxDiff.TextChanged += new System.EventHandler(this._textBoxDiff_TextChanged);
            // 
            // _radioButtonWeek
            // 
            this._radioButtonWeek.AutoSize = true;
            this._radioButtonWeek.Checked = true;
            this._radioButtonWeek.Location = new System.Drawing.Point(124, 75);
            this._radioButtonWeek.Name = "_radioButtonWeek";
            this._radioButtonWeek.Size = new System.Drawing.Size(58, 19);
            this._radioButtonWeek.TabIndex = 12;
            this._radioButtonWeek.TabStop = true;
            this._radioButtonWeek.Text = "週間";
            this._radioButtonWeek.UseVisualStyleBackColor = true;
            this._radioButtonWeek.CheckedChanged += new System.EventHandler(this._radioButtonWeek_CheckedChanged);
            // 
            // _radioButtonDay
            // 
            this._radioButtonDay.AutoSize = true;
            this._radioButtonDay.Location = new System.Drawing.Point(188, 75);
            this._radioButtonDay.Name = "_radioButtonDay";
            this._radioButtonDay.Size = new System.Drawing.Size(58, 19);
            this._radioButtonDay.TabIndex = 13;
            this._radioButtonDay.Text = "日間";
            this._radioButtonDay.UseVisualStyleBackColor = true;
            this._radioButtonDay.CheckedChanged += new System.EventHandler(this._radioButtonDay_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(798, 447);
            this.Controls.Add(this._radioButtonDay);
            this.Controls.Add(this._radioButtonWeek);
            this.Controls.Add(this._textBoxDiff);
            this.Controls.Add(this._textBoxEndDate);
            this.Controls.Add(this._textBoxStartDate);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.endCalendar);
            this.Controls.Add(this.startCalendar);
            this.Controls.Add(this._TextBox_FileName);
            this.Controls.Add(this._Button_OpenFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button _Button_OpenFile;
        private System.Windows.Forms.TextBox _TextBox_FileName;
        private System.Windows.Forms.MonthCalendar startCalendar;
        private System.Windows.Forms.MonthCalendar endCalendar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox _textBoxStartDate;
        private System.Windows.Forms.TextBox _textBoxEndDate;
        private System.Windows.Forms.TextBox _textBoxDiff;
        private System.Windows.Forms.RadioButton _radioButtonWeek;
        private System.Windows.Forms.RadioButton _radioButtonDay;
    }
}

