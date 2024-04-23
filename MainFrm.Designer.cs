namespace MCD
{
    partial class MainFrm
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
            try
            {
                if (disposing && (components != null))
                {
                    components.Dispose();
                }
                base.Dispose(disposing);
            }
            catch (System.Exception ex )
            {

                System.Windows.Forms.MessageBox.Show(ex.Message + ex.StackTrace);
            }
            
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainFrm));
            this.StopBttn = new System.Windows.Forms.Button();
            this.ChkDbAcadTimer = new System.Windows.Forms.Timer(this.components);
            this.FilePathLabel = new System.Windows.Forms.Label();
            this.ChkDbDocTimer = new System.Windows.Forms.Timer(this.components);
            this.GrpBxLabl = new System.Windows.Forms.GroupBox();
            this.GrpBxLabl.SuspendLayout();
            this.SuspendLayout();
            // 
            // StopBttn
            // 
            this.StopBttn.Location = new System.Drawing.Point(155, 69);
            this.StopBttn.Name = "StopBttn";
            this.StopBttn.Size = new System.Drawing.Size(112, 29);
            this.StopBttn.TabIndex = 0;
            this.StopBttn.Text = "Stop execution";
            this.StopBttn.UseVisualStyleBackColor = false;
            this.StopBttn.Click += new System.EventHandler(this.StopBttn_Click);
            // 
            // ChkDbAcadTimer
            // 
            this.ChkDbAcadTimer.Interval = 12000;
            this.ChkDbAcadTimer.Tick += new System.EventHandler(this.ChkDbTimer_Tick);
            // 
            // FilePathLabel
            // 
            this.FilePathLabel.AutoSize = true;
            this.FilePathLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FilePathLabel.Location = new System.Drawing.Point(6, 16);
            this.FilePathLabel.Name = "FilePathLabel";
            this.FilePathLabel.Size = new System.Drawing.Size(0, 16);
            this.FilePathLabel.TabIndex = 1;
            // 
            // ChkDbDocTimer
            // 
            this.ChkDbDocTimer.Interval = 10000;
            this.ChkDbDocTimer.Tick += new System.EventHandler(this.ChkDbDocTimer_Tick);
            // 
            // GrpBxLabl
            // 
            this.GrpBxLabl.Controls.Add(this.FilePathLabel);
            this.GrpBxLabl.Location = new System.Drawing.Point(3, 5);
            this.GrpBxLabl.Name = "GrpBxLabl";
            this.GrpBxLabl.Size = new System.Drawing.Size(407, 44);
            this.GrpBxLabl.TabIndex = 2;
            this.GrpBxLabl.TabStop = false;
            this.GrpBxLabl.Enter += new System.EventHandler(this.GrpBxLabl_Enter);
            // 
            // MainFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 108);
            this.Controls.Add(this.GrpBxLabl);
            this.Controls.Add(this.StopBttn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainFrm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Mcd Building Plan";
            this.Load += new System.EventHandler(this.MainFrm_Load);
            this.Shown += new System.EventHandler(this.MainFrm_Shown);
            this.GrpBxLabl.ResumeLayout(false);
            this.GrpBxLabl.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button StopBttn;
        private System.Windows.Forms.Timer ChkDbAcadTimer;
        private System.Windows.Forms.Label FilePathLabel;
        private System.Windows.Forms.Timer ChkDbDocTimer;
        private System.Windows.Forms.GroupBox GrpBxLabl;
    }
}

