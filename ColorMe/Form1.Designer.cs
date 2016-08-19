namespace ColorMe
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
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
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
            this.Pl_FileLoad = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Pl_FileLoad.SuspendLayout();
            this.SuspendLayout();
            // 
            // Pl_FileLoad
            // 
            this.Pl_FileLoad.AllowDrop = true;
            this.Pl_FileLoad.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Pl_FileLoad.Controls.Add(this.label2);
            this.Pl_FileLoad.Controls.Add(this.label1);
            this.Pl_FileLoad.Location = new System.Drawing.Point(12, 12);
            this.Pl_FileLoad.Name = "Pl_FileLoad";
            this.Pl_FileLoad.Size = new System.Drawing.Size(397, 292);
            this.Pl_FileLoad.TabIndex = 1;
            this.Pl_FileLoad.DragDrop += new System.Windows.Forms.DragEventHandler(this.Pl_FileLoad_DragDrop);
            this.Pl_FileLoad.DragEnter += new System.Windows.Forms.DragEventHandler(this.Pl_FileLoad_DragEnter);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("MS UI Gothic", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.Color.Blue;
            this.label2.Location = new System.Drawing.Point(62, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(265, 64);
            this.label2.TabIndex = 1;
            this.label2.Text = "カラーミー";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS UI Gothic", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(35, 148);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(309, 27);
            this.label1.TabIndex = 0;
            this.label1.Text = "ここにCSVファイルをドラッグ";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(424, 312);
            this.Controls.Add(this.Pl_FileLoad);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Pl_FileLoad.ResumeLayout(false);
            this.Pl_FileLoad.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel Pl_FileLoad;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}

