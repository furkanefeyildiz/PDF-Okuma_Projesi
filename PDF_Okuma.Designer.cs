
namespace PDF_OkumaProjesi
{
    partial class PDF_Okuma
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
            this.btnSelectPDF = new System.Windows.Forms.Button();
            this.btnExportExcel = new System.Windows.Forms.Button();
            this.txtPDFContent = new System.Windows.Forms.TextBox();
            this.cmbBelgeTuru = new System.Windows.Forms.ComboBox();
            this.cmbSirketSecimi = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnSelectPDF
            // 
            this.btnSelectPDF.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnSelectPDF.Location = new System.Drawing.Point(199, 197);
            this.btnSelectPDF.Name = "btnSelectPDF";
            this.btnSelectPDF.Size = new System.Drawing.Size(306, 78);
            this.btnSelectPDF.TabIndex = 0;
            this.btnSelectPDF.Text = "PDF Seç";
            this.btnSelectPDF.UseVisualStyleBackColor = false;
            this.btnSelectPDF.Click += new System.EventHandler(this.btnSelectPDF_Click);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnExportExcel.Enabled = false;
            this.btnExportExcel.Location = new System.Drawing.Point(199, 311);
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(306, 78);
            this.btnExportExcel.TabIndex = 1;
            this.btnExportExcel.Text = "Kaydet";
            this.btnExportExcel.UseVisualStyleBackColor = false;
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // txtPDFContent
            // 
            this.txtPDFContent.Location = new System.Drawing.Point(12, 395);
            this.txtPDFContent.Multiline = true;
            this.txtPDFContent.Name = "txtPDFContent";
            this.txtPDFContent.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtPDFContent.Size = new System.Drawing.Size(930, 238);
            this.txtPDFContent.TabIndex = 2;
            // 
            // cmbBelgeTuru
            // 
            this.cmbBelgeTuru.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBelgeTuru.FormattingEnabled = true;
            this.cmbBelgeTuru.Items.AddRange(new object[] {
            "Poliçe",
            "Credit Note",
            "Vergi Levhası"});
            this.cmbBelgeTuru.Location = new System.Drawing.Point(199, 44);
            this.cmbBelgeTuru.Name = "cmbBelgeTuru";
            this.cmbBelgeTuru.Size = new System.Drawing.Size(306, 40);
            this.cmbBelgeTuru.TabIndex = 3;
            this.cmbBelgeTuru.SelectedIndexChanged += new System.EventHandler(this.cmbBelgeTuru_SelectedIndexChanged);
            // 
            // cmbSirketSecimi
            // 
            this.cmbSirketSecimi.BackColor = System.Drawing.Color.White;
            this.cmbSirketSecimi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSirketSecimi.Enabled = false;
            this.cmbSirketSecimi.FormattingEnabled = true;
            this.cmbSirketSecimi.Location = new System.Drawing.Point(199, 130);
            this.cmbSirketSecimi.Name = "cmbSirketSecimi";
            this.cmbSirketSecimi.Size = new System.Drawing.Size(306, 40);
            this.cmbSirketSecimi.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(49, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 32);
            this.label1.TabIndex = 5;
            this.label1.Text = "Belge Türü:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 133);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(168, 32);
            this.label2.TabIndex = 6;
            this.label2.Text = "Şirket Seçimi:";
            // 
            // PDF_Okuma
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 32F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.ClientSize = new System.Drawing.Size(954, 645);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbSirketSecimi);
            this.Controls.Add(this.cmbBelgeTuru);
            this.Controls.Add(this.txtPDFContent);
            this.Controls.Add(this.btnExportExcel);
            this.Controls.Add(this.btnSelectPDF);
            this.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.Name = "PDF_Okuma";
            this.Text = " PDF Reader";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectPDF;
        private System.Windows.Forms.Button btnExportExcel;
        private System.Windows.Forms.TextBox txtPDFContent;
        private System.Windows.Forms.ComboBox cmbBelgeTuru;
        private System.Windows.Forms.ComboBox cmbSirketSecimi;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

