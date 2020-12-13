
namespace Beadando_projekt
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtCause = new System.Windows.Forms.TextBox();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.txtCountry = new System.Windows.Forms.TextBox();
            this.txtAcres = new System.Windows.Forms.TextBox();
            this.txtStructures = new System.Windows.Forms.TextBox();
            this.txtDeaths = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.btnDone = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(139, 37);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(157, 20);
            this.txtName.TabIndex = 0;
            this.txtName.Validating += new System.ComponentModel.CancelEventHandler(this.txtName_Validating);
            this.txtName.Validated += new System.EventHandler(this.txtName_Validated);
            // 
            // txtCause
            // 
            this.txtCause.Location = new System.Drawing.Point(139, 87);
            this.txtCause.Name = "txtCause";
            this.txtCause.Size = new System.Drawing.Size(157, 20);
            this.txtCause.TabIndex = 1;
            this.txtCause.Validating += new System.ComponentModel.CancelEventHandler(this.txtCause_Validating);
            this.txtCause.Validated += new System.EventHandler(this.txtCause_Validated);
            // 
            // txtMonth
            // 
            this.txtMonth.Location = new System.Drawing.Point(139, 132);
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(157, 20);
            this.txtMonth.TabIndex = 2;
            this.txtMonth.Validating += new System.ComponentModel.CancelEventHandler(this.txtMonth_Validating);
            this.txtMonth.Validated += new System.EventHandler(this.txtMonth_Validated);
            // 
            // txtYear
            // 
            this.txtYear.Location = new System.Drawing.Point(139, 175);
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(157, 20);
            this.txtYear.TabIndex = 3;
            this.txtYear.Validating += new System.ComponentModel.CancelEventHandler(this.txtYear_Validating);
            this.txtYear.Validated += new System.EventHandler(this.txtYear_Validated);
            // 
            // txtCountry
            // 
            this.txtCountry.Location = new System.Drawing.Point(139, 220);
            this.txtCountry.Name = "txtCountry";
            this.txtCountry.Size = new System.Drawing.Size(157, 20);
            this.txtCountry.TabIndex = 4;
            this.txtCountry.Validating += new System.ComponentModel.CancelEventHandler(this.txtCountry_Validating);
            this.txtCountry.Validated += new System.EventHandler(this.txtCountry_Validated);
            // 
            // txtAcres
            // 
            this.txtAcres.Location = new System.Drawing.Point(139, 265);
            this.txtAcres.Name = "txtAcres";
            this.txtAcres.Size = new System.Drawing.Size(157, 20);
            this.txtAcres.TabIndex = 5;
            this.txtAcres.Validating += new System.ComponentModel.CancelEventHandler(this.txtAcres_Validating);
            this.txtAcres.Validated += new System.EventHandler(this.txtAcres_Validated);
            // 
            // txtStructures
            // 
            this.txtStructures.Location = new System.Drawing.Point(139, 309);
            this.txtStructures.Name = "txtStructures";
            this.txtStructures.Size = new System.Drawing.Size(157, 20);
            this.txtStructures.TabIndex = 6;
            this.txtStructures.Validating += new System.ComponentModel.CancelEventHandler(this.txtStructures_Validating);
            this.txtStructures.Validated += new System.EventHandler(this.txtStructures_Validated);
            // 
            // txtDeaths
            // 
            this.txtDeaths.Location = new System.Drawing.Point(139, 348);
            this.txtDeaths.Name = "txtDeaths";
            this.txtDeaths.Size = new System.Drawing.Size(157, 20);
            this.txtDeaths.TabIndex = 7;
            this.txtDeaths.Validating += new System.ComponentModel.CancelEventHandler(this.txtDeaths_Validating);
            this.txtDeaths.Validated += new System.EventHandler(this.txtDeaths_Validated);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label1.Location = new System.Drawing.Point(93, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 16);
            this.label1.TabIndex = 8;
            this.label1.Text = "Név:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label2.Location = new System.Drawing.Point(93, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 16);
            this.label2.TabIndex = 9;
            this.label2.Text = "Oka:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label3.Location = new System.Drawing.Point(75, 136);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 16);
            this.label3.TabIndex = 10;
            this.label3.Text = "Hónap:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label4.Location = new System.Drawing.Point(103, 179);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 16);
            this.label4.TabIndex = 11;
            this.label4.Text = "Év:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label5.Location = new System.Drawing.Point(80, 224);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 16);
            this.label5.TabIndex = 12;
            this.label5.Text = "Város:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label6.Location = new System.Drawing.Point(72, 269);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(61, 16);
            this.label6.TabIndex = 13;
            this.label6.Text = "Terület:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label7.Location = new System.Drawing.Point(11, 313);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(122, 16);
            this.label7.TabIndex = 14;
            this.label7.Text = "Épületek száma:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label8.Location = new System.Drawing.Point(14, 352);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(119, 16);
            this.label8.TabIndex = 15;
            this.label8.Text = "Halottak száma:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label9.Location = new System.Drawing.Point(135, 9);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(176, 24);
            this.label9.TabIndex = 16;
            this.label9.Text = "Adatok megadása";
            // 
            // btnDone
            // 
            this.btnDone.CausesValidation = false;
            this.btnDone.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnDone.Location = new System.Drawing.Point(139, 402);
            this.btnDone.Name = "btnDone";
            this.btnDone.Size = new System.Drawing.Size(157, 36);
            this.btnDone.TabIndex = 18;
            this.btnDone.Text = "Excel megjelenítése";
            this.btnDone.UseVisualStyleBackColor = true;
            this.btnDone.Click += new System.EventHandler(this.btnDone_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnAdd.Location = new System.Drawing.Point(318, 335);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(116, 33);
            this.btnAdd.TabIndex = 19;
            this.btnAdd.Text = "Hozzáadás";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.CausesValidation = false;
            this.btnDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnDelete.Location = new System.Drawing.Point(318, 296);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(116, 33);
            this.btnDelete.TabIndex = 20;
            this.btnDelete.Text = "Mezők törlése";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 450);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnDone);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDeaths);
            this.Controls.Add(this.txtStructures);
            this.Controls.Add(this.txtAcres);
            this.Controls.Add(this.txtCountry);
            this.Controls.Add(this.txtYear);
            this.Controls.Add(this.txtMonth);
            this.Controls.Add(this.txtCause);
            this.Controls.Add(this.txtName);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtCause;
        private System.Windows.Forms.TextBox txtMonth;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.TextBox txtCountry;
        private System.Windows.Forms.TextBox txtAcres;
        private System.Windows.Forms.TextBox txtStructures;
        private System.Windows.Forms.TextBox txtDeaths;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnDone;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.ErrorProvider errorProvider1;
    }
}

