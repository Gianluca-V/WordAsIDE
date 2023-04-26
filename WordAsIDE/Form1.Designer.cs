namespace WordAsIDE
{
    partial class WordAsIDE
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.CompileButton = new System.Windows.Forms.Button();
            this.ExecuteButton = new System.Windows.Forms.Button();
            this.CompileExecuteButton = new System.Windows.Forms.Button();
            this.OpenFileButton = new System.Windows.Forms.Button();
            this.MinGwPathButton = new System.Windows.Forms.Button();
            this.SaveFileButton = new System.Windows.Forms.Button();
            this.LanguageButton = new System.Windows.Forms.Button();
            this.ThemeButton = new System.Windows.Forms.Button();
            this.NewFileButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CompileButton
            // 
            this.CompileButton.Font = new System.Drawing.Font("Agency FB", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.CompileButton.Location = new System.Drawing.Point(90, 41);
            this.CompileButton.Name = "CompileButton";
            this.CompileButton.Size = new System.Drawing.Size(263, 186);
            this.CompileButton.TabIndex = 1;
            this.CompileButton.Text = "Compile";
            this.CompileButton.UseVisualStyleBackColor = true;
            this.CompileButton.Click += new System.EventHandler(this.compileTextOnClick);
            // 
            // ExecuteButton
            // 
            this.ExecuteButton.Font = new System.Drawing.Font("Agency FB", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ExecuteButton.Location = new System.Drawing.Point(90, 281);
            this.ExecuteButton.Name = "ExecuteButton";
            this.ExecuteButton.Size = new System.Drawing.Size(263, 186);
            this.ExecuteButton.TabIndex = 2;
            this.ExecuteButton.Text = "Execute";
            this.ExecuteButton.UseVisualStyleBackColor = true;
            this.ExecuteButton.Click += new System.EventHandler(this.ExecuteButtonOnClick);
            // 
            // CompileExecuteButton
            // 
            this.CompileExecuteButton.Font = new System.Drawing.Font("Agency FB", 44F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.CompileExecuteButton.Location = new System.Drawing.Point(397, 41);
            this.CompileExecuteButton.Name = "CompileExecuteButton";
            this.CompileExecuteButton.Size = new System.Drawing.Size(263, 186);
            this.CompileExecuteButton.TabIndex = 3;
            this.CompileExecuteButton.Text = "Compile and execute";
            this.CompileExecuteButton.UseVisualStyleBackColor = true;
            this.CompileExecuteButton.Click += new System.EventHandler(this.CompileExecuteOnCLick);
            // 
            // OpenFileButton
            // 
            this.OpenFileButton.Font = new System.Drawing.Font("Agency FB", 28F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.OpenFileButton.Location = new System.Drawing.Point(535, 281);
            this.OpenFileButton.Name = "OpenFileButton";
            this.OpenFileButton.Size = new System.Drawing.Size(125, 85);
            this.OpenFileButton.TabIndex = 8;
            this.OpenFileButton.Text = "Open file";
            this.OpenFileButton.UseVisualStyleBackColor = true;
            this.OpenFileButton.Click += new System.EventHandler(this.OpenFileButton_Click);
            // 
            // MinGwPathButton
            // 
            this.MinGwPathButton.Font = new System.Drawing.Font("Agency FB", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.MinGwPathButton.Location = new System.Drawing.Point(397, 382);
            this.MinGwPathButton.Name = "MinGwPathButton";
            this.MinGwPathButton.Size = new System.Drawing.Size(125, 85);
            this.MinGwPathButton.TabIndex = 9;
            this.MinGwPathButton.Text = "MinGw Path";
            this.MinGwPathButton.UseVisualStyleBackColor = true;
            this.MinGwPathButton.Click += new System.EventHandler(this.MinGwPathButton_Click);
            // 
            // SaveFileButton
            // 
            this.SaveFileButton.Font = new System.Drawing.Font("Agency FB", 28F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.SaveFileButton.Location = new System.Drawing.Point(535, 382);
            this.SaveFileButton.Name = "SaveFileButton";
            this.SaveFileButton.Size = new System.Drawing.Size(125, 85);
            this.SaveFileButton.TabIndex = 10;
            this.SaveFileButton.Text = "Save file";
            this.SaveFileButton.UseVisualStyleBackColor = true;
            this.SaveFileButton.Click += new System.EventHandler(this.SaveFileButton_Click);
            // 
            // LanguageButton
            // 
            this.LanguageButton.Font = new System.Drawing.Font("Agency FB", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.LanguageButton.Location = new System.Drawing.Point(1, 1);
            this.LanguageButton.Name = "LanguageButton";
            this.LanguageButton.Size = new System.Drawing.Size(65, 37);
            this.LanguageButton.TabIndex = 11;
            this.LanguageButton.Text = "EN";
            this.LanguageButton.UseVisualStyleBackColor = true;
            this.LanguageButton.Click += new System.EventHandler(this.LanguageButton_Click);
            // 
            // ThemeButton
            // 
            this.ThemeButton.Font = new System.Drawing.Font("Agency FB", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ThemeButton.Location = new System.Drawing.Point(1, 55);
            this.ThemeButton.Name = "ThemeButton";
            this.ThemeButton.Size = new System.Drawing.Size(65, 37);
            this.ThemeButton.TabIndex = 12;
            this.ThemeButton.Text = "Light";
            this.ThemeButton.UseVisualStyleBackColor = true;
            this.ThemeButton.Click += new System.EventHandler(this.ThemeButton_Click);
            // 
            // NewFileButton
            // 
            this.NewFileButton.Font = new System.Drawing.Font("Agency FB", 28F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.NewFileButton.Location = new System.Drawing.Point(397, 281);
            this.NewFileButton.Name = "NewFileButton";
            this.NewFileButton.Size = new System.Drawing.Size(125, 85);
            this.NewFileButton.TabIndex = 13;
            this.NewFileButton.Text = "New file";
            this.NewFileButton.UseVisualStyleBackColor = true;
            this.NewFileButton.Click += new System.EventHandler(this.NewFileButton_Click);
            // 
            // WordAsIDE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(694, 510);
            this.Controls.Add(this.NewFileButton);
            this.Controls.Add(this.ThemeButton);
            this.Controls.Add(this.LanguageButton);
            this.Controls.Add(this.SaveFileButton);
            this.Controls.Add(this.MinGwPathButton);
            this.Controls.Add(this.OpenFileButton);
            this.Controls.Add(this.CompileExecuteButton);
            this.Controls.Add(this.ExecuteButton);
            this.Controls.Add(this.CompileButton);
            this.Name = "WordAsIDE";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private Button viewText;
        private Button CompileButton;
        private Button ExecuteButton;
        private Button CompileExecuteButton;
        private Button OpenFileButton;
        private Button MinGwPathButton;
        private Button SaveFileButton;
        private Button LanguageButton;
        private Button ThemeButton;
        private Button NewFileButton;
    }
}