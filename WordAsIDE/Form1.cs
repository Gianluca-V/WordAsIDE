using System.Text;
using Microsoft.Office.Interop.Word;
using MicrosoftWord = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Text.RegularExpressions;


namespace WordAsIDE
{
    public partial class WordAsIDE : Form
    {
        private string language = "EN";
        private string MinGwInstalledText = "MinGW is not installed or is not installed on the default folder, please install it or provide your MinGW\\bin path. \n\n (if you already have MinGW\\bin on the PATH environment variable you can ignore this message)";
        private string folderDescriptionText = "Please select the MinGW\\Bin folder (if not C:\\MinGW\\bin)";

        private string theme = "light";
        private string prevTheme = "light";

        private string text;
        private Document wordDoc;
        private MicrosoftWord.Application wordApp;
        private string mingwPath = @"C:\\MinGW\\bin";
        private string filePath = @"..\..\..\..\";
        private string fileName = "";
        private string executableName;
        private System.Threading.Timer cooldownTimer;
        private bool onCooldown = false;
        string prevText = "";
        private string[] blueKeywords = { "private", "public", "protected", "using", "namespace", "class", "int", "char", "float", "double", "bool", "new", "for", "if", "while", "switch", "case", "default", "break", "continue", "const","void","enum" };

        public WordAsIDE() {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists(mingwPath))
            {
                MessageBox.Show(MinGwInstalledText);
            }
                OpenWord();
            wordDoc.ActiveWindow.View.Type = WdViewType.wdWebView;
        }

        private void OpenWord()
        {
            wordApp = new MicrosoftWord.Application();
            wordApp.Visible = true;


            wordDoc = wordApp.Documents.Add();
            wordDoc.Content.Select();
            fileName = wordDoc.Name;
            GetWordText();

            wordApp.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(WindowSelectionChange);
        }

        private void OpenWord(string filePath)
        {
            wordApp = new MicrosoftWord.Application();
            wordApp.Visible = true;


            wordDoc = wordApp.Documents.Open(filePath);
            wordDoc.Content.Select();

            GetWordText();

            wordApp.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(WindowSelectionChange);
        }

        private void WindowSelectionChange(Selection sel)
        {
            if (wordDoc != null)
            {

                if (cooldownTimer == null)
                {
                    cooldownTimer = new System.Threading.Timer(manageCooldown, null, 2000, 2000); //cololdown to not change colors too often
                }

                GetWordText();

                if ((!onCooldown && wordDoc.Content.Text != prevText) || theme != prevTheme)
                {
                    ManageColors(sel);
                }
            }
        }

        private void ManageColors(Selection sel)
        {
            prevTheme = theme;
            prevText = wordDoc.Content.Text;
            onCooldown = true;


            if (theme == "Dark")        //set the color of text based on theme
            {
                sel.Range.Paragraphs[1].Range.Font.Color = WdColor.wdColorWhite;
            }
            else
            {
                sel.Range.Paragraphs[1].Range.Font.Color = WdColor.wdColorBlack;
            }

            MicrosoftWord.Range docRange = wordDoc.Content;

            string hashtagRegex = @"#(.+)";
            string slashRegex = @"//(.+)";
            string keywordRegex = "\\b(" + string.Join("|", blueKeywords) + ")\\b";

            foreach (Paragraph para in docRange.Paragraphs)
            {
                MatchCollection hashtagRegexMatches = Regex.Matches(para.Range.Text, hashtagRegex);
                MatchCollection slashRegexMatches = Regex.Matches(para.Range.Text, slashRegex);
                MatchCollection keywordRegexMatches = Regex.Matches(para.Range.Text, keywordRegex);

                //change color to blue to every keyword
                foreach (Match match in keywordRegexMatches)
                {
                    int start = match.Index + 1;
                    int length = match.Length;
                    MicrosoftWord.Range matchRange = para.Range.Characters[start];
                    matchRange.MoveEnd(WdUnits.wdCharacter, length);
                    if (theme == "light") { matchRange.Font.Color = WdColor.wdColorBlue; }
                    else { matchRange.Font.Color = WdColor.wdColorLightBlue; }
                }

                //change color to green to every word after #
                foreach (Match match in hashtagRegexMatches)
                {
                    int start = match.Index + 1;
                    int length = match.Length - 1;
                    MicrosoftWord.Range matchRange = para.Range.Characters[start];
                    matchRange.MoveEnd(WdUnits.wdCharacter, length);
                    matchRange.Font.Color = WdColor.wdColorSeaGreen;
                }

                //change color to gray to every word after //
                foreach (Match match in slashRegexMatches)
                {
                    int start = match.Index + 1;
                    int length = match.Length;
                    MicrosoftWord.Range matchRange = para.Range.Characters[start];
                    matchRange.MoveEnd(WdUnits.wdCharacter, length);
                    matchRange.Font.Color = WdColor.wdColorGray50;
                }
            }

            // check if the line contains ", ' or ”
            Regex quotationRegex = new Regex("([\"”'])(.*?)\\1");

            MatchCollection quotationRegexMatches = quotationRegex.Matches(wordDoc.Content.Text);

            //change color to gray to every word between "
            foreach (Match match in quotationRegexMatches)
            {
                // set the font color of the text after between quotation marks to orange
                MicrosoftWord.Range range = wordDoc.Range(match.Index, match.Index + match.Length);
                range.Font.Color = WdColor.wdColorOrange;
            }
        }

        private void manageCooldown(object state)
        {
            onCooldown = false;
        }

        private void GetWordText()
        {   //get text and set it to UTF8
            text = wordDoc.Content.Text;
            Encoding iso = Encoding.GetEncoding("ISO-8859-1");
            Encoding utf8 = Encoding.UTF8;
            byte[] isoBytes = iso.GetBytes(text);
            byte[] utf8Bytes = Encoding.Convert(iso, utf8, isoBytes);
            text = utf8.GetString(utf8Bytes);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (wordDoc != null)
            {
                try
                {
                    wordDoc.Close();
                }
                catch (Exception ex) { }
            }
            if (wordApp != null)
            {
                try
                {
                    wordApp.Quit();
                }
                catch (Exception ex) { }
            }
            System.Windows.Forms.Application.Exit();
        }

        private void compileTextOnClick(object sender, EventArgs e)
        {
            GetWordText();
            
            string cppFileName = fileName + ".cpp";
            executableName = fileName +".exe";


            
            Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + ";" + mingwPath);



            using (StreamWriter archivo = new StreamWriter(cppFileName))
            {
                archivo.Write(text);
            }

            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = false;
            process.Start();

            process.StandardInput.WriteLine($"g++ {cppFileName} -o {executableName}");
            process.StandardInput.Flush();
            //process.StandardInput.Close();
            process.WaitForExit();

        }

        private void ExecuteButtonOnClick(object sender, EventArgs e)
        {
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = false;
            process.Start();

            process.StandardInput.WriteLine($"{executableName}");
            process.StandardInput.Flush();
            //process.StandardInput.Close();
            process.WaitForExit();

        }

        private void CompileExecuteOnCLick(object sender, EventArgs e)
        {
            compileTextOnClick(sender, e);
            ExecuteButtonOnClick(sender, e);
        }

        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            //open word doc
            OpenFileDialog openFileDialog = new OpenFileDialog();       //open a OpenFileDialog 

            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "Word documents (*.doc;*.docx)|*.doc;*.docx";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
                OpenWord(openFileDialog.FileName);
            }
        }

        private void SaveFileButton_Click(object sender, EventArgs e)
        {
            //save word doc
            SaveFileDialog saveFileDialog = new SaveFileDialog();           //open a saveFileDialog
            saveFileDialog.Title = "Save a text file";
            saveFileDialog.Filter = "Word documents (*.doc;*.docx)|*.doc;*.docx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog.FileName;

                wordDoc.SaveAs2(fileName);
            }
        }

        private void MinGwPathButton_Click(object sender, EventArgs e)
        {
            //set the MinGW path to a diferent folder than the default
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = folderDescriptionText;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                mingwPath = folderBrowserDialog.SelectedPath;
            }
        }

        private void LanguageButton_Click(object sender, EventArgs e)
        {
            //change the language
            if (language == "EN")  //change the language to Spanish
            {
                language = "ES";
                LanguageButton.Text = language;
                CompileButton.Text = "Compilar";
                ExecuteButton.Text = "Ejecutar";
                CompileExecuteButton.Text = "Compilar y ejecutar";
                NewFileButton.Text = "Nuevo archivo";
                NewFileButton.Font = new System.Drawing.Font("Agency FB", 22);
                OpenFileButton.Text = "Abrir archivo";
                OpenFileButton.Font = new System.Drawing.Font("Agency FB", 22);
                SaveFileButton.Text = "Guardar archivo";
                SaveFileButton.Font = new System.Drawing.Font("Agency FB", 22);
                MinGwPathButton.Text = "Ruta MinGW";
                MinGwInstalledText = "MinGW no está instalado o no está instalado en la carpeta predeterminada, instálelo o proporcione su ruta MinGW\\bin. \n\n(si ya tiene MinGW\\bin en la variable de entorno PATH, puede ignorar este mensaje)";
                folderDescriptionText = "Seleccione la carpeta MinGW\\Bin (si no es C:\\MinGW\\bin)";
            }
            else                   //change the language to English
            {
                language = "EN";
                LanguageButton.Text = language;
                CompileButton.Text = "Compile";
                ExecuteButton.Text = "Execute";
                CompileExecuteButton.Text = "Compile and execute";
                NewFileButton.Text = "New file";
                NewFileButton.Font = new System.Drawing.Font("Agency FB", 28);
                OpenFileButton.Text = "Open file";
                OpenFileButton.Font = new System.Drawing.Font("Agency FB", 28);
                SaveFileButton.Text = "Save file";
                SaveFileButton.Font = new System.Drawing.Font("Agency FB", 28);
                MinGwPathButton.Text = "MinGW path";
                MinGwInstalledText = "MinGW is not installed or is not installed on the default folder, please install it or provide your MinGW\\bin path. \n\n (if you already have MinGW\\bin on the PATH environment variable you can ignore this message)";
                folderDescriptionText = "Please select the MinGW\\Bin folder (if not C:\\MinGW\\bin)";
            }
        }

        private void ThemeButton_Click(object sender, EventArgs e)
        {
            //change the word theme
            if(theme == "Light")    //cahnge to Light Theme
            {
                theme = "Dark";
                prevTheme = "light";
                ThemeButton.Text = theme;
                wordDoc.Background.Fill.ForeColor.RGB = (int)WdColor.wdColorGray90;
                wordDoc.Background.Fill.Solid();
                wordDoc.Content.Font.Color = WdColor.wdColorWhite;
                Selection sel = wordApp.Selection;
                ManageColors(sel);
            }
            else                   //cahnge to Dark Theme
            {
                theme = "Light";
                prevTheme = "Dark";
                ThemeButton.Text = theme;
                wordDoc.Background.Fill.Solid();
                wordDoc.Background.Fill.ForeColor.RGB = (int)WdColor.wdColorWhite;
                wordDoc.Content.Font.Color = WdColor.wdColorBlack;
                Selection sel = wordApp.Selection;
                ManageColors(sel);
            }
        }

        private void NewFileButton_Click(object sender, EventArgs e)
        {
            OpenWord();
        }
    }
}