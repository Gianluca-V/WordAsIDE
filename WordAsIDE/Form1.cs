using Microsoft.Office.Core;
using System.Text;
using Microsoft.Office.Interop.Word;
using MicrosoftWord = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Text.RegularExpressions;


namespace WordAsIDE
{
    public partial class Form1 : Form
    {
        private string text;
        private Document wordDoc;
        private MicrosoftWord.Application wordApp;
        private string fileName;
        private string executableName;
        private System.Threading.Timer cooldownTimer;
        private bool onCooldown = false;
        private string prevLine = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            wordApp = new MicrosoftWord.Application();
            wordApp.Visible = true;


            wordDoc = wordApp.Documents.Add();
            wordDoc.SaveAs(@"..\Document.docx");
            wordDoc.Content.Select();
            text = wordDoc.Content.Text;
            byte[] utf8Bytes = Encoding.UTF8.GetBytes(text);
            text = Encoding.UTF8.GetString(utf8Bytes);
            wordApp.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(WindowSelectionChange);

        }

        private void WindowSelectionChange(Selection sel)
        {
            if (wordDoc != null)
            {
                if (cooldownTimer == null)
                {
                    cooldownTimer = new System.Threading.Timer(manageCooldown, null, 2000,2000); //cololdown to not change colors too often
                }

                //format the text to UTF8
                text = wordDoc.Content.Text;
                Encoding iso = Encoding.GetEncoding("ISO-8859-1");
                Encoding utf8 = Encoding.UTF8;
                byte[] isoBytes = iso.GetBytes(text);
                byte[] utf8Bytes = Encoding.Convert(iso, utf8, isoBytes);
                text = utf8.GetString(utf8Bytes);

                #region manageColors
                
                int docLine = 1;       
                string lineText = sel.Range.Paragraphs[docLine].Range.Text; //get the current line
                if (!onCooldown && (lineText != prevLine))
                {
                    onCooldown = true;
                    prevLine = lineText;

                    sel.Range.Paragraphs[docLine].Range.Font.Color = WdColor.wdColorBlack;

                    string[] blueKeywords = { "private", "public", "protected", "using", "namespace", "class", "int", "char", "float", "double", "bool", "new", "for", "if", "while", "switch", "case", "default", "break", "continue" };

                    // check if the line contains a blue keyword
                    foreach (string keyword in blueKeywords)
                    {
                        if (lineText.Contains(keyword))
                        {
                            // set the font color of the keyword to blue
                            MicrosoftWord.Range keywordPos = sel.Range.Paragraphs[docLine].Range;
                            keywordPos.Find.Execute(keyword);
                            keywordPos.Font.Color = WdColor.wdColorBlue;
                        }
                    }

                    // check if the line contains "#"
                    if (lineText.Contains("#"))
                    {
                        // set the font color of the text after "#" to greeen
                        MicrosoftWord.Range textAfterHashtag = sel.Range.Paragraphs[docLine].Range;
                        textAfterHashtag.Find.Execute("#", Forward: true);
                        textAfterHashtag.SetRange(textAfterHashtag.End - 1, sel.Range.Paragraphs[docLine].Range.End);
                        textAfterHashtag.Font.Color = WdColor.wdColorGreen;
                    }

                    // check if the line contains "//"
                    if (lineText.Contains("//"))
                    {
                        // set the font color of the text after "//" to gray
                        MicrosoftWord.Range textAfterSlash = sel.Range.Paragraphs[docLine].Range;
                        textAfterSlash.Find.Execute("//", Forward: true);
                        textAfterSlash.SetRange(textAfterSlash.End - 2, sel.Range.Paragraphs[docLine].Range.End);
                        textAfterSlash.Font.Color = WdColor.wdColorGray50;
                    }

                    // check if the line contains ", ' or ”
                    if (lineText.Contains("\"") || lineText.Contains("\'") || lineText.Contains("”"))
                    {
                        Regex regex = new Regex("([\"”'])(.*?)\\1"); 

                        MatchCollection matches = regex.Matches(wordDoc.Content.Text);

                        foreach (Match match in matches)
                        {
                            // set the font color of the text after between quotation marks to orange
                            MicrosoftWord.Range range = wordDoc.Range(match.Index, match.Index + match.Length);
                            range.Font.Color = WdColor.wdColorOrange;
                        }
                    }
                    
                }
                #endregion
            }
        }

        private void manageCooldown(object state)
        {
            onCooldown = false;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (wordDoc != null)
            {
                wordDoc.Close();
            }
            if (wordApp != null)
            {
                wordApp.Quit();
            }
        }

        private void compileTextOnClick(object sender, EventArgs e)
        {
            fileName = @"..WordAsIDE\" + "code.cpp";
            executableName = @"..\WordAsIDE\" + "code.exe";

            string mingwPath = "C:\\MinGW\\bin";
            Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + ";" + mingwPath);


            using (StreamWriter archivo = new StreamWriter(fileName))
            {
                archivo.Write(text);
            }

            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;
            process.Start();

            process.StandardInput.WriteLine($"g++ {fileName} -o {executableName}");
            process.StandardInput.Flush();
            process.StandardInput.Close();
            process.WaitForExit();

        }

        private void ExecuteButtonOnClick(object sender, EventArgs e)
        {
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;
            process.Start();

            process.StandardInput.WriteLine($"{executableName}");
            process.StandardInput.Flush();
            process.StandardInput.Close();
            process.WaitForExit();

        }

        private void CompileExecuteOnCLick(object sender, EventArgs e)
        {
            compileTextOnClick(sender, e);
            ExecuteButtonOnClick(sender, e);
        }
    }
}