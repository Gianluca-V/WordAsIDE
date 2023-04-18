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
        string prevText = "";
        private string[] blueKeywords = { "private", "public", "protected", "using", "namespace", "class", "int", "char", "float", "double", "bool", "new", "for", "if", "while", "switch", "case", "default", "break", "continue", "const","void","enum" };

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
                    cooldownTimer = new System.Threading.Timer(manageCooldown, null, 2000, 2000); //cololdown to not change colors too often
                }

                //format the text to UTF8
                text = wordDoc.Content.Text;
                Encoding iso = Encoding.GetEncoding("ISO-8859-1");
                Encoding utf8 = Encoding.UTF8;
                byte[] isoBytes = iso.GetBytes(text);
                byte[] utf8Bytes = Encoding.Convert(iso, utf8, isoBytes);
                text = utf8.GetString(utf8Bytes);

                #region manageColors
                
                if (!onCooldown && wordDoc.Content.Text != prevText)
                {
                    prevText = wordDoc.Content.Text;
                    onCooldown = true;

                    sel.Range.Paragraphs[1].Range.Font.Color = WdColor.wdColorBlack;
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
                            int start = match.Index+1;
                            int length = match.Length;
                            MicrosoftWord.Range matchRange = para.Range.Characters[start];
                            matchRange.MoveEnd(WdUnits.wdCharacter, length);
                            matchRange.Font.Color = WdColor.wdColorBlue;
                        }

                        //change color to green to every word after #
                        foreach (Match match in hashtagRegexMatches)
                        {
                            int start = match.Index + 1;
                            int length = match.Length - 1;
                            MicrosoftWord.Range matchRange = para.Range.Characters[start];
                            matchRange.MoveEnd(WdUnits.wdCharacter, length);
                            matchRange.Font.Color = WdColor.wdColorGreen;
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
            fileName = @"..\..\..\..\..\WordAsIDE\" + "code.cpp";
            executableName = @"..\..\..\..\..\WordAsIDE\" + "code.exe";

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
            process.StartInfo.CreateNoWindow = false;
            process.Start();

            process.StandardInput.WriteLine($"g++ {fileName} -o {executableName}");
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
            process.StartInfo.CreateNoWindow = true;
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
    }
}