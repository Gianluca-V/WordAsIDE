using Microsoft.Office.Core;
using System.Text;
using Microsoft.Office.Interop.Word;
using MicrosoftWord = Microsoft.Office.Interop.Word;
using System.Diagnostics;


namespace WordAsIDE
{
    
    public partial class Form1 : Form
    {
        private string text;
        private Document wordDoc;
        private MicrosoftWord.Application wordApp;
        private string fileName;
        private string executableName;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            wordApp = new MicrosoftWord.Application();
            wordApp.Visible = true;


            wordDoc = wordApp.Documents.Add();
            wordDoc.SaveAs(@"C:\Users\gianl\Desktop\WordAsIDE\MiDocumento.docx");
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
                text = wordDoc.Content.Text;
                Encoding iso = Encoding.GetEncoding("ISO-8859-1");
                Encoding utf8 = Encoding.UTF8;
                byte[] isoBytes = iso.GetBytes(text);
                byte[] utf8Bytes = Encoding.Convert(iso, utf8, isoBytes);
                text = utf8.GetString(utf8Bytes);

                MicrosoftWord.Range range = wordDoc.Range();

                Find findObj = wordDoc.Content.Find;
                findObj.ClearFormatting();
                findObj.Text = "//*^p";
                findObj.Forward = true;
                findObj.Wrap = WdFindWrap.wdFindContinue;
                findObj.Format = true;
                findObj.Font.Color = WdColor.wdColorGray25;
                findObj.Execute();

                // Buscar y cambiar el color de las líneas que comienzan por "#"
                findObj.ClearFormatting();
                findObj.Text = "#*^p";
                findObj.Forward = true;
                findObj.Wrap = WdFindWrap.wdFindContinue;
                findObj.Format = true;
                findObj.Font.Color = WdColor.wdColorGreen;
                findObj.Execute();
            }
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
            string exampleCode = "#include <iostream>\n\nusing namespace std;\n\nint main()\n{\n   cout << \"Hello, World!\" << endl;\n   return 0;\n}\n";
            fileName = @"C:\Users\gianl\Desktop\WordAsIDE\" + "codigo.cpp";
            executableName = @"C:\Users\gianl\Desktop\WordAsIDE\" + "codigo.exe";

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