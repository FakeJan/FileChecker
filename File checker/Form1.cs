using Aspose.Words;
using iTextSharp.text.exceptions;
using iTextSharp.text.pdf;
using Syncfusion.Pdf.Parsing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace File_checker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetFileStatus();
        }

        public static void GetFileStatus()
        {
            string path = ReturnPath();
            string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);

            TableBuilder tb = new TableBuilder();
            tb.AddRow("FileName", "Status");

            // if first true -> access restriction, if both true -> password protected
            foreach (var file in files)
            {
                //word, excel, powerpoint, pdf, onenote
                string[] words = file.Split('\\');
                bool accessRestricted, passWordProtected;
                if (words[words.Length - 1] == "~$istemi.docx") continue;
                if (words[words.Length - 1].Contains(".pdf"))
                {
                    string status = IsPdfPasswordProtected(file);
                    if (status == "Bad user password") status = "Password protected";
                    tb.AddRow(words[words.Length - 1], status);
                }
                else if (words[words.Length - 1].Contains(".one") || words[words.Length - 1].Contains(".onetoc2"))
                {
                    var result = Path.ChangeExtension(file, ".txt");
                    if (!new FileInfo(result).Exists)
                    {
                        File.Copy(file, result);
                    }

                    string x = File.ReadAllText(result);
                    if (x.Contains("encryption"))
                    {
                        tb.AddRow(words[words.Length - 1], "Password protected");
                    }
                    else
                    {
                        tb.AddRow(words[words.Length - 1], "Not restricted");
                    }
                    File.SetAttributes(result, FileAttributes.Normal);
                    File.Delete(result);
                }
                else
                {
                    // detect access restriction
                    FileFormatInfo info = null;
                    try
                    {
                        info = FileFormatUtil.DetectFileFormat(file);
                    }
                    catch (Exception)
                    {
                        // file is corrupted
                        tb.AddRow(words[words.Length - 1], "File is corrupted");
                        continue;
                    }
                    accessRestricted = info.IsEncrypted;
                    passWordProtected = IsPassworded(file);
                    string status = GetStatus(accessRestricted, passWordProtected);
                    tb.AddRow(words[words.Length - 1], status);
                }

            }
            File.WriteAllText(path + @"\test.txt", tb.Output());
        }

        // returns path the user chose with dialog option
        public static string ReturnPath()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            string newSavePath = "";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                newSavePath = fbd.SelectedPath;
            }
            return newSavePath;
        }

        public static string GetStatus(bool accessRestricted, bool passWordProtected)
        {
            if (accessRestricted && passWordProtected) return "Password protected";
            if (accessRestricted && !passWordProtected) return "Access restriction";
            if (!accessRestricted && !passWordProtected) return "Not restricted";
            return "";
        }

        public static string IsPdfPasswordProtected(string pdfFullname)
        {
            try
            {
                PdfReader pdfReader = new PdfReader(pdfFullname);
                string status = IsPDFHeader(pdfFullname);
                return status;
            }
            catch (BadPasswordException ex)
            {
                return ex.Message;
            }
        }

        public static string IsPDFHeader(string fileName)
        {
            //Load the PDF file as stream.
            using (FileStream pdfStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                //Create a new instance of PDF document syntax analyzer.
                PdfDocumentAnalyzer analyzer = new PdfDocumentAnalyzer(pdfStream);
                //Analyze the syntax and return the results.
                SyntaxAnalyzerResult analyzerResult = analyzer.AnalyzeSyntax();
                analyzer.Close();
                //Check whether the document is corrupted or not.
                if (analyzerResult.IsCorrupted)
                {
                    return "The PDF document is corrupted.";
                }
                return "Not restricted";
            }
        }

        public static bool IsPassworded(string file)
        {
            var bytes = File.ReadAllBytes(file);
            return IsPassworded(bytes);
        }
        public static bool IsPassworded(byte[] bytes)
        {
            var prefix = Encoding.Default.GetString(bytes.Take(2).ToArray());
            if (prefix == "PK")
            {
                //ZIP and not password protected
                return false;
            }
            if (prefix == "ĐĎ")
            {
                //Office format.

                //Flagged with password
                if (bytes.Skip(0x20c).Take(1).ToArray()[0] == 0x2f) return true; //XLS 2003
                if (bytes.Skip(0x214).Take(1).ToArray()[0] == 0x2f) return true; //XLS 2005
                if (bytes.Skip(0x20B).Take(1).ToArray()[0] == 0x13) return true; //DOC 2005

                if (bytes.Length < 2000) return false; //Guessing false
                var start = Encoding.Default.GetString(bytes.Take(2000).ToArray()); //DOC/XLS 2007+
                start = start.Replace("\0", " ");
                if (start.Contains("E n c r y p t e d P a c k a g e")) return true;
                return false;
            }

            //Unknown.
            return false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = ReturnPath();
            string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                FileInfo oFileInfo = new FileInfo(file);
                //RichTextBox textBox = new RichTextBox();
                string atr = oFileInfo.Attributes;
            }

        }
    }

    //https://stackoverflow.com/a/14698822
    public interface ITextRow
    {
        String Output();
        void Output(StringBuilder sb);
        Object Tag { get; set; }
    }

    public class TableBuilder : IEnumerable<ITextRow>
    {
        protected class TextRow : List<String>, ITextRow
        {
            protected TableBuilder owner = null;
            public TextRow(TableBuilder Owner)
            {
                owner = Owner;
                if (owner == null) throw new ArgumentException("Owner");
            }
            public String Output()
            {
                StringBuilder sb = new StringBuilder();
                Output(sb);
                return sb.ToString();
            }
            public void Output(StringBuilder sb)
            {
                sb.AppendFormat(owner.FormatString, this.ToArray());
            }
            public Object Tag { get; set; }
        }

        public String Separator { get; set; }

        protected List<ITextRow> rows = new List<ITextRow>();
        protected List<int> colLength = new List<int>();

        public TableBuilder()
        {
            Separator = "  ";
        }

        public TableBuilder(String separator)
            : this()
        {
            Separator = separator;
        }

        public ITextRow AddRow(params object[] cols)
        {
            TextRow row = new TextRow(this);
            foreach (object o in cols)
            {
                String str = o.ToString().Trim();
                row.Add(str);
                if (colLength.Count >= row.Count)
                {
                    int curLength = colLength[row.Count - 1];
                    if (str.Length > curLength) colLength[row.Count - 1] = str.Length;
                }
                else
                {
                    colLength.Add(str.Length);
                }
            }
            rows.Add(row);
            return row;
        }

        protected String _fmtString = null;
        public String FormatString
        {
            get
            {
                if (_fmtString == null)
                {
                    String format = "";
                    int i = 0;
                    foreach (int len in colLength)
                    {
                        format += String.Format("{{{0},-{1}}}{2}", i++, len, Separator);
                    }
                    format += "\r\n";
                    _fmtString = format;
                }
                return _fmtString;
            }
        }

        public String Output()
        {
            StringBuilder sb = new StringBuilder();
            foreach (TextRow row in rows)
            {
                row.Output(sb);
            }
            return sb.ToString();
        }

        #region IEnumerable Members

        public IEnumerator<ITextRow> GetEnumerator()
        {
            return rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return rows.GetEnumerator();
        }

        #endregion
    }

}
