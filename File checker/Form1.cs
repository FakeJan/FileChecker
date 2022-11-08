using Aspose.Words;
using CsvHelper;
using CsvHelper.Configuration;
using iTextSharp.text.exceptions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Org.BouncyCastle.Utilities;
using Syncfusion.Pdf.Parsing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static File_checker.Form1;
using File = System.IO.File;
using Path = System.IO.Path;

namespace File_checker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // file status
        private void button1_Click(object sender, EventArgs e)
        {
            GetFileStatus();
        }

        // works for word, excel, powerpoint, pdf, onenote
        public void GetFileStatus()
        {
            string path = ReturnPath();
            string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
            // if first true -> access restriction, if both true -> password protected
            var fileObject = new List<FileStatus>();
            foreach (var file in files)
            {
                string[] words = file.Split('\\');
                string[] ListOfSuspiciousFileExtensions = { "..zip", ".cawwcca", ".ecc", ".ezz", ".exx", ".zzz", ".xyz", ".aaa", ".abc", ".ccc", ".vvv", ".xxx", ".ttt", ".micro", ".encrypted", ".locked", ".crypto", ".crinf", ".r5a", ".XRNT", ".XTBL", ".crypt", ".R16M01D05", ".pzdc", ".good", ".LOL!", ".OMG!", ".RDM", ".RRK", ".encryptedRSA", ".crjoker", ".EnCiPhErEd", ".LeChiffre", ".keybtc@inbox_com", ".0x0", ".bleep", ".1999", ".vault", ".HA3", ".toxcrypt", ".magic", ".SUPERCRYPT", ".CTBL", ".CTB2", ".locky", ".cerber", ".coverton", ".cryp1", ".crypz", ".encrypt", ".frtrss", ".locky", ".rsnslocked", ".silent", ".zcrypt", ".zepto" };
                string[] ListOfSuspiciousFileNames = { "!recover!", ".cryptotorlocker", ".hydracrypt_ID", "_recover_", "decrypt my file", "decryptmyfiles", "files_are_encrypted", "rec0ver", "recover", "restore_fi", "want your files back", "warning-!!", "_crypt", "_help_instruct", "confirmation.key", "cryptolocker", "de_crypt_readme", "decrypt_instruct", "decrypt-instruct", "enc_files.txt", "help_decrypt", "help_file_", "help_instructions.", "help_recover", "help_restore", "help_your_file", "how to decrypt", "how_recover", "how_to_decrypt", "how_to_recover", "howto_restore", "howtodecrypt", "install_tor", "last_chance.txt", "message.txt", "readme_decrypt", "readme_for_decrypt", "recovery_file.txt", "recovery_key.txt", "recovery+", "vault.hta", "vault.key", "vault.txt", "your_files.url" };
                string[] ListOfDeniedFileExtensions = { ".dll", ".exe", ".bat", ".cmd", ".vbs", ".reg", ".url", ".msu", ".zip", ".7z", ".tmp" };
                bool accessRestricted, passWordProtected;

                if (words[words.Length - 1] == "~$istemi.docx") continue;

                if (ContainsAny(words[words.Length - 1], ListOfSuspiciousFileNames))
                {
                    fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = "Suspicious file name" });
                    continue;
                }
                else if (ContainsAny(words[words.Length - 1], ListOfSuspiciousFileExtensions))
                {
                    fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = "Suspicious file extension" });
                    continue;
                }
                else if (ContainsAny(words[words.Length - 1], ListOfDeniedFileExtensions))
                {
                    fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = "Denied file extension" });
                    continue;
                }

                if (words[words.Length - 1].Contains(".pdf"))
                {
                    string status = IsPdfPasswordProtected(file);
                    if (status == "Bad user password") status = "Password protected";

                    fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = status });
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
                        fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = "Password protected" });
                    }
                    else
                    {
                        fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = "Not restricted" });
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
                        fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = "File is corrupted" });
                        continue;
                    }
                    accessRestricted = info.IsEncrypted;
                    passWordProtected = IsPassworded(file);
                    string status = GetStatus(accessRestricted, passWordProtected);
                    fileObject.Add(new FileStatus { Name = words[words.Length - 1], Status = status });
                }
            }
            WriteCsv(fileObject, path);
            textBox1.Clear();
        }

        public bool ContainsAny(string haystack, string[] needles)
        {
            foreach (string needle in needles)
            {
                if (haystack.Contains(needle))
                    return true;
            }
            return false;
        }

        // returns path the user chose with dialog option
        public string ReturnPath()
        {
            string path = textBox1.Text;
            if (path != "")
            {
                return path;
            }
            else
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                string newSavePath = "";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    newSavePath = fbd.SelectedPath;
                }
                return newSavePath;
            }
        }

        public string GetStatus(bool accessRestricted, bool passWordProtected)
        {
            if (accessRestricted && passWordProtected) return "Password protected";
            if (accessRestricted && !passWordProtected) return "Access restriction";
            if (!accessRestricted && !passWordProtected) return "Not restricted";
            return "";
        }

        public string IsPdfPasswordProtected(string pdfFullname)
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

        public string IsPDFHeader(string fileName)
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

        public bool IsPassworded(string file)
        {
            var bytes = File.ReadAllBytes(file);
            return IsPassworded(bytes);
        }
        public bool IsPassworded(byte[] bytes)
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

        // file info
        private void button2_Click(object sender, EventArgs e)
        {
            string path = ReturnPath();
            string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
            var fileObject = new List<InfoFile>();
            foreach (var file in files)
            {
                string[] words = file.Split('\\');
                if (words[words.Length - 1] == "~$istemi.docx") continue;
                FileInfo oFileInfo = new FileInfo(file);
                var size = oFileInfo.Length / 1024;
                fileObject.Add(new InfoFile { Path = file, Name = oFileInfo.Name, LastWriteTime = oFileInfo.LastWriteTime.ToString(), CreationTime = oFileInfo.CreationTime.ToString(), Size_KB = size.ToString(), Attributes = oFileInfo.Attributes.ToString() });
            }
            WriteCsv(fileObject, path);
            textBox1.Clear();
        }

        // write file info into csv
        public void WriteCsv(List<InfoFile> fileObject, string path)
        {
            path += "\\fileInfo.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        // write file status into csv
        public void WriteCsv(List<FileStatus> fileObject, string path)
        {
            path += "\\fileStatus.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        // write file hash into csv
        public void WriteCsv(List<FileHash> fileObject, string path)
        {
            path += "\\fileHash.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        // write file hash into csv
        public void WriteCsv(List<FSInfo> fileObject, string path)
        {
            path += "\\fsInfo.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        // get fs info
        private void button3_Click(object sender, EventArgs e)
        {
            string path = ReturnPath();
            int numOfDirs = Directory.GetDirectories(path, "*", SearchOption.AllDirectories).Length;
            int numOfFiles = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Length;

            var size = DirSize(new DirectoryInfo(path));
            size /= 1024 * 1024;

            string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
            DateTime dateTime = new DateTime();
            string newest = "";
            foreach (var file in files)
            {
                string[] words = file.Split('\\');
                FileInfo oFileInfo = new FileInfo(file);

                var newDate = oFileInfo.LastWriteTime;
                if (newDate > dateTime)
                {
                    dateTime = newDate;
                    newest = oFileInfo.Name;
                }
            }
            var fileObject = new List<FSInfo>()
            {
                new FSInfo { Name = path, Size_MB = size.ToString(), Number_Of_Folders = numOfDirs, Number_Of_Files = numOfFiles, Last_Changed_File = newest }
            };

            WriteCsv(fileObject, path);
            textBox1.Clear();
            //foreach (var dir in dirs)
            //{
            //    DirectoryInfo dirInfo = new DirectoryInfo(dir);
            //    var size = new DirectoryInfo(path).GetDirectorySize();
            //    //richTextBox1.AppendText(dirInfo.Attributes + "\n");
            //}
        }

        public long DirSize(DirectoryInfo d)
        {
            long size = 0;
            // Add file sizes.
            FileInfo[] fis = d.GetFiles();
            foreach (FileInfo fi in fis)
            {
                size += fi.Length;
            }
            // Add subdirectory sizes.
            DirectoryInfo[] dis = d.GetDirectories();
            foreach (DirectoryInfo di in dis)
            {
                size += DirSize(di);
            }
            return size;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string path = ReturnPath();
            string[] temp = path.Split('\\');
            string subpath = temp[temp.Length - 1];
            string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
            var fileObject = new List<FileHash>();
            foreach (var file in files)
            {
                string[] words = file.Split('\\');
                if (words[words.Length - 1] == "~$istemi.docx") continue;
                int index = file.IndexOf(subpath);
                string name = file.Substring(index);
                var hash = GetMD5Checksum(file);

                fileObject.Add(new FileHash { Name = name, Hash = hash });
            }
            WriteCsv(fileObject, path);
            textBox1.Clear();
        }

        public string GetMD5Checksum(string filename)
        {
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                using (var stream = System.IO.File.OpenRead(filename))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string firstFile = ReturnFilePath();
            string secondFile = ReturnFilePath();

            List<string> csvANames = new List<string>();
            List<string> csvAHash = new List<string>();
            List<string> csvBNames = new List<string>();
            List<string> csvBHash = new List<string>();

            using (var reader = new StreamReader(firstFile))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    csvANames.Add(values[0]);
                    csvAHash.Add(values[1]);
                }
            }
            using (var reader = new StreamReader(secondFile))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    csvBNames.Add(values[0]);
                    csvBHash.Add(values[1]);
                }
            }

            if (csvAHash.Count != csvBHash.Count) return;
            List<bool> result = new List<bool>();
            for (int i = 0; i < csvAHash.Count; i++)
            {
                if (csvAHash[i] == csvBHash[i])
                {
                    result.Add(true);
                }
                else
                {
                    result.Add(false);
                }
            }
            string[] dirs = firstFile.Split('\\');
            List<string> list = new List<string>(dirs);
            list.RemoveAt(dirs.Length - 1);
            dirs = list.ToArray();
            string path = String.Join("\\", dirs);

            var fileObject = new List<FileHashExtended>();
            for (int i = 1; i < result.Count; i++)
            {
                fileObject.Add(new FileHashExtended { Name = csvANames[i], Hash = csvAHash[i], Matching = result[i] });
            }
            WriteCsv(fileObject, path);
        }

        // write file hash into csv
        public void WriteCsv(List<FileHashExtended> fileObject, string path)
        {
            path += "\\hashCompareResult.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        private string ReturnFilePath()
        {
            OpenFileDialog open = new OpenFileDialog();
            //open.Filter = "All Files *.txt | *.txt";
            open.Multiselect = true;
            open.Title = "Select files";
            string filePath = "";
            if (open.ShowDialog() == DialogResult.OK)
            {
                filePath = open.FileName;
            }
            return filePath;
        }

        public class FileHashExtended : FileHash
        {
            public bool Matching { get; set; }
        }

        public class FSInfo
        {
            public string Name { get; set; }

            public string Size_MB { get; set; }

            public int Number_Of_Folders { get; set; }

            public int Number_Of_Files { get; set; }

            public string Last_Changed_File { get; set; }
        }

        public class FileHash
        {
            public string Name { get; set; }

            public string Hash { get; set; }
        }

        public class FileStatus
        {
            public string Name { get; set; }

            public string Status { get; set; }
        }

        public class InfoFile
        {
            public string Path { get; set; }
            public string Name { get; set; }
            public string LastWriteTime { get; set; }
            public string CreationTime { get; set; }
            public string Size_KB { get; set; }
            public string Attributes { get; set; }

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
}
