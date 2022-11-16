using Aspose.Words;
using Azure;
using CsvHelper;
using CsvHelper.Configuration;
using iTextSharp.text.exceptions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Org.BouncyCastle.Utilities;
using Syncfusion.Pdf.Parsing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
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
            string[] temp = path.Split('\\');
            string subpath = temp[temp.Length - 1];
            string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
            var fileObject = new List<InfoFile>();
            foreach (var file in files)
            {
                string[] words = file.Split('\\');
                if (words[words.Length - 1] == "~$istemi.docx") continue;
                FileInfo oFileInfo = new FileInfo(file);
                var size = oFileInfo.Length / 1024;
                int index = file.IndexOf(subpath);
                string name = file.Substring(index);
                string attributes = oFileInfo.Attributes.ToString().Replace(",", ";");
                fileObject.Add(new InfoFile { Path = name, Name = oFileInfo.Name, LastWriteTime = oFileInfo.LastWriteTime.ToString(), CreationTime = oFileInfo.CreationTime.ToString(), Size_KB = size.ToString(), Attributes = attributes });
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

        // compare file hash
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

            var fileObject = new List<FileHashExtended>();

            //check if A is in B
            for (int i = 1; i < csvANames.Count; i++)
            {
                string[] temp = csvANames[i].Split('\\');
                string fileName = temp[temp.Length - 1];
                for (int j = 1; j < csvBNames.Count; j++)
                {
                    if (csvBNames.Contains(csvANames[i]))
                    {
                        //check if hash is the same
                        if (csvAHash[i] == csvBHash[csvBNames.IndexOf(csvANames[i])])
                        {
                            // true
                            fileObject.Add(new FileHashExtended { Name = csvANames[i], Hash = csvAHash[i], Matching = "Hash matches", FileName = fileName });
                        }
                        else
                        {
                            // false
                            fileObject.Add(new FileHashExtended { Name = csvANames[i], Hash = csvAHash[i], Matching = "Hash doesn't match", FileName = fileName });
                        }
                        break;
                    }
                    else if (!csvBNames.Contains(csvANames[i]))
                    {
                        // doesn't contain
                        fileObject.Add(new FileHashExtended { Name = csvANames[i], Hash = csvAHash[i], Matching = "File not in second FS", FileName = fileName });
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }

            //check if B is in A
            for (int i = 1; i < csvBNames.Count; i++)
            {
                string[] temp = csvBNames[i].Split('\\');
                string fileName = temp[temp.Length - 1];
                for (int j = 1; j < csvANames.Count; j++)
                {
                    if (csvANames.Contains(csvBNames[i]))
                    {
                        //check if hash is the same
                        if (csvBHash[i] == csvAHash[csvANames.IndexOf(csvBNames[i])])
                        {
                            // true
                            fileObject.Add(new FileHashExtended { Name = csvBNames[i], Hash = csvBHash[i], Matching = "Hash matches", FileName = fileName });
                        }
                        else
                        {
                            // false
                            fileObject.Add(new FileHashExtended { Name = csvBNames[i], Hash = csvBHash[i], Matching = "Hash doesn't match", FileName = fileName });
                        }
                        break;
                    }
                    else if (!csvANames.Contains(csvBNames[i]))
                    {
                        // doesn't contain
                        fileObject.Add(new FileHashExtended { Name = csvBNames[i], Hash = csvBHash[i], Matching = "File not in first FS", FileName = fileName });
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            var noDupes = fileObject.GroupBy(x => x.Name).Select(x => x.First()).ToList();

            string[] dirs = firstFile.Split('\\');
            List<string> list = new List<string>(dirs);
            list.RemoveAt(dirs.Length - 1);
            dirs = list.ToArray();
            string path = String.Join("\\", dirs);

            WriteCsv(noDupes, path);
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

        public void WriteCsv(List<InfoFileExtended> fileObject, string path)
        {
            path += "\\fileInfoCompareResult.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        // compare files
        private void button6_Click(object sender, EventArgs e)
        {
            string firstFile = ReturnFilePath();
            string secondFile = ReturnFilePath();

            List<string> csvAPath = new List<string>();
            List<string> csvANames = new List<string>();
            List<string> csvALastWrite = new List<string>();
            List<string> csvACreateTime = new List<string>();
            List<string> csvASize = new List<string>();
            List<string> csvAAttributes = new List<string>();

            List<string> csvBPath = new List<string>();
            List<string> csvBNames = new List<string>();
            List<string> csvBLastWrite = new List<string>();
            List<string> csvBCreateTime = new List<string>();
            List<string> csvBSize = new List<string>();
            List<string> csvBAttributes = new List<string>();

            using (var reader = new StreamReader(firstFile))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    csvAPath.Add(values[0]);
                    csvANames.Add(values[1]);
                    csvALastWrite.Add(values[2]);
                    csvACreateTime.Add(values[3]);
                    csvASize.Add(values[4]);
                    csvAAttributes.Add(values[5]);
                }
            }
            using (var reader = new StreamReader(secondFile))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    csvBPath.Add(values[0]);
                    csvBNames.Add(values[1]);
                    csvBLastWrite.Add(values[2]);
                    csvBCreateTime.Add(values[3]);
                    csvBSize.Add(values[4]);
                    csvBAttributes.Add(values[5]);
                }
            }

            var fileObject = new List<InfoFileExtended>();

            //check if A is in B
            for (int i = 1; i < csvANames.Count; i++)
            {
                string[] temp = csvANames[i].Split('\\');
                string fileName = temp[temp.Length - 1];
                for (int j = 1; j < csvBNames.Count; j++)
                {
                    if (csvBNames.Contains(csvANames[i]))
                    {
                        //check if all columns are the same
                        if (csvALastWrite[i] == csvBLastWrite[csvBNames.IndexOf(csvANames[i])] &&
                            csvACreateTime[i] == csvBCreateTime[csvBNames.IndexOf(csvANames[i])] &&
                            csvASize[i] == csvBSize[csvBNames.IndexOf(csvANames[i])] &&
                            csvAAttributes[i] == csvBAttributes[csvBNames.IndexOf(csvANames[i])])
                        {
                            fileObject.Add(new InfoFileExtended { Path = csvAPath[i], Name = csvANames[i], LastWriteTime = csvALastWrite[i], CreationTime = csvACreateTime[i], Size_KB = csvASize[i], Attributes = csvAAttributes[i], Matching = "Files are the same" });
                        }
                        else
                        {
                            fileObject.Add(new InfoFileExtended { Path = csvAPath[i], Name = csvANames[i], LastWriteTime = csvALastWrite[i], CreationTime = csvACreateTime[i], Size_KB = csvASize[i], Attributes = csvAAttributes[i], Matching = "Files aren't the same" });
                        }
                        break;
                    }
                    else if (!csvBNames.Contains(csvANames[i]))
                    {
                        // doesn't contain
                        fileObject.Add(new InfoFileExtended { Path = csvAPath[i], Name = csvANames[i], LastWriteTime = csvALastWrite[i], CreationTime = csvACreateTime[i], Size_KB = csvASize[i], Attributes = csvAAttributes[i], Matching = "File not in second FS" });
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }

            //check if B is in A
            for (int i = 1; i < csvBNames.Count; i++)
            {
                string[] temp = csvBNames[i].Split('\\');
                string fileName = temp[temp.Length - 1];
                for (int j = 1; j < csvANames.Count; j++)
                {
                    if (csvANames.Contains(csvBNames[i]))
                    {
                        //check if all columns are the same
                        if (csvBLastWrite[i] == csvALastWrite[csvANames.IndexOf(csvBNames[i])] &&
                            csvBCreateTime[i] == csvACreateTime[csvANames.IndexOf(csvBNames[i])] &&
                            csvBSize[i] == csvASize[csvANames.IndexOf(csvBNames[i])] &&
                            csvBAttributes[i] == csvAAttributes[csvANames.IndexOf(csvBNames[i])])
                        {
                            fileObject.Add(new InfoFileExtended { Path = csvBPath[i], Name = csvBNames[i], LastWriteTime = csvBLastWrite[i], CreationTime = csvBCreateTime[i], Size_KB = csvBSize[i], Attributes = csvBAttributes[i], Matching = "Files are the same" });
                        }
                        else
                        {
                            fileObject.Add(new InfoFileExtended { Path = csvBPath[i], Name = csvBNames[i], LastWriteTime = csvBLastWrite[i], CreationTime = csvBCreateTime[i], Size_KB = csvBSize[i], Attributes = csvBAttributes[i], Matching = "Files aren't the same" });
                        }
                        break;
                    }
                    else if (!csvANames.Contains(csvBNames[i]))
                    {
                        // doesn't contain
                        fileObject.Add(new InfoFileExtended { Path = csvBPath[i], Name = csvBNames[i], LastWriteTime = csvBLastWrite[i], CreationTime = csvBCreateTime[i], Size_KB = csvBSize[i], Attributes = csvBAttributes[i], Matching = "File not in first FS" });
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            var noDupes = fileObject.GroupBy(x => x.Name).Select(x => x.First()).ToList();

            string[] dirs = firstFile.Split('\\');
            List<string> list = new List<string>(dirs);
            list.RemoveAt(dirs.Length - 1);
            dirs = list.ToArray();
            string path = String.Join("\\", dirs);

            WriteCsv(noDupes, path);
        }

        // get folder membership
        private void button7_Click(object sender, EventArgs e)
        {
            string path = ReturnPath();
            string[] temp = path.Split('\\');
            string subpath = temp[temp.Length - 1];
            string[] dirs = Directory.GetDirectories(path, "*", SearchOption.AllDirectories);
            var fileObject = new List<Membership>();
            foreach (var dir in dirs)
            {
                PowerShell ps = PowerShell.Create();

                ps.AddCommand("icacls");
                ps.AddArgument(dir);
                //ps.AddArgument("/T");
                Collection<PSObject> output = ps.Invoke();
                output.RemoveAt(output.Count - 1);
                output.RemoveAt(output.Count - 1);
                List<string> result = new List<string>();
                int index = dir.IndexOf(subpath);
                string name = dir.Substring(index);
                foreach (var item in output)
                {
                    string user;
                    List<string> properties;

                    if (item == output[0])
                    {
                        string[] tmp = item.ToString().Split(' ');
                        //name = tmp[0];
                        string[] tmp2 = tmp[2].ToString().Split('\\');
                        string[] user_settings = tmp2[1].Split(':');
                        user = user_settings[0];
                        properties = MapSettings(user_settings[1]);
                        fileObject.Add(new Membership { Name = name, User = user, Properties = String.Join(";", properties) });
                    }
                    else
                    {
                        string trimmed = String.Concat(item.ToString().Where(c => !Char.IsWhiteSpace(c)));
                        string[] tmp = trimmed.ToString().Split(':');
                        user = tmp[0];
                        properties = MapSettings(tmp[1]);

                        fileObject.Add(new Membership { Name = name, User = user, Properties = String.Join(";", properties) });
                    }
                }

            }
            WriteCsv(fileObject, path);

        }

        public void WriteCsv(List<Membership> fileObject, string path)
        {
            path += "\\membership.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        private List<string> MapSettings(string sh)
        {
            List<string> mappedValues = new List<string>();

            string[] values = sh.Split(')');
            values = values.Select(x => x.Replace("(", "")).ToArray();
            foreach (var value in values)
            {
                string permission = "";
                switch (value)
                {
                    // iCACLS inheritance settings:
                    case "OI":
                        permission = "object inherit";
                        break;
                    case "CI":
                        permission = "container inherit";
                        break;
                    case "IO":
                        permission = "inherit only";
                        break;
                    case "NP":
                        permission = "don't propagate inherit";
                        break;
                    case "I":
                        permission = "permission inherited from the parent container";
                        break;
                    // list of basic access permissions
                    case "D":
                        permission = "delete access";
                        break;
                    case "F":
                        permission = "full access";
                        break;
                    case "N":
                        permission = "no access";
                        break;
                    case "M":
                        permission = "modify access";
                        break;
                    case "RX":
                        permission = "read and execute access";
                        break;
                    case "R":
                        permission = "read-only access";
                        break;
                    case "W":
                        permission = "write-only access";
                        break;
                    case "":
                        permission = "skip";
                        break;
                    default:
                        permission = value;
                        break;
                }
                if (permission == "skip") break;
                else mappedValues.Add(permission);
            }
            return mappedValues;
        }

        // compare membership
        private void button8_Click(object sender, EventArgs e)
        {
            string firstFile = ReturnFilePath();
            string secondFile = ReturnFilePath();

            List<string> csvAProperties = new List<string>();
            List<string> csvBProperties = new List<string>();

            List<MembershipReduced> A = new List<MembershipReduced>();
            List<MembershipReduced> B = new List<MembershipReduced>();

            using (var reader = new StreamReader(firstFile))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    A.Add(new MembershipReduced { Name = values[0], User = values[1] });
                    csvAProperties.Add(values[2]);
                }
            }
            using (var reader = new StreamReader(secondFile))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    B.Add(new MembershipReduced { Name = values[0], User = values[1] });
                    csvBProperties.Add(values[2]);
                }
            }

            var fileObject = new List<MembershipExtended>();

            // check if A is in B
            for (int i = 1; i < A.Count; i++)
            {
                for (int j = 1; j < B.Count; j++)
                {
                    if (B.Any(prod => prod.Name == A[i].Name && prod.User == A[i].User))
                    {
                        int index = B.FindIndex(x => (x.Name == A[i].Name) && (x.User == A[i].User));
                        if (csvAProperties[i] == csvBProperties[index])
                        {
                            fileObject.Add(new MembershipExtended { Name = A[i].Name, User = A[i].User, Properties = csvAProperties[i], Matching = "Membership is the same" });
                        }
                        else
                        {
                            fileObject.Add(new MembershipExtended { Name = A[i].Name, User = A[i].User, Properties = csvAProperties[i], Matching = "Membership isn't the same" });
                        }
                        break;
                    }
                    else if (!B.Any(prod => prod.Name == A[i].Name && prod.User == A[i].User))
                    {
                        fileObject.Add(new MembershipExtended { Name = A[i].Name, User = A[i].User, Properties = csvAProperties[i], Matching = "Membership not in second FS" });
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }

            // check if A is in B
            for (int i = 1; i < B.Count; i++)
            {
                for (int j = 1; j < A.Count; j++)
                {
                    if (A.Any(prod => prod.Name == B[i].Name && prod.User == B[i].User))
                    {
                        int index = A.FindIndex(x => (x.Name == B[i].Name) && (x.User == B[i].User));
                        if (csvBProperties[i] == csvAProperties[index])
                        {
                            fileObject.Add(new MembershipExtended { Name = B[i].Name, User = B[i].User, Properties = csvBProperties[i], Matching = "Membership is the same" });
                        }
                        else
                        {
                            fileObject.Add(new MembershipExtended { Name = B[i].Name, User = B[i].User, Properties = csvBProperties[i], Matching = "Membership isn't the same" });
                        }
                        break;
                    }
                    else if (!A.Any(prod => prod.Name == B[i].Name && prod.User == B[i].User))
                    {
                        fileObject.Add(new MembershipExtended { Name = B[i].Name, User = B[i].User, Properties = csvBProperties[i], Matching = "Membership not in first FS" });
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }

            var noDupes = fileObject.GroupBy(x => new { x.Name, x.User, x.Properties }).Select(x => x.First()).ToList();
            string[] dirs = firstFile.Split('\\');
            List<string> list = new List<string>(dirs);
            list.RemoveAt(dirs.Length - 1);
            dirs = list.ToArray();
            string path = String.Join("\\", dirs);

            WriteCsv(noDupes, path);
        }

        public void WriteCsv(List<MembershipExtended> fileObject, string path)
        {
            path += "\\membershipCompare.csv";

            using (var stream = File.OpenWrite(path))
            using (var writer = new StreamWriter(stream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(fileObject);
            }
        }

        public class MembershipReduced : IEquatable<MembershipReduced>
        {
            public string Name { get; set; }
            public string User { get; set; }

            public bool Equals(MembershipReduced other)
            {
                return this.Name == other.Name && this.User == other.User;
            }

        }

        public class MembershipExtended : Membership
        {
            public string Matching { get; set; }
        }

        public class Membership
        {
            public string Name { get; set; }
            public string User { get; set; }
            public string Properties { get; set; }
        }

        public class InfoFileExtended : InfoFile
        {
            public string Matching { get; set; }
        }

        public class FileHashExtended : FileHash
        {
            public string Matching { get; set; }

            public string FileName { get; set; }
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
