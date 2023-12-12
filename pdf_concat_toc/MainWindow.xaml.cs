using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;

namespace pdf_concat_toc
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<PdfFile> PdfFiles;

        public MainWindow()
        {
            InitializeComponent();
            PdfFiles = new ObservableCollection<PdfFile>();
            listView_pdf_files.ItemsSource = PdfFiles;
            FILENAME_TXTBOX.Text = DateTime.Now.ToString("yymmddHHMMss");
        }

        //private string SaveFileName;
        private string SaveDirectoryPath;

        public class PdfFile
        {
            public string Path
            {
                get;
                set;
            }

            public string Title
            {
                get;
                set;
            }

            public int ID
            {
                get;
                set;
            }

            public PdfFile(string path, string title, int id)
            {
                Path = path;
                Title = title;
                this.ID = id;
            }
        }

        public void Create_PDF_List(string[] files)
        {
            Array.Sort(files);
            if (files == null)
            {
                return;
            }
            for (int i = 0; i < files.Length; i++)
            {
                if (System.IO.Directory.Exists(files[i]))
                {
                    Create_PDF_List(Directory.GetFiles(files[i], "*", SearchOption.AllDirectories));
                    continue;
                }
                if (System.IO.Path.GetExtension(files[i]) != ".pdf")
                {
                    continue;
                }
                PdfFiles.Add(
                    new PdfFile(Path.GetFullPath(files[i]),
                    Path.GetFileName(files[i]),
                    PdfFiles.Count + 1
                    )
                );
            }
        }

        public void PDF_Drop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files.Length == 0) return;
            if (Directory.Exists(files[0]))
            {
                //フォルダをドロップしたときに、保存ファイル名をフォルダ名に変更する
                FILENAME_TXTBOX.Text = new DirectoryInfo(files[0]).Name;
                //SaveFileName = new DirectoryInfo(files[0]).Name;
            }
            //保存先フォルダを更新する
            SaveDirectoryPath = Path.GetDirectoryName(files[0]);
            Create_PDF_List(e.Data.GetData(DataFormats.FileDrop) as string[]);
        }

        private void Increase_Item_Order(object sender, RoutedEventArgs e)
        {
            int selectedID = listView_pdf_files.SelectedIndex;
            if (selectedID == (PdfFiles.Count - 1) || selectedID == -1)
            {
                return;
            }
            PdfFiles[selectedID].ID += 1;
            PdfFiles[selectedID + 1].ID -= 1;
            PdfFiles = new ObservableCollection<PdfFile>(PdfFiles.OrderBy(n => n.ID));
            listView_pdf_files.ItemsSource = PdfFiles;
        }

        private void Decrease_Item_Order(object sender, RoutedEventArgs e)
        {
            int selectedID = listView_pdf_files.SelectedIndex;
            if (selectedID == 0 || selectedID == -1)
            {
                return;
            }
            PdfFiles[selectedID].ID -= 1;
            PdfFiles[selectedID - 1].ID += 1;
            PdfFiles = new ObservableCollection<PdfFile>(PdfFiles.OrderBy(n => n.ID));
            listView_pdf_files.ItemsSource = PdfFiles;
        }

        private void Remove_Item(object sender, RoutedEventArgs e)
        {
            int selectedID = listView_pdf_files.SelectedIndex;
            if (selectedID != -1)
            {
                PdfFiles.RemoveAt(selectedID);
            }
        }

        private void Concat_PDF_Click(object sender, RoutedEventArgs e)
        {
            if (PdfFiles.Count == 0)
            {
                return;
            }
            Concat_PDF();
            PdfFiles.Clear();
        }

        private void clear_all_item(object sender, RoutedEventArgs e)
        {
            PdfFiles.Clear();
        }


        //目次の作成
        private void Create_TOC(ref Document documents, ref PdfWriter writer)
        {
            var fontName = "meiryo";
            if (!FontFactory.IsRegistered(fontName))
            {
                var fontPath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\meiryo.ttc";
                FontFactory.Register(fontPath, fontName);
            }
            Font pfont =
                    FontFactory.GetFont(fontName,
                    BaseFont.IDENTITY_H,    //横書き
                    BaseFont.EMBEDDED,  //フォントをPDFファイルに組み込まない（重要）
                    10f,                    //フォントサイズ
                    Font.UNDERLINE,           //フォントスタイル
                    BaseColor.BLUE);       //フォントカラー
            Font h1font =
              FontFactory.GetFont(fontName,
              BaseFont.IDENTITY_H,    //横書き
              BaseFont.EMBEDDED,  //フォントをPDFファイルに組み込まない（重要）
              12f,                    //フォントサイズ
              Font.NORMAL,           //フォントスタイル
              BaseColor.BLACK);

            //2ページ目からスタート
            int pageCount = 2;

            documents.NewPage();
            string directoryNameHeaderBuff = string.Empty;
            documents.Add(new Paragraph(directoryNameHeaderBuff, h1font));


            var docHeader = new Paragraph("目次", h1font);
            docHeader.Alignment = Element.ALIGN_CENTER;
            documents.Add(docHeader);

            // 目次の作成
            foreach (PdfFile pdf in PdfFiles)
            {
                //PdfReader.unethicalreading = true;
                PdfReader pdfReader = new PdfReader(File.ReadAllBytes(pdf.Path));
                string directoryNameHeader = new DirectoryInfo(pdf.Path).Parent.Name;

                if (directoryNameHeader != directoryNameHeaderBuff)
                {
                    documents.Add(new Paragraph(20f, directoryNameHeader, h1font));
                    directoryNameHeaderBuff = directoryNameHeader;
                }
                Chunk link = new Chunk(pdf.Title, pfont);
                Chunk dottedLine = new Chunk(new DottedLineSeparator());
                dottedLine.Font = pfont;
                Chunk pageNum = new Chunk(pageCount.ToString(), pfont);
                PdfAction action = PdfAction.GotoLocalPage(pageCount, new PdfDestination(PdfDestination.FIT), writer);
                link.SetAction(action);
                dottedLine.SetAction(action);
                pageNum.SetAction(action);
                var p = new Paragraph(link);
                p.Add(dottedLine);
                p.Add(pageNum);
                documents.Add(p);
                pageCount += pdfReader.NumberOfPages;
            }

        }

        private string Get_Save_PDF_Path()
        {
            SaveFileDialog sdialog = new SaveFileDialog();
            sdialog.InitialDirectory = SaveDirectoryPath;
            sdialog.Title = "保存する場所を指定してください";
            sdialog.DefaultExt = "pdf";
            sdialog.FileName = FILENAME_TXTBOX.Text ?? DateTime.Now.ToString("yymmddHHMMss") + ".pdf";
            var result = sdialog.ShowDialog();
            if (result == false) return string.Empty;
            return sdialog.FileName;
        }

        private void Concat_PDF()
        {
            var saveFileName = Get_Save_PDF_Path();
            if (string.IsNullOrEmpty(saveFileName)) return;

            FileStream stream = new FileStream(saveFileName, FileMode.OpenOrCreate);
            Document documents = new Document();
            PdfWriter writer = PdfWriter.GetInstance(documents, stream);
            documents.Open();
            PdfContentByte joinPcb = writer.DirectContent;
            PdfImportedPage page;
            int rotation;

            if (CREATE_TOC_CHECKBOX.IsChecked ?? false)
            {
                Create_TOC(ref documents, ref writer);
            }
            foreach (PdfFile pdf in PdfFiles)
            {
                PdfReader.unethicalreading = true;
                PdfReader pdfReader = new PdfReader(File.ReadAllBytes(pdf.Path));
                for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                {
                    documents.SetPageSize(pdfReader.GetPageSizeWithRotation(i));
                    documents.NewPage();
                    page = writer.GetImportedPage(pdfReader, i);
                    rotation = pdfReader.GetPageRotation(i);
                    if (rotation == 90)
                        joinPcb.AddTemplate(page, 0, -1, 1, 0, 0, pdfReader.GetPageSizeWithRotation(i).Height);
                    else if (rotation == 180)
                        joinPcb.AddTemplate(page, -1, 0, 1, -1, pdfReader.GetPageSizeWithRotation(i).Width, pdfReader.GetPageSizeWithRotation(i).Height);
                    else if (rotation == 270)
                        joinPcb.AddTemplate(page, 0, 1, -1, 0, pdfReader.GetPageSizeWithRotation(i).Width, 0);
                    else
                        joinPcb.AddTemplate(page, 1, 0, 0, 1, 0, 0);
                }
            }
            documents.Close();
            stream.Close();
            System.Diagnostics.Process.Start(saveFileName);
        }
    }
}
