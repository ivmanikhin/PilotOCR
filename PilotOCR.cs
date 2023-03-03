using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.CreateObjectSample;
using Ascon.Pilot.SDK.Menu;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;
using PdfiumViewer;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
//using System.Windows.Forms;
using TesseractOCR;

namespace PilotOCR
{
    public class PiPage
    {
        //public Guid DocID { get; set; }
        //public string DocName { get; set; }
        public string FileName { get; set; }
        public int PageNum { get; set; }
        public string Text { get; set; }
        public Image Image { get; set; }
        public string CorruptedFilePath { get; set; }

        public PiPage() { }

        public PiPage(/*Guid docID, string docName,*/ string fileName, int pageNum, string text, Image image)
        {
            //this.DocID = docID;
            //this.DocName = docName;
            this.FileName = fileName;
            this.PageNum = pageNum;
            this.Text = text;
            this.Image = image;
        }
    }

    public class PiLetter
    {
        public IDataObject DataObject { get; set; }
        public Guid DocId { get; set; }
        public string OutNo { get; set; }
        public string DocDate { get; set; }
        public string DocSubject { get; set; }
        public string DocCorrespondent { get; set; }
        public List<PiPage> Pages { get; set; }
        public string Text { get; set; }
        public List<string> CorruptedFiles { get; set; }
        public int PagesQtt { get; set; }
        public void SetText()
        {
            foreach (PiPage piPage in Pages)
            {
                this.Text += "\n" + piPage.FileName + " " + piPage.PageNum.ToString() + ":\n\n" + piPage.Text + "\n\n=======================================================================================================================\n\n";
            };
            PagesQtt = Pages.Count;
        }
    }


    public class PiLetterInbox : PiLetter
    {
        public string InputNo { get; set; }

        public PiLetterInbox() { }

        //public PiLetterInbox(Guid docId, string inputNo, string outNo, string docDate, string docSubject, string docCorrespondent)
        //{
        //    this.docId = docId;
        //    this.inputNo = inputNo;
        //    this.outNo = outNo;
        //    this.docDate = docDate;
        //    this.docSubject = docSubject;
        //    this.docCorrespondent = docCorrespondent;
        //    this.text = "";
        //}

        public void ReadAttributes()
        {
            foreach (KeyValuePair<string, object> attribute in (IEnumerable<KeyValuePair<string, object>>) DataObject.Attributes)
            {
                if (attribute.Value != null)
                {
                    Text += attribute.Key.ToString() + ":\n     " + attribute.Value.ToString() + "\n";
                    switch (attribute.Key.ToString())
                    {
                        case "ECM_inbound_letter_counter":
                            InputNo = attribute.Value.ToString();
                            break;
                        case "ECM_inbound_letter_number":
                            OutNo = attribute.Value.ToString();
                            break;
                        case "ECM_inbound_letter_sending_date":
                            DocDate = attribute.Value.ToString();
                            break;
                        case "ECM_letter_subject":
                            DocSubject = attribute.Value.ToString();
                            break;
                        case "ECM_letter_correspondent":
                            DocCorrespondent = attribute.Value.ToString();
                            break;
                        default:
                            break;
                    }
                }
            };
        }

    }

    public class PiLetterSent : PiLetter
    {

    }


    [Export(typeof(IMenu<ObjectsViewContext>))]


    public class ModifyObjectsPlugin : IMenu<ObjectsViewContext>
    {
        private const string CONNECTION_PARAMETERS = "datasource=localhost;port=3306;username=root;password=C@L0P$Ck;charset=utf8";
//        private const string PATH = "D:\\TEMP\\Recognized\\";
        private TaskFactory _taskFactoryRecognition;
        private readonly IXpsRender _xpsRender;
        private readonly IFileProvider _fileProvider;
        private readonly IObjectModifier _modifier;
        private readonly IObjectsRepository _objectsRepository;
        private readonly ObjectLoader _loader;
        private List<IDataObject> _dataObjects = new List<IDataObject>();
        private readonly LimitedConcurrencyLevelTaskScheduler lctsRecognition = new LimitedConcurrencyLevelTaskScheduler(8);
        private int docsCount = 0; 
        private int pagesCount = 0;
        private bool cancelled = false;


        [ImportingConstructor]
        public ModifyObjectsPlugin(IObjectModifier modifier, IXpsRender xpsRender, IFileProvider fileProvider, IObjectsRepository objectsRepository)
        {
            this._modifier = modifier;
            this._xpsRender = xpsRender;
            this._fileProvider = fileProvider;
            this._objectsRepository = objectsRepository;
            this._loader = new ObjectLoader(_objectsRepository);
        }

        public void Build(IMenuBuilder builder, ObjectsViewContext context)
        {
            if (context.IsContext)
                return;
            this._dataObjects = context.SelectedObjects.ToList<Ascon.Pilot.SDK.IDataObject>();
            if (this._dataObjects.Count<Ascon.Pilot.SDK.IDataObject>() < 1)
                return;
            builder.AddItem("RecognizeItemName", 0).WithHeader("Распознать");
        }

        private bool IsXpsFile(string fileName) => Path.GetExtension(fileName) == ".xps";

        private bool IsDocFile(string fileName) => ((IEnumerable<string>)new string[2]
        {
      ".doc",
      ".docx"
        }).Contains<string>(Path.GetExtension(fileName));

        private bool IsXlsFile(string fileName) => ((IEnumerable<string>)new string[2]
        {
      ".xls",
      ".xlsx"
        }).Contains<string>(Path.GetExtension(fileName));

        private bool IsTxtFile(string fileName) => Path.GetExtension(fileName) == ".txt";

        private bool IsPdfFile(string fileName) => Path.GetExtension(fileName) == ".pdf";

        public void KillThemAll()
        {
            cancelled = true;
        }

        public void OnMenuItemClick(string name, ObjectsViewContext context)
        {

            if (!(name == "RecognizeItemName"))
                return;
            ProgressDialog progressDialog = new ProgressDialog(this);
            cancelled = false;
            _taskFactoryRecognition = new TaskFactory(lctsRecognition);
            docsCount = _dataObjects.Count;
            pagesCount = 0;
            _dataObjects = MakeRecognitionList(_dataObjects);
            progressDialog.SetMax(_dataObjects.Count);
            Task progressDialogTask = Task.Run(() => System.Windows.Forms.Application.Run(progressDialog));
            Task.Run(async () =>
            {
                foreach (Ascon.Pilot.SDK.IDataObject dataObject in _dataObjects)
                {
                    if (cancelled) break;
                    _objectsRepository.Mount(dataObject.Id);
                    await Task.Delay(50);
                };
                await Task.Delay(3000);
                MySqlConnection connection = new MySqlConnection(CONNECTION_PARAMETERS);
                connection.Open();
                foreach (Ascon.Pilot.SDK.IDataObject dataObject in _dataObjects)
                {
                    if (cancelled) break;
                    PiLetterInbox piLetter = new PiLetterInbox();
                    piLetter.DataObject = await _loader.Load(dataObject.Id);
                    piLetter.ReadAttributes();
                    progressDialog.SetCurrentDocName(piLetter.InputNo + " - " + piLetter.OutNo + " - " + piLetter.DocDate + " - " + piLetter.DocSubject);
                    piLetter.Pages = RecognizeWholeDoc(piLetter.DataObject);
                    if (cancelled) break;
                    piLetter.SetText();
                    pagesCount += piLetter.PagesQtt;
                    DocToDB(connection, piLetter.InputNo, piLetter.OutNo, piLetter.DocId.ToString(), piLetter.DocDate, piLetter.DocSubject, piLetter.DocCorrespondent, piLetter.Text);
                    progressDialog.UpdateProgress();
                };

                connection.Close();
                System.Windows.Forms.MessageBox.Show(pagesCount.ToString() + " страниц распознано\n в " + _dataObjects.Count.ToString() + " документах");
                pagesCount = 0;
                if (!cancelled) progressDialog.CloseRemotely();
            });
        }

        public List<IDataObject> MakeRecognitionList(List<IDataObject> dataObjects)
        {
            List<Ascon.Pilot.SDK.IDataObject> recognitionList = new List<Ascon.Pilot.SDK.IDataObject>();
            foreach (Ascon.Pilot.SDK.IDataObject dataObject in dataObjects)
            {
                if (dataObject.Attributes.Count < 1 || !dataObject.Type.IsMountable)
                    continue;
                //object letterSubject;
                //object letterDate;
                //string fullFileName = "";
                //string letterInboxNum = dataObject.Attributes.FirstOrDefault().Value.ToString();
                //bool letterSubjectExists = dataObject.Attributes.TryGetValue("ECM_letter_subject", out letterSubject);
                //bool letterDateExists = dataObject.Attributes.TryGetValue("ECM_inbound_letter_sending_date", out letterDate);
                //if (letterSubjectExists & letterDateExists)
                //    fullFileName = PATH
                //                    + letterInboxNum
                //                    + " - " + letterDate.ToString().Substring(0, 10)
                //                    + " - " + letterSubject.ToString().Replace('/', '-').Replace('|', '-').Replace('*', ' ').Replace('\\', '-')
                //                                                    .Replace('"', ' ').Replace('?', ' ').Replace('\t', ' ')
                //                                                    .Replace('<', ' ')
                //                                                    .Replace('>', ' ')
                //                                                    .Replace(':', ' ') + ".txt";
                //else if (letterSubjectExists)
                //    fullFileName = PATH
                //        + letterInboxNum + " - " 
                //        + letterSubject.ToString().Replace('/', '-').Replace('|', '-').Replace('*', ' ').Replace('\\', '-')
                //                                                    .Replace('"', ' ').Replace('?', ' ').Replace('\t', ' ')
                //                                                    .Replace('<', ' ')
                //                                                    .Replace('>', ' ')
                //                                                    .Replace(':', ' ') + ".txt";
                //else
                //    fullFileName = PATH + letterInboxNum + ".txt";
                //if (File.Exists(fullFileName) || File.Exists(PATH + letterInboxNum + ".txt"))
                //    continue;
                //dataObject.Attributes.Add("fullFileName", fullFileName);
                recognitionList.Add(dataObject);

            }
            return recognitionList;
        }

     

        public List<PiPage> RecognizeWholeDoc(IDataObject dataObject)
        {
            var ctsRecognition = new CancellationTokenSource();
            var token = ctsRecognition.Token;
            var recognitionTasks = new ConcurrentBag<Task>();
            List<PiPage> pieceOfDoc = new List<PiPage>();
            foreach (IFile file in dataObject.ActualFileSnapshot.Files)
            {
                if (cancelled) break;
                if (this.IsPdfFile(file.Name))
                {
                    try
                    {
                        using (Stream pdfStream = this._fileProvider.OpenRead(file))
                        {
                            foreach (PiPage page in this.PdfToPages(pdfStream, file.Name))
                                pieceOfDoc.Add(page);
                            pdfStream.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        pieceOfDoc.Add(new PiPage()
                        {
                            FileName = file.Name,
                            Text = "FILE IS CORRUPTED " + ex.Message
                        });
                    }
                }
                else if (this.IsXpsFile(file.Name))
                {
                    try
                    {
                        using (Stream xpsStream = this._fileProvider.OpenRead(file))
                        {
                            foreach (PiPage page in this.XpsToPages(xpsStream, file.Name))
                                pieceOfDoc.Add(page);
                            xpsStream.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        pieceOfDoc.Add(new PiPage()
                        {
                            FileName = file.Name,
                            Text = "FILE IS CORRUPTED " + ex.Message
                        });
                    }
                }
            };
            foreach (Guid child in dataObject.Children)
            {
                if (cancelled) break;
                Guid fileGuid = child;
                string storagePath = this._objectsRepository.GetStoragePath(fileGuid);
                if (this.IsPdfFile(storagePath))
                {
                    string fileName = Path.GetFileName(storagePath);
                    try
                    {
                        using (FileStream pdfStream = File.OpenRead(storagePath))
                        {
                            foreach (PiPage page in this.PdfToPages((Stream)pdfStream, fileName))
                                pieceOfDoc.Add(page);
                            pdfStream.Close();
                        };
                    }
                    catch (Exception ex)
                    {
                        pieceOfDoc.Add(new PiPage()
                        {
                            FileName = fileName,
                            CorruptedFilePath = storagePath,
                            Text = "FILE IS CORRUPTED " + ex.Message
                        });
                    }
                }
                else if (this.IsDocFile(storagePath))
                {
                    pieceOfDoc.Add(DocToPage(storagePath));
                }
                else if (this.IsTxtFile(storagePath))
                {
                    pieceOfDoc.Add(TxtToPage(storagePath));
                }
                else if (this.IsXlsFile(storagePath))
                {
                    pieceOfDoc.Add(XlsToPage(storagePath));
                }
            };
            foreach (PiPage piPage in pieceOfDoc)
            {
                if (cancelled)
                {
                    ctsRecognition.Cancel();
                }
                else if (piPage.Image != null)
                {
                    Task recognitionTask = _taskFactoryRecognition.StartNew(() =>
                    {
                        if (!token.IsCancellationRequested)
                        {
                            piPage.Text = PageToText(piPage.Image);
                            piPage.Image = null;
                        }
                    }, token);
                    recognitionTasks.Add(recognitionTask);
                }
                //piPage.DocID = dataObject.Id;
                //piPage.DocName = dataObject.Attributes.FirstOrDefault<KeyValuePair<string, object>>().Value.ToString();
            };
            Task.WaitAll(recognitionTasks.ToArray());
            ctsRecognition.Dispose();
            return pieceOfDoc;
        }


        private PiPage XlsToPage(string storagePath)
        {
            string fileName = Path.GetFileName(storagePath);
            PiPage piPage = new PiPage();
            piPage.FileName = fileName;
            piPage.PageNum = 1;
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(storagePath, false))
                {
                    SharedStringTable sharedStringTable = spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable;
                    foreach (WorksheetPart worksheetPart in spreadsheetDocument.WorkbookPart.WorksheetParts)
                    {
                        foreach (SheetData element1 in worksheetPart.Worksheet.Elements<SheetData>())
                        {
                            if (element1.HasChildren)
                            {
                                foreach (OpenXmlElement element2 in element1.Elements<Row>())
                                {
                                    foreach (Cell element3 in element2.Elements<Cell>())
                                    {
                                        string innerText = element3.InnerText;
                                        if (element3.DataType != null)
                                        {
                                            if ((CellValues)element3.DataType == CellValues.SharedString)
                                            {
                                                piPage.Text = piPage.Text + sharedStringTable.ElementAt<OpenXmlElement>(int.Parse(innerText)).InnerText + " ";
                                            }
                                            else
                                            {
                                                piPage.Text = piPage.Text + innerText + " ";
                                            };
                                        };
                                    };
                                };
                            };
                        };
                    };
                };
            }
            catch (Exception ex)
            {
                piPage.CorruptedFilePath = storagePath;
                piPage.Text = "FILE IS CORRUPTED " + ex.Message.ToString();
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }


        private PiPage DocToPage(string storagePath)
        {
            string fileName = Path.GetFileName(storagePath);
            PiPage piPage = new PiPage();
            piPage.FileName = fileName;
            piPage.PageNum = 1;
            try
            {
                using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(storagePath, false))
                {
                    piPage.Text = wordprocessingDocument.MainDocumentPart.Document.Body.InnerText;
                }
            }
            catch (Exception ex)
            {
                piPage.CorruptedFilePath = storagePath;
                piPage.Text = "FILE IS CORRUPTED " + ex.Message.ToString();
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }


        private PiPage TxtToPage(string storagePath)
        {
            string fileName = Path.GetFileName(storagePath);
            PiPage piPage = new PiPage();
            piPage.FileName = fileName;
            piPage.PageNum = 1;
            try
            {
                piPage.Text = File.ReadAllText(storagePath, Encoding.Default);
            }
            catch (Exception ex)
            {
                piPage.CorruptedFilePath = storagePath;
                piPage.Text = "FILE IS CORRUPTED " + ex.Message.ToString();
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }

        private List<PiPage> PdfToPages(Stream pdfStream, string fileName)
        {
            List<PiPage> pages = new List<PiPage>();
            using (PdfDocument pdfDocument1 = PdfDocument.Load(pdfStream))
            {
                for (int index = 0; index < pdfDocument1.PageCount; ++index)
                {
                    PiPage page = new PiPage();
                    page.PageNum = index + 1;
                    page.FileName = fileName;
                    PdfDocument pdfDocument2 = pdfDocument1;
                    int page1 = index;
                    SizeF pageSiz = pdfDocument1.PageSizes[index];
                    int width = pageSiz.ToSize().Width * 2;
                    pageSiz = pdfDocument1.PageSizes[index];
                    int height = pageSiz.ToSize().Height * 2;
                    page.Image = pdfDocument2.Render(page1, width, height, 96f, 96f, false);
                    pages.Add(page);
                };
                return pages;
            };
        }

        public List<PiPage> XpsToPages(Stream xpsStream, string fileName)
        {
            List<PiPage> pages = new List<PiPage>();
            IEnumerable<Stream> bitmap = this._xpsRender.RenderXpsToBitmap(xpsStream, 2.0);
            int num = 0;
            foreach (Stream stream in bitmap.ToList<Stream>())
            {
                PiPage page = new PiPage();
                page.Image = (System.Drawing.Image)new Bitmap(stream);
                page.PageNum = num + 1;
                page.FileName = fileName;
                pages.Add(page);
                ++num;
            };
            return pages;
        }

        public string GetResourcesPath()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            int length = location.LastIndexOf('\\');
            return location.Substring(0, length);
        }

        public string PageToText(Image pageImg)
        {
            string text = (string)null;
            if (pageImg != null)
            {
                using (Engine engine = new Engine(this.GetResourcesPath() + "\\tessdata\\", "rus(best)"))
                {
                    Image image = (Image)this.TiltDocument((Bitmap)pageImg);
                    //Random random = new Random();
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        image.Save((Stream)memoryStream, ImageFormat.Png);
                        using (TesseractOCR.Page page = engine.Process(TesseractOCR.Pix.Image.LoadFromMemory(memoryStream)))
                            //text = Regex.Replace(Regex.Replace(page.Text, " +", " "), "(?<!\\d)-+[ +]?\\n+[ +]?|(?<=[\\d\\wа-яА-я])\\s+(?=\\.)|(?<=\\.)\\s+(?=\\d)", "");
                            text = page.Text;
                    };
                };
            };
            return text;
        }

        private void DocToDB(MySqlConnection connection, string inputNo, string outNo, string docId, string docDate, string docSubject, string docCorrespondent, string docText)
        {
            docText = MySqlHelper.EscapeString(docText);
            outNo = MySqlHelper.EscapeString(outNo);
            docId = MySqlHelper.EscapeString(docId);
            docDate = MySqlHelper.EscapeString(docDate);
            docSubject = MySqlHelper.EscapeString(docSubject);
            docCorrespondent = MySqlHelper.EscapeString(docCorrespondent);
            //docSubject = docSubject.Replace("'", " ").Replace("\"", " ");
            string commandText = "INSERT INTO pilotsql.inbox(input_no, out_no, doc_id, date, subject, correspondent, text) VALUES (@inputNo, @outNo, @docId, @docDate, @docSubject, @docCorrespondent, @docText)";

            MySqlCommand command = new MySqlCommand(commandText, connection);
            command.Parameters.AddWithValue("@inputNo", inputNo);
            command.Parameters.AddWithValue("@outNo", outNo);
            command.Parameters.AddWithValue("@docId", docId);
            command.Parameters.AddWithValue("@docDate", docDate);
            command.Parameters.AddWithValue("@docSubject", docSubject);
            command.Parameters.AddWithValue("@docCorrespondent", docCorrespondent);
            command.Parameters.AddWithValue("@docText", docText);
            //MessageBox.Show(commandText);
            try
            {

            command.ExecuteNonQuery();
            }
            catch (MySqlException ex){ System.Windows.Forms.MessageBox.Show(ex.Message); };
            //    MessageBox.Show("Data Inserted");
            //else
            //    MessageBox.Show("Failed");
            
        }

        private Bitmap TiltDocument(Bitmap inputPic)
        {
            Bitmap bitmap1 = new Bitmap(inputPic.Width * 2 / 3, inputPic.Height / 2);
            Bitmap bitmap2 = new Bitmap(inputPic.Width, inputPic.Height);
            int height = 640;
            Bitmap inputBmp = new Bitmap(height * 2 / 3, height);
            Rectangle rect = new Rectangle((inputPic.Width - bitmap1.Width) / 2, (inputPic.Height - bitmap1.Height) / 2, (inputPic.Width + bitmap1.Width) / 2, (inputPic.Height + bitmap1.Height) / 2);
            Bitmap bitmap3 = inputPic.Clone(rect, inputPic.PixelFormat);
            //Task graphicsTask = _taskFactoryGrapgic.StartNew(() =>
            //{
                using (Graphics graphics = Graphics.FromImage((Image)inputBmp))
                    graphics.DrawImage((Image)bitmap3, 0, 0, inputBmp.Width, inputBmp.Height);
                float angle = this.GetOptimumAngle(inputBmp, 10f, 1f) + this.GetOptimumAngle(inputBmp, 1.25f, 0.25f);
                using (Graphics graphics = Graphics.FromImage((Image)bitmap2))
                {
                    graphics.RotateTransform(angle);
                    graphics.Clear(System.Drawing.Color.White);
                    graphics.DrawImage((Image)inputPic, (int)((double)inputPic.Height * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5), -(int)((double)inputPic.Width * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5));
                };
            //}, ctsGrapgic.Token);
            //graphicsTask.Wait();
            return bitmap2;
        }

        private float GetOptimumAngle(Bitmap inputBmp, float amplitude, float stepSize)
        {
            int width = inputBmp.Width;
            int height = inputBmp.Height;
            Bitmap bitmap1 = new Bitmap(1, height);
            int[] numArray1 = new int[height];
            Dictionary<float, int> source = new Dictionary<float, int>();
            for (float num1 = -amplitude; (double)num1 < (double)amplitude; num1 += stepSize)
            {
                int num2 = 765;
                Bitmap bitmap2 = new Bitmap(width, height);
                using (Graphics graphics = Graphics.FromImage((Image)bitmap2))
                {
                    graphics.RotateTransform(num1);
                    graphics.Clear(System.Drawing.Color.White);
                    graphics.DrawImage((Image)inputBmp, (int)((double)height * Math.Sin((double)num1 * 3.1415 / 180.0) * 0.5), -(int)((double)width * Math.Sin((double)num1 * 3.1415 / 180.0) * 0.5));
                };
                using (Graphics graphics = Graphics.FromImage((Image)bitmap1))
                    graphics.DrawImage((Image)bitmap2, 0, 0, 1, height);
                for (int y = 0; y < height; ++y)
                {
                    int[] numArray2 = numArray1;
                    int index = y;
                    int num3 = num2;
                    System.Drawing.Color pixel = bitmap1.GetPixel(0, y);
                    int r1 = (int)pixel.R;
                    pixel = bitmap1.GetPixel(0, y);
                    int g1 = (int)pixel.G;
                    int num4 = r1 + g1;
                    pixel = bitmap1.GetPixel(0, y);
                    int b1 = (int)pixel.B;
                    int num5 = num4 + b1;
                    int num6 = Math.Abs(num3 - num5);
                    numArray2[index] = num6;
                    pixel = bitmap1.GetPixel(0, y);
                    int r2 = (int)pixel.R;
                    pixel = bitmap1.GetPixel(0, y);
                    int g2 = (int)pixel.G;
                    int num7 = r2 + g2;
                    pixel = bitmap1.GetPixel(0, y);
                    int b2 = (int)pixel.B;
                    num2 = num7 + b2;
                };
                System.Array.Sort<int>(numArray1);
                System.Array.Reverse((System.Array)numArray1);
                source.Add(num1, ((IEnumerable<int>)numArray1).Take<int>(10).Sum());
            }
            return source.Aggregate<KeyValuePair<float, int>>((Func<KeyValuePair<float, int>, KeyValuePair<float, int>, KeyValuePair<float, int>>)((l, r) => l.Value <= r.Value ? r : l)).Key;
        }
    }
}

