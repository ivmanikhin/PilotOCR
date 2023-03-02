using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.CreateObjectSample;
using Ascon.Pilot.SDK.Menu;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
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
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
//using System.Windows.Forms;
using TesseractOCR;

namespace PilotOCR
{
    public class PiPage
    {
        public Guid docID { get; set; }
        public string docName { get; set; }
        public string fileName { get; set; }
        public int pageNum { get; set; }
        public string text { get; set; }
        public Image image { get; set; }

        public PiPage()
        {

        }
        public PiPage(Guid docID, string docName, string fileName, int pageNum, string text, Image image)
        {
            this.docID = docID;
            this.docName = docName;
            this.fileName = fileName;
            this.pageNum = pageNum;
            this.text = text;
            this.image = image;
        }
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
        private List<Ascon.Pilot.SDK.IDataObject> _dataObjects = new List<Ascon.Pilot.SDK.IDataObject>();
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
            var netTasks = new ConcurrentBag<Task>();
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
                Task.WaitAll(netTasks.ToArray());
                await Task.Delay(3000);
                MySqlConnection connection = new MySqlConnection(CONNECTION_PARAMETERS);
                connection.Open();
                foreach (Ascon.Pilot.SDK.IDataObject dataObject in _dataObjects)
                {
                    if (cancelled) break;
                    string inputNo = "";
                    string outNo = "";
                    string docId = dataObject.Id.ToString();
                    string docDate = "";
                    string docSubject = "";
                    string docCorrespondent = "";
                    string str = "";
                    List<PiPage> recognizedDoc = new List<PiPage>();
                    Ascon.Pilot.SDK.IDataObject dataObjectMounted;
                    dataObjectMounted = await _loader.Load(dataObject.Id);
                    foreach (KeyValuePair<string, object> attribute in (IEnumerable<KeyValuePair<string, object>>)dataObjectMounted.Attributes)
                    {
                        if (attribute.Value != null)
                        {
                            str = str + attribute.Key.ToString() + ":\n     " + attribute.Value.ToString() + "\n";
                            switch (attribute.Key.ToString())
                            {
                                case "ECM_inbound_letter_counter":
                                    inputNo = attribute.Value.ToString();
                                    break;
                                case "ECM_inbound_letter_number":
                                    outNo = attribute.Value.ToString();
                                    break;
                                case "ECM_inbound_letter_sending_date":
                                    docDate = attribute.Value.ToString();
                                    break;
                                case "ECM_letter_subject":
                                    docSubject = attribute.Value.ToString();
                                    break;
                                case "ECM_letter_correspondent":
                                    docCorrespondent = attribute.Value.ToString();
                                    break;
                                default:
                                    break;
                            }
                        }
                    };
                    progressDialog.SetCurrentDocName(inputNo + " - " + outNo + " - " + docDate + " - " + docSubject);
                    recognizedDoc = RecognizeWholeDoc(dataObjectMounted);
                    if (cancelled) break;
                    string contents = str + "\n";
                    foreach (PiPage piPage in recognizedDoc)
                    {
                        ++pagesCount;
                        contents = contents + piPage.fileName + " " + piPage.pageNum.ToString() + ":\n\n" + piPage.text + "\n\n=======================================================================================================================\n\n";
                    };
                    DocToDB(connection, inputNo, outNo, docId, docDate, docSubject, docCorrespondent, contents);
                    progressDialog.UpdateProgress();
                };

                connection.Close();
                System.Windows.Forms.MessageBox.Show(pagesCount.ToString() + " страниц распознано\n в " + _dataObjects.Count.ToString() + " документах");
                pagesCount = 0;
                if (!cancelled) progressDialog.CloseRemotely();
            });
        }

        public List<Ascon.Pilot.SDK.IDataObject> MakeRecognitionList(List<Ascon.Pilot.SDK.IDataObject> dataObjects)
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

     

        public List<PiPage> RecognizeWholeDoc(Ascon.Pilot.SDK.IDataObject dataObject)
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
                        };
                    }
                    catch (Exception ex)
                    {
                        pieceOfDoc.Add(new PiPage()
                        {
                            fileName = file.Name,
                            text = "FILE IS CORRUPTED " + ex.Message
                        });
                    };
                };
                if (this.IsXpsFile(file.Name))
                {
                    try
                    {
                        using (Stream xpsStream = this._fileProvider.OpenRead(file))
                        {
                            foreach (PiPage page in this.XpsToPages(xpsStream, file.Name))
                                pieceOfDoc.Add(page);
                            xpsStream.Close();
                        };
                    }
                    catch (Exception ex)
                    {
                        pieceOfDoc.Add(new PiPage()
                        {
                            fileName = file.Name,
                            text = "FILE IS CORRUPTED " + ex.Message
                        });
                    };
                };
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
                            fileName = fileName,
                            text = "FILE IS CORRUPTED " + ex.Message
                        });
                    };
                };
                if (this.IsDocFile(storagePath))
                {
                    pieceOfDoc.Add(DocToPage(storagePath));
                };
                if (this.IsTxtFile(storagePath))
                {
                    pieceOfDoc.Add(TxtToPage(storagePath));
                };
                if (this.IsXlsFile(storagePath))
                {
                    pieceOfDoc.Add(XlsToPage(storagePath));
                };
            };
            foreach (PiPage piPage in pieceOfDoc)
            {
                if (cancelled)
                {
                    ctsRecognition.Cancel();
                    break;
                }
                if (piPage.image != null)
                {
                    Task recognitionTask = _taskFactoryRecognition.StartNew(() =>
                    {
                        if (!token.IsCancellationRequested)
                        {
                            piPage.text = PageToText(piPage.image);
                            piPage.image = null;
                        }
                    }, token);
                    recognitionTasks.Add(recognitionTask);
                }
                piPage.docID = dataObject.Id;
                piPage.docName = dataObject.Attributes.FirstOrDefault<KeyValuePair<string, object>>().Value.ToString();
            };
            Task.WaitAll(recognitionTasks.ToArray());
            ctsRecognition.Dispose();
            return pieceOfDoc;
        }


        private PiPage XlsToPage(string storagePath)
        {
            string fileName = Path.GetFileName(storagePath);
            PiPage piPage = new PiPage();
            piPage.fileName = fileName;
            piPage.pageNum = 1;
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
                                                piPage.text = piPage.text + sharedStringTable.ElementAt<OpenXmlElement>(int.Parse(innerText)).InnerText + " ";
                                            }
                                            else
                                            {
                                                piPage.text = piPage.text + innerText + " ";
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
                piPage.text = "FILE IS CORRUPTED " + ex.Message.ToString();
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }


        private PiPage DocToPage(string storagePath)
        {
            string fileName = Path.GetFileName(storagePath);
            PiPage piPage = new PiPage();
            piPage.fileName = fileName;
            piPage.pageNum = 1;
            try
            {
                using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(storagePath, false))
                {
                    piPage.text = wordprocessingDocument.MainDocumentPart.Document.Body.InnerText;
                }
            }
            catch (Exception ex)
            {
                piPage.text = "FILE IS CORRUPTED " + ex.Message.ToString();
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }


        private PiPage TxtToPage(string storagePath)
        {
            string fileName = Path.GetFileName(storagePath);
            PiPage piPage = new PiPage();
            piPage.fileName = fileName;
            piPage.pageNum = 1;
            try
            {
                piPage.text = File.ReadAllText(storagePath, Encoding.Default);
            }
            catch (Exception ex)
            {
                piPage.text = "FILE IS CORRUPTED " + ex.Message.ToString();
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
                    page.pageNum = index + 1;
                    page.fileName = fileName;
                    PdfDocument pdfDocument2 = pdfDocument1;
                    int page1 = index;
                    SizeF pageSiz = pdfDocument1.PageSizes[index];
                    int width = pageSiz.ToSize().Width * 2;
                    pageSiz = pdfDocument1.PageSizes[index];
                    int height = pageSiz.ToSize().Height * 2;
                    page.image = pdfDocument2.Render(page1, width, height, 96f, 96f, false);
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
                page.image = (System.Drawing.Image)new Bitmap(stream);
                page.pageNum = num + 1;
                page.fileName = fileName;
                pages.Add(page);
                ++num;
            };
            return pages;
        }

        public string getResourcesPath()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            int length = location.LastIndexOf('\\');
            return location.Substring(0, length);
        }

        public string PageToText(System.Drawing.Image pageImg)
        {
            string text = (string)null;
            if (pageImg != null)
            {
                using (Engine engine = new Engine(this.getResourcesPath() + "\\tessdata\\", "rus(best)"))
                {
                    System.Drawing.Image image = (System.Drawing.Image)this.TiltDocument((Bitmap)pageImg);
                    //Random random = new Random();
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        image.Save((Stream)memoryStream, ImageFormat.Png);
                        using (TesseractOCR.Page page = engine.Process(TesseractOCR.Pix.Image.LoadFromMemory(memoryStream)))
                            text = Regex.Replace(Regex.Replace(page.Text, " +", " "), "(?<!\\d)-+[ +]?\\n+[ +]?|(?<=[\\d\\wа-яА-я])\\s+(?=\\.)|(?<=\\.)\\s+(?=\\d)", "");
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
            command.ExecuteNonQuery();
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
                using (Graphics graphics = Graphics.FromImage((System.Drawing.Image)inputBmp))
                    graphics.DrawImage((System.Drawing.Image)bitmap3, 0, 0, inputBmp.Width, inputBmp.Height);
                float angle = this.GetOptimumAngle(inputBmp, 10f, 1f) + this.GetOptimumAngle(inputBmp, 1.25f, 0.25f);
                using (Graphics graphics = Graphics.FromImage((System.Drawing.Image)bitmap2))
                {
                    graphics.RotateTransform(angle);
                    graphics.Clear(System.Drawing.Color.White);
                    graphics.DrawImage((System.Drawing.Image)inputPic, (int)((double)inputPic.Height * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5), -(int)((double)inputPic.Width * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5));
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
                using (Graphics graphics = Graphics.FromImage((System.Drawing.Image)bitmap2))
                {
                    graphics.RotateTransform(num1);
                    graphics.Clear(System.Drawing.Color.White);
                    graphics.DrawImage((System.Drawing.Image)inputBmp, (int)((double)height * Math.Sin((double)num1 * 3.1415 / 180.0) * 0.5), -(int)((double)width * Math.Sin((double)num1 * 3.1415 / 180.0) * 0.5));
                };
                using (Graphics graphics = Graphics.FromImage((System.Drawing.Image)bitmap1))
                    graphics.DrawImage((System.Drawing.Image)bitmap2, 0, 0, 1, height);
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

