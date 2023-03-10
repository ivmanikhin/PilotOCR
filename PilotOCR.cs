//using System;
//using System.Collections.Generic;
//using System.Linq;
////using System.Text;
////using System.Threading.Tasks;
//using System.ComponentModel.Composition;
//using Ascon.Pilot.SDK;
//using Ascon.Pilot.SDK.Menu;
////using Ascon.Pilot.SDK.CreateObjectSample;
//using System.IO;
//using System.Text.RegularExpressions;
//using TesseractOCR;
//using System.Drawing;
//using System.Threading;
//using System.Threading.Tasks;
//using Spire.Pdf;
////using Aspose.Pdf.
////using System.Windows;

using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.CreateObjectSample;
using Ascon.Pilot.SDK.Menu;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PdfiumViewer;
using System;
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

    public class LimitedConcurrencyLevelTaskScheduler : TaskScheduler
    {
        // Indicates whether the current thread is processing work items.
        [ThreadStatic]
        private static bool _currentThreadIsProcessingItems;

        // The list of tasks to be executed
        private readonly LinkedList<Task> _tasks = new LinkedList<Task>(); // protected by lock(_tasks)

        // The maximum concurrency level allowed by this scheduler.
        private readonly int _maxDegreeOfParallelism;

        // Indicates whether the scheduler is currently processing work items.
        private int _delegatesQueuedOrRunning = 0;

        // Creates a new instance with the specified degree of parallelism.
        public LimitedConcurrencyLevelTaskScheduler(int maxDegreeOfParallelism)
        {
            if (maxDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException("maxDegreeOfParallelism");
            _maxDegreeOfParallelism = maxDegreeOfParallelism;
        }

        // Queues a task to the scheduler.
        protected sealed override void QueueTask(Task task)
        {
            // Add the task to the list of tasks to be processed.  If there aren't enough
            // delegates currently queued or running to process tasks, schedule another.
            lock (_tasks)
            {
                _tasks.AddLast(task);
                if (_delegatesQueuedOrRunning < _maxDegreeOfParallelism)
                {
                    ++_delegatesQueuedOrRunning;
                    NotifyThreadPoolOfPendingWork();
                }
            }
        }

        // Inform the ThreadPool that there's work to be executed for this scheduler.
        private void NotifyThreadPoolOfPendingWork()
        {
            ThreadPool.UnsafeQueueUserWorkItem(_ =>
            {
                // Note that the current thread is now processing work items.
                // This is necessary to enable inlining of tasks into this thread.
                _currentThreadIsProcessingItems = true;
                try
                {
                    // Process all available items in the queue.
                    while (true)
                    {
                        Task item;
                        lock (_tasks)
                        {
                            // When there are no more items to be processed,
                            // note that we're done processing, and get out.
                            if (_tasks.Count == 0)
                            {
                                --_delegatesQueuedOrRunning;
                                break;
                            }

                            // Get the next item from the queue
                            item = _tasks.First.Value;
                            _tasks.RemoveFirst();
                        }

                        // Execute the task we pulled out of the queue
                        base.TryExecuteTask(item);
                    }
                }
                // We're done processing items on the current thread
                finally { _currentThreadIsProcessingItems = false; }
            }, null);
        }

        // Attempts to execute the specified task on the current thread.
        protected sealed override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
        {
            // If this thread isn't already processing a task, we don't support inlining
            if (!_currentThreadIsProcessingItems) return false;

            // If the task was previously queued, remove it from the queue
            if (taskWasPreviouslyQueued)
                // Try to run the task.
                if (TryDequeue(task))
                    return base.TryExecuteTask(task);
                else
                    return false;
            else
                return base.TryExecuteTask(task);
        }

        // Attempt to remove a previously scheduled task from the scheduler.
        protected sealed override bool TryDequeue(Task task)
        {
            lock (_tasks) return _tasks.Remove(task);
        }

        // Gets the maximum concurrency level supported by this scheduler.
        public sealed override int MaximumConcurrencyLevel { get { return _maxDegreeOfParallelism; } }

        // Gets an enumerable of the tasks currently scheduled on this scheduler.
        protected sealed override IEnumerable<Task> GetScheduledTasks()
        {
            bool lockTaken = false;
            try
            {
                Monitor.TryEnter(_tasks, ref lockTaken);
                if (lockTaken) return _tasks;
                else throw new NotSupportedException();
            }
            finally
            {
                if (lockTaken) Monitor.Exit(_tasks);
            }
        }
    }


    [Export(typeof(IMenu<ObjectsViewContext>))]


    public class ModifyObjectsPlugin : IMenu<ObjectsViewContext>
    {
        private const string PATH = "D:\\TEMP\\Recognized\\";
        private TaskFactory _taskFactory, _taskFactoryNet;
        private readonly IXpsRender _xpsRender;
        private readonly IFileProvider _fileProvider;
        private readonly IObjectModifier _modifier;
        private readonly IObjectsRepository _objectsRepository;
        private readonly ObjectLoader _loader;
        private const string RECOGNIZE_ITEM_NAME = "RecognizeItemName";
        private List<Ascon.Pilot.SDK.IDataObject> _dataObjects = new List<Ascon.Pilot.SDK.IDataObject>();
        private readonly LimitedConcurrencyLevelTaskScheduler lctsNet = new LimitedConcurrencyLevelTaskScheduler(4);
        private readonly LimitedConcurrencyLevelTaskScheduler lcts = new LimitedConcurrencyLevelTaskScheduler(6);
        private List<PiPage> recognizedDoc = new List<PiPage>();
        private int pagesCount = 0;
        private List<Task> taskList = new List<Task>();
        private bool letterSubjectExists = false;
        private bool letterDateExists = false;
        private string letterInboxNum;
        private object letterSubject;
        private object letterDate;
        private string fullFileName;
        

        [ImportingConstructor]
        public ModifyObjectsPlugin(
          IObjectModifier modifier,
          IXpsRender xpsRender,
          IFileProvider fileProvider,
          IObjectsRepository objectsRepository)
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

        public void OnMenuItemClick(string name, ObjectsViewContext context)
        {

            int docsCount = _dataObjects.Count();
            if (!(name == "RecognizeItemName"))
                return;
            CancellationTokenSource cts = new CancellationTokenSource();
            this._taskFactory = new TaskFactory(lcts);
            this._taskFactoryNet = new TaskFactory(lctsNet);

            Task.Run(async () =>
            {
                foreach (Ascon.Pilot.SDK.IDataObject dataObject in _dataObjects)
                {
                    if (dataObject.Attributes.Count > 0)
                    {
                        letterInboxNum = dataObject.Attributes.FirstOrDefault().Value.ToString();
                        letterSubjectExists = dataObject.Attributes.TryGetValue("ECM_letter_subject", out letterSubject);
                        letterDateExists = dataObject.Attributes.TryGetValue("ECM_inbound_letter_sending_date", out letterDate);
                        if (letterSubjectExists & letterDateExists)
                            fullFileName = PATH + letterInboxNum + " - " + letterDate.ToString().Substring(0, 10) + " - " + letterSubject.ToString().Replace('/', '-').Replace('"', ' ').Replace('<', ' ').Replace('>', ' ').Replace(':', ' ') + ".txt";
                        else if (letterSubjectExists)
                            fullFileName = PATH + letterInboxNum + " - " + letterSubject.ToString().Replace('/', '-').Replace('"', ' ').Replace('<', ' ').Replace('>', ' ').Replace(':', ' ') + ".txt";
                        else
                            fullFileName = PATH + letterInboxNum + ".txt";
                        if (!File.Exists(fullFileName) && !File.Exists(PATH + letterInboxNum + ".txt"))
                        {
                            if (dataObject.Type.IsMountable)
                            {
                                _objectsRepository.Mount(dataObject.Id);
                                await Task.Delay(500);
                            }
                            DoTheJob(dataObject.Id);
                            //string str = "";
                            //recognizedDoc = RecognizeWholeDoc(dataObject);
                            //foreach (KeyValuePair<string, object> attribute in (IEnumerable<KeyValuePair<string, object>>)dataObject.Attributes)
                            //{
                            //    if (attribute.Value != null)
                            //        str = str + attribute.Key.ToString() + ":\n     " + attribute.Value.ToString() + "\n";
                            //}
                            //string contents = str + "\n";
                            //foreach (PiPage piPage in recognizedDoc)
                            //{
                            //    ++pagesCount;
                            //    contents = contents + piPage.fileName + " " + piPage.pageNum.ToString() + ":\n\n" + piPage.text + "\n\n=======================================================================================================================\n\n";
                            //}
                            //try
                            //{
                            //    File.WriteAllText(fullFileName, contents);
                            //}
                            //catch
                            //{
                            //    File.WriteAllText(PATH + letterInboxNum + ".txt", contents);
                            //}
                        }
                        else if (File.Exists(PATH + letterInboxNum + ".txt"))
                        {
                            try
                            {
                                File.Move(PATH + letterInboxNum + ".txt", fullFileName);
                            }
                            catch
                            {
                            }
                        }
                    }
                }
                cts.Dispose();
                int num = (int)System.Windows.Forms.MessageBox.Show(pagesCount.ToString() + " страниц найдено.\n в " + docsCount.ToString() + " документах");
                this._dataObjects = (List<Ascon.Pilot.SDK.IDataObject>)null;
                recognizedDoc = (List<PiPage>)null;
                GC.Collect();
            });
        }

        public async void DoTheJob(Guid guid)
        {
            Ascon.Pilot.SDK.IDataObject dataObject = await _loader.Load(guid);
            string str = "";
            recognizedDoc = RecognizeWholeDoc(dataObject);
            foreach (KeyValuePair<string, object> attribute in (IEnumerable<KeyValuePair<string, object>>)dataObject.Attributes)
            {
                if (attribute.Value != null)
                    str = str + attribute.Key.ToString() + ":\n     " + attribute.Value.ToString() + "\n";
            }
            string contents = str + "\n";
            foreach (PiPage piPage in recognizedDoc)
            {
                ++pagesCount;
                contents = contents + piPage.fileName + " " + piPage.pageNum.ToString() + ":\n\n" + piPage.text + "\n\n=======================================================================================================================\n\n";
            }
            try
            {
                File.WriteAllText(fullFileName, contents);
            }
            catch
            {
                File.WriteAllText(PATH + letterInboxNum + ".txt", contents);
            }
        }

        public List<PiPage> RecognizeWholeDoc(Ascon.Pilot.SDK.IDataObject dataObject)
        {
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            List<Task> taskList = new List<Task>();
            List<PiPage> pieceOfDoc = new List<PiPage>();
            foreach (IFile file1 in dataObject.ActualFileSnapshot.Files)
            {
                IFile file = file1;
                Task task = this._taskFactory.StartNew((Action)(() =>
                {
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
                                fileName = file.Name,
                                text = "FILE IS CORRUPTED " + ex.Message
                            });
                        }
                    }
                    if (!this.IsXpsFile(file.Name))
                        return;
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
                            fileName = file.Name,
                            text = "FILE IS CORRUPTED " + ex.Message
                        });
                    }
                }), cancellationTokenSource.Token);
                taskList.Add(task);
            }
            foreach (Guid child in dataObject.Children)
            {
                Guid fileGuid = child;
                Task task = this._taskFactory.StartNew((Action)(() =>
                {
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
                            }
                        }
                        catch (Exception ex)
                        {
                            pieceOfDoc.Add(new PiPage()
                            {
                                fileName = fileName,
                                text = "FILE IS CORRUPTED " + ex.Message
                            });
                        }
                    }
                    if (this.IsDocFile(storagePath))
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
                                pieceOfDoc.Add(piPage);
                            }
                        }
                        catch (Exception ex)
                        {
                            piPage.text = "FILE IS CORRUPTED " + ex.Message;
                            pieceOfDoc.Add(piPage);
                        }
                    }
                    if (this.IsTxtFile(storagePath))
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
                            piPage.text = "FILE IS CORRUPTED " + ex.Message;
                        }
                        pieceOfDoc.Add(piPage);
                    }
                    if (!this.IsXlsFile(storagePath))
                        return;
                    string fileName1 = Path.GetFileName(storagePath);
                    PiPage piPage1 = new PiPage();
                    piPage1.fileName = fileName1;
                    piPage1.pageNum = 1;
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
                                                        PiPage piPage2 = piPage1;
                                                        piPage2.text = piPage2.text + sharedStringTable.ElementAt<OpenXmlElement>(int.Parse(innerText)).InnerText + " ";
                                                    }
                                                    else
                                                    {
                                                        PiPage piPage3 = piPage1;
                                                        piPage3.text = piPage3.text + innerText + " ";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        piPage1.text = "FILE IS CORRUPTED " + ex.Message;
                    }
                    pieceOfDoc.Add(piPage1);
                }), cancellationTokenSource.Token);
                taskList.Add(task);
            }
            Task.WaitAll(taskList.ToArray());
            foreach (PiPage piPage in pieceOfDoc)
            {
                piPage.docID = dataObject.Id;
                piPage.docName = dataObject.Attributes.FirstOrDefault<KeyValuePair<string, object>>().Value.ToString();
            }
            return pieceOfDoc;
        }

        public List<PiPage> PdfToPages(Stream pdfStream, string fileName)
        {
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            List<Task> taskList = new List<Task>();
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
                    System.Drawing.Image img = pdfDocument2.Render(page1, width, height, 96f, 96f, false);
                    Task task = this._taskFactory.StartNew((Action)(() => page.text = this.PageToText(img)), cancellationTokenSource.Token);
                    pages.Add(page);
                    taskList.Add(task);
                }
                Task.WaitAll(taskList.ToArray());
                return pages;
            }
        }

        public List<PiPage> XpsToPages(Stream xpsStream, string fileName)
        {
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            List<Task> taskList = new List<Task>();
            List<PiPage> pages = new List<PiPage>();
            IEnumerable<Stream> bitmap = this._xpsRender.RenderXpsToBitmap(xpsStream, 2.0);
            int num = 0;
            foreach (Stream stream in bitmap.ToList<Stream>())
            {
                System.Drawing.Image image = (System.Drawing.Image)new Bitmap(stream);
                PiPage page = new PiPage();
                page.pageNum = num + 1;
                page.fileName = fileName;
                pages.Add(page);
                Task task = this._taskFactory.StartNew((Action)(() => page.text = this.PageToText(image)), cancellationTokenSource.Token);
                taskList.Add(task);
                ++num;
            }
            Task.WaitAll(taskList.ToArray());
            return pages;
        }

        public string getResourcesPath()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            int length = location.LastIndexOf('\\');
            return location.Substring(0, length);
        }

        public string PageToText(System.Drawing.Image page)
        {
            string text = (string)null;
            if (page != null)
            {
                using (Engine engine = new Engine(this.getResourcesPath() + "\\tessdata\\", "rus(best)"))
                {
                    System.Drawing.Image image = (System.Drawing.Image)this.TiltDocument((Bitmap)page);
                    Random random = new Random();
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        image.Save((Stream)memoryStream, ImageFormat.Png);
                        using (TesseractOCR.Page page1 = engine.Process(TesseractOCR.Pix.Image.LoadFromMemory(memoryStream)))
                            text = Regex.Replace(Regex.Replace(page1.Text, " +", " "), "(?<!\\d)-+[ +]?\\n+[ +]?|(?<=[\\d\\wа-яА-я])\\s+(?=\\.)|(?<=\\.)\\s+(?=\\d)", "");
                    }
                }
            }
            return text;
        }

        private Bitmap TiltDocument(Bitmap inputPic)
        {
            Bitmap bitmap1 = new Bitmap(inputPic.Width * 2 / 3, inputPic.Height / 2);
            Bitmap bitmap2 = new Bitmap(inputPic.Width, inputPic.Height);
            int height = 640;
            Bitmap inputBmp = new Bitmap(height * 2 / 3, height);
            Rectangle rect = new Rectangle((inputPic.Width - bitmap1.Width) / 2, (inputPic.Height - bitmap1.Height) / 2, (inputPic.Width + bitmap1.Width) / 2, (inputPic.Height + bitmap1.Height) / 2);
            Bitmap bitmap3 = inputPic.Clone(rect, inputPic.PixelFormat);
            using (Graphics graphics = Graphics.FromImage((System.Drawing.Image)inputBmp))
                graphics.DrawImage((System.Drawing.Image)bitmap3, 0, 0, inputBmp.Width, inputBmp.Height);
            float angle = this.GetOptimumAngle(inputBmp, 10f, 1f) + this.GetOptimumAngle(inputBmp, 1.25f, 0.25f);
            using (Graphics graphics = Graphics.FromImage((System.Drawing.Image)bitmap2))
            {
                graphics.RotateTransform(angle);
                graphics.Clear(System.Drawing.Color.White);
                graphics.DrawImage((System.Drawing.Image)inputPic, (int)((double)inputPic.Height * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5), -(int)((double)inputPic.Width * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5));
            }
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
                }
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
                }
                System.Array.Sort<int>(numArray1);
                System.Array.Reverse((System.Array)numArray1);
                source.Add(num1, ((IEnumerable<int>)numArray1).Take<int>(10).Sum());
            }
            return source.Aggregate<KeyValuePair<float, int>>((Func<KeyValuePair<float, int>, KeyValuePair<float, int>, KeyValuePair<float, int>>)((l, r) => l.Value <= r.Value ? r : l)).Key;
        }
    }
}

