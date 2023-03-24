using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.CreateObjectSample;
using Ascon.Pilot.SDK.Menu;
using Ascon.Pilot.SDK.Toolbar;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;
using PdfiumViewer;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
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
    //Дополнение для распознавания сканированных документов (писем) и вложений
    //(в т.ч. чтение вложений txt, xls, doc) и составление БД для возможности поиска по тексту писем

    //страница документа:
    public class PiPage
    {
        //имя файла:
        public string FileName { get; set; }

        //номер листа:
        public int PageNum { get; set; }
        
        //текст (распознанный или прочитанный из doc, txt или xls):
        public string Text { get; set; }

        //скан или рендер документа:
        public Image Image { get; set; }

        //ссылка на файл, если он не прочитался:
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

    //письмо:
    public class PiLetter
    {
        //сам объект, из которого будет читаться письмо и вложения:
        public IDataObject DataObject { get; set; }

        //номер письма (ECM_inbound_letter_counter или ECM_letter_counter):
        public string LetterCounter { get; set; }
        
        //ID письма:
        public string DocId { get; set; }

        //исходящий номер:
        public string OutNo { get; set; }

        //дата документа:
        public string DocDate { get; set; }
        
        //тема письма:
        public string DocSubject { get; set; }
        
        //адресат для исходящего или отправитель для входящего
        public string DocCorrespondent { get; set; }

        //страницы письма и вложений:
        public List<PiPage> Pages { get; set; }

        //текст письма и вложений:
        public string Text { get; set; }

        //перечень не прочитавшихся файлов:
        public string CorruptedFiles { get; set; }

        //количество листов письма и вложений:
        public int PagesQtt { get; set; }
        
        //запись текста письма и вложений, присвоение количества листов:
        public void SetText()
        {
            foreach (PiPage piPage in Pages)
            {
                this.Text += "\n" + piPage.FileName + " " + piPage.PageNum.ToString() + ":\n\n" + piPage.Text + "\n\n=======================================================================================================================\n\n";
            };
            PagesQtt = Pages.Count;
        }

        //составление перечня не прочитавшихся файлов:
        public void SetCorruptedFiles()
        {
            foreach (PiPage piPage in Pages)
            {
                if (piPage.CorruptedFilePath != null)
                {
                    if (CorruptedFiles == null) CorruptedFiles = piPage.CorruptedFilePath + "\n";
                    else CorruptedFiles += piPage.CorruptedFilePath;
                }
            }
        }
    }



    //входящее письмо:
    public class PiLetterInbound : PiLetter
    {
        public PiLetterInbound() { }

        //запись атрибутов из DataObject'а в PiLetterInbound:
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
                            LetterCounter = attribute.Value.ToString();
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


    //исходящее письмо:
    public class PiLetterSent : PiLetter
    {
        public PiLetterSent() { }

        //запись атрибутов из DataObject'а в PiLetterSent:
        public void ReadAttributes()
        {
            foreach (KeyValuePair<string, object> attribute in (IEnumerable<KeyValuePair<string, object>>)DataObject.Attributes)
            {
                if (attribute.Value != null)
                {
                    Text += attribute.Key.ToString() + ":\n     " + attribute.Value.ToString() + "\n";
                    switch (attribute.Key.ToString())
                    {
                        case "ECM_letter_counter":
                            LetterCounter = attribute.Value.ToString();
                            break;
                        case "ECM_letter_number":
                            OutNo = attribute.Value.ToString();
                            break;
                        case "ECM_letter_date":
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
                    if (LetterCounter == null)
                        LetterCounter = OutNo;
                }
            };
        }
    }



    [Export(typeof(IToolbar<ObjectsViewContext>))]

    public class SearchButton : IToolbar<ObjectsViewContext>
    {
        public void Build(IToolbarBuilder builder, ObjectsViewContext context)
        {
            //кнопка поиска по тексту писем и вложений:
            builder.AddButtonItem("SearchByContent", 0)
                   .WithHeader("Поиск по тексту");
        }

        public void OnToolbarItemClick(string name, ObjectsViewContext context)
        {
            if (name == "SearchByContent")
            {
                //запуск GUI поиска по тексту писем и вложений:
                SearchByContext searchByContext = new SearchByContext();
                Task searchByContextTask = Task.Run(() => System.Windows.Forms.Application.Run(searchByContext));
            }
        }
    }


    [Export(typeof(IMenu<ObjectsViewContext>))]
    public class ModifyObjectsPlugin : IMenu<ObjectsViewContext>
    {
        #if DEBUG
        private string logs = "";
        #endif
        //чтение настроек соединения с SQL базой:
        private readonly string connectionParameters = File.ReadAllText($"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\\ASCON\\Pilot-ICE Enterprise\\PilotOCR\\connection_settings.txt");
        //определение типов документов, подлежащих распознаванию:
        private readonly List<string> acceptableDocTypes = File.ReadAllLines($"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\\ASCON\\Pilot-ICE Enterprise\\PilotOCR\\acceptable_doc_types.txt").ToList();

        private TaskFactory _taskFactoryRecognition;
        private readonly IXpsRender _xpsRender;
        private readonly IFileProvider _fileProvider;
        private readonly IObjectModifier _modifier;
        private readonly IObjectsRepository _objectsRepository;
        private readonly ObjectLoader _loader;
        private List<IDataObject> _dataObjects = new List<IDataObject>();
        //task scheduler для задач распознавания с ограничением параллельных задач - не более 8 шт.:
        private readonly LimitedConcurrencyLevelTaskScheduler lctsRecognition = new LimitedConcurrencyLevelTaskScheduler(8);
        //количество распознаваемых документов:
        private int docsCount = 0; 
        //суммарное кол-во листов распознаваемых документов и вложений:
        private int pagesCount = 0;
        //переменная для остановки процесса распознавания:
        private bool cancelled = false;
        //переменная для предотвращения запуска пользователем нового задания на распознавание:
        private bool isBusy = false;


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
            //пункт меню для распознавания:
            if (context.IsContext)
                return;
            this._dataObjects = context.SelectedObjects.ToList<Ascon.Pilot.SDK.IDataObject>();
            if (this._dataObjects.Count<Ascon.Pilot.SDK.IDataObject>() < 1)
                return;
            builder.AddItem("RecognizeItemName", 0).WithHeader("Распознать").WithIsEnabled(!isBusy);
        }
        public void OnMenuItemClick(string name, ObjectsViewContext context)
        {
            if (name == "RecognizeItemName")
            {
                docsCount = 0;
                //окно отображения хода распознавания:
                ProgressDialog progressDialog = new ProgressDialog(this);
                cancelled = false;
                _taskFactoryRecognition = new TaskFactory(lctsRecognition);
                pagesCount = 0;
                //определение списка документов, подлежащих распознаванию:
                _dataObjects = MakeRecognitionList(_dataObjects);
                //назначение общего количества распознаваемых документов для отображения в окне хода распознавания:
                progressDialog.SetMax(_dataObjects.Count);
                Task progressDialogTask = Task.Run(() => System.Windows.Forms.Application.Run(progressDialog));
                //непосредственно запуск распознавания:
                Task.Run(async () =>
                {
                    isBusy = true;
                    foreach (Ascon.Pilot.SDK.IDataObject dataObject in _dataObjects)
                    {
                        //проверка, не отменена ли работа (будет на всех этапах)
                        if (cancelled)
                            break;
                        //проверка, подлежит ли распознаванию документ
                            //TODO 1:
                            //перенести это в MakeRecognitionList
                        else if (!acceptableDocTypes.Contains(dataObject.Type.Name))
                            continue;
                        //запрос на выгрузку вложений (исходных файлов) в хранилище:
                        else
                        {
                            _objectsRepository.Mount(dataObject.Id);
                        }
                    };
                    //ожидание (3 секунды), чтобы дать загрузиться в хранилище хотя бы первому документу из списка.
                    //остальные документы успеют загрузиться по ходу поочерёдного распознавания. 
                    await Task.Delay(3000);
 
                    foreach (Ascon.Pilot.SDK.IDataObject dataObject in _dataObjects)
                    {
                        if (cancelled) break;
                        //если документ является входящим письмом или служебкой:
                        else if (acceptableDocTypes.GetRange(0,2).Contains(dataObject.Type.Name))
                        {
                            PiLetterInbound piLetter = new PiLetterInbound();
                            //загрузка актуального DataObject'а:
                            piLetter.DataObject = await _loader.Load(dataObject.Id);
                            //заполнение атрибутов документа:
                            piLetter.DocId = dataObject.Id.ToString();
                            piLetter.ReadAttributes();
                            //отображение номера и темы распознаваемого в данный момент письма: 
                            progressDialog.SetCurrentDocName(piLetter.LetterCounter + " - " + piLetter.OutNo + " - " + piLetter.DocDate + " - " + piLetter.DocSubject);
                            //распознавание документа и вложений:
                            piLetter.Pages = RecognizeWholeDoc(piLetter.DataObject);
                            if (cancelled) break;
                            //запись текста распознанных страниц в один большой текстовый атрибут, по которому в дальнейшем можно будет проводить поиск:
                            piLetter.SetText();
                            //перечень непрочитанных файлов:
                            piLetter.SetCorruptedFiles();
                            //обновление счётчика листов:
                            pagesCount += piLetter.PagesQtt;
                            //запись распознанного письма в SQL БД:
                            DocToDB("inbox", piLetter.LetterCounter, piLetter.OutNo, piLetter.DocId, piLetter.DocDate, piLetter.DocSubject, piLetter.DocCorrespondent, piLetter.Text, piLetter.CorruptedFiles);
                        }
                        else
                        {
                            //всё то же самое для исходящих:
                            PiLetterSent piLetter = new PiLetterSent();
                            piLetter.DataObject = await _loader.Load(dataObject.Id);
                            piLetter.DocId = dataObject.Id.ToString();
                            piLetter.ReadAttributes();
                            progressDialog.SetCurrentDocName(piLetter.LetterCounter + " - " + piLetter.OutNo + " - " + piLetter.DocDate + " - " + piLetter.DocSubject);
                            piLetter.Pages = RecognizeWholeDoc(piLetter.DataObject);
                            if (cancelled) break;
                            piLetter.SetText();
                            piLetter.SetCorruptedFiles();
                            pagesCount += piLetter.PagesQtt;
                            DocToDB("sent", piLetter.LetterCounter, piLetter.OutNo, piLetter.DocId, piLetter.DocDate, piLetter.DocSubject, piLetter.DocCorrespondent, piLetter.Text, piLetter.CorruptedFiles);
                        }
                        //клац по прогрессбару:
                        progressDialog.UpdateProgress();
                        //клац по счётчику распознанных документов:
                        docsCount++;
                    };
                    //по окончанию работы окно с результатами:
                    System.Windows.Forms.MessageBox.Show(pagesCount.ToString() + " страниц распознано\n в " + _dataObjects.Count.ToString() + " документах");
                    pagesCount = 0;
                    isBusy = false;
                    #if DEBUG
                    File.WriteAllText("D:\\TEMP\\logs.txt", logs);
                    #endif
                    //закрывание окна прогрессбара:
                    if (!cancelled) progressDialog.CloseRemotely();
                });
            }
        }

        //определение типа файла:
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


        //отмена:
        public void KillThemAll()
        {
            cancelled = true;
        }

        //составление списка документов, подлежащих распознаванию (по типу, mountable, по наличию в SQL БД):
        public List<IDataObject> MakeRecognitionList(List<IDataObject> dataObjects)
        {
            //переменные для названия SQL таблицы и названий атрибутов DataObject'а
            string tableName = "";
            string searchConditions = "";
            string letterCounter = "";
            string letterDate = "";
            string docNumType = "";
            string docDateType = "";
            MySqlConnection connection = new MySqlConnection(connectionParameters);
            var recognitionDict = new Dictionary<string, IDataObject>();
            foreach (Ascon.Pilot.SDK.IDataObject dataObject in dataObjects)
            {
                if (dataObject.Attributes.Count < 1 || !dataObject.Type.IsMountable)
                    continue;
                //определение таблицы, в которой может быть искомый документ и присвоение наиаенований атрибутов
                if (acceptableDocTypes.GetRange(0, 2).Contains(dataObject.Type.Name))
                {
                    docDateType = "ECM_inbound_letter_sending_date";
                    docNumType = "ECM_inbound_letter_counter";
                    tableName = "inbox";
                }
                else
                {
                    docDateType = "ECM_letter_date";
                    docNumType = "ECM_letter_counter";
                    tableName = "sent";
                }

                //присвоение атрибутов "номер" и "дата"
                foreach (KeyValuePair<string, object> attribute in (IEnumerable<KeyValuePair<string, object>>)dataObject.Attributes)
                {
                    if (attribute.Value != null)
                        if (attribute.Key == docNumType)
                            letterCounter = attribute.Value.ToString();
                        else if (attribute.Key == docDateType)
                            letterDate = attribute.Value.ToString();
                };

                //добавление в поисковый запрос (SELECT) искомых писем по номеру и дате
                searchConditions += $"(letter_counter = '{letterCounter}' and date = '{letterDate}') or ";
                recognitionDict.Add(letterCounter + letterDate, dataObject);
            }
            connection.Open();
            using (var command = new MySqlCommand($"select letter_counter, date from pilotsql.{tableName} where {searchConditions.Remove(searchConditions.Length - 4)}", connection))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        //исключение из выделенных пользователем DataObject'ов писем, имеющихся в SQL БД
                        recognitionDict.Remove(reader.GetString(0)+reader.GetString(1));
                    }
                }
            };
            connection.Close();
            return recognitionDict.Values.ToList();
        }
 

        //распознавание письма и всех вложений:
        public List<PiPage> RecognizeWholeDoc(IDataObject dataObject)
        {
            var ctsRecognition = new CancellationTokenSource();
            var token = ctsRecognition.Token;
            var recognitionTasks = new ConcurrentBag<Task>();
            List<PiPage> pieceOfDoc = new List<PiPage>();
            //рендер всех страниц документа в Image:
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
                            CorruptedFilePath = file.Name,
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
                            CorruptedFilePath = file.Name,
                            Text = "FILE IS CORRUPTED " + ex.Message
                        });
                    }
                }
            };
            //рендер всех вложений:
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
                            CorruptedFilePath = fileName,
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

            //параллельное распознавание отрендеренных страниц:
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
                            piPage.Text = PageToText(piPage.Image, piPage.FileName);
                            piPage.Image = null;
                        }
                    }, token);
                    recognitionTasks.Add(recognitionTask);
                }
            };
            Task.WaitAll(recognitionTasks.ToArray());
            ctsRecognition.Dispose();
            return pieceOfDoc;
        }


        //извлечение текста из XLS 
        private PiPage XlsToPage(string storagePath)
        {
            string fileName = Path.GetFileName(storagePath);
            PiPage piPage = new PiPage
            {
                FileName = fileName,
                PageNum = 1
            };
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
                piPage.CorruptedFilePath = fileName;
                piPage.Text = "FILE IS CORRUPTED " + ex.Message;
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }


        //извлечение текста из DOC
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
                piPage.CorruptedFilePath = fileName;
                piPage.Text = "FILE IS CORRUPTED " + ex.Message.ToString();
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }

        //извлечение текста из TXT
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
                piPage.CorruptedFilePath = fileName;
                piPage.Text = "FILE IS CORRUPTED " + ex.Message;
                //System.Windows.Forms.MessageBox.Show("File corrupted in\n" + piPage.docID.ToString() + " " + piPage.fileName);
            };
            return piPage;
        }

        //преобразование PDF документа в объект PiPage c рендерами:
        private List<PiPage> PdfToPages(Stream pdfStream, string fileName)
        {
            List<PiPage> pages = new List<PiPage>();
            using (PdfDocument pdfDocument1 = PdfDocument.Load(pdfStream))
            {
                for (int index = 0; index < pdfDocument1.PageCount; ++index)
                {
                    PiPage page = new PiPage
                    {
                        PageNum = index + 1,
                        FileName = fileName
                    };
                    PdfDocument pdfDocument2 = pdfDocument1;
                    SizeF pageSiz = pdfDocument1.PageSizes[index];
                    int width = pageSiz.ToSize().Width * 2;
                    pageSiz = pdfDocument1.PageSizes[index];
                    int height = pageSiz.ToSize().Height * 2;
                    page.Image = pdfDocument2.Render(index, width, height, 96f, 96f, false);
                    pages.Add(page);
                };
                return pages;
            };
        }

        //преобразование XPS в объект PiPage с рендерами:
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

        //определение адреса, где хранятся файлы сборки:
        public string GetResourcesPath()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            int length = location.LastIndexOf('\\');
            return location.Substring(0, length);
        }

        //распознавание Image страницы и запись распознанного текста в PiPage.Text:
        public string PageToText(Image pageImg, string filename)
        {
            string text = (string)null;
            if (pageImg != null)
            {
                using (Engine engine = new Engine(this.GetResourcesPath() + "\\tessdata\\", "rus(best)"))
                {
                    Image image = (Image)this.TiltDocument((Bitmap)pageImg, filename);
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        image.Save((Stream)memoryStream, ImageFormat.Png);
                        using (TesseractOCR.Page page = engine.Process(TesseractOCR.Pix.Image.LoadFromMemory(memoryStream)))
                            text = page.Text;
                    };
                };
            };
            return text;
        }

        //запись распознанного документа в SQL БД:
        private void DocToDB(string tableName, string letterCounter, string outNo, string docId, string docDate, string docSubject, string docCorrespondent, string docText, string corruptedFiles)
        {
            if (docText != null)
                docText = MySqlHelper.EscapeString(docText);
            if (outNo != null)
                outNo = MySqlHelper.EscapeString(outNo);
            if (docId != null)
                docId = MySqlHelper.EscapeString(docId);
            if (docDate != null)
                docDate = MySqlHelper.EscapeString(docDate);
            if (docSubject != null)
                docSubject = MySqlHelper.EscapeString(docSubject);
            if (docCorrespondent != null)
                docCorrespondent = MySqlHelper.EscapeString(docCorrespondent);
            if (corruptedFiles != null)
                corruptedFiles = MySqlHelper.EscapeString(corruptedFiles);
            MySqlConnection connection = new MySqlConnection(connectionParameters);
            connection.Open();
            string commandText = $"INSERT INTO pilotsql.{tableName}(letter_counter, out_no, doc_id, date, subject, correspondent, text, unrecognized) VALUES (@letterCounter, @outNo, @docId, @docDate, @docSubject, @docCorrespondent, @docText, @corruptedFiles)";
            MySqlCommand command = new MySqlCommand(commandText, connection);
            command.Parameters.AddWithValue("@letterCounter", letterCounter);
            command.Parameters.AddWithValue("@outNo", outNo);
            command.Parameters.AddWithValue("@docId", docId);
            command.Parameters.AddWithValue("@docDate", docDate);
            command.Parameters.AddWithValue("@docSubject", docSubject);
            command.Parameters.AddWithValue("@docCorrespondent", docCorrespondent);
            command.Parameters.AddWithValue("@docText", docText);
            command.Parameters.AddWithValue("@corruptedFiles", corruptedFiles);
            try
            {
                command.ExecuteNonQuery();
            }
            catch (MySqlException ex){ System.Windows.Forms.MessageBox.Show(ex.Message); };
            connection.Close();
        }


        //выравнивание криво отсканированного документа (поворот документа на -10...10 градусов):
        private Bitmap TiltDocument(Bitmap inputPic, string filename)
        {
            //создание Bitmapa с размерами меньше исходного листа - для обрезки полей и шапки письма:
            Bitmap croppedPic = new Bitmap(inputPic.Width * 2/3, inputPic.Height / 2);
            //создание Bitmapа для идеально горизонтального документа:
            Bitmap horisontalPic = new Bitmap(inputPic.Width, inputPic.Height);
            //высота одностолбцового битмапа, подлежащего анализу для определения угла поворота:
            int colImageHeight = 480;
            if (croppedPic.Height < colImageHeight) colImageHeight = croppedPic.Height; 
            float croppedAspectRatio = (float)croppedPic.Width / (float)croppedPic.Height;
            //bitmap для хранения уменьшенного (ужатого) рендера/скана документа:
            Bitmap shrinkPic = new Bitmap((int)((float)colImageHeight * croppedAspectRatio), colImageHeight);
            RectangleF rect = new RectangleF((inputPic.Width - croppedPic.Width) / 2,
                                           (inputPic.Height - croppedPic.Height) / 2,
                                           croppedPic.Width,
                                           croppedPic.Height);
            croppedPic = inputPic.Clone(rect, inputPic.PixelFormat);
            using (Graphics graphics = Graphics.FromImage((Image)shrinkPic))
                graphics.DrawImage((Image)croppedPic, 0, 0, shrinkPic.Width, shrinkPic.Height);
#if DEBUG
            Stopwatch stopwatch = Stopwatch.StartNew();
            stopwatch.Start();
#endif
            //определение угла поворота документа:
            float optimumAngle = this.GetOptimumAngle(shrinkPic, 0f, 5f, 1f);
            optimumAngle = this.GetOptimumAngle(shrinkPic, (float)optimumAngle, 0.6f, 0.2f);
#if DEBUG
            stopwatch.Stop();
            Debug.Write($"{stopwatch.ElapsedMilliseconds}, {optimumAngle}\n");
#endif
            //поворот документа:
            using (Graphics graphics = Graphics.FromImage((Image)horisontalPic))
            {
                graphics.RotateTransform(optimumAngle);
                graphics.Clear(System.Drawing.Color.White);
                graphics.DrawImage(inputPic, (int)(inputPic.Height * Math.Sin(optimumAngle * 3.1415 / 180.0) * 0.5), -(int)(inputPic.Width * Math.Sin(optimumAngle * 3.1415 / 180.0) * 0.5));
            };
#if DEBUG
            if (Math.Abs(optimumAngle) > 1.5)
                logs += $"{optimumAngle} - {filename}\n";
#endif
            return horisontalPic;
        }


        //определение угла поворота документа путём нахождения угла, при котором преобразованный в 
        //вертикальную линию документ будет иметь максимальную разницу между соседними пикселями:
        private float GetOptimumAngle(Bitmap inputBmp, float initAngle, float amplitude, float stepSize)
        {
            int width = inputBmp.Width;
            int height = inputBmp.Height;
            //переменная для хранения документа, преобразованного в bitmap шириной в 1 пиксель:
            Bitmap columnPic = new Bitmap(1, height);
            //переменные для хранения значений суммарной яркости каждого пикселя:
            int brightnessPrevPix = 255 * 3;
            int brightnessCurrPix = 255 * 3;
            //int maxGradient = 0;
            //переменная для храниния вектора разниц яркости соседних пикселей (градиентов) отдностолбцовой картинки (одномерной проекции): 
            int[] brightnessGradients = new int[height];
            //int[] brightness = new int[height];
            //словарь соответствия характеристики контраста углу поворота:
            Dictionary<float, int> angles = new Dictionary<float, int>();
            //переменная для хранения изображения письма, повёрнутого на пробный угол:
            Bitmap rotatedPic = new Bitmap(width, height);
            //переменная для хранения суммы "top 5%" значений градиента яркости:
            int topGradsSum = 0;
            //цикличный поворот документа с заданным шагом в заданном диапазоне и проверка контрастности его одномерной проекции:
            for (float angle = (-amplitude + initAngle); angle < (amplitude + initAngle); angle += stepSize)
            {
                //поворот картинки:
                using (Graphics graphics = Graphics.FromImage(rotatedPic))
                {
                    graphics.RotateTransform(angle);
                    graphics.Clear(System.Drawing.Color.White);
                    graphics.DrawImage(inputBmp, (int)((double)height * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5), -(int)(width * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5));
                };
                //создание одномерной проекции:
                using (Graphics graphics = Graphics.FromImage(columnPic))
                    graphics.DrawImage(rotatedPic, 0, 0, 1, height);
                //сохранение начального значения яркости предыдущего пикселя:
                brightnessPrevPix = columnPic.GetPixel(0, 0).R + columnPic.GetPixel(0, 0).G + columnPic.GetPixel(0, 0).B;
                //составление вектора градиентов:
                for (int y = 0; y < height; ++y)
                {
                    brightnessCurrPix = columnPic.GetPixel(0, y).R + columnPic.GetPixel(0, y).G + columnPic.GetPixel(0, y).B;
                    //brightness[y] = (255 * 3) - brightnessCurrPix;
                    brightnessGradients[y] = Math.Abs(brightnessCurrPix - brightnessPrevPix);
                    brightnessPrevPix = brightnessCurrPix;
                };
//Рисование диаграмм яркости и градиента на изображении письма для визуализации и лучшего понимания при отладке алгоритма:
                //using (Graphics graphics = Graphics.FromImage((Image)rotatedPic))
                //{
                //    //graphics.Clear(System.Drawing.Color.White);
                //    Pen penRed = new Pen(System.Drawing.Color.Red);
                //    Pen penBlue = new Pen(System.Drawing.Color.Navy);
                //    for (int y = 0; y < brightnessGradients.Length; ++y)
                //    {
                //        graphics.DrawLine(penRed, 0, y, brightnessGradients[y]/3, y);
                //        graphics.DrawLine(penBlue, width, y, width - brightness[y]/3, y);
                //    }
                //выявление top 5% значений градиентов:
                Array.Sort<int>(brightnessGradients);
                Array.Reverse((System.Array)brightnessGradients);
                //maxGradient = brightnessGradients[0];
                //определение критерия контрастности (суммы top 5%) градиентов проекции:
                topGradsSum = (brightnessGradients).Take((int)(0.04 * height)).Sum();
                //    for (int y = 0; y < brightnessGradients.Length; ++y)
                //    {
                //        graphics.DrawString(topGradsSum.ToString(), new System.Drawing.Font("Times New Roman", 18.0f), new SolidBrush(System.Drawing.Color.DarkGreen), new Point(50, 30));
                //        graphics.DrawString(maxGradient.ToString(), new System.Drawing.Font("Times New Roman", 18.0f), new SolidBrush(System.Drawing.Color.DarkGreen), new Point(50, 10));
                //    }
                //};
                //пополнение словарая соответствия характеристики контраста углу поворота:
                angles.Add(angle, topGradsSum);
            }
            //выбор оптимального угла по наибольшей характеристике контраста:
            return angles.Aggregate((l, r) => l.Value > r.Value ? l : r).Key;
        }

        //private float GetOptimumAngle(Bitmap inputBmp, float initAngle, float amplitude, float stepSize)
        //{
        //    int width = inputBmp.Width;
        //    int height = inputBmp.Height;
        //    //переменная для хранения документа, преобразованного в bitmap шириной в 1 пиксель:
        //    Bitmap columnPic = new Bitmap(1, height);
        //    //
        //    //int brightnessPrevPix = 255 * 3;
        //    //int brightnessCurrPix = 255 * 3;
        //    int maxBrightness = 0;
        //    var topBrightnesses = new List<int>();
        //    //int[] brightnessGradients = new int[height];
        //    int[] brightness = new int[height];
        //    Dictionary<float, int> angles = new Dictionary<float, int>();
        //    Bitmap rotatedPic = new Bitmap(width, height);
        //    int midTopBrights = 0;
        //    for (float angle = (-amplitude + initAngle); angle < (amplitude + initAngle); angle += stepSize)
        //    {
        //        using (Graphics graphics = Graphics.FromImage(rotatedPic))
        //        {
        //            graphics.RotateTransform(angle);
        //            graphics.Clear(System.Drawing.Color.White);
        //            graphics.DrawImage(inputBmp, (int)((double)height * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5), -(int)(width * Math.Sin((double)angle * 3.1415 / 180.0) * 0.5));
        //        };
        //        //rotatedPic.Save($"D:\\TEMP\\{angle}.jpg");
        //        using (Graphics graphics = Graphics.FromImage(columnPic))
        //            graphics.DrawImage(rotatedPic, 0, 0, 1, height);
        //        //brightnessPrevPix = columnPic.GetPixel(0, 0).R + columnPic.GetPixel(0, 0).G + columnPic.GetPixel(0, 0).B;
        //        for (int y = 0; y < height; ++y)
        //            brightness[y] = (255 * 3) - (columnPic.GetPixel(0, y).R + columnPic.GetPixel(0, y).G + columnPic.GetPixel(0, y).B);


        //        //Array.Sort<int>(brightnessGradients);
        //        maxBrightness = System.Linq.Enumerable.Max(brightness);
        //        //Bitmap gradGraph = new Bitmap(height, maxBrightness);

        //        //Array.Reverse((System.Array)brightnessGradients);
        //        foreach (int item in brightness)
        //            if (item > 0.7 * maxBrightness) topBrightnesses.Add(item);
        //        midTopBrights = topBrightnesses.Sum() / topBrightnesses.Count;
        //        using (Graphics graphics = Graphics.FromImage((Image)rotatedPic))
        //        {
        //            //graphics.Clear(System.Drawing.Color.White);
        //            Pen penRed = new Pen(System.Drawing.Color.Red);
        //            Pen penBlue = new Pen(System.Drawing.Color.Navy);
        //            for (int y = 0; y < brightness.Length; ++y)
        //            {
        //                //graphics.DrawLine(penRed, 0, y, brightnessGradients[y]/3, y);
        //                graphics.DrawLine(penBlue, width, y, width - brightness[y] / 3, y);
        //                graphics.DrawString(midTopBrights.ToString(), new System.Drawing.Font("Times New Roman", 24.0f), new SolidBrush(System.Drawing.Color.Red), new Point(50, 50));
        //                graphics.DrawString(maxBrightness.ToString(), new System.Drawing.Font("Times New Roman", 24.0f), new SolidBrush(System.Drawing.Color.Red), new Point(50, 10));
        //            }
        //        };
        //        angles.Add(angle, midTopBrights);
        //        topBrightnesses.Clear();
        //        rotatedPic.Save($"D:\\TEMP\\anomaly\\{midTopBrights}-{angle}-graph.jpg");
        //    }
        //    float optimumAngle = angles.Aggregate((l, r) => l.Value > r.Value ? l : r).Key;
        //    //if (Math.Abs(optimumAngle) > 1)
        //    return optimumAngle;
        //}



    }
}

