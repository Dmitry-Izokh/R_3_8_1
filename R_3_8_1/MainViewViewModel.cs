using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Win32;
using System.Windows.Forms;

namespace R_3_8_1
{
    public class MainViewViewModel
    {
        // Свойства
        private ExternalCommandData _commandData;
        private Document _doc;
        public DelegateCommand ExportIFC { get;}
        public DelegateCommand ExportNWC { get; }
        public DelegateCommand ExportImmage { get; }
        // Конструктор
        public MainViewViewModel(ExternalCommandData commandData)        
        {
            _commandData = commandData;
            _doc = _commandData.Application.ActiveUIDocument.Document;
            ExportIFC = new DelegateCommand(OnExportIFC);
            ExportNWC = new DelegateCommand(OnExportNWC);
            ExportImmage = new DelegateCommand(OnExportImmage);
        }

        // Методы
        public void OnExportIFC()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            using (Transaction ts = new Transaction(doc, "Создание листов"))
            {
                ts.Start();
                var IfcOption = new IFCExportOptions();
                var saveDialog = new System.Windows.Forms.SaveFileDialog
                {
                    OverwritePrompt = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    Filter = "All files (*.*)|*.*",
                    FileName = "Export.ifc",
                    DefaultExt = ".ifc"
                };
                string selectedFilePath = string.Empty;
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = saveDialog.FileName;
                }
                if (string.IsNullOrEmpty(selectedFilePath))
                    return;
                doc.Export(selectedFilePath, "Экспорт.ifc", IfcOption);
                ts.Commit();
            }
            RaiseCloseRequest();
        }  

        public void OnExportNWC()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (var ts = new Transaction(doc, "Создание листов"))
            {

                ts.Start();
                var nwcOption = new NavisworksExportOptions();                               
                var saveDialog = new System.Windows.Forms.SaveFileDialog
                {
                    OverwritePrompt = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    Filter = "All files (*.*)|*.*",
                    FileName = "Export.nwc",
                    DefaultExt = ".nwc"
                };
                string selectedFilePath = string.Empty;
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = saveDialog.FileName;
                }
                if (string.IsNullOrEmpty(selectedFilePath))
                    return;
                doc.Export(selectedFilePath, "Экспорт.nwc", nwcOption);
                ts.Commit();

            }
            RaiseCloseRequest();
        }

        public void OnExportImmage()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (var ts = new Transaction(doc, "Создание листов"))
            {

                ts.Start();
                // Выбираем виды - планы этажей и находим первый соответствующий типу вида "План этажа" и имеющий наименование "Level 1".
                // Таким образом мы могли бы выбрать любые другие виды.
                // Как мне кажется это приложение либо должно иметь ListBox для выбора нужных видов, либо экспортировать активный вид.
                ViewPlan viewPlan = new FilteredElementCollector(doc)
                                    .OfClass(typeof(ViewPlan))
                                    .Cast<ViewPlan>()
                                    .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan && v.Name.Equals("Level 1"));

                var saveDialog = new System.Windows.Forms.SaveFileDialog
                {
                    OverwritePrompt = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    Filter = "All files (*.*)|*.*",
                    FileName = "Export.png",
                    DefaultExt = ".PNG"
                };
                string selectedFilePath = string.Empty;
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = saveDialog.FileName;
                }
                if (string.IsNullOrEmpty(selectedFilePath))
                    return;

                var imageOption = new ImageExportOptions
                // Ниже описание настроек экспорта
                {
                    ZoomType = ZoomFitType.FitToPage, // размер изображения "вписать в лист"
                    PixelSize = 2024, // Максимальное разрешение по одной из сторон 
                    FilePath = selectedFilePath, // Путь сохранения
                    FitDirection = FitDirectionType.Horizontal, // Горизонтальная ориентация листа
                    HLRandWFViewsFileType = ImageFileType.PNG, // Расширение экспортируемого файла (по умолчанию .jpeg)
                    ShadowViewsFileType = ImageFileType.PNG, // Тоже тип расширения, но тут речь идет про теневые виды, что это?
                    ImageResolution = ImageResolution.DPI_600, // Разрешение изображения в точках на дюйм
                    ExportRange = ExportRange.CurrentView // тут речь о том какие виды пойдут на экспорт, в данном случае экспортируется только текущий вид,
                                                          // тот который мы отобрали с помощью FilteredElementCollectorю
                };
                // тут странный синтаксис со знаком ";"
                
                doc.ExportImage(imageOption);                
                ts.Commit();

            }
            RaiseCloseRequest();
        }

        //событие создается для скрытие окна на время выбора.
        public event EventHandler HideRequest;
        private void RaiseHideRequest()
        {
            HideRequest?.Invoke(this, EventArgs.Empty);
        }
        //событие создается для повторного открытия окна после отработки программы.
        public event EventHandler ShowRequest;
        private void RaiseShowRequest()
        {
            ShowRequest?.Invoke(this, EventArgs.Empty);
        }
        //событие создается для закрытия программы.
        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }
    }
}

