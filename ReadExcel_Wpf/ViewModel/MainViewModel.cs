using ExcelDataReader;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using LicenseContext = OfficeOpenXml.LicenseContext;
using WPF_Exercise.Model;

namespace WPF_ReadExcel.ViewModel
{
    public class MainViewModel : Base
    {
        public MainModel Model { get; set; }
        public ObservableCollection<SheetModel> Sheets
        {
            get
            {
                return Model.Sheets;
            }
            set
            {
                Model.Sheets = value;
                OnPropertyChanged(nameof(Sheets));
            }
        }
        public SheetModel SelectSheet
        {
            get
            {
                return Model.SelectSheet;
            }
            set
            {
                Model.SelectSheet = value;
                OnPropertyChanged(nameof(SelectSheet));
            }
        }

        public PathModel Name
        {
            get
            {
                return Model.Name;
            }
            set
            {
                Model.Name = value;
                OnPropertyChanged(nameof(Name));
            }
        }
        private ObservableCollection<NamePathViewModel> _Namepath;
        public ObservableCollection<NamePathViewModel> NamePath 
        { 
            get { return _Namepath; }
            set {
                _Namepath = value; 
                OnPropertyChanged(nameof(NamePath)); 
            }
        }
        //public string Namepath2
        //{            
        //    get;
        //    set; 
        //}

        public RelayCommand CancelCommand { get; set; }
        public RelayCommand ImportCommand { get; set; }
        public RelayCommand ExportCommand { get; set; }

        private ObservableCollection<SheetViewModel> _Sheets;
        public ObservableCollection<SheetViewModel> Item 
        {
            get { return _Sheets; }
            set
            {
                _Sheets = value;
                OnPropertyChanged(nameof(Item));
            }
        }
        private SheetViewModel _SelectSheet; 
        public SheetViewModel SelectedItem
        {
            get { return _SelectSheet; }
            set 
            {
                _SelectSheet = value;
                OnPropertyChanged(nameof(SelectedItem));
            }
        }

        private NamePathViewModel _NamePathTest;
        public NamePathViewModel NamePathTest
        {
            get { return _NamePathTest; }
            set
            {
                _NamePathTest = value;
                OnPropertyChanged(nameof(NamePathTest));
            }
        }

        public MainViewModel()
        {
            CancelCommand = new RelayCommand(Cancel_event);
            ImportCommand = new RelayCommand(ImportExcel);
            ExportCommand = new RelayCommand(CanExportExcel);
            Item = new ObservableCollection<SheetViewModel>();
            NamePath = new ObservableCollection<NamePathViewModel>();
      
            NamePath.Add(new NamePathViewModel(new PathModel("...")));

        }
        public void Cancel_event(object obj)
        {
            System.Windows.Application.Current.Shutdown();
        }
        public void ImportExcel(object obj)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "All Files (*.*)|*.*";
            
            if (openFile.ShowDialog() == true)
            {
                try
                {
                    
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var fs = File.Open(openFile.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true 
                                }
                            });
                            int i = 0;

                            NamePath = new ObservableCollection<NamePathViewModel>();
                            PathModel path = new PathModel(openFile.FileName.ToString());
                            NamePathViewModel namepath = new NamePathViewModel(path);              
                            if (namepath == null) throw new Exception("Path null");
                            NamePath.Add(namepath);
                            NamePathTest = NamePath.First();

                            Item = new ObservableCollection<SheetViewModel>();                         
                            foreach (DataTable dt in result.Tables)
                            {                              
                                SheetModel sheet = new SheetModel(dt.TableName.ToString(),dt);
                                SheetViewModel sheetView = new SheetViewModel(sheet);
                                Item.Add(sheetView);                              
                            }                            
                            SelectedItem = Item.First();
                            reader.Close();                            
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi!\n{ex.Message}");
                }
            }
        }
        public void CanExportExcel(object obj)
        {
            SaveFileDialog Excelfile = new SaveFileDialog();
            Excelfile.Filter = "Excel (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            if (Excelfile.ShowDialog() == true)
            {             
                    ExportExcel(Excelfile.FileName);                
            }
        }
        public void ExportExcel(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelfile = new ExcelPackage())
            {
                foreach(var item in Item)
                {
                    ExcelWorksheet ws = excelfile.Workbook.Worksheets.Add(item.Namesheet);
                    ws.Cells["A1"].LoadFromDataTable(item.DataSheet, true, TableStyles.Custom);
                }                
                excelfile.SaveAs(new FileInfo(path));
            }
        }

    }
}
