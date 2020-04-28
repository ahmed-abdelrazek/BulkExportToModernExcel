using ExcelBulkExport.Core;
using Ookii.Dialogs.Wpf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelBulkExport.ViewModels
{
    public class MainWindowViewModel : ViewModel
    {
        private string filesDirectory;
        private string extensions;
        private string newFilesDirectory;
        private int progressBarMaxValue;
        private int progressBarCurrentValue;
        private bool isError;
        private bool isWorking;

        public string FilesDirectory
        {
            get => filesDirectory;
            set
            {
                NotifyPropertyChange(ref filesDirectory, value);
                IsError = false;
                if (!Directory.Exists(filesDirectory))
                {
                    MessageBox.Show("The Folder you chose doesn't exist", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    IsError = true;
                }
            }
        }

        public string Extensions { get => extensions; set => NotifyPropertyChange(ref extensions, value); }

        public string NewFilesDirectory
        {
            get => newFilesDirectory;
            set
            {
                NotifyPropertyChange(ref newFilesDirectory, value);
                IsError = false;
                if (!Directory.Exists(newFilesDirectory))
                {
                    MessageBox.Show("The Folder you chose doesn't exist", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    IsError = true;
                }
                if (!Directory.Exists(newFilesDirectory))
                {
                    MessageBox.Show("Choose a different folder to save to", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        public int ProgressBarMaxValue { get => progressBarMaxValue; set => NotifyPropertyChange(ref progressBarMaxValue, value); }

        public int ProgressBarCurrentValue { get => progressBarCurrentValue; set => NotifyPropertyChange(ref progressBarCurrentValue, value); }

        public bool IsError { get => isError; set => NotifyPropertyChange(ref isError, value); }

        public bool IsWorking { get => isWorking; set => NotifyPropertyChange(ref isWorking, value); }

        public IAsyncCommand ChooseFilesDirectory { get; private set; }
        public IAsyncCommand ChooseNewFilesDirectory { get; private set; }
        public IAsyncCommand Start { get; private set; }
        public IAsyncCommand Stop { get; private set; }
        public IAsyncCommand Exit { get; private set; }

        public MainWindowViewModel()
        {
            ProgressBarMaxValue = 100;
            Extensions = ".xml, .xls";

            ChooseFilesDirectory = new AsyncCommand(DoChooseFilesDirectory);
            ChooseNewFilesDirectory = new AsyncCommand(DoChooseNewFilesDirectory);
            Start = new AsyncCommand(DoStart, CanStart);
            Stop = new AsyncCommand(DoStop, CanStop);
            Exit = new AsyncCommand(DoExit, CanExit);
        }

        private async Task DoChooseFilesDirectory()
        {
            await Task.Delay(1);
            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog
            {
                Description = "Please select a folder.",
                UseDescriptionForTitle = true // This applies to the Vista style dialog only, not the old dialog.
            };
            if (dialog.ShowDialog() == true)
            {
                FilesDirectory = dialog.SelectedPath;
            }
        }

        private async Task DoChooseNewFilesDirectory()
        {
            await Task.Delay(1);
            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog
            {
                Description = "Please select a folder.",
                UseDescriptionForTitle = true // This applies to the Vista style dialog only, not the old dialog.
            };
            if (dialog.ShowDialog() == true)
            {
                NewFilesDirectory = dialog.SelectedPath;
            }
        }

        private bool CanStart()
        {
            return !IsError;
        }

        private async Task DoStart()
        {
            IsWorking = true;
            try
            {
                var di = new DirectoryInfo(FilesDirectory);
                var newExts = new List<string>();
                foreach (var ext in Extensions.Split(','))
                {
                    newExts.Add(ext.Trim());
                }

                var AllFiles = di.GetFiles().Where(file => newExts.Any(file.FullName.ToLower().EndsWith)).ToList();

                if (AllFiles.Count > 0)
                {
                    progressBarMaxValue = AllFiles.Count;
                    progressBarCurrentValue = 0;

                    await Task.Run(() =>
                    {
                        foreach (var file in AllFiles)
                        {
                            if (!IsWorking)
                            {
                                break;
                            }
                            var excelApp = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel.Workbook currentWorkbook = excelApp.Workbooks.Open(file.FullName);
                            currentWorkbook.SaveAs($"{NewFilesDirectory}\\{Path.GetFileNameWithoutExtension(file.FullName)}.xlsx",
                            Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
                            Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                            Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true,
                            Missing.Value, Missing.Value, Missing.Value);
                            excelApp.Quit();
                            progressBarCurrentValue++;
                        }
                    });
                }
                else
                {
                    MessageBox.Show($" There is no {Extensions} Files", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            await DoStop();
        }

        private bool CanStop()
        {
            return IsWorking;
        }

        private async Task DoStop()
        {
            await Task.Delay(1);
            IsWorking = false;
        }

        private bool CanExit()
        {
            return true;
        }

        private async Task DoExit()
        {
            if (CanStop())
            {
                await DoStop();
            }

            System.Environment.Exit(0);
        }
    }
}
