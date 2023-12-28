using ConductivityReportAlgo;
using Microsoft.UI;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Media;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Text;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ConductivityGenerator
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        private SolidColorBrush Color1 = new SolidColorBrush(Colors.Green);
        private SolidColorBrush Color2 = new SolidColorBrush(Colors.Yellow);
        private SolidColorBrush Color3 = new SolidColorBrush(Colors.Orange);
        private SolidColorBrush Color4 = new SolidColorBrush(Colors.Red);
        private SolidColorBrush Color5 = new SolidColorBrush(Colors.DarkRed);

        public MainWindow()
        {
            this.AppWindow.MoveAndResize(new Windows.Graphics.RectInt32(400, 200, 650, 600));
            this.InitializeComponent();
            setAppIcon();
        }

        private async void setAppIcon()
        {
            StorageFolder installedLocation = Package.Current.InstalledLocation;
            StorageFolder assetsFolder = await installedLocation.GetFolderAsync("Assets");
            AppWindow.SetIcon(Path.Combine(assetsFolder.Path, "appIcon.ico"));
        }

        public Windows.UI.Color TextColorPicker1
        {
            get { return Color1.Color; }
            set { Color1.Color = value; }
        }
        public SolidColorBrush TextColorShow1
        {
            get { return Color1; }
        }      
        public Windows.UI.Color TextColorPicker2
        {
            get { return Color2.Color; }
            set { Color2.Color = value; }
        }
        public SolidColorBrush TextColorShow2
        {
            get { return Color2; }
        }
        public Windows.UI.Color TextColorPicker3
        {
            get { return Color3.Color; }
            set { Color3.Color = value; }
        }
        public SolidColorBrush TextColorShow3
        {
            get { return Color3; }
        }
        public Windows.UI.Color TextColorPicker4
        {
            get { return Color4.Color; }
            set { Color4.Color = value; }
        }
        public SolidColorBrush TextColorShow4
        {
            get { return Color4; }
        }
        public Windows.UI.Color TextColorPicker5
        {
            get { return Color5.Color; }
            set { Color5.Color = value; }
        }
        public SolidColorBrush TextColorShow5
        {
            get { return Color5; }
        }

        private void ChangeColor1(ColorPicker sender, ColorChangedEventArgs args)
        {
            this.Color1.Color = args.NewColor;
        }

        private void ChangeColor2(ColorPicker sender, ColorChangedEventArgs args)
        {
            this.Color2.Color = args.NewColor;
        }
        private void ChangeColor3(ColorPicker sender, ColorChangedEventArgs args)
        {
            this.Color3.Color = args.NewColor;
        }
        private void ChangeColor4(ColorPicker sender, ColorChangedEventArgs args)
        {
            this.Color4.Color = args.NewColor;
        }
        private void ChangeColor5(ColorPicker sender, ColorChangedEventArgs args)
        {
            this.Color5.Color = args.NewColor;
        }
        private void addConsoleLine(string text)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run
            {
                Text = "* -->" + text
            };
            paragraph.Inlines.Add(run);
            consoleRTB.Blocks.Add(paragraph);
        }
        private async void semiAutoClick(object sender, RoutedEventArgs e)
        {
            consoleRTB.Blocks.Clear();
            addConsoleLine("Generating Report Sample");
            await Task.Delay(1000);
            pBar.ShowError = false;
            pBar.Value = 0;
            var dir = new Windows.Storage.Pickers.FolderPicker();
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(dir, hWnd);
            dir.SuggestedStartLocation = PickerLocationId.Desktop;
            pBar.Value = 10;
            await Task.Delay(10);
            dir.FileTypeFilter.Add("*");
            addConsoleLine("Getting User defined Folder");
            StorageFolder folder = await dir.PickSingleFolderAsync();
            if (folder != null)
            {
                addConsoleLine("User Chosse Folder: " + folder.Path);
                await Task.Delay(500);
                addConsoleLine("Generating sample file");
                string fileName = Algo.semiPackage(folder.Path);
                var path = System.IO.Path.Combine(folder.Path, fileName);
                addConsoleLine("File Created!");
                var filer = new Windows.Storage.Pickers.FileOpenPicker();
                WinRT.Interop.InitializeWithWindow.Initialize(filer, hWnd);
                filer.SuggestedStartLocation = PickerLocationId.Desktop;
                filer.FileTypeFilter.Add(".xlsx");
                filer.FileTypeFilter.Add(".csv");
                addConsoleLine("Optinal, add Top file");

                ContentDialog popupTopDialog = new ContentDialog
                {
                    Title = "Add Top File?",
                    Content = "would you like to add a top file csv?",
                    CloseButtonText = "Skip",
                    PrimaryButtonText = "Select"
                };
                //set the XamlRoot property
                popupTopDialog.XamlRoot = MainTab.XamlRoot;
                ContentDialogResult resultTop = await popupTopDialog.ShowAsync();
                if(resultTop == ContentDialogResult.Primary)
                {
                    var topFile = await filer.PickSingleFileAsync();
                    if (topFile != null)
                    {
                        Algo.addTopFile(folder.Path, topFile.Path);
                        addConsoleLine($"{topFile.Path} Found adding to file");

                    }
                    else
                    {
                        addConsoleLine("User skiped adding the top file");
                    }
                }
                addConsoleLine("Optinal, add Bottom file");
                ContentDialog popupBottomDialog = new ContentDialog
                {
                    Title = "Add Bottom File?",
                    Content = "would you like to add a bottom file csv?",
                    CloseButtonText = "Skip",
                    PrimaryButtonText = "Select"
                };
                //set the XamlRoot property
                popupTopDialog.XamlRoot = MainTab.XamlRoot;
                ContentDialogResult resultBot = await popupTopDialog.ShowAsync();
                if( resultBot == ContentDialogResult.Primary)
                {
                    var botFile = await filer.PickSingleFileAsync();
                    if (botFile != null)
                    {
                        Algo.addBottomFile(folder.Path, botFile.Path);
                        addConsoleLine($"{botFile.Path} Found adding to file");
                    }
                    else
                    {
                        addConsoleLine("User skipped addding the bottom file");
                    }
                }
                ProcessStartInfo info = new ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                };
                pBar.Value = 50;
                await Task.Delay(10);
                Process process = new Process { StartInfo = info };
                process.Start();
                pBar.Value = 100;
                addConsoleLine("Trying to open file");
                await Task.Delay(10);
                openPopup(title:"Done!",content:"sample file created! Fill the file and use it in the next step");
            }
            else
            {
                addConsoleLine("Operation Aborted!");
                pBar.Value = 100;
                pBar.ShowError = true;
            }
        }
        private async void semiAutoCalcClick(object sender, RoutedEventArgs e)
        {
            consoleRTB.Blocks.Clear();
            addConsoleLine("Generating Report from Sample file");
            pBar.ShowError = false;
            pBar.Value = 0;
            addConsoleLine("Waiting for user to choose file:");
            var filePicker = new Windows.Storage.Pickers.FileOpenPicker();
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(filePicker, hWnd);
            filePicker.SuggestedStartLocation = PickerLocationId.Desktop;
            filePicker.FileTypeFilter.Add(".xlsx");
            var file = await filePicker.PickSingleFileAsync();
            if(file != null)
            {
                addConsoleLine("User Choose file:"+file.Path);
                await Task.Delay(500);
                addConsoleLine("Parsing Data");
                float.TryParse(RuleMax1.Text.ToString(), out float rule1);
                float.TryParse(RuleMax2.Text.ToString(), out float rule2);
                float.TryParse(RuleMax3.Text.ToString(), out float rule3);
                float.TryParse(RuleMax4.Text.ToString(), out float rule4);
                float.TryParse(RuleMax5.Text.ToString(), out float rule5);
                await Task.Delay(500);
                addConsoleLine("Running Report Algorithem");
                Algo algo = new Algo(selectColor1.IsChecked.Value, selectColor2.IsChecked.Value, selectColor3.IsChecked.Value, selectColor4.IsChecked.Value, selectColor5.IsChecked.Value
                    , Color1, Color2,Color3,Color4,Color5
                    , file.Path, rule1, rule2, rule3, rule4, rule5, System.IO.Path.GetDirectoryName(file.Path));
                bool integrity = algo.testIntegrity(file.Path);
                if (!integrity)
                {
                    openFormatPopup();
                    return;
                }
                addConsoleLine("Adding Scan Temperature to the Report");
                algo.scanTempSummery(file.Path);
                pBar.Value = 15;
                await Task.Delay(10);
                addConsoleLine("Adding Scan Details to the Report");
                algo.scanDetailsSummery(file.Path);
                pBar.Value = 30;
                await Task.Delay(10);
                addConsoleLine("Adding Part Details to the Report");
                algo.partDetailsSummery(file.Path);
                pBar.Value = 45;
                await Task.Delay(10);
                addConsoleLine("Adding Titles to the Report");
                algo.titleSummery(file.Path);
                pBar.Value = 60;
                await Task.Delay(10);
                addConsoleLine("Adding Top Data Tabel to the Report");
                algo.dataTableTop(file.Path, algo.getTopCorner(file.Path), algo.getBottomCorner(file.Path));
                pBar.Value = 80;
                await Task.Delay(10);
                addConsoleLine("Adding Bottom Data Tabel to the Report");
                algo.dataTableBottom(file.Path, algo.getTopCorner(file.Path), algo.getBottomCorner(file.Path));
                pBar.Value = 90;
                await Task.Delay(10);
                addConsoleLine("Adding Scan Value Summery to the Report");
                algo.scanValueSummery(file.Path);
                pBar.Value = 100;
                addConsoleLine("Done!");
                await Task.Delay(10);
                openPopup(title:"Report Generated!",content:"Report file have been created succesfully!");
            }
        }
        private async void openPopup(bool isError=false,string title="",string content="")
        {
            pBar.ShowError = isError;
            ContentDialog popupDialog = new ContentDialog
            {
                Title = title,
                Content = content,
                CloseButtonText = "Ok"
            };
            //set the XamlRoot property
            popupDialog.XamlRoot = MainTab.XamlRoot;
            ContentDialogResult result = await popupDialog.ShowAsync();
        }
        private async void openFormatPopup()
        {
            pBar.ShowError = true;
            ContentDialog popupDialog = new ContentDialog
            {
                Title = "Wrong Format",
                Content = "the file you choose is in the wrong sample file format, make sure its the file generated from the software itself",
                CloseButtonText = "Ok"
            };
            //set the XamlRoot property
            popupDialog.XamlRoot = MainTab.XamlRoot;
            ContentDialogResult result = await popupDialog.ShowAsync();
        }
    }
}
