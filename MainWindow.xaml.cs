using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

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
        public MainWindow()
        {
            this.InitializeComponent();
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
    }
}
