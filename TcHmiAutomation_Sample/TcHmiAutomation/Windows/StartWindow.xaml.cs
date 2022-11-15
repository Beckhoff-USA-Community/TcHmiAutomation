using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace TcHmiAutomationInterface.Windows
{
    /// <summary>
    /// Interaction logic for TcWindow.xaml
    /// </summary>
    public partial class StartWindow : Window
    {
        AutomationInterface automation = null;
       
        public StartWindow()
        {
            InitializeComponent();
            automation = new AutomationInterface();        
        }

        private void Btn_TriggerTest_Click(object sender, RoutedEventArgs e)
        {
            automation.TriggerTest();
        }

        private void StartWindow_Exit(object sender, System.ComponentModel.CancelEventArgs e)
        {
            automation.CloseSolution();
            automation.ExitVisualStudio();
            automation = null;           
            Application.Current.Shutdown();
        }


    }
}
