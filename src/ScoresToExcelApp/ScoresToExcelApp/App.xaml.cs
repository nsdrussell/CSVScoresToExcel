﻿using System.Windows;

namespace ScoresToExcelApp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static string[] Args;

        private void App_Startup(object sender, StartupEventArgs e)
        {
            // If no command line arguments were provided, don't process them if (e.Args.Length == 0) return;
            if (e.Args.Length > 0)
            {
                Args = e.Args;
            }
        }
    }
}