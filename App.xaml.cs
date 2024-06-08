﻿using System.Configuration;
using System.Data;
using System.Windows;

using static protocolPlus.Core.DatabaseUtils;

namespace protocolPlus
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            InitDatabase();
        }
    }
}