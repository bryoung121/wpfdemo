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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using NetOffice.OfficeApi;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;


namespace wpfFileDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FileSystemWatcher _folderWatcher;

        private readonly string _inputFolder = @"C:\temp";

        public Presentation PresentationFl;


        public MainWindow()
        {
            InitializeComponent();

            _folderWatcher = new FileSystemWatcher(_inputFolder, "*.*")
            {
                NotifyFilter = NotifyFilters.CreationTime | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName
            };

            _folderWatcher.Created += new FileSystemEventHandler(Input_OnCreated);
            _folderWatcher.Deleted += new FileSystemEventHandler(Input_OnDeleted);
            _folderWatcher.Changed += new FileSystemEventHandler(Input_OnChanged);

            _folderWatcher.EnableRaisingEvents = true;


            SetMyText("Start Up");
        }

        public void Input_OnDeleted(object source, FileSystemEventArgs e)
        {
            string msg = "";
            FileInfo file = new FileInfo(e.FullPath);
            if (file.Name.Contains("~"))
            {
                msg = string.Format("You closed a powerpoint {0}.", file.Name);
            }
            this.Dispatcher.Invoke((Action)(() =>
            {
                SetMyText(msg);
            }));


        }
        public void Input_OnCreated(object source, FileSystemEventArgs e)
        {
            string msg;

            if (e.ChangeType == WatcherChangeTypes.Created)
            {
                FileInfo file = new FileInfo(e.FullPath);
                if (file.Name.Contains("~"))
                {
                    msg = string.Format("You opened a powerpoint {0}.", file.Name);
                }
                else
                {
                    msg = string.Format("A new file has been created named {0}.", file.Name);
                }

                this.Dispatcher.Invoke((Action)(() =>
                {
                    SetMyText(msg);
                }));

            }

        }
        public void Input_OnChanged(object source, FileSystemEventArgs e)
        {

            if (e.ChangeType == WatcherChangeTypes.Changed)
            {

                FileInfo file = new FileInfo(e.FullPath);
                if (!file.Name.Contains("~"))
                {
                    //Call the move file function for this file. 
                    this.Dispatcher.Invoke((Action)(() =>
                    {
                        SetMyText("Noticed the file has changed. We need to push it around.");
                    }));
                }
            }
        }

        public void SetMyText(string msg)
        {
            string finalmsg = string.Format("{0} \n", msg);
            tbLogFile.Text += finalmsg;
            return;
        }

        private void Btn_FileReceived(object sender, RoutedEventArgs e)
        {
            //call the conductor service that a file as been received and needs to be processed for Sync
            ServiceReference2.Service1Client client = new ServiceReference2.Service1Client();
            string retStr;
            //retStr = "test";
            retStr = client.PostFileEventNotificationAsync("test", "test", "test").Result.ToString();
            SetMyText(retStr);

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ServiceReference2.Service1Client client = new ServiceReference2.Service1Client();
            string retStr;
            /*string returnString;
            returnString = client.GetDataAsync("test").Result.ToString();
            SetMyText(returnString);
            */

            retStr = client.MoveFilesAsync("C:\\temp\\a.txt", "C:\\temp_destination\\a.txt").Result.ToString();
            SetMyText(retStr);
        }

        //Launch a powerpoint from CODE
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SetMyText("Opening a powerpoint from code.");

            string filePath = @"C:\temp\testPowerPoint.pptx";

            NetOffice.PowerPointApi.Application app = new NetOffice.PowerPointApi.Application();
            var file = app.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            try
            {
                int.TryParse(app.Version, out int AppVerNum);

                if (AppVerNum > 14)
                {
                    foreach (CommandBar bar in app.CommandBars)
                    {
                        CommandBarControl button = app.CommandBars.FindControl(MsoControlType.msoControlButton, 7458, Type.Missing, Type.Missing);

                        if (button != null)

                        { button.Visible = false; }

                        //Hide the HidePresenterView button
                        button = app.CommandBars.FindControl(MsoControlType.msoControlButton, 24160, Type.Missing, Type.Missing);

                        if (button != null){ button.Visible = false; }

                        foreach (CommandBarControl ctrl in bar.Controls)
                        {
                            if (ctrl.Id == 24160)
                            {
                                ctrl.Enabled = false; ctrl.Visible = false;
                            }

                            //24493 = see all slides
                            if (ctrl.Id == 24493)
                            {
                                ctrl.Enabled = false; ctrl.Visible = false;
                            }

                            //30224 = right-click context: screen
                            if (ctrl.Id == 30224)
                            {
                                CommandBarPopup btn = (CommandBarPopup)ctrl;

                                foreach (CommandBarControl ct in btn.Controls)
                                {
                                    //ID 7458 = Show Taskbar
                                    if (ct.Id == 7458)
                                    { ct.Enabled = false; ct.Visible = false; }

                                    if (ct.Id == 24160)
                                    { ct.Enabled = false; ct.Visible = false; }

                                }

                            }

                        }


                    }

                }
            }
            catch (Exception ex)
            {
                SetMyText(ex.ToString());
            }

            file.SlideShowSettings.Run();
            
            //call a function to say we started

            PresentationFl = file;

            //watch the next slide show event 
            app.SlideShowNextSlideEvent += App_SlideShowNextSlideEvent;

        }
        public void App_SlideShowNextSlideEvent(SlideShowWindow Wn)
        {

            //need to create a call to an api endpoint 

            this.Dispatcher.Invoke(() =>
            {
                SetMyText(string.Format("Next slide event processed for{0}", Wn.View.CurrentShowPosition));
            });
            
        }

    }

}

