using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NotifyIcon = System.Windows.Forms.NotifyIcon;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.IO;
using NCrontab;
using SimpleJson;

namespace KKDaemon
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string PyExecFile
        {
            get
            {
                return Config["PyExe"] as string;
            }
        }
        private const string ConfigFilePath = "config.json";
        private NotifyIcon notifyIcon;
        short hotkey_flag = 1;

        JsonObject Config;

        readonly Dictionary<System.Windows.Controls.MenuItem, string> MenuItem2CmdMap = new Dictionary<MenuItem, string>();
        CronTaskManager _cronTask = new CronTaskManager();

        public MainWindow()
        {
            InitializeComponent();

            ChDirToThisExe();
            InitConfigs();

            InitNotifyIcon();

            ShowMenus();
            ResetTimer();

            this.Show();
            this.Hide();
        }
        System.Windows.Threading.DispatcherTimer dTimer = null;
        private void ResetTimer()
        {
            if (dTimer != null)
                dTimer.Stop();

            int tipsInterval = 5;  // 默認5秒
            if (Config.ContainsKey("TipsInterval"))
            {
                var o = Config["TipsInterval"]; // 間隔時間
                tipsInterval = int.Parse(o.ToString());
            }

            dTimer = new System.Windows.Threading.DispatcherTimer();
            dTimer.Tick += new EventHandler((object sender, EventArgs e) =>
            {
                this.notifyIcon.BalloonTipText = GetRandTips();
                this.notifyIcon.ShowBalloonTip(500);
            });
            dTimer.Interval = new TimeSpan(0, 0, tipsInterval);
            dTimer.Start();
        }


        private void InitConfigs()
        {
            if (!File.Exists(ConfigFilePath))
            {
                MessageBox.Show("找不到配置文件 " + ConfigFilePath);
            }
            string jsonTxt = File.ReadAllText(ConfigFilePath);

            var jsonObj = SimpleJson.SimpleJson.DeserializeObject(jsonTxt) as JsonObject;

            Config = jsonObj;

            StringBuilder sb = new StringBuilder();
            foreach (string tips in Config["Tips"] as JsonArray)
            {
                sb.AppendLine(tips);
            }

            int tipsInterval = 500;
            int.TryParse(Config["TipsInterval"].ToString(), out tipsInterval);

            this.TextBlockTipsInterval_.Text = tipsInterval.ToString();
            this.TextBlockTips_.Text = sb.ToString();

            this.TextBlockTipsInterval_.TextChanged += TextBlockTipsInterval__TextChanged;
            this.TextBlockTips_.TextChanged += TextBlockTips__Changed;

            ResetMenuCfgs();

            this.TextBox_MenuKeyCfg.TextChanged += TextBox_MenuCfg_TextChanged;
            this.TextBox_MenuValueCfg.TextChanged += TextBox_MenuCfg_TextChanged;

        }

        /// <summary>
        /// 菜单显示
        /// </summary>
        void ResetMenuCfgs()
        {
            // 指令性菜单
            this.ListBox_Menus.Items.Clear();
            foreach (KeyValuePair<string, object> kv in Config["Menus"] as JsonObject)
            {
                ListBoxItem listBoxItem = new ListBoxItem();
                listBoxItem.Selected += OnListBoxItem_Selected;
                listBoxItem.Content = kv.Key;
                this.ListBox_Menus.Items.Add(listBoxItem);
            }

        }
        void SaveConfigFile()
        {
            ChDirToThisExe();

            // Write
            File.WriteAllText(ConfigFilePath, Config.ToString());

            ShowMenus();
            ResetTimer();
        }

        string selectingMenuConf = "";

        void TextBlockTipsInterval__TextChanged(object sender, TextChangedEventArgs e)
        {
            string intervalTxt = this.TextBlockTipsInterval_.Text;
            if (!string.IsNullOrEmpty(intervalTxt))
            {
                // TipsInterval
                Config["TipsInterval"] = this.TextBlockTipsInterval_.Text;

                SaveConfigFile();
            }
        }

        void TextBox_MenuCfg_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(selectingMenuConf))
            {
                
                string key = this.TextBox_MenuKeyCfg.Text;

                JsonObject menuConf = Config["Menus"] as JsonObject;
                menuConf.Remove(selectingMenuConf);
                menuConf[key] = this.TextBox_MenuValueCfg.Text;
                selectingMenuConf = key;

                SaveConfigFile();
                ResetMenuCfgs();
            }
        }

        void OnListBoxItem_Selected(object sender, RoutedEventArgs e)
        {
            selectingMenuConf = null;

            ListBoxItem listBoxItem = sender as ListBoxItem;

            JsonObject menuConf = Config["Menus"] as JsonObject;
            string key = listBoxItem.Content as string;
            string menuScript = menuConf[key].ToString();
            this.TextBox_MenuValueCfg.Text = menuScript;
            this.TextBox_MenuKeyCfg.Text = key;

            selectingMenuConf = key; // 选中的
        }

        void TextBlockTips__Changed(object sender, TextChangedEventArgs e)
        {
            //this.ButtonApplyConfig_.IsEnabled = true;

            // Tips
            string[] tips = this.TextBlockTips_.Text.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            JsonArray tipsArray = new JsonArray();
            foreach (string tip in tips)
            {
                tipsArray.Add(tip);
            }
            Config["Tips"] = tipsArray;

            SaveConfigFile();
        }

        /// <summary>
        /// 註冊全局快捷鍵
        /// </summary>
        /// <param name="e"></param>
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            // 快捷键注册
            WindowInteropHelper wih = new WindowInteropHelper(this);
            HotKey.RegisterHotKey(wih.Handle, hotkey_flag, HotKey.KeyModifiers.Alt, (int)System.Windows.Forms.Keys.Oemtilde);  // alt + ~
            HotKey.RegisterHotKey(wih.Handle, hotkey_flag + 1, HotKey.KeyModifiers.Ctrl, (int)System.Windows.Forms.Keys.Oemtilde);  // ctrl + ~
            HotKey.RegisterHotKey(wih.Handle, hotkey_flag + 2, HotKey.KeyModifiers.Shift | HotKey.KeyModifiers.Ctrl, (int)System.Windows.Forms.Keys.Oemtilde);  // shift + ~
            HotKey.RegisterHotKey(wih.Handle, hotkey_flag + 3, HotKey.KeyModifiers.Shift | HotKey.KeyModifiers.Alt, (int)System.Windows.Forms.Keys.Oemtilde);  // shift + ~
            HwndSource hs = HwndSource.FromHwnd(wih.Handle);
            hs.AddHook(new HwndSourceHook(OnHotKey));
        }

        /// <summary>
        /// 系统快捷键响应
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="msg"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <param name="handled"></param>
        /// <returns></returns>
        private IntPtr OnHotKey(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            switch (msg)
            {
                case HotKey.WM_HOTKEY:
                    {
                        int sid = wParam.ToInt32();
                        if (sid >= hotkey_flag && sid <= (hotkey_flag + 10))
                        {
                            //MessageBox.Show("按下Alt+S");
                            ShowMyMenu();
                        }
                        handled = true;
                        break;
                    }
            }
            return IntPtr.Zero;
        }

        // 工作目錄切換到這個exe文件夾內
        void ChDirToThisExe()
        {
            string myExePath = System.Reflection.Assembly.GetEntryAssembly().Location; // 切換本exe的目錄
            string myExeDir = System.IO.Path.GetDirectoryName(myExePath);
            System.IO.Directory.SetCurrentDirectory(myExeDir);
        }

        void ShowMenus()
        {
            this.ButtonApplyConfig_.ContextMenu = new System.Windows.Controls.ContextMenu();
            var items = this.ButtonApplyConfig_.ContextMenu.Items;

            var title = new System.Windows.Controls.MenuItem();
            title.Header = @"== KKDaemon ==";
            title.Click += (__, __2) =>
            {
                ShowAbout(null, null);
            };
            this.ButtonApplyConfig_.ContextMenu.Items.Add(title);
            this.ButtonApplyConfig_.ContextMenu.Items.Add(new System.Windows.Controls.Separator());

            // Crons
            if (Config.ContainsKey("Cron"))
            {
                foreach (var jCron in (JsonArray)Config["Cron"])
                {
                    var jCronArray2 = (JsonArray) jCron;
                    var cron = (string)jCronArray2[0];
                    var desc = (string)jCronArray2[1];
                    var cmd = (string)jCronArray2[2];
                    
                    // 不能按
                    var menuItem = new MenuItem {Header = desc, IsEnabled = false};
                    _cronTask.BeginTask(cron, () =>
                    {
                        DoCmd(cmd);
                    });
                    this.ButtonApplyConfig_.ContextMenu.Items.Add(menuItem);
                    items.Add(new Separator());
                }

            }


            foreach (KeyValuePair<string, object> kv in (JsonObject)Config["Menus"])
            {
                string text = kv.Key;
                string cmd = kv.Value as string;
                if (text == "-" || cmd == "-")
                {
                    this.ButtonApplyConfig_.ContextMenu.Items.Add(new System.Windows.Controls.Separator());
                }
                else
                {
                    System.Windows.Controls.MenuItem menuItem = new System.Windows.Controls.MenuItem();
                    menuItem.Header = text;
                    MenuItem2CmdMap[menuItem] = cmd;
                    menuItem.Click += OnClickMenu;
                    this.ButtonApplyConfig_.ContextMenu.Items.Add(menuItem);
                }

            }

            System.Windows.Controls.MenuItem exit2 = new System.Windows.Controls.MenuItem();
            this.ButtonApplyConfig_.ContextMenu.Items.Add(new System.Windows.Controls.Separator());
            this.ButtonApplyConfig_.ContextMenu.Items.Add(exit2);
            exit2.Header = @"殘忍地退出";
            exit2.Click += ExitApp;

            this.ButtonApplyConfig_.ContextMenu.IsOpen = false;
        }

        void DoCmd(string scriptFile)
        {

            ChDirToThisExe();

            string fullPath = System.IO.Path.GetFullPath(scriptFile);  // 完整路徑
            string dirPath = System.IO.Path.GetDirectoryName(fullPath);
            string pyExeFullPath = System.IO.Path.GetFullPath(PyExecFile);
            System.IO.Directory.SetCurrentDirectory(dirPath); // chdir
            string pyExeName = System.IO.Path.GetFileName(PyExecFile);

            try
            {
                if (Directory.Exists(fullPath))
                {
                    System.Diagnostics.Process.Start("explorer.exe", fullPath);
                }
                else
                {
                    switch (System.IO.Path.GetExtension(scriptFile))
                    {
                        case ".py":
                            System.Diagnostics.Process.Start(pyExeFullPath, fullPath);
                            break;
                        case ".bat":
                            System.Diagnostics.Process.Start(fullPath);
                            break;
                        default:
                            throw new Exception("UnSupport file " + fullPath);
                    }
                }
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message);
            }


        }
        void OnClickMenu(object sender, RoutedEventArgs e)
        {
            string scriptFile = this.MenuItem2CmdMap[(MenuItem)sender];
            DoCmd(scriptFile);

        }
        /// <summary>
        /// 菜单~
        /// </summary>
        void ShowMyMenu()
        {

            this.Activate();  // 激活窗口~ 失焦時關閉菜單

            this.ButtonApplyConfig_.ContextMenu.IsOpen = true;

        }

        private void ShowAbout(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = System.Windows.WindowState.Normal;
        }


        private void HideAbout(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void ExitApp(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            e.Cancel = true;
            HideAbout(null, null);

        }

        string GetRandTips()
        {
            var rand = new System.Random();

            JsonArray tipsArray = Config["Tips"] as JsonArray;
            int randNum = rand.Next(0, tipsArray.Count);

            return tipsArray[randNum] as string;
        }
        private void InitNotifyIcon()
        {
            this.notifyIcon = new NotifyIcon();

            this.notifyIcon.BalloonTipText = GetRandTips();
            this.notifyIcon.Text = GetRandTips();

            this.notifyIcon.Icon = Properties.Resources.icon;
            this.notifyIcon.Visible = true;
            this.notifyIcon.ShowBalloonTip(500);

            System.Windows.Application.Current.Exit += (object sender, ExitEventArgs exitEvent) =>
            {
                this.notifyIcon.Dispose();
            };

            this.notifyIcon.MouseClick += notifyIcon_MouseClick;
            this.notifyIcon.MouseDoubleClick += (o, _e) =>
            {
                if (_e.Button == System.Windows.Forms.MouseButtons.Left) this.ShowAbout(o, _e);
            };
        }

        void notifyIcon_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            ShowMyMenu();
        }

        private void Onhyperlink_Click(object sender, RoutedEventArgs e)
        {            
            Hyperlink link = sender as Hyperlink;
            try { System.Diagnostics.Process.Start(link.NavigateUri.AbsoluteUri); }
            catch { }
        }
    }
}

class HotKey
{
    /// <summary> 
    /// 如果函数执行成功，返回值不为0。 
    /// 如果函数执行失败，返回值为0。要得到扩展错误信息，调用GetLastError。.NET方法:Marshal.GetLastWin32Error() 
    /// </summary> 
    /// <param name="hWnd">要定义热键的窗口的句柄</param> 
    /// <param name="id">定义热键ID（不能与其它ID重复） </param> 
    /// <param name="fsModifiers">标识热键是否在按Alt、Ctrl、Shift、Windows等键时才会生效</param> 
    /// <param name="vk">定义热键的内容,WinForm中可以使用Keys枚举转换， 
    /// WPF中Key枚举是不正确的,应该使用System.Windows.Forms.Keys枚举，或者自定义正确的枚举或int常量</param> 
    /// <returns></returns> 
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(
    IntPtr hWnd,
    int id,
    KeyModifiers fsModifiers,
    int vk
    );
    /// <summary> 
    /// 取消注册热键 
    /// </summary> 
    /// <param name="hWnd">要取消热键的窗口的句柄</param> 
    /// <param name="id">要取消热键的ID</param> 
    /// <returns></returns> 
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnregisterHotKey(
    IntPtr hWnd,
    int id
    );
    /// <summary> 
    /// 向全局原子表添加一个字符串，并返回这个字符串的唯一标识符,成功则返回值为新创建的原子ID,失败返回0 
    /// </summary> 
    /// <param name="lpString"></param> 
    /// <returns></returns> 
    [DllImport("kernel32", SetLastError = true)]
    public static extern short GlobalAddAtom(string lpString);
    [DllImport("kernel32", SetLastError = true)]
    public static extern short GlobalDeleteAtom(short nAtom);
    /// <summary> 
    /// 定义了辅助键的名称（将数字转变为字符以便于记忆，也可去除此枚举而直接使用数值） 
    /// </summary> 
    [Flags()]
    public enum KeyModifiers
    {
        None = 0,
        Alt = 1,
        Ctrl = 2,
        Shift = 4,
        WindowsKey = 8
    }
    /// <summary> 
    /// 热键的对应的消息ID 
    /// </summary> 
    public const int WM_HOTKEY = 0x312;
}
