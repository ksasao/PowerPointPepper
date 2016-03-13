using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.IO;
using Baku.LibqiDotNet;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Win32;

namespace PowerPointPepper
{
    public partial class ThisAddIn
    {
        bool _started = false;
        string _address = "";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //[注] 実行する前に PowerPointPepperプロジェクトのプロパティ>公開>発行するバージョン 
            // のリビジョンを変更してください。
            
            //インストーラープロジェクト(Setup)で指定
            //HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointPepper
            string keyName = @"SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointPepper";
            string val = "Manifest";

            // レジストリの取得
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(keyName);
                string installPath = (string)key.GetValue(val);
                if(installPath == null || installPath.IndexOf("\\")==-1)
                {
                    ShowError("Visual Studio から直接起動することはできません。レジストリが変更されたため、もう一度インストーラーで再インストールしてください。");
                }
                else
                {
                    string path = Path.GetDirectoryName(installPath);
                    PathModifier.AddEnvironmentPaths(path);
                }
                key.Close();
            }
            catch (NullReferenceException)
            {
                ShowError("実行に必要なファイルが不足しています。再インストールしてください。");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.SlideShowNextClick += Application_SlideShowNextClick;
        }

        #endregion

        // スライドショーでページが切り替わった場合に発生するイベント
        void Application_SlideShowNextClick(PowerPoint.SlideShowWindow Wn, PowerPoint.Effect nEffect)
        {
            // PowerPointスライドの 「ノート」部分を取得する。
            // http://msdn.microsoft.com/ja-jp/library/office/ff744720(v=office.15).aspx
            // 等を参照
            string note = Wn.View.Slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text.Trim();

            if (note != "")
            {
                var data = Parse(note);

                foreach (var d in data)
                {
                    switch (d.Command)
                    {
                        case Command.Start:
                            _address = d.Data;
                            _started = true;
                            break;
                        default:
                            if (_started)
                            {
                                SayText(d.Data);
                            }
                            break;
                    }
                }

            }
        }

        void SayText(string data)
        {
            string adr = "tcp://" + _address + ":9559";
            var session = QiSession.Create(adr);

            if (!session.IsConnected)
            {
                ShowError("Pepperに接続できませんでした。[start:(PepperのIPアドレス)] をPowerPointのノート部分に記載してください。");
                return;
            }

            string text = data.Replace(@"\\", @"\").Replace("\r", "").Replace("\n", "");
            var tts = session.GetService("ALAnimatedSpeech");
            //var mode = new KeyValuePair<QiString, QiString>[]
            //{
            //    new KeyValuePair<QiString, QiString>(new QiString("bodyLanguageMode"), new QiString("contextual"))
            //};
            //var map = QiMap<QiString, QiString>.Create(mode);

            tts.Post("say", new QiString(text));//, map);
            session.Close();
            session.Destroy();
        }
        Message[] Parse(string command)
        {
            List<Message> result = new List<Message>();

            int from = 0;
            int to = 0;
            while (to < command.Length)
            {
                if (command[to] == '[')
                {
                    if (to - from > 0)
                    {
                        Message message = new Message
                        {
                            Command = Command.Speech,
                            Data = command.Substring(from, to - from)
                        };
                        result.Add(message);
                    }

                    from = to;

                    while (to < command.Length)
                    {
                        if (command[to] == ']')
                        {
                            string parsed = command.Substring(from + 1, to - from - 1);
                            string[] data = parsed.Split(':');
                            if(data.Length == 2)
                            {
                                var str = data[0].Trim().ToLower();
                                var d = data[1].Trim();

                                if(str == "start")
                                {
                                    Message message = new Message
                                    {
                                        Command = Command.Start,
                                        Data = d
                                    };
                                    result.Add(message);
                                }
                            }
                            to++;
                            break;
                        }
                        to++;
                    }
                    from = to;
                }
                else
                {
                    to++;
                }
            }
            if (to - from > 0)
            {
                Message message = new Message
                {
                    Command = Command.Speech,
                    Data = command.Substring(from, to - from)
                };
                result.Add(message);
            }

            return result.ToArray();
        }

        void ShowError(string message)
        {
            MessageBox.Show(message, "Pepper接続エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        class PathModifier
        {
            public static void AddEnvironmentPaths(params string[] paths)
            {
                var path = new[] { Environment.GetEnvironmentVariable("PATH") ?? "" };
                string newPath = string.Join(Path.PathSeparator.ToString(), path.Concat(paths));
                Environment.SetEnvironmentVariable("PATH", newPath);
            }
        }
    }
}
