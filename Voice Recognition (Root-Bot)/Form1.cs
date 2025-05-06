using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Voice_Recognition__Root_Bot_
{
    public partial class RootMain : Form
    {
        //Enabled System Command
        SpeechSynthesizer s = new SpeechSynthesizer();
        Choices list = new Choices();
        SpeechRecognitionEngine rec = new SpeechRecognitionEngine();

        //If System Command is Disabled
        Choices listed = new Choices();
        SpeechRecognitionEngine enable = new SpeechRecognitionEngine();

        //System Commands
        [DllImport("user32")]
        public static extern void LockWorkStation();
        [DllImport("PowrProf.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool SetSuspendState(bool hiberate, bool forceCritical, bool disableWakeEvent);

        //Launching Application
        Choices launch = new Choices();
        SpeechRecognitionEngine launched = new SpeechRecognitionEngine();

        //KeyPressing
        Choices keypress = new Choices();
        SpeechRecognitionEngine keys = new SpeechRecognitionEngine();

        //VOLUME
        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg,
            IntPtr wParam, IntPtr lParam);

        public RootMain()
        {
            listed.Add(new String[] {
                                     //Enabling Root
                                      "Listen to me now, Root", "Okay Root, Listen to me now","Okay Root, Enable Key pressing control", "Okay Root, Enable Launching Application",
                                      "Okay Root, Show me the available Commands", "Okay Root, Hide the list of Commands", "I Love You, Root"
                                      });
            //If System Command is disabled
            Grammar gr2 = new Grammar(new GrammarBuilder(listed));

            try
            {
                enable.RequestRecognizerUpdate();
                enable.LoadGrammar(gr2);
                enable.SpeechRecognized += enable_SpeechRecognized;
                enable.SetInputToDefaultAudioDevice();
                enable.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch { return; }

            s.SelectVoiceByHints(VoiceGender.Female);
            s.Speak("Root Initiated");

            list.Add(new String[] {
                //About Root - 7
                "Who is your bestfriend, root?", "Who is your owner, root?", "Okay Root, Show yourself",
                "Okay Root, tell me details about you", "Okay Root, Hide yourself", "Root, Are you listening to me?", "Who are you?",
                "Okay Root, Show me the available Commands", "Okay Root, Hide the list of Commands",

                //Root Voice Type - 2
                "Okay Root, use female voice", "Okay Root, use male voice",

                //Root Greetings - 4
                 "Hello, Good Morning Root", "Hello, Good Afternoon Root", "Hello, Good Evening Root", "Hello, How are you root?",

                 //Disabling/Closing Root - 3
                  "Okay Root, Do not Listen to me", "Do not listen to me, root", "Okay Root, terminate yourself",

                //Root system Command - 6
                "Okay Root, Shutdown in one minute", "Okay Root, Shutdown immediately", "Okay Root, Cancel Shutdown", 
                "Okay Root, Take a sleep", "Okay Root, Restart", "Okay Root, Lock yourself",
               
                //Launching Application
                "Do not listen root, and enable launching application", "Okay root, do not listen and enable launching application",

                //Keypress
                "Do not listen root, and enable Keypressing control", "Okay root, do not listen and enable Keypressing control",

            });

            //Enabled System Command
            Grammar gr = new Grammar(new GrammarBuilder(list));

            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeechRecognized;
                rec.SetInputToDefaultAudioDevice();
            }
            catch { return; }

            InitializeComponent();
            //ADDITIONAL COMPONENTS START HERE //Launch Application, Go to, Search for, and Keypress

            launch.Add(new String[] {
                //Disabling launching and enabling listening
                "Listen to me Root and disable launching application", "Okay Root, Listen to me and disable launching application",

                //Launching Application
                "Okay Root, Open Google Chrome Browser", "Okay Root, Open System Notepad", "Okay Root, Open Microsoft Word", "Okay Root, Open Microsoft Powerpoint", "Okay Root, Open Microsoft Excel",
                "Okay Root, Open System Calculator", "Okay Root, Open File Manager", "Okay Root, Open System Camera",
                "Okay Root, Open Windows Settings", "Okay Root, Open Command Prompt", "Okay Root, Open Notepad plus plus",

                //Other Term (Launch instead of "open")

                "Okay Root, Launch Google Chrome Browser", "Okay Root, Launch System Notepad", "Okay Root, Launch Microsoft Word", "Okay Root, Launch Microsoft Powerpoint",
                "Okay Root, Launch Microsoft Excel", "Okay Root, Launch System Calculator", "Okay Root, Launch File Manager",
                "Okay Root, Launch System Camera", "Okay Root, Launch Windows Settings", "Okay Root, Launch Command Prompt", "Okay Root, Launch Notepad plus plus",

                //Commands
                "Okay Root, Show me the available Commands", "Okay Root, Hide the list of Commands",
                });
            //Grammar
            Grammar gr3 = new Grammar(new GrammarBuilder(launch));

            try
            {
                launched.RequestRecognizerUpdate();
                launched.LoadGrammar(gr3);
                launched.SpeechRecognized += launchedApp;
                launched.SetInputToDefaultAudioDevice();
            }
            catch { return; }

            //KEYPRESS
            keypress.Add(new String[]
            {
                //Diabling keypress and enabling root
                "Listen to me Root and disable key pressing", "Okay Root, Listen to me and disable key pressing",

                //Commands
                "Okay Root, Show me the available Commands", "Okay Root, Hide the list of Commands",

                //Enter, Spacebar, Backspace, Escape, Capslock
                "Okay Root, Please press Enter Key", "Okay Root, Please press Spacebar Key", "Okay Root, resume please", "Okay Root, pause please",
                "Okay Root, Please press Delete Key", "Okay Root, Please press Escape Key", "Okay Root, Please press Capslock Key",

                //Volume controls
                "Okay Root, reduce the volume please","Okay Root, increase the volume please", "Okay Root, Mute volume please",

                //ARROW KEYS
                "Okay Root, Next Please", "Okay Root, Back Please", "Okay Root, Move down please", "Okay Root, Move up please"
            });
            //Grammar for keypress
            Grammar gr4 = new Grammar(new GrammarBuilder(keypress));

            try
            {
                keys.RequestRecognizerUpdate();
                keys.LoadGrammar(gr4);
                keys.SpeechRecognized += keypressing;
                keys.SetInputToDefaultAudioDevice();
            }
            catch { return; }
        }

        private void enable_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String t = e.Result.Text;
             
            //Enabling Root
            if (t == "Listen to me now, Root")
            {
                enable.RecognizeAsyncStop();
                rec.RecognizeAsync(RecognizeMode.Multiple);
                say("Root is now listening");
            }

            if (t == "Okay Root, Listen to me now")
            {
                enable.RecognizeAsyncStop();
                rec.RecognizeAsync(RecognizeMode.Multiple);
                say("Root is now listening");
            }

            //"Okay Root, Enable Key pressing control", "Okay Root, Enable Launching Application"
            if (t == "Okay Root, Enable Key pressing control")
            {
                enable.RecognizeAsyncStop();
                rec.RecognizeAsyncStop();
                keys.RecognizeAsync(RecognizeMode.Multiple);
                say("Key Pressing Control Enabled");
            }

            if (t == "Okay Root, Enable Launching Application")
            {
                enable.RecognizeAsyncStop();
                rec.RecognizeAsyncStop();
                launched.RecognizeAsync(RecognizeMode.Multiple);
                say("Launching Application Enabled");
            }

            if (t == "Okay Root, Show me the available Commands")
            {
                say("List of commands should be there");
                CmdList cmdshow = new CmdList();
                cmdshow.Show();
            }

            if (t == "Okay Root, Hide the list of Commands")
            {
                say("Okay, Will hide it");
            }

            
            if (t == "I Love You, Root")
            {
                say("I Love You too, John Venice");
            }
        }

        public void say(String h)
        {
            s.Speak(h);

        }
        private void rec_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;

            //Speak after me -- not completed

           // if (r == "Okay root, Speak after me")
           // {
           //     
           // }

            //Launching Application

            //"Do not listen root, and enable launching application", "Okay root, do not listen and enable launching application"
            if (r == "Okay root, do not listen and enable launching application")
            {
                say("Root will not listen, and will enable launching application");
                rec.RecognizeAsyncStop();
                launched.RecognizeAsync(RecognizeMode.Multiple);
                say("Launching Application Enabled");
            }

            if (r == "Do not listen root, and enable launching application")
            {
                say("Root will not listen, and will enable launching application");
                rec.RecognizeAsyncStop();
                launched.RecognizeAsync(RecognizeMode.Multiple);
                say("Launching Application Enabled");
            }

            //Keypressing
            //"Do not listen root, and enable Keypressing control", "Okay root, do not listen and enable Keypressing control"
            if (r == "Okay root, do not listen and enable Keypressing control")
            {
                say("Root will not listen, and will enable Key pressing");
                rec.RecognizeAsyncStop();
                keys.RecognizeAsync(RecognizeMode.Multiple);
                say("Key Pressing Enabled");
            }

            if (r == "Do not listen root, and enable Keypressing control")
            {
                say("Root will not listen, and will enable Key pressing");
                rec.RecognizeAsyncStop();
                keys.RecognizeAsync(RecognizeMode.Multiple);
                say("Key Pressing Enabled");
            }


            //About Root

            //what you say/command
            if (r == "Okay Root, tell me details about you")
            {
                say("I am Developed by John Venice Almazan(Kalikot), Exclusively for his computer(root)");
            }

            if (r == "Root, Are you listening to me?")
            {
                say("Yes, I'm Listening");
            }

            if (r == "Okay Root, Show me the available Commands")
            {
                say("List of commands should be there");
                CmdList cmdshow = new CmdList();
                cmdshow.Show();
            }

            if (r == "Okay Root, Hide the list of Commands")
            {
                say("Okay, Will hide it");
            }

                if (r == "Who is your bestfriend, root?")
            {
                say("My bestfriend is Kalikot");
            }

            if (r == "Who is your owner, root?")
            {
                say("My owner is John Venice Almazan");
            }

            if (r == "Okay Root, Show yourself")
            {
                this.WindowState = FormWindowState.Normal;
                this.ShowInTaskbar = true;
                notifyicon.Visible = false;
                this.Visible = true;
            }

            if (r == "Okay Root, Hide yourself")
            {
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                notifyicon.Visible = false;
                this.Hide();
            }

            //Disabling/closing root
            if (r == "Okay Root, Do not Listen to me")
            {
                say("I will not be listening");
                rec.RecognizeAsyncStop();
                enable.RecognizeAsync(RecognizeMode.Multiple);
                say("Voice Command Disabled");
            }

            if (r == "Do not listen to me, root")
            {
                say("I will not be listening");
                rec.RecognizeAsyncStop();
                enable.RecognizeAsync(RecognizeMode.Multiple);
                say("Voice Command Disabled");
            }

            if (r == "Okay Root, terminate yourself")
            {
                say("say Good bye to root-bot");
                this.Close();
            }


            //Root Voice Type
            if (r == "Okay Root, use female voice")
            {
                s.SelectVoiceByHints(VoiceGender.Female);
                say("I am now in female voice");
            }

            if (r == "Okay Root, use male voice")
            {
                s.SelectVoiceByHints(VoiceGender.Male);
                say("I am now in male voice");
            }

            //Root System Commands
            if (r == "Okay Root, Shutdown in one minute")
                {
                    Process.Start("shutdown", "/s /t 60");
                    say("Root will shutdown in a minute");
                }

                if (r == "Okay Root, Shutdown immediately")
                {
                    Process.Start("shutdown", "/s /t 0");
                    say("Root is shuttingdown");
                }

                if (r == "Okay Root, Cancel Shutdown")
                {
                    Process.Start("shutdown", "/a");
                    say("Root Shutdown canceled");
                }

                if (r == "Okay Root, Take a sleep")
                {
                say("Good Night Dear");
                SetSuspendState(false, true, true);
                }

                if (r == "Okay Root, Lock yourself")
                {
                say("Unlock me soon");
                LockWorkStation();
                }

                if (r == "Okay Root, Restart")
                {
                say("Root will Reboot");
                Process.Start("shutdown", "/r /t 0");
                
                }


                //Greetings
            if (r == "Hello, Good Morning Root")
            {
                say("Good Morning");
            }

            if (r == "Hello, Good Afternoon Root")
            {
                say("Good Afternoon");
            }

            if (r == "Hello, Good Evening Root")
            {
                say("Good evening");
            }

            if (r == "Hello, How are you root?")
            {
                say("I'm fine");
            }

            if (r == "Perform Enter")
            {
                ;
            }
        }

        //LAUNCHING APPLICATION
        private void launchedApp(object sender, SpeechRecognizedEventArgs e)
        {
            String l = e.Result.Text;

            if (l == "Okay Root, Show me the available Commands")
            {
                say("List of commands should be there");
                CmdList cmdshow = new CmdList();
                cmdshow.Show();
            }

            if (l == "Okay Root, Hide the list of Commands")
            {
                say("Okay, Will hide it");
            }

            if (l == "Listen to me Root and disable launching application")
            {
                say("Launching Application will be disabled and root will listen");
                launched.RecognizeAsyncStop();
                rec.RecognizeAsync(RecognizeMode.Multiple);
                say("Root is now listening");
            }

            if (l == "Okay Root, Listen to me and disable launching application")
            {
                say("Launching Application will be disabled and root will listen");
                launched.RecognizeAsyncStop();
                rec.RecognizeAsync(RecognizeMode.Multiple);
                say("Root is now listening");
            }

            //Starts Here
            if (l == "Okay Root, Open Google Chrome Browser")
            {
                Process.Start("Chrome.exe");
            }

            if (l == "Okay Root, Open System Notepad")
            {
                Process.Start("Notepad.exe");
            }

            if (l == "Okay Root, Open System Calculator")
            {
                Process.Start("Calc");
            }

            if (l == "Okay Root, Open System Camera")
            {
                Process.Start("microsoft.windows.camera:");
            }

            if (l == "Okay Root, Open File Manager")
            {
                Process.Start("explorer.exe");
            }

            if (l == "Okay Root, Open Windows Settings")
            {
                Process.Start("ms-settings:home");
            }

            if (l == "Okay Root, Open Command Prompt")
            {
                Process.Start("cmd.exe");
            }

            //Installed
            if (l == "Okay Root, Open Notepad plus plus")
            {
                Process.Start("notepad++");
            }

            if (l == "Okay Root, Open Microsoft Word")
            {
                Process.Start("winword");
            }

            if (l == "Okay Root, Open Microsoft Powerpoint")
            {
                Process.Start("powerpnt");
            }

            if (l == "Okay Root, Open Microsoft Excel")
            {
                Process.Start("excel.exe");
            }

            //Using Launch instead of "open"

            if (l == "Okay Root, Launch Google Chrome Browser")
            {
                Process.Start("Chrome.exe");
            }

            if (l == "Okay Root, Launch System Notepad")
            {
                Process.Start("Notepad.exe");
            }

            if (l == "Okay Root, Launch System Calculator")
            {
                Process.Start("Calc");
            }

            if (l == "Okay Root, Launch System Camera")
            {
                Process.Start("microsoft.windows.camera:");
            }

            if (l == "Okay Root, Launch File Manager")
            {
                Process.Start("explorer.exe");
            }

            if (l == "Okay Root, Launch System Settings")
            {
                Process.Start("ms-settings:home");
            }

            if (l == "Okay Root, Launch Command Prompt")
            {
                Process.Start("cmd.exe");
            }

            //Installed
            if (l == "Okay Root, Launch Notepad plus plus")
            {
                Process.Start("notepad++");
            }

            if (l == "Okay Root, Launch Microsoft Word")
            {
                Process.Start("winword");
            }

            if (l == "Okay Root, Launch Microsoft Powerpoint")
            {
                Process.Start("powerpnt");
            }

            if (l == "Okay Root, Launch Microsoft Excel")
            {
                Process.Start("excel.exe");
            }

        }

        private void keypressing(object sender, SpeechRecognizedEventArgs e)
        {
            String k = e.Result.Text;

            if (k == "Okay Root, Show me the available Commands")
            {
                say("List of commands should be there");
                CmdList cmdshow = new CmdList();
                cmdshow.Show();
            }

            if (k == "Okay Root, Hide the list of Commands")
            {
                say("Okay, Will hide it");
            }

            if (k == "Listen to me Root and disable key pressing")
            {
                say("Key Pressing will be disabled and root will listen");
                keys.RecognizeAsyncStop();
                rec.RecognizeAsync(RecognizeMode.Multiple);
                say("Root is now listening");
            }

            if (k == "Okay Root, Listen to me and disable key pressing")
            {
                say("Key Pressing will be disabled and root will listen");
                keys.RecognizeAsyncStop();
                rec.RecognizeAsync(RecognizeMode.Multiple);
                say("Root is now listening");
            }

            //Enter, Spacebar, Backspace, Escape,

            if (k == "Okay Root, Please press Enter Key")
            {
                SendKeys.Send("{ENTER}");
            }

            if (k == "Okay Root, Please press Spacebar Key")
            {
                SendKeys.Send(" ");
            }

            if (k == "Okay Root, Please press Delete Key")
            {
                SendKeys.Send("{BACKSPACE}");
            }

            if (k == "Okay Root, Please press Escape Key")
            {
                SendKeys.Send("{ESC}");
            }


            //ARROW KEYS
            //      "Okay Root, Next Please", "Okay Root, Back Please", "Okay Root, Move down please", "Okay Root, Move up please"

            if (k == "Okay Root, Next Please")
            {
                SendKeys.Send("{RIGHT}");
            }

            if (k == "Okay Root, Back Please")
            {
                SendKeys.Send("{LEFT}");
            }

            if (k == "Okay Root, Move down please")
            {
                SendKeys.Send("{DOWN}");
            }

            if (k == "Okay Root, Move up please")
            {
                SendKeys.Send("{UP}");
            }

            //Volume Controls

            if (k == "Okay Root, Mute volume please")
            {
                SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
               (IntPtr)APPCOMMAND_VOLUME_MUTE);
            }

            if (k == "Okay Root, reduce the volume please")
            {   
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                
            }

            
            if (k == "Okay Root, increase the volume please")
            {            
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_UP);
            }

            //Media Controls

            if (k == "Okay Root, pause please")
            {
                SendKeys.Send(" ");
            }

            if (k == "Okay Root, resume please")
            {
                SendKeys.Send(" ");
            }

            

            //Media Play/Pause
            //  if (k == "Okay Root, pause please")
            //  {
            //      SendKeys.Send(" ");
            //  }

            //  if (k == "Okay Root, resume please")
            //  {
            //      SendKeys.Send(" ");
            //  }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            // Determines whether the cursor is in the task bar
            bool cursorNotInBar = Screen.GetWorkingArea(this).Contains(Cursor.Position);

            if (this.WindowState == FormWindowState.Minimized && cursorNotInBar)
            {
                this.ShowInTaskbar = false;
                notifyicon.Visible = true;
                this.Hide();
            }
        }

        private void notifyicon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            notifyicon.Visible = false;
            this.Visible = true;
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void terminateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
