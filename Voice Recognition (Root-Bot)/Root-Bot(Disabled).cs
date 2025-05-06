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

namespace Voice_Recognition__Root_Bot_
{
    public partial class Root_Bot_Disabled_ : Form
    {
        SpeechSynthesizer s = new SpeechSynthesizer();
        Choices list = new Choices();
        SpeechRecognitionEngine rec = new SpeechRecognitionEngine();

        public Root_Bot_Disabled_()
        {
            InitializeComponent();
            list.Add(new String[] { "Okay Root, Listen", "Listen root"});

            Grammar gr = new Grammar(new GrammarBuilder(list));

            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeechRecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch { return; }
        }

        public void say(String h)
        {
            s.Speak(h);
        }

        private void rec_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;

            //what you say/command
            if (r == "Listen Root")
            {
                //respond or to execute
                RootMain enable = new RootMain();
                this.Close();
            }

            //what you say/command
            if (r == "Okay Root, Listen")
            {
                //respond or to execute
                RootMain enable = new RootMain();
                this.Hide();
            }

        }

        private void Root_Bot_Disabled__Load(object sender, EventArgs e)
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
    }
}