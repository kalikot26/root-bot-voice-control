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
    public partial class CmdList : Form
    {
        SpeechSynthesizer s = new SpeechSynthesizer();
        Choices listed = new Choices();
        SpeechRecognitionEngine hide = new SpeechRecognitionEngine();

        public CmdList()
        {
            InitializeComponent();

            listed.Add(new String[] {
                                     //Hiding Commands
                                      "Okay Root, Hide your Commands", "Okay Root, use Female voice", "Okay Root, use Male voice"
                                      });
            //If System Command is disabled
            Grammar gr2 = new Grammar(new GrammarBuilder(listed));

            try
            {
                hide.RequestRecognizerUpdate();
                hide.LoadGrammar(gr2);
                hide.SpeechRecognized += enable_SpeechRecognized;
                hide.SetInputToDefaultAudioDevice();
                hide.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch { return; }
        }

        public void say(String h)
        {
            s.Speak(h);

        }

        private void enable_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;
            if (r == "Okay Root, Hide your Commands")
            {
                this.Hide();
                hide.RecognizeAsyncStop();
            }
        }

        private void regardingRootToolStripMenuItem_Click(object sender, EventArgs e)
        {
            abtRoot.Visible = true;
            sysCommand.Visible = false;
            vType.Visible = false;
            Greetings.Visible = false;
            controls.Visible = false;
            launching.Visible = false;
            Keypress.Visible = false;
            Others.Visible = false;
        }

        private void abtRoot_TextChanged(object sender, EventArgs e)
        {

        }

        private void systemCommandToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sysCommand.Visible = true;
            abtRoot.Visible = false;
            vType.Visible = false;
            Greetings.Visible = false;
            controls.Visible = false;
            launching.Visible = false;
            Keypress.Visible = false;
            Others.Visible = false;
        }

        private void greetingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Greetings.Visible = true;
            abtRoot.Visible = false;
            sysCommand.Visible = false;
            vType.Visible = false;
            controls.Visible = false;
            launching.Visible = false;
            Keypress.Visible = false;
        }

        private void voiceTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            vType.Visible = true;
            abtRoot.Visible = false;
            sysCommand.Visible = false;
            Greetings.Visible = false;
            controls.Visible = false;
            launching.Visible = false;
            Keypress.Visible = false;
            Others.Visible = false;
        }

        private void enablingDisablingClosingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            controls.Visible = true;
            abtRoot.Visible = false;
            sysCommand.Visible = false;
            vType.Visible = false;
            Greetings.Visible = false;
            launching.Visible = false;
            Keypress.Visible = false;
            Others.Visible = false;
        }
        private void launchingApplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            launching.Visible = true;
            controls.Visible = false;
            abtRoot.Visible = false;
            sysCommand.Visible = false;
            vType.Visible = false;
            Greetings.Visible = false;
            Keypress.Visible = false;
            Others.Visible = false;
        }

        private void CmdList_Load(object sender, EventArgs e)
        {

        }

        private void menuToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void keyPressingControlsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            launching.Visible = false;
            controls.Visible = false;
            abtRoot.Visible = false;
            sysCommand.Visible = false;
            vType.Visible = false;
            Greetings.Visible = false;
            Keypress.Visible = true;
            Others.Visible = false;
        }

        private void enablingDisablingOtherControlsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            launching.Visible = false;
            controls.Visible = false;
            abtRoot.Visible = false;
            sysCommand.Visible = false;
            vType.Visible = false;
            Greetings.Visible = false;
            Keypress.Visible = false;
            Others.Visible = true;
        }
    }
    }
