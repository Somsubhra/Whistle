using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using QuartzTypeLib;
using System.IO;
using System.Windows;
using System.Diagnostics;

namespace Whistle
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine speechEngine = null;
        List<Word> words = new List<Word>();

        private FilgraphManager m_objFilterGraph = null;
        private IBasicAudio m_objBasicAudio = null;
        private IVideoWindow m_objVideoWindow = null;
        private IMediaEvent m_objMediaEvent = null;
        private IMediaEventEx m_objMediaEventEx = null;
        private IMediaPosition m_objMediaPosition = null;
        private IMediaControl m_objMediaControl = null;

        enum MediaStatus { None, Stopped, Paused, Running };
        private MediaStatus m_CurrentStatus = MediaStatus.None;

        private const int WM_APP = 0x8000;
        private const int WM_GRAPHNOTIFY = WM_APP + 1;
        private const int EC_COMPLETE = 0x01;
        private const int WS_CHILD = 0x40000000;
        private const int WS_CLIPCHILDREN = 0x2000000;

        public Form1()
        {
            InitializeComponent();

            try
            {
                speechEngine = createSpeechEngine("en-US");
                speechEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(engine_SpeechRecognized);
                loadGrammarAndCommands();
                speechEngine.SetInputToDefaultAudioDevice();
                speechEngine.RecognizeAsync(RecognizeMode.Multiple);

            }

            catch(Exception ex) {
                MessageBox.Show(ex.Message,"Voice Recognition failed");
            }
        }

        private SpeechRecognitionEngine createSpeechEngine(string preferredCulture) {
            foreach (RecognizerInfo config in SpeechRecognitionEngine.InstalledRecognizers()) {
                if (config.Culture.ToString() == preferredCulture) {
                    speechEngine = new SpeechRecognitionEngine(config);
                    break;
                }
            }

            if (speechEngine == null) {
                MessageBox.Show("The desired culture is not installed on this machine, the speech-engine will continue using " +
                    SpeechRecognitionEngine.InstalledRecognizers()[0].Culture.ToString() + " as the default culture.",
                    "Culture " + preferredCulture + " not found");
                speechEngine = new SpeechRecognitionEngine(SpeechRecognitionEngine.InstalledRecognizers()[0]);
            }

            return speechEngine;
        }

        private void loadGrammarAndCommands() {
            try
            {
                Choices texts = new Choices();
                string[] lines = File.ReadAllLines(Environment.CurrentDirectory + "\\words.txt");
                foreach (string line in lines)
                {
                    if (line.StartsWith("--") || line == String.Empty) continue;
                    var parts = line.Split(new char[] { '|' });
                    words.Add(new Word() { Text = parts[0], AttachedText = parts[1], IsShellCommand = (parts[2] == "true") });
                    texts.Add(parts[0]);
                }
                Grammar wordsList = new Grammar(new GrammarBuilder(texts));
                speechEngine.LoadGrammar(wordsList);
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        private string getKnownTextOrExecute(string command) {
            try
            {
                var cmd = words.Where(c => c.Text == command).First();

                if (cmd.IsShellCommand)
                {
                    //Add all the actions here.
                    if (cmd.AttachedText == ("open")) {
                        openMedia();
                    }

                    if (cmd.AttachedText == "play") {
                        playMedia();
                    }

                    if (cmd.AttachedText == "pause")
                    {
                        pauseMedia();
                    }

                    if (cmd.AttachedText == "stop")
                    {
                        stopMedia();
                    }

                    if (cmd.AttachedText == "exit")
                    {
                        this.Close();
                    }
                    return "You said: " + cmd.AttachedText;
                }
                else
                {
                    return cmd.AttachedText;
                }
            }

            catch(Exception) {
                return command;
            }
        }

        void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
            string txtSpoken = getKnownTextOrExecute(e.Result.Text);
        }

        private void Form_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
            CleanUp();
            speechEngine.RecognizeAsyncStop();
            speechEngine.Dispose();
        }

        private void openMedia() {
            openFileDialog1.Filter = "Media Files|*.mpg;*.avi;*.wma;*.mov;*.wav;*.mp2;*.mp3|All Files|*.*";
            if (DialogResult.OK == openFileDialog1.ShowDialog()) {
                CleanUp();
                m_objFilterGraph = new FilgraphManager();
                m_objFilterGraph.RenderFile(openFileDialog1.FileName);
                m_objBasicAudio = m_objFilterGraph as IBasicAudio;

                try
                {
                    m_objVideoWindow = m_objFilterGraph as IVideoWindow;
                    m_objVideoWindow.Owner = (int)panel1.Handle;
                    m_objVideoWindow.WindowStyle = WS_CHILD | WS_CLIPCHILDREN;
                    m_objVideoWindow.SetWindowPosition(panel1.ClientRectangle.Left,
                        panel1.ClientRectangle.Top,
                        panel1.ClientRectangle.Width,
                        panel1.ClientRectangle.Height);
                }
                catch(Exception) {
                    m_objVideoWindow = null;
                }

                m_objMediaEvent = m_objFilterGraph as IMediaEvent;
                m_objMediaEventEx = m_objFilterGraph as IMediaEventEx;
                m_objMediaEventEx.SetNotifyWindow((int) this.Handle, WM_GRAPHNOTIFY, 0);
                m_objMediaPosition = m_objFilterGraph as IMediaPosition;
                m_objMediaControl = m_objFilterGraph as IMediaControl;
                this.Text = "Whistle- " + openFileDialog1.FileName + "]";
                m_objMediaControl.Run();
                m_CurrentStatus = MediaStatus.Running;

            }
        }

        private void playMedia() {
            m_objMediaControl.Run();
            m_CurrentStatus = MediaStatus.Running;
        }

        private void pauseMedia() {
            m_objMediaControl.Pause();
            m_CurrentStatus = MediaStatus.Paused;
        }

        private void stopMedia() {
            m_objMediaControl.Stop();
            m_objMediaPosition.CurrentPosition = 0;
            m_CurrentStatus = MediaStatus.Stopped;
        }

        private void CleanUp() {
            if (m_objMediaControl != null) {
                m_objMediaControl.Stop();
            }

            m_CurrentStatus = MediaStatus.Stopped;

            if (m_objMediaEventEx != null) {
                m_objMediaEventEx.SetNotifyWindow(0, 0, 0);
            }

            if (m_objVideoWindow != null) {
                m_objVideoWindow.Visible = 0;
                m_objVideoWindow.Owner = 0;
            }

            if (m_objMediaControl != null) {
                m_objMediaControl = null;
            }

            if (m_objMediaPosition != null) {
                m_objMediaPosition = null;
            }

            if (m_objMediaEventEx != null) {
                m_objMediaEventEx = null;
            }

            if (m_objMediaEvent != null) {
                m_objMediaEvent = null;
            }

            if (m_objVideoWindow != null) {
                m_objVideoWindow = null;
            }

            if (m_objBasicAudio != null) {
                m_objBasicAudio = null;
            }

            if (m_objFilterGraph != null) {
                m_objFilterGraph = null;
            }
        }

        private void Form1_SizeChanged(object sender, System.EventArgs e) {
            if (m_objVideoWindow != null) {
                m_objVideoWindow.SetWindowPosition(panel1.ClientRectangle.Left,
                    panel1.ClientRectangle.Top,
                    panel1.ClientRectangle.Width,
                    panel1.ClientRectangle.Height);
            }
        }

    }
}
