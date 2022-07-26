using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Threading;
using System.IO;
using System.Speech.AudioFormat;
using System.Speech.Recognition.SrgsGrammar;
using System.Speech.Synthesis.TtsEngine;
//sing System.Windows.Inp

using System.Runtime.InteropServices;


namespace Text_SpeechAssist_App
{
    public partial class Form1 : Form
    {
        //all the object decleration 
        public SpeechSynthesizer VoiceTTS = new SpeechSynthesizer();//voice for talking
        SpeechRecognitionEngine SRE = new SpeechRecognitionEngine();//voice command catcher 
        SaveFileDialog s = new SaveFileDialog();//file save dialogue ,not implemented yet
        OpenFileDialog file = new OpenFileDialog();//open file dialogue ,|| || || 
        string Title = "Text speech assist";
        string a = null;
        public Form1()
        {
            InitializeComponent();
            
        


        }
       
        
        private void Form1_Load(object sender, EventArgs e)
        {
            VoiceTTS.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Child);//deafult voice selection 
            VoiceTTS.SetOutputToDefaultAudioDevice();//default audio device selection
            //VoiceTTS.SER

            //to show the default speed and volume ,had to call these here to show immidiately after openning the form
            Volshow.Text = VoiceTTS.Volume.ToString();
            speedshow.Text = VoiceTTS.Rate.ToString();
            ///all these lines are for speech recognization and command setup
            ///
         
            Choices command = new Choices();
            command.Add(new string[] { "read", "pause", "resume", "stop", "paste", "clear","undo" }); // voice commands 
            GrammarBuilder gb = new GrammarBuilder();
            Grammar g = new Grammar(gb);
            SRE.LoadGrammarAsync(g);
            SRE.SpeechRecognized += SRE_SpeechRecognized;// recognization event
         

             Volshow.Update();
             Volshow.Show();
             speedshow.Show();
             speedshow.Update();
        }

        #region voice command implement
        private void SRE_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if(e.Result.Text=="read")
            {
                Read.PerformClick();
            }
            else if (e.Result.Text == "pause")
            {
                Pause.PerformClick();
            }
           else if (e.Result.Text == "resume")
            {
                Pause.PerformClick();
            }
            else if (e.Result.Text == "stop")
            {
                Stop.PerformClick();
            }
          else  if (e.Result.Text == "paste")
            {
                Paste.PerformClick();
            }
            else if (e.Result.Text == "clear")
            {
                Clear.PerformClick();
            }
            else if (e.Result.Text == "undo")
            {
                UndoAction(sender,e);
            }

            
        }
       
        private void speech(object sender, EventArgs e)
        {
            Read.Enabled = true;
            
            MaleVoiceradioButton.Enabled = false;
            FemaleVoiceradioButton.Enabled = false;

            VoiceTTS.SpeakAsync(a);

            if (a == null)
            {
                Read.Enabled = false;
            }
            else
            {
                Read.Enabled = true;
            }
            Thread.Sleep(250);
            
            MaleVoiceradioButton.Enabled = true;
            FemaleVoiceradioButton.Enabled = true;
  }
       


        private void richTextBox1_MouseLeave(object sender, EventArgs e)
        {
            a = richTextBox1.Text;
            Read.Enabled = true;

        }
      
        private void Clear_Button_click(object sender, EventArgs e)//function name states the function :P
        {
            richTextBox1.Clear();
            Read.Enabled = false;
        }

        private void Paste_Button_Click(object sender, EventArgs e)//function name states the function :P
        {
            richTextBox1.Paste();
            a = richTextBox1.Text;
            Read.Enabled = true;


        }
        
        private void richTextBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {   //click event
               
                ContextMenu contextMenu = new System.Windows.Forms.ContextMenu();
                MenuItem menuItem = new MenuItem("Cut");
                menuItem.Click += new EventHandler(CutAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new MenuItem("Copy");
                menuItem.Click += new EventHandler(CopyAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new MenuItem("Paste");
                menuItem.Click += new EventHandler(PasteAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new MenuItem("Select All");
                menuItem.Click += new EventHandler(SelectAllAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new MenuItem("Clear");
                menuItem.Click += new EventHandler(ClearAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new MenuItem("ClearAll");
                menuItem.Click += new EventHandler(ClearAllAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new MenuItem("Undo");
                menuItem.Click += new EventHandler(UndoAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new MenuItem("Redo");
                menuItem.Click += new EventHandler(RedoAction);
                contextMenu.MenuItems.Add(menuItem);

                richTextBox1.ContextMenu = contextMenu;


            }
        }
        

      
        private void ClearAllAction(object sender, EventArgs e)//used in menustrip in the top
        {
            richTextBox1.Clear();
        }

        private void RedoAction(object sender, EventArgs e)//used in menustrip in the top
        {
            richTextBox1.Redo();
        }

        private void UndoAction(object sender, EventArgs e)//used in menustrip in the top
        {
            richTextBox1.Undo();
            if (richTextBox1.CanUndo == false)
            {
                richTextBox1.Text = Clipboard.GetText().ToString();
                richTextBox1.Undo();
            }
        }

        private void SelectAllAction(object sender, EventArgs e)//used in menustrip in the top
        {
            richTextBox1.SelectAll();
        }

        private void ClearAction(object sender, EventArgs e)//used in menustrip in the top
        {
            if (richTextBox1.SelectedText != null)
            {
                richTextBox1.SelectedText = "";
            }
            else
            {
                Clipboard.SetText(richTextBox1.Text);
                richTextBox1.Clear();
            }




        }

        void CutAction(object sender, EventArgs e)//used in menustrip in the top
        {
            richTextBox1.Cut();

        }

        void CopyAction(object sender, EventArgs e)//used in menustrip in the top
        {


            if (richTextBox1.SelectedText != null && richTextBox1.SelectedText != "")
            {
                Clipboard.SetData(DataFormats.Rtf, richTextBox1.SelectedRtf);
            }

        }

        void PasteAction(object sender, EventArgs e)//used in menustrip in the top
        {
            if (Clipboard.ContainsText(TextDataFormat.Rtf))
            {
                richTextBox1.SelectedRtf
                    = Clipboard.GetData(DataFormats.Rtf).ToString();
            }
            else if (Clipboard.ContainsText())
            {
                richTextBox1.Paste();
            }

        }
     

      
        private void Pause_button_Click(object sender, EventArgs e)
        {
            if (VoiceTTS.State == SynthesizerState.Speaking)
            {
                Pause.Text = "Resume";
                VoiceTTS.Pause();
                Read.Enabled = false;
            }
            else if (VoiceTTS.State == SynthesizerState.Paused)
            {
                Pause.Text = "Pause";
                VoiceTTS.Resume();
                Read.Enabled = true;
            }

        }

        private void Stop_Button_Click(object sender, EventArgs e)//function name states the function :P
        {
            if (VoiceTTS.State == SynthesizerState.Paused)
            {
                VoiceTTS.SpeakAsyncCancelAll();
                Read.Enabled = true;
                Pause.Text = "Pause";
                VoiceTTS.Resume();

            }
            else if (VoiceTTS.State == SynthesizerState.Speaking)
            {
                VoiceTTS.SpeakAsyncCancelAll();
            }

        }
        #endregion

       
        private void voicefemale(object sender, EventArgs e)
        {
            VoiceTTS.SelectVoiceByHints(VoiceGender.Female);
            MaleVoiceradioButton.ForeColor = System.Drawing.Color.DarkGray;
            FemaleVoiceradioButton.ForeColor = System.Drawing.Color.Black;
        }

        private void voicemale(object sender, EventArgs e)
        {
            VoiceTTS.SelectVoiceByHints(VoiceGender.Male);
            // VoiceTTS.SelectVoice("Microsoft Mark Mobile");
            FemaleVoiceradioButton.ForeColor = System.Drawing.Color.DarkGray;
            MaleVoiceradioButton.ForeColor = System.Drawing.Color.Black;
        }
       

       
        private void ratedown_Click(object sender, EventArgs e)
        {
            if (VoiceTTS.Rate > -10 && VoiceTTS.Rate <= 10)
            {
                VoiceTTS.Rate--;
                speedtrackBar.Value = VoiceTTS.Rate;
                speedshow.Text = VoiceTTS.Rate.ToString();
                speedshow.Update();
                speedtrackBar.Update();
            }

        }

        private void volup_Click(object sender, EventArgs e)
        {
            if (VoiceTTS.Volume >= 0 && VoiceTTS.Volume < 100)
            {
                VoiceTTS.Volume++;
                volumetrackBar.Value = VoiceTTS.Volume;
                Volshow.Text = VoiceTTS.Volume.ToString();
                Volshow.Update();
                volumetrackBar.Update();
            }
        }

        private void voldown_Click(object sender, EventArgs e)
        {
            if (VoiceTTS.Volume > 0 && VoiceTTS.Volume <= 100)
            {
                VoiceTTS.Volume--;
                volumetrackBar.Value = VoiceTTS.Volume;
                Volshow.Text = VoiceTTS.Volume.ToString();
                Volshow.Update();
                volumetrackBar.Update();
            }
        }

        private void rateup_Click(object sender, EventArgs e)
        {
            if (VoiceTTS.Rate >= -10 && VoiceTTS.Rate < 10)
            {
                VoiceTTS.Rate++;
                speedtrackBar.Value = VoiceTTS.Rate;
                speedshow.Text = VoiceTTS.Rate.ToString();
                speedshow.Update();
                speedtrackBar.Update();
            }
        }

        private void Volshow_TextChanged(object sender, EventArgs e)
        {

            Volshow.Update();
        }

        private void speedshow_TextChanged(object sender, EventArgs e)
        {
            speedshow.Update();
        }
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            VoiceTTS.Volume = volumetrackBar.Value;

            Volshow.Text = volumetrackBar.Value.ToString();
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            VoiceTTS.Rate = speedtrackBar.Value;
            speedshow.Text = speedtrackBar.Value.ToString();
        }
        //tooltip message , used in vollume control page
        string tt1 = "Any change in speed\nwill be effective after the\n current paragraph is spoken ";
        string tt2 = "Any change in Volume\nwill be effective after the\n current paragraph is spoken ";
        string tt3 = "Change in Voice\n will be effective after the\n current paragraph is spoken ";

        private void Speedlabel_MouseHover(object sender, EventArgs e)
        {
            ToolTip a = new ToolTip();
            a.ToolTipIcon = ToolTipIcon.None;
            a.IsBalloon = true;
            a.ShowAlways = true;
            a.ReshowDelay = 300;
            a.UseAnimation = true;
            a.UseFading = true;

            a.SetToolTip(Speedlabel, tt1);

        }

        private void Volumelabel_MouseHover(object sender, EventArgs e)
        {
            ToolTip a = new ToolTip();
            a.ToolTipIcon = ToolTipIcon.None;
            a.IsBalloon = true;
            a.ShowAlways = true;
            a.ReshowDelay = 300;
            a.UseAnimation = true;
            a.UseFading = true;

            a.SetToolTip(volumelabel, tt2);
        }

        private void Voicecgenderlabel_MouseHover(object sender, EventArgs e)
        {
            ToolTip a = new ToolTip();
            a.ToolTipIcon = ToolTipIcon.None;
            a.IsBalloon = true;
            a.ShowAlways = true;
            a.ReshowDelay = 300;
            a.UseAnimation = true;
            a.UseFading = true;

            a.SetToolTip(selectvoicelabel, tt3);
        }
       


      
        private void saveAsToolStripMenuItem1_Click(object sender, EventArgs e) //save as button er function, not implemented 
        {

            s.DefaultExt = "*.txt";
            s.Filter = "Text Files|*.txt";
            if (s.ShowDialog() == System.Windows.Forms.DialogResult.OK && s.FileName.Length > 0)
            {
                richTextBox1.SaveFile(s.FileName, RichTextBoxStreamType.PlainText);
            }

        }


        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (this.Text == Title || this.Text == "Untitled-VoicePad")
            {
                SaveFile();


            }
            else
            {
                System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(file.FileName);
                //Writes the text to the file
                SaveFile.WriteLine(richTextBox1.Text);
                //Closes the proccess
                SaveFile.Close();
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);
                this.Text = fileNameWithoutExtension + "-" + Title;
            }

        }

        private void NewToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Modified == true && file.FileName == null)
            {
                this.Text = "*Untitled-" + Title;
            }
            else
            {
                this.Text = "Untitled-" + Title;
                richTextBox1.Clear();
            }

        }

        private void openToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            file.Filter = "Text Files|*.txt";
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);
                this.Text = fileNameWithoutExtension + "-" + Title;
                richTextBox1.LoadFile(file.FileName, RichTextBoxStreamType.PlainText);

            }
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Modified == true)

            {

                DialogResult dr = MessageBox.Show("Save the file before exiting?", "unsaved file", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)

                {
                    SaveFile();

                    richTextBox1.Modified = false;

                    Application.Exit();

                }

                else

                {

                    richTextBox1.Modified = false;

                    Application.Exit();

                }

            }
            else

            {


                DialogResult dr = MessageBox.Show("Really wanna Quit?", "Quit", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    Application.Exit();
                }


            }
        }
        private void SaveFile()

        {

            //setting title of savefiledialog to Save As  

            s.Title = "Save As";

            s.Filter = "Text Document|*.txt";//applied filter       

            s.DefaultExt = "txt";//applied default extension    

            if (s.ShowDialog() == DialogResult.OK)

            {

                

                {

                    

                }


                {

                    richTextBox1.SaveFile(s.FileName, RichTextBoxStreamType.PlainText);
                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(s.FileName);
                    this.Text = fileNameWithoutExtension + "-" + Title;

                                

                }

            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)//close button dialogue
        {
            if (e.CloseReason == CloseReason.WindowsShutDown) return;
            switch (MessageBox.Show(this, "Are you sure you want to Quit?", "Quit", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                case DialogResult.No:
                    e.Cancel = true;

                    break;
                default:
                    break;
            }
        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)//about dialogue
        {

            AboutBox1 box = new AboutBox1();
            box.ShowDialog();
        }

       
        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog font = new FontDialog();
            font.Font = richTextBox1.SelectionFont;
            font.ShowDialog();
            richTextBox1.Font = font.Font;

        }

        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)

        {

            //performing wordwrap operation         

            if (wordWrapToolStripMenuItem1.Checked == false)

            {

                wordWrapToolStripMenuItem1.Checked = true;

                richTextBox1.WordWrap = true;

            }

            else

            {

                wordWrapToolStripMenuItem1.Checked = false;

                richTextBox1.WordWrap = false;

            }

        }

        private void colorToolStripMenuItem_Click(object sender, EventArgs e)//0 reference mane event e connect kora hoy nai 
        {
            ColorDialog c = new ColorDialog();

            c.Color = richTextBox1.SelectionColor;
            c.ShowDialog();
            richTextBox1.ForeColor = c.Color;

        }

        private void backgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog c = new ColorDialog();

            c.Color = richTextBox1.SelectionColor;
            c.ShowDialog();
            richTextBox1.BackColor = c.Color;
        }
       

        private void showToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.BalloonTipText = "Voicepad is now in Normal mode. Minimize to put in system tray";
            notifyIcon1.BalloonTipTitle = "Voicepad Normal";
            notifyIcon1.ShowBalloonTip(1000);
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Move(object sender, EventArgs e)
        {
            if(this.WindowState==FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon1.BalloonTipText = "Voicepad is now in system tray. Double click to go back";
                notifyIcon1.BalloonTipTitle = "Voicepad minimized";
                notifyIcon1.ShowBalloonTip(50);
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
          //  notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
            notifyIcon1.BalloonTipText = "Text speech assist is now in Normal mode. Minimize to put in system tray";
            notifyIcon1.BalloonTipTitle = "Text speech assist Normal";
            notifyIcon1.ShowBalloonTip(500);
        }

        private void voiceCommandsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("read-For reading the contents of Textbox.\npause- For pausing the reading.\nresume- For resuming.\npaste-For pasteing the clipboard contents.\nclear- For clearing the textbox.\nundo-For undoing the last action. ", "Voice Commands ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void MaleVoiceradioButton_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void selectvoicelabel_Click(object sender, EventArgs e)
        {

        }

        private void FemaleVoiceradioButton_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void volumelabel_Click(object sender, EventArgs e)
        {

        }

        private void Speedlabel_Click(object sender, EventArgs e)
        {

        }


        private void button2_Click(object sender, EventArgs e)
        {
           
            
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
