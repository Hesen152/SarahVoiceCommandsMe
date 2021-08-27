using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Management;


namespace SarahSpeakRecognition
{
    public partial class Form1 : Form
    {

        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer Sarah = new SpeechSynthesizer();
        
       
        SpeechRecognitionEngine startlistening = new SpeechRecognitionEngine();
        Random rnd = new Random();
        int RecTimeOut = 0;
        DateTime TimeNow = DateTime.Now;



        public Form1()
        {
            Sarah.SelectVoiceByHints(VoiceGender.Female);
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"DefaultCommands.txt")))));
            _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Default_SpeechRecognized);
            _recognizer.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(_recognizer_SpeechRecognized);
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);

            startlistening.SetInputToDefaultAudioDevice();
            startlistening.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"DefaultCommands.txt")))));
            startlistening.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(startlistening_SpeechRecognized);
        }
         private void Default_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int ranNum;
            string speech = e.Result.Text;
            //I know this code is a bit of spaghetti. It's not something I like or do. but i wrote to try :))

            if (speech == "Hello Sarah")
            {
                Sarah.SpeakAsync("Hello my creator");

            }
            if  (speech =="How are you")
            {
                Sarah.SpeakAsync("I am working normaly");

            }


            if  (speech =="What time is it")

            {
                Sarah.SpeakAsync(DateTime.Now.ToString("h mm tt"));
            }

            if (speech=="Stop talking")
            {
                Sarah.SpeakAsyncCancelAll();
                ranNum = rnd.Next(1, 2);
                if (ranNum==1)
                {
                    Sarah.SpeakAsync("Yes sir");

                }

                  if (ranNum==2)
                    {
                    Sarah.SpeakAsync("I am sorry i will be quiet");

                    }

                

            }
            if (speech=="Stop Listening")
            {
                Sarah.SpeakAsync(" if you need me just ask");
                _recognizer.RecognizeAsyncCancel();
                startlistening.RecognizeAsync(RecognizeMode.Multiple);

            }

            if(speech=="Show Commands")
            {
                string[] commands = (File.ReadAllLines(@"DefaultCommands.txt"));
                LstCommands.Items.Clear();
                LstCommands.SelectionMode = SelectionMode.None;
                LstCommands.Visible = true;
                foreach (string command in commands)
                {
                    LstCommands.Items.Add(command);


                }
            }

            if(speech=="Hide Commands")
            {
                LstCommands.Visible = false;

            }


            if (speech=="How much battery power do i have")
            {
                BateryStatus();
            }
        }
       

        private void _recognizer_SpeechRecognized(object sender, SpeechDetectedEventArgs e)
        {
            RecTimeOut = 0;
        }
         private void startlistening_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;

            if (speech=="Wake up")
            {
                startlistening.RecognizeAsyncCancel();
                Sarah.SpeakAsync("Yes,i am here");
                _recognizer.RecognizeAsync(RecognizeMode.Multiple);
            }
        }

        private void TmrSpeaking_Tick(object sender, EventArgs e)
        {
            if (RecTimeOut == 10)
            {
                _recognizer.RecognizeAsyncCancel();

            }

            else if (RecTimeOut==11)
            {
                TmrSpeaking.Stop();
                startlistening.RecognizeAsync(RecognizeMode.Multiple);
                RecTimeOut = 0;
            }

        }


        private void BateryStatus()
        {
            System.Management.ManagementClass wmi = new System.Management.ManagementClass("win32_Battery");
            var allBatteries = wmi.GetInstances();
            foreach (var battery in allBatteries)
            {
                int estimatedChargeRemaninig = Convert.ToInt32(battery["EstimatedChargeRemaining"]);
                if (estimatedChargeRemaninig==100)
                {
                    Sarah.SpeakAsync("the estimated charge is "+estimatedChargeRemaninig+"%,please remove the charge ,the system is full charged");


                }
                if (estimatedChargeRemaninig <100)
                {
                    Sarah.SpeakAsync("the estimated charge is " + estimatedChargeRemaninig + "  please charge it");


                }
                if (estimatedChargeRemaninig <70)
                {
                    Sarah.SpeakAsync("the estimated charge is " + estimatedChargeRemaninig + "less than 70%");


                }
              


            }
        }
    }
}
