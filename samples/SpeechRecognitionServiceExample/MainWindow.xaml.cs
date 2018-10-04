//SpeechRecognitionEngine_SpeechRecognized does the main stuff of recognising commands
// LoadGrammarKeyboard 
//   ??? <copyright file="MainWindow.xaml.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
//
// Microsoft Cognitive Services (formerly Project Oxford): https://www.microsoft.com/cognitive-services
//
// Microsoft Cognitive Services (formerly Project Oxford) GitHub:
// https://github.com/Microsoft/Cognitive-Speech-STT-Windows
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// </copyright>
//This is a test
// a handle to an application window.

namespace Microsoft.CognitiveServices.SpeechRecognition
{
    using System;
    using Microsoft.Cognitive.LUIS;
    using System.Windows.Forms;
    using System.ComponentModel;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.IO.IsolatedStorage;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Windows;
    using System.Collections.Generic;
    using System.Threading;
    using SpeechToTextWPFSample.Models;
    using System.Linq;
    using System.Windows.Automation.Peers;
    using System.Windows.Automation.Provider;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Serialization;
    using System.Speech.Recognition;
    using System.Speech.Synthesis;
    using System.Threading.Tasks;
    using System.Drawing;
    using SpeechToTextWPFSample;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName,
    string lpWindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();


        private string languageMatched;
        private string categoryMatched;
        private bool languageAndCategoryAlreadyMatched = false;

        private SpeechRecognitionEngine speechRecognitionEngine = new SpeechRecognitionEngine();
        private SpeechSynthesizer SpeechSynthesizer = new SpeechSynthesizer();
        private SpeechResponseEventArgs lastResult = null;
        /// <summary>
        /// The isolated storage subscription key file name.
        /// </summary>
        private const string IsolatedStorageSubscriptionKeyFileName = "Subscription.txt";

        /// <summary>
        /// The default subscription key prompt message
        /// </summary>
        private const string DefaultSubscriptionKeyPromptMessage = "Paste your subscription key here to start";

        /// <summary>
        /// You can also put the primary key in app.config, instead of using UI.
        /// string subscriptionKey = ConfigurationManager.AppSettings["primaryKey"];
        /// </summary>
        private string subscriptionKey;

        /// <summary>
        /// The data recognition client
        /// </summary>
        private DataRecognitionClient dataClient;

        /// <summary>
        /// The microphone client
        /// </summary>
        private MicrophoneRecognitionClient micClient;

        private System.Windows.Threading.DispatcherTimer dispatcherTimer;
        Process currentProcess;
        private bool isKeyboard=false;
        private string lastCommand;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
        {
            speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognitionEngine_SpeechRecognized);
            this.InitializeComponent();
            this.Initialize();
        }

        private void LoadGrammarCustomIntellisense(string specificLanguage)
        {
            Choices choices = new Choices();
            using (var db = new MyDatabase())
            {
                List<tblLanguage> languages = null;
                if (specificLanguage!=  null )
                {
                    languages=db.tblLanguages.Where(l => l.Language ==specificLanguage).OrderBy(l => l.Language).ToList();
                }
                else
                {
                    languages=db.tblLanguages.Where(l => l.tblCustomIntelliSenses.Count > 0 && l.Active == true).OrderBy(l => l.Language).ToList();
                }
                foreach (var language in languages)
                {
                    List<tblCategory> categories1 = db.tblCategories.OrderBy(c => c.Category).Where(c => c.tblCustomIntelliSenses.Count > 0 && c.Category_Type == "IntelliSense Command"  && c.tblCustomIntelliSenses.Where(s => s.Language_ID== language.ID && (s.ComputerID== null  || s.ComputerID==4 )).Count()>0).ToList();
                    foreach (var category in categories1)
                    {
                        var tempLanguage = language.Language;
                        if (tempLanguage=="Not Applicable")
                        {
                            tempLanguage = "Intellisense";
                        }
                        var count = db.tblCustomIntelliSenses.Where(s => s.Category_ID == category.MenuNumber && s.Language_ID == language.ID && (s.ComputerID == null || s.ComputerID == 4)).Count();
                        choices.Add($"{tempLanguage} {category.Category}");
                        this.WriteCommandLine($"{tempLanguage} {category.Category} ({count})");
                    }
                }
                this.WriteLine($"IntelliSense grammars loaded...");
            }
            choices.Add("Stop IntelliSense");
            this.WriteCommandLine("Stop IntelliSense");
            this.WriteCommandLine("Go Dormant");
            choices.Add("Go Dormant");
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognitionEngine.LoadGrammarAsync(grammar);
        }
        private void LoadGrammarKeyboard()
        {
            //dispatcherTimer.Stop();
            List<string> phoneticAlphabet = new List<string> { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet", "Kilo", "Lima", "Mike", "November", "Oscar", "Papa", "Qubec", "Romeo", "Sierra", "Tango", "Uniform", "Victor", "Whiskey", "X-ray", "Yankee", "Zulu" };
            Choices choices = new Choices();
            this.WriteCommandLine($"Keyboard Commands:");
            this.WriteCommandLine("Press Alpha-Zulu");
            foreach (var item in phoneticAlphabet)
            {
                choices.Add($"Press {item}");
            }
            this.WriteCommandLine("Press 1-9");
            AddChoiceAndWriteCommandline(choices, "Press Zero");
            for (int i = 1; i < 10; i++)
            {
                choices.Add($"Press {i}");
            }
            this.WriteCommandLine("Press Control Alpha-Zulu");
            this.WriteCommandLine("Press Alt Alpha-Zulu");
            this.WriteCommandLine("Press Shift Alpha-Zulu");
            Modifier modifier = new Modifier();
            List<Modifier> modifiers = modifier.GetModifiers();
            foreach (var modifier1 in modifiers)
            {
                foreach (var item in phoneticAlphabet)
                {
                    choices.Add($"Press {modifier1.ModifierName} {item}");
                }
            }
            var mod = modifiers[0];
            foreach (var otherModifier in modifiers)
            {
                if (otherModifier.ModifierName != mod.ModifierName)
                {
                    foreach (var item in phoneticAlphabet)
                    {
                        choices.Add($"Press {mod.ModifierName} {otherModifier.ModifierName} {item}");
                    }
                }
            }
            mod = modifiers[1];
            foreach (var otherModifier in modifiers)
            {
                if (otherModifier.ModifierName != mod.ModifierName)
                {
                    foreach (var item in phoneticAlphabet)
                    {
                        choices.Add($"Press {mod.ModifierName} {otherModifier.ModifierName} {item}");
                    }
                }
            }
            mod = modifiers[2];
            foreach (var otherModifier in modifiers)
            {
                if (otherModifier.ModifierName != mod.ModifierName)
                {
                    foreach (var item in phoneticAlphabet)
                    {
                        choices.Add($"Press {mod.ModifierName} {otherModifier.ModifierName} {item}");
                    }
                }
            }

            AddChoiceAndWriteCommandline(choices, "Press Down");
            AddChoiceAndWriteCommandline(choices, "Press Up");
            AddChoiceAndWriteCommandline(choices, "Press Left");
            AddChoiceAndWriteCommandline(choices, "Press Right");
            AddChoiceAndWriteCommandline(choices, "Press Enter");
            AddChoiceAndWriteCommandline(choices, "Press Backspace");
            AddChoiceAndWriteCommandline(choices, "Press Alt Tab");
            AddChoiceAndWriteCommandline(choices, "Press Tab");
            AddChoiceAndWriteCommandline(choices, "Press Shift Tab");
            AddChoiceAndWriteCommandline(choices, "Press Escape");
            AddChoiceAndWriteCommandline(choices, "Press Delete");
            AddChoiceAndWriteCommandline(choices, "Press Space");
            AddChoiceAndWriteCommandline(choices, "Press Alt Space");
            AddChoiceAndWriteCommandline(choices, "Press Home");
            AddChoiceAndWriteCommandline(choices, "Press Control Home");
            AddChoiceAndWriteCommandline(choices, "Press Control Shift Home");
            AddChoiceAndWriteCommandline(choices, "Press End");
            AddChoiceAndWriteCommandline(choices, "Press Control End");
            AddChoiceAndWriteCommandline(choices, "Press Control Shift End");
            AddChoiceAndWriteCommandline(choices, "Press Page Down");
            AddChoiceAndWriteCommandline(choices, "Press Page Up");
            AddChoiceAndWriteCommandline(choices, "Press Hyphen");
            AddChoiceAndWriteCommandline(choices, "Press _");


            this.WriteCommandLine("Press Function 1-12");
            for (int i = 1; i < 13; i++)
            {
                choices.Add($"Press Function {i}");
            }
            foreach (var item in modifiers)
            {
                for (int i = 1; i < 13; i++)
                {
                    choices.Add($"Press {item.ModifierName} Function {i}");
                }
                choices.Add($"Press {item.ModifierName} Page Down");
                choices.Add($"Press {item.ModifierName} Page Up");
            }            

            mod = modifiers[0];
            foreach (var otherModifier in modifiers)
            {
                if (otherModifier.ModifierName != mod.ModifierName)
                {
                    for (int i = 1; i < 13; i++)
                    {
                        choices.Add($"Press {mod.ModifierName} {otherModifier.ModifierName} Function {i}");
                    }
                }
            }

            mod = modifiers[1];
            foreach (var otherModifier in modifiers)
            {
                if (otherModifier.ModifierName != mod.ModifierName)
                {
                    for (int i = 1; i < 13; i++)
                    {
                        choices.Add($"Press {mod.ModifierName} {otherModifier.ModifierName} Function {i}");
                    }
                }
            }
            mod = modifiers[2];
            foreach (var otherModifier in modifiers)
            {
                if (otherModifier.ModifierName != mod.ModifierName)
                {
                    for (int i = 1; i < 13; i++)
                    {
                        choices.Add($"Press {mod.ModifierName} {otherModifier.ModifierName} Function {i}");
                    }
                }
            }
            Choices numberChoices = new Choices(new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" });
            GrammarBuilder grammarBuilderUp = new GrammarBuilder("Press Up");
            grammarBuilderUp.Append(numberChoices);
            GrammarBuilder grammarBuilderDown = new GrammarBuilder("Press Down");
            grammarBuilderDown.Append(numberChoices);
            GrammarBuilder grammarBuilderLeft = new GrammarBuilder("Press Left");
            grammarBuilderLeft.Append(numberChoices);
            GrammarBuilder grammarBuilderRight = new GrammarBuilder("Press Right");
            grammarBuilderRight.Append(numberChoices);
            Choices directionChoices = new Choices(new GrammarBuilder[] { grammarBuilderUp, grammarBuilderDown, grammarBuilderLeft, grammarBuilderRight });
            Grammar grammarDirections = new Grammar((GrammarBuilder)directionChoices);
            grammarDirections.Name = "Directions";

            choices.Add("Stop Keyboard");
            this.WriteCommandLine("Stop Keyboard");
            this.WriteCommandLine("Go Dormant");
            choices.Add("Go Dormant");
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognitionEngine.LoadGrammarAsync(grammarDirections);
            speechRecognitionEngine.LoadGrammarAsync(grammar);

            UpdateCurrentProcess();
        }

        private void UpdateCurrentProcess()
        {
            IntPtr hwnd = GetForegroundWindow();
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            currentProcess = Process.GetProcessById((int)pid);
            this.WriteLine($"****Current Process: {currentProcess.ProcessName}****");
        }

        private void AddChoiceAndWriteCommandline(Choices choices,string phrase)
        {
            choices.Add(phrase);
            this.WriteCommandLine(phrase);
        }
        private void SpeechRecognitionEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            UpdateCurrentProcess();
            Dispatcher.Invoke(() =>
            {
                Result.Text = ($"Recognised: {e.Result.Text}");
            });
            if (e.Result.Text.ToLower() == "keyboard" && e.Result.Confidence > 0.5)
            {
                this.WriteLine($"Keyboard mode...");
                isKeyboard = true;
                speechRecognitionEngine.UnloadAllGrammars();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                LoadGrammarKeyboard();
            }
            else if (e.Result.Text.ToLower().StartsWith("press ") && e.Result.Confidence > 0.5 && isKeyboard == true)
            {
                UpdateCurrentProcess();
                var value = "";
                this.WriteLine($"****Heard Words***{e.Result.Text}*****************");
                value = e.Result.Text;
                List<string> phoneticAlphabet = new List<string> { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet", "Kilo", "Lima", "Mike", "November", "Oscar", "Papa", "Qubec", "Romeo", "Sierra", "Tango", "Uniform", "Victor", "Whiskey", "X-ray", "Yankee", "Zulu" };
                foreach (var item in phoneticAlphabet)
                {
                    if (value.IndexOf("Shift") > 0)
                    {
                        value = value.Replace(item, item.ToUpper().Substring(0, 1));
                    }
                    else
                    {
                        value = value.Replace(item, item.ToLower().Substring(0, 1));
                    }
                }
                value = value.Replace("Press ", "");
                value = value.Replace("Control", "^");
                value = value.Replace("Alt Space", "% ");
                value = value.Replace("Alt", "%");
                value = value.Replace("Escape", "{Esc}");
                value = value.Replace("Zero", "0");
                value = value.Replace("Tab", "{Tab}");
                value = value.Replace("Backspace", "{Backspace}");
                value = value.Replace("Enter", "{Enter}");
                value = value.Replace("Page Down", "{PgDn}");
                if (value.IndexOf("Page Up") >= 0)
                {
                    value = value.Replace("Page Up", "{PgUp}");
                }
                else
                {
                    value = value.Replace("Up", "{Up}");
                }
                value = value.Replace("Right", "{Right}");
                value = value.Replace("Left", "{Left}");
                value = value.Replace("Down", "{Down}");
                value = value.Replace("Delete", "{Del}");
                value = value.Replace("Home", "{Home}");
                value = value.Replace("End", "{End}");
                value = value.Replace("Hyphen", "-");



                for (int i = 12; i > 0; i--)
                {
                    value = value.Replace($"Function {i}", "{F" + i + "}");
                }
                value = value.Replace("Shift", "+");
                if (value != "% ")
                {
                    value = value.Replace(" ", "");
                }
                if (value == "Space")
                {
                    value = value.Replace("Space", " ");
                }
                this.WriteLine($"*****Sending Keys: {value.Replace("{", "").Replace("}", "").ToString()}*******");
                if (value.Contains("{Up}")  && IsNumber(value.Substring( value.IndexOf("}") +1 )))
                {
                    value="{Up " + value.Substring(value.IndexOf("}") + 1) + "}";
                }
                if (value.Contains("{Down}") && IsNumber(value.Substring(value.IndexOf("}") + 1)))
                {
                    value = "{Down " + value.Substring(value.IndexOf("}") + 1) + "}";
                }
                if (value.Contains("{Left}") && IsNumber(value.Substring(value.IndexOf("}") + 1)))
                {
                    value = "{Left " + value.Substring(value.IndexOf("}") + 1) + "}";
                }
                if (value.Contains("{Right}")  && IsNumber(value.Substring( value.IndexOf("}") +1 )))
                {
                    value="{Right " + value.Substring(value.IndexOf("}") + 1) + "}";
                }

                List<string> keys = new List<string>(new string[] { value });
                SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            }
            else if (e.Result.Text.ToLower() == "toggle microphone" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                this.WriteLine("*************Keys sent to toggle the microphone*****************");
                List<string> keys = new List<string>(new string[] { "{ADD}" });
                SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            }
            else if (e.Result.Text.ToLower() == "restart dragon" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                this.WriteLine("*************Restarting Dragon and KnowBrainer*****************");
                var name = "Dragon Naturally speaking";
                KillAllProcesses(name);

                name = "Dragonbar";
                KillAllProcesses(name);
                name = "Command Browser";
                KillAllProcesses(name);
                name = "KnowBrainer";
                KillAllProcesses(name);
                try
                {
                    Process process = new Process();
                    var filename = "C:\\Program Files(x86)\\KnowBrainer\\KnowBrainer Professional 2017\\KBPro.exe";
                    if (File.Exists(filename))
                    {
                        process.StartInfo.UseShellExecute = true;
                        process.StartInfo.WorkingDirectory = "C:\\Program Files (x86)\\KnowBrainer\\KnowBrainer Professional 2017\\";
                        process.StartInfo.FileName = filename;
                        process.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                        process.Start();

                    }

                    else
                    {
                        List<string> keysKB = new List<string>(new string[] { "^+k" });
                        SendKeysCustom(null,  null , keysKB, currentProcess.ProcessName);

                    }
                }
                catch (Exception exception)
                {
                    Dispatcher.Invoke(() =>
                    {
                        _logText.Text += $"An error has occurred: {exception.Message}";
                    });
                }
                //List<string> keys = new List<string>(new string[] { "^%v" });
                //SendKeysCustom(null, "Untitled - Notepad", keys, "notepad", "Notepad.exe");
            }
            else if (e.Result.Text.ToLower().StartsWith("kill") && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                var name = e.Result.Text.ToLower();
                name = name.Replace("kill", "");
                name = name.Trim();
                KillAllProcesses(name);
            }
            else if (e.Result.Text.ToLower() == "quit application" && e.Result.Confidence > 0.5)
            {
                List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
                SendKeysCustom(null, null, keys, currentProcess.ProcessName);
                try
                {
                    System.Windows.Application.Current.Shutdown();
                }
                catch (Exception)
                {
                }
            }
            else if ((e.Result.Text.ToLower() == "emergency speech" && e.Result.Confidence > 0.9) || (e.Result.Text.ToLower()=="list commands" && e.Result.Confidence>0.5))
            {

                speechRecognitionEngine.UnloadAllGrammars();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                if (lastCommand== null )
                {
                    ListCommands();
                }
                else if (lastCommand.ToLower()=="global intellisense")
                {
                    LoadGrammarCustomIntellisense( null ); 
                }
                else if (lastCommand.ToLower()=="visual basic intellisense")
                {
                    LoadGrammarCustomIntellisense("Visual Basic");
                }
                else if (lastCommand.ToLower()=="c sharp intellisense")
                {
                    LoadGrammarCustomIntellisense("C Sharp");
                }
                else if (lastCommand.ToLower()=="javascript intellisense")
                {
                    LoadGrammarCustomIntellisense("JavaScript");
                }
                else if (lastCommand.ToLower().StartsWith("press "))
                {
                    this.WriteLine($"Keyboard mode...");
                    isKeyboard = true;
                    LoadGrammarKeyboard();
                }
                else if (languageAndCategoryAlreadyMatched==true)
                {
                    MatchLanguageAndCategory(languageMatched + " " + categoryMatched);
                }
                else
                {
                    ListCommands();
                }
                this.WriteLine("****Pressing the / Key on the number pad****");
                List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
                SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            }
            else if (e.Result.Text.ToLower() == "short phrase mode" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                SetDragonToSleep();

                speechRecognitionEngine.UnloadAllGrammars();
                speechRecognitionEngine.RecognizeAsyncStop();

                this.IsMicrophoneClientShortPhrase = true;
                this.IsMicrophoneClientWithIntent = false;
                this.IsMicrophoneClientDictation = false;
                this.IsDataClientShortPhrase = false;
                this.IsDataClientWithIntent = false;
                this.IsDataClientDictation = false;

                // Set the default choice for the group of checkbox.
                this._micIntentRadioButton.IsChecked = false;
                this._micDictationRadioButton.IsChecked = false;
                this._micRadioButton.IsChecked = true;
                ResetEverything();
                //Click the button
                ButtonAutomationPeer peer = new ButtonAutomationPeer(this._startButton);
                IInvokeProvider invokeProvider = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                invokeProvider.Invoke();
                dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 0, 100);
                dispatcherTimer.Start();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                this.WriteCommandLine("Grammar Mode");
                this.WriteCommandLine("Transfer to Notepad");
                this.WriteCommandLine("Transfer as Paragraph");

                Choices choices = new Choices();
                choices.Add("Grammar Mode");
                choices.Add("Transfer to Notepad");
                choices.Add("Transfer as Paragraph");
                for (int i = 0; i < 6; i++)
                {
                    choices.Add($"Choose {i}");
                    this.WriteCommandLine($"Choose {i}");
                }

                Grammar grammar = new Grammar(new GrammarBuilder(choices));
                speechRecognitionEngine.LoadGrammarAsync(grammar);
                speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
            }
            else if (e.Result.Text.ToLower().StartsWith("choose ") && lastResult != null)
            {
                var choiceNumber = Int32.Parse(e.Result.Text.Substring(7));
                isKeyboard = false;
                if (choiceNumber <= (lastResult.PhraseResponse.Results.Length - 1))
                {
                    Dispatcher.Invoke(() => { finalResult.Text = lastResult.PhraseResponse.Results[choiceNumber].DisplayText; });
                }
            }
            else if (e.Result.Text.ToLower() == "transfer as paragraph" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                if (this.TransferAsParagraph.IsChecked == false)
                {
                    Dispatcher.Invoke(() => { TransferAsParagraph.IsChecked = true; });
                }
                else
                {
                    Dispatcher.Invoke(() => { TransferAsParagraph.IsChecked = false; });
                }
            }
            else if (e.Result.Text.ToLower() == "transfer to notepad" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                var resultText = "";
                if (this.TransferAsParagraph.IsChecked == true)
                {
                    resultText = Environment.NewLine + this.finalResult.Text;
                }
                else
                {
                    resultText = "  " + this.finalResult.Text;
                }
                List<string> keys = new List<string>(new string[] { resultText });
                SendKeysCustom(null, "Untitled - Notepad", keys, "notepad", "Notepad.exe");
            }
            else if (e.Result.Text.ToLower() == "grammar mode" && e.Result.Confidence > 0.5)
            {
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                ListCommands();
                isKeyboard = false;
            }
            else if (e.Result.Text.ToLower() == "long dictation mode" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                speechRecognitionEngine.RecognizeAsyncStop();
                this.IsMicrophoneClientShortPhrase = false;
                this.IsMicrophoneClientWithIntent = false;
                this.IsMicrophoneClientDictation = true;
                this.IsDataClientShortPhrase = false;
                this.IsDataClientWithIntent = false;
                this.IsDataClientDictation = false;

                // Set the default choice for the grou'p of checkbox.
                this._micIntentRadioButton.IsChecked = false;
                this._micDictationRadioButton.IsChecked = true;
                this._micRadioButton.IsChecked = false;
                ResetEverything();
                ButtonAutomationPeer peer = new ButtonAutomationPeer(this._startButton);
                IInvokeProvider invokeProvider = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                invokeProvider.Invoke();
                dispatcherTimer.Start();
                this.WriteCommandLine("Grammar Mode");
            }
            else if (e.Result.Text.ToLower() == "with intent mode" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                speechRecognitionEngine.RecognizeAsyncStop();
                this.IsMicrophoneClientShortPhrase = false;
                this.IsMicrophoneClientWithIntent = true;
                this.IsMicrophoneClientDictation = false;
                this.IsDataClientShortPhrase = false;
                this.IsDataClientWithIntent = false;
                this.IsDataClientDictation = false;

                // Set the default choice for the grou'p of checkbox.
                this._micIntentRadioButton.IsChecked = true;
                this._micDictationRadioButton.IsChecked = false;
                this._micRadioButton.IsChecked = false;
                ResetEverything();
                ButtonAutomationPeer peer = new ButtonAutomationPeer(this._startButton);
                IInvokeProvider invokeProvider = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                invokeProvider.Invoke();
                dispatcherTimer.Start();
                this.WriteCommandLine("Grammar Mode");
            }
            else if ((e.Result.Text.ToLower() == "visual basic intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower()=="go back" && languageMatched.ToLower()=="visual basic" && e.Result.Confidence>0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                speechRecognitionEngine.UnloadAllGrammars();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                LoadGrammarCustomIntellisense("Visual Basic");
                languageAndCategoryAlreadyMatched = false;
            }
            else if ((e.Result.Text.ToLower() == "c sharp intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower() == "go back" && languageMatched.ToLower() == "c sharp" && e.Result.Confidence > 0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                speechRecognitionEngine.UnloadAllGrammars();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                LoadGrammarCustomIntellisense("C Sharp");
                languageAndCategoryAlreadyMatched = false;
            }
            else if ((e.Result.Text.ToLower() == "javascript intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower() == "go back" && languageMatched.ToLower() == "javascript" && e.Result.Confidence > 0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                speechRecognitionEngine.UnloadAllGrammars();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                LoadGrammarCustomIntellisense("JavaScript");
                languageAndCategoryAlreadyMatched = false;
            }
            else if ((e.Result.Text.ToLower() == "global intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower() == "go back" && languageMatched.ToLower() != "visual basic" && e.Result.Confidence > 0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                speechRecognitionEngine.UnloadAllGrammars();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                LoadGrammarCustomIntellisense(null);
                languageAndCategoryAlreadyMatched = false;
            }
            else if (e.Result.Text.ToLower() == "go dormant" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                speechRecognitionEngine.UnloadAllGrammars();
                Dispatcher.Invoke(() =>
                {
                    Commands.Text = "";
                });
                Choices choices = new Choices();
                choices.Add("emergency speech");
                this.WriteCommandLine("Emergency Speech");
                Grammar grammar = new Grammar(new GrammarBuilder(choices));
                speechRecognitionEngine.LoadGrammarAsync(grammar);
                this.WriteLine("****Pressing the / Key on the number pad****");
                List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
                SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            }
            else if (e.Result.Text == "Stop IntelliSense" || e.Result.Text == "Stop Keyboard")
            {
                isKeyboard = false;
                speechRecognitionEngine.RecognizeAsyncCancel();
                ListCommands();
                speechRecognitionEngine.SetInputToDefaultAudioDevice();
                speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
                languageAndCategoryAlreadyMatched = false;
            }
            else if (languageAndCategoryAlreadyMatched == true)
            {
                isKeyboard = false;
                PerformGlobalIntelliSense(e);
            }
            else
            {
                isKeyboard = false;
                MatchLanguageAndCategory(e.Result.Text);
            }

            if (e.Result.Text.ToLower()!="go dormant")
            {
                lastCommand = e.Result.Text.ToLower();
            }
        }

        private void SetDragonToSleep()
        {
            //Send Dragon to sleep
            this.WriteLine("*************Key sent to sleep the microphone*****************");
            List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
            SendKeysCustom(null, null, keys, currentProcess.ProcessName, null);
        }

        public Boolean IsNumber(String value)
        {
            return value.All(Char.IsDigit);
        }



        #region Events

        /// <summary>
        /// Implement INotifyPropertyChanged interface
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Events

        /// <summary>
        /// Gets or sets a value indicating whether this instance is microphone client short phrase.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is microphone client short phrase; otherwise, <c>false</c>.
        /// </value>
        public bool IsMicrophoneClientShortPhrase { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is microphone client dictation.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is microphone client dictation; otherwise, <c>false</c>.
        /// </value>
        public bool IsMicrophoneClientDictation { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is microphone client with intent.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is microphone client with intent; otherwise, <c>false</c>.
        /// </value>
        public bool IsMicrophoneClientWithIntent { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is data client short phrase.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is data client short phrase; otherwise, <c>false</c>.
        /// </value>
        public bool IsDataClientShortPhrase { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is data client with intent.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is data client with intent; otherwise, <c>false</c>.
        /// </value>
        public bool IsDataClientWithIntent { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is data client dictation.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is data client dictation; otherwise, <c>false</c>.
        /// </value>
        public bool IsDataClientDictation { get; set; }

        /// <summary>
        /// Gets or sets subscription key
        /// </summary>
        public string SubscriptionKey
        {
            get
            {
                return this.subscriptionKey;
            }

            set
            {
                this.subscriptionKey = value;
                this.OnPropertyChanged<string>();
            }
        }

        /// <summary>
        /// Gets the LUIS endpoint URL.
        /// </summary>
        /// <value>
        /// The LUIS endpoint URL.
        /// </value>
        private string LuisEndpointUrl
        {
            get { return ConfigurationManager.AppSettings["LuisEndpointUrl"]; }
        }

        /// <summary>
        /// Gets a value indicating whether or not to use the microphone.
        /// </summary>
        /// <value>
        ///   <c>true</c> if [use microphone]; otherwise, <c>false</c>.
        /// </value>
        private bool UseMicrophone
        {
            get
            {
                return this.IsMicrophoneClientWithIntent ||
                    this.IsMicrophoneClientDictation ||
                    this.IsMicrophoneClientShortPhrase;
            }
        }

        /// <summary>
        /// Gets a value indicating whether LUIS results are desired.
        /// </summary>
        /// <value>
        ///   <c>true</c> if LUIS results are to be returned otherwise, <c>false</c>.
        /// </value>
        private bool WantIntent
        {
            get
            {
                return !string.IsNullOrEmpty(this.LuisEndpointUrl) &&
                    (this.IsMicrophoneClientWithIntent || this.IsDataClientWithIntent);
            }
        }

        /// <summary>
        /// Gets the current speech recognition mode.
        /// </summary>
        /// <value>
        /// The speech recognition mode.
        /// </value>
        private SpeechRecognitionMode Mode
        {
            get
            {
                if (this.IsMicrophoneClientDictation ||
                    this.IsDataClientDictation)
                {
                    return SpeechRecognitionMode.LongDictation;
                }

                return SpeechRecognitionMode.ShortPhrase;
            }
        }

        /// <summary>
        /// Gets the default locale.
        /// </summary>
        /// <value>
        /// The default locale.
        /// </value>
        private string DefaultLocale
        {
            get { return "en-GB"; }
        }

        /// <summary>
        /// Gets the short wave file path.
        /// </summary>
        /// <value>
        /// The short wave file.
        /// </value>
        private string ShortWaveFile
        {
            get
            {
                return ConfigurationManager.AppSettings["ShortWaveFile"];
            }
        }

        /// <summary>
        /// Gets the long wave file path.
        /// </summary>
        /// <value>
        /// The long wave file.
        /// </value>
        private string LongWaveFile
        {
            get
            {
                return ConfigurationManager.AppSettings["LongWaveFile"];
            }
        }

        /// <summary>
        /// Gets the Cognitive Service Authentication Uri.
        /// </summary>
        /// <value>
        /// The Cognitive Service Authentication Uri.  Empty if the global default is to be used.
        /// </value>
        private string AuthenticationUri
        {
            get
            {
                return ConfigurationManager.AppSettings["AuthenticationUri"];
            }
        }

        /// <summary>
        /// Raises the System.Windows.Window.Closed event.
        /// </summary>
        /// <param name="e">An System.EventArgs that contains the event data.</param>
        protected override void OnClosed(EventArgs e)
        {
            if (null != this.dataClient)
            {
                this.dataClient.Dispose();
            }

            if (null != this.micClient)
            {
                this.micClient.Dispose();
            }

            base.OnClosed(e);
        }

        /// <summary>
        /// Saves the subscription key to isolated storage.
        /// </summary>
        /// <param name="subscriptionKey">The subscription key.</param>
        private static void SaveSubscriptionKeyToIsolatedStorage(string subscriptionKey)
        {
            using (IsolatedStorageFile isoStore = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Assembly, null, null))
            {
                using (var oStream = new IsolatedStorageFileStream(IsolatedStorageSubscriptionKeyFileName, FileMode.Create, isoStore))
                {
                    using (var writer = new StreamWriter(oStream))
                    {
                        writer.WriteLine(subscriptionKey);
                    }
                }
            }
        }

        /// <summary>
        /// Initializes a fresh audio session.
        /// </summary>
        private void Initialize()
        {
            this.IsMicrophoneClientShortPhrase = true;
            this.IsMicrophoneClientWithIntent = false;
            this.IsMicrophoneClientDictation = false;
            this.IsDataClientShortPhrase = false;
            this.IsDataClientWithIntent = false;
            this.IsDataClientDictation = false;

            // Set the default choice for the group of checkbox.
            //this._micIntentRadioButton.IsChecked = true;
            this._micRadioButton.IsChecked = true;
            this.SubscriptionKey = this.GetSubscriptionKeyFromIsolatedStorage();

            //ButtonAutomationPeer peer = new ButtonAutomationPeer(this._startButton);
            //IInvokeProvider invokeProvider = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
            //invokeProvider.Invoke();

            dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
            dispatcherTimer.Tick += dispatcherTimer_Tick;
            dispatcherTimer.Interval = new TimeSpan(0, 0,0,0, 100);

            Screen s = System.Windows.Forms.Screen.AllScreens[1];


            System.Drawing.Rectangle r = s.WorkingArea;
            this.Top = r.Top;
            this.Left = r.Left;

            ListCommands();
            speechRecognitionEngine.SetInputToDefaultAudioDevice();
            speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
            UpdateCurrentProcess();
            List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
        }

        // Send a series of key presses to the Calculator application.
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            //List<string> keys = new List<string>(new string[] { "^`" });
            //List<string> keys = new List<string>(new string[] { "^+%\\" });
            List<string> keys = new List<string>(new string[] { "111", "*", "2", "=" });
            SendKeysCustom("ApplicationFrameWindow","Calculator",keys,"calc");
        }

        private  void SendKeysCustom(string applicationClass, string applicationCaption, List<string> keys, string processName,string applicationToLaunch="",int delay=0)
        {
            // Get a handle to the application. The window class
            // and window name can be obtained using the Spy++ tool.
            IntPtr applicationHandle=IntPtr.Zero;
            while (true)
            {
                if (applicationClass!= null  || applicationCaption!= null )
                {
                    applicationHandle = FindWindow(applicationClass, applicationCaption);
                }

                // Verify that Application is a running process.
                if (applicationHandle == IntPtr.Zero)
                {
                    if (applicationToLaunch!= null && applicationToLaunch.Length>0)
                    {
                        Process.Start(applicationToLaunch);
                        Thread.Sleep(1000);
                    }
                    else
                    {
                 //       System.Windows.MessageBox.Show($"{applicationCaption} is not running.");
                        //ActivateApp(processName);
                        Process process = Process.GetProcessesByName(processName)[0];
                        applicationHandle = process.Handle;
                        break;
                    }
                }
                else
                {
                    break;
                }
            }

            // Make Application the foreground application and send it 
            // a set of Keys.
            SetForegroundWindow(applicationHandle);
            foreach (var item in keys)
            {
                Thread.Sleep(delay);
                try
                {
                    var temporary = item.Replace("(", "{(}");
                    temporary = temporary.Replace(")", "{)}");

                    SendKeys.SendWait(temporary);
                }
                catch (Exception exception)
                {
                    this.WriteLine($"An error has occurred: {exception.Message}");
                }
            }
        }

        /// <summary>
        /// Handles the Click event of the _startButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            ListenForSpeech(sender,e);
        }

        private void ListenForSpeech(object sender, RoutedEventArgs e)
        {
            this._startButton.IsEnabled = false;
            this._radioGroup.IsEnabled = false;

            this.LogRecognitionStart();

            if (this.UseMicrophone)
            {
                if (this.micClient == null)
                {
                    if (this.WantIntent)
                    {
                        this.CreateMicrophoneRecoClientWithIntent();
                    }
                    else
                    {
                        this.CreateMicrophoneRecoClient();
                    }
                }

                this.micClient.StartMicAndRecognition();
            }
            else
            {
                if (null == this.dataClient)
                {
                    if (this.WantIntent)
                    {
                        this.CreateDataRecoClientWithIntent();
                    }
                    else
                    {
                        this.CreateDataRecoClient();
                    }
                }

                this.SendAudioHelper((this.Mode == SpeechRecognitionMode.ShortPhrase) ? this.ShortWaveFile : this.LongWaveFile);
            }
        }

        /// <summary>
        /// Logs the recognition start.
        /// </summary>
        private void LogRecognitionStart()
        {
            string recoSource;
            if (this.UseMicrophone)
            {
                recoSource = "microphone";
            }
            else if (this.Mode == SpeechRecognitionMode.ShortPhrase)
            {
                recoSource = "short wav file";
            }
            else
            {
                recoSource = "long wav file";
            }

            this.WriteLine("\n--- Start speech recognition using " + recoSource + " with " + this.Mode + " mode in " + this.DefaultLocale + " language ----\n\n");
        }

        /// <summary>
        /// Creates a new microphone reco client without LUIS intent support.
        /// </summary>
        private void CreateMicrophoneRecoClient()
        {
            this.micClient = SpeechRecognitionServiceFactory.CreateMicrophoneClient(
                this.Mode,
                this.DefaultLocale,
                this.SubscriptionKey);
            this.micClient.AuthenticationUri = this.AuthenticationUri;

            // Event handlers for speech recognition results
            this.micClient.OnMicrophoneStatus += this.OnMicrophoneStatus;
            this.micClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            if (this.Mode == SpeechRecognitionMode.ShortPhrase)
            {
                this.micClient.OnResponseReceived += this.OnMicShortPhraseResponseReceivedHandler;
            }
            else if (this.Mode == SpeechRecognitionMode.LongDictation)
            {
                this.micClient.OnResponseReceived += this.OnMicDictationResponseReceivedHandler;
            }

            this.micClient.OnConversationError += this.OnConversationErrorHandler;
        }

        /// <summary>
        /// Creates a new microphone reco client with LUIS intent support.
        /// </summary>
        private void CreateMicrophoneRecoClientWithIntent()
        {
            this.WriteLine("--- Start microphone dictation with Intent detection ----");

            this.micClient =
                SpeechRecognitionServiceFactory.CreateMicrophoneClientWithIntentUsingEndpointUrl(
                    this.DefaultLocale,
                    this.SubscriptionKey,
                    this.LuisEndpointUrl);
            this.micClient.AuthenticationUri = this.AuthenticationUri;
            this.micClient.OnIntent += this.OnIntentHandler;

            // Event handlers for speech recognition results
            this.micClient.OnMicrophoneStatus += this.OnMicrophoneStatus;
            this.micClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            this.micClient.OnResponseReceived += this.OnMicShortPhraseResponseReceivedHandler;
            this.micClient.OnConversationError += this.OnConversationErrorHandler;
        }

        /// <summary>
        /// Handles the Click event of the HelpButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://www.projectoxford.ai/doc/general/subscription-key-mgmt");
        }

        /// <summary>
        /// Creates a data client without LUIS intent support.
        /// Speech recognition with data (for example from a file or audio source).  
        /// The data is broken up into buffers and each buffer is sent to the Speech Recognition Service.
        /// No modification is done to the buffers, so the user can apply their
        /// own Silence Detection if desired.
        /// </summary>
        private void CreateDataRecoClient()
        {
            this.dataClient = SpeechRecognitionServiceFactory.CreateDataClient(
                this.Mode,
                this.DefaultLocale,
                this.SubscriptionKey);
            this.dataClient.AuthenticationUri = this.AuthenticationUri;

            // Event handlers for speech recognition results
            if (this.Mode == SpeechRecognitionMode.ShortPhrase)
            {
                this.dataClient.OnResponseReceived += this.OnDataShortPhraseResponseReceivedHandler;
            }
            else
            {
                this.dataClient.OnResponseReceived += this.OnDataDictationResponseReceivedHandler;
            }

            this.dataClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            this.dataClient.OnConversationError += this.OnConversationErrorHandler;
        }

        /// <summary>
        /// Creates a data client with LUIS intent support.
        /// Speech recognition with data (for example from a file or audio source).  
        /// The data is broken up into buffers and each buffer is sent to the Speech Recognition Service.
        /// No modification is done to the buffers, so the user can apply their
        /// own Silence Detection if desired.
        /// </summary>
        private void CreateDataRecoClientWithIntent()
        {
            this.dataClient = SpeechRecognitionServiceFactory.CreateDataClientWithIntentUsingEndpointUrl(
                this.DefaultLocale,
                this.SubscriptionKey,
                this.LuisEndpointUrl);
            this.dataClient.AuthenticationUri = this.AuthenticationUri;

            // Event handlers for speech recognition results
            this.dataClient.OnResponseReceived += this.OnDataShortPhraseResponseReceivedHandler;
            this.dataClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            this.dataClient.OnConversationError += this.OnConversationErrorHandler;

            // Event handler for intent result
            this.dataClient.OnIntent += this.OnIntentHandler;
        }

        /// <summary>
        /// Sends the audio helper.
        /// </summary>
        /// <param name="wavFileName">Name of the wav file.</param>
        private void SendAudioHelper(string wavFileName)
        {
            using (FileStream fileStream = new FileStream(wavFileName, FileMode.Open, FileAccess.Read))
            {
                // Note for wave files, we can just send data from the file right to the server.
                // In the case you are not an audio file in wave format, and instead you have just
                // raw data (for example audio coming over bluetooth), then before sending up any 
                // audio data, you must first send up an SpeechAudioFormat descriptor to describe 
                // the layout and format of your raw audio data via DataRecognitionClient's sendAudioFormat() method.
                int bytesRead = 0;
                byte[] buffer = new byte[1024];

                try
                {
                    do
                    {
                        // Get more Audio data to send into byte buffer.
                        bytesRead = fileStream.Read(buffer, 0, buffer.Length);

                        // Send of audio data to service. 
                        this.dataClient.SendAudio(buffer, bytesRead);
                    }
                    while (bytesRead > 0);
                }
                finally
                {
                    // We are done sending audio.  Final recognition results will arrive in OnResponseReceived event call.
                    this.dataClient.EndAudio();
                }
            }
        }

        /// <summary>
        /// Called when a final response is received;
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnMicShortPhraseResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            Dispatcher.Invoke((System.Action)(() =>
            {
                this.WriteLine("--- OnMicShortPhraseResponseReceivedHandler ---");

                // we got the final result, so it we can end the mic reco.  No need to do this
                // for dataReco, since we already called endAudio() on it as soon as we were done
                // sending all the data.
                this.micClient.EndMicAndRecognition();

                this.WriteResponseResult(e);
            if (e.PhraseResponse.Results.Length > 0)
            {

                //if (e.PhraseResponse.Results[0].LexicalForm.StartsWith("launch ") && e.PhraseResponse.RecognitionStatus == RecognitionStatus.RecognitionSuccess)
                //{
                //    var name = e.PhraseResponse.Results[0].LexicalForm;
                //    name = name.Replace("launch ", "");
                //    var commandline = GetCommandLineFromLauncherName(name);
                //    if (commandline != null && commandline.Length > 0)
                //    {
                //        try
                //        {
                //            Process.Start(commandline);
                //        }
                //        catch (Exception exception)
                //        {
                //            this.WriteLine($"*************Failed to Launch {commandline} *****************");
                //            this.WriteLine(exception.Message);
                //        }
                //        this.WriteLine($"*************Launching  {commandline} *****************");

                //    }
                //}



                //else if (globalIntelliSenseOn == true && languageMatched != null && categoryMatched != null)
                //{
                //    PerformGlobalIntelliSense(e);
                //}
                //else
                //{
                //    MatchLanguageAndCategory(e);
                //}
                }
                _startButton.IsEnabled = true;
                _radioGroup.IsEnabled = true;
            }));
        }

        private void PerformGlobalIntelliSense(SpeechRecognizedEventArgs e)
        {
            var phrase = e.Result.Text.ToLower();
            this.WriteLine($"Heard Phrase: {phrase}");

            if (languageMatched=="intellisense")
            {
                languageMatched = "not applicable";
            }
            using (var db=new MyDatabase())                             
            {
                var globalIntellisense = db.tblCustomIntelliSenses.Where(c => c.tblLanguage.Language == languageMatched && c.tblCategory.Category == categoryMatched && c.Display_Value == phrase).FirstOrDefault();
                if (globalIntellisense!= null )
                {
                    ActivateApp(currentProcess.MainWindowHandle);
                    if (globalIntellisense.DeliveryType=="Copy and Paste")
                    {
                        System.Windows.Clipboard.SetText( globalIntellisense.SendKeys_Value);
                        this.WriteLine($"Copied to Clipboard: {globalIntellisense.Display_Value}");
                        List<string> keys = new List<string>(new string[] { "^v" });
                        SendKeysCustom("", "", keys, currentProcess.ProcessName);
                    }
                    else
                    {
                        List<string> keys = new List<string>(new string[] { globalIntellisense.SendKeys_Value });
                        SendKeysCustom("", "", keys, currentProcess.ProcessName);
                    }
                    this.WriteLine($"Send Keys Value: {globalIntellisense.SendKeys_Value}");
                }

            }

        }

        private void MatchLanguageAndCategory(string phrase )
        {
            phrase = phrase.ToLower().Replace("intellisense", "not applicable");
            using (var db = new MyDatabase())
            {
                var languages = db.tblLanguages.OrderBy(l => l.Language).ToList();
                foreach (var language in languages)
                {
                    if (phrase.StartsWith(language.Language.ToLower()))
                    {
                        languageMatched = language.Language;
                        if (languageMatched == "intellisense")
                        {
                            languageMatched = "not applicable";
                        }
                        var categories = db.tblCategories.Where(c => c.Category_Type == "IntelliSense Command").ToList();
                        foreach (var category in categories)
                        {
                            var temporary = phrase.Replace(language.Language.ToLower(), "").Trim();
                            if (category.Category.ToLower().EndsWith( temporary) )
                            {
                                categoryMatched = category.Category;
                                languageAndCategoryAlreadyMatched = true;
                                var commands = db.tblCustomIntelliSenses.Where(i => i.Language_ID == language.ID && i.Category_ID == category.MenuNumber).OrderBy(s => s.Display_Value).ToList();
                                speechRecognitionEngine.UnloadAllGrammars();
                                Choices choices = new Choices();
                                Dispatcher.Invoke(() =>
                                {
                                    Commands.Text = "";
                                });
                                foreach (var command in commands)
                                {
                                    this.WriteCommandLine(command.Display_Value);
                                    //Load a new grammar here
                                    choices.Add(command.Display_Value);
                                }

                                choices.Add("Stop IntelliSense");
                                this.WriteCommandLine("Stop IntelliSense");
                                choices.Add("Global IntelliSense");
                                this.WriteCommandLine("Global IntelliSense");
                                choices.Add("Go Back");
                                this.WriteCommandLine("Go Back");
                                this.WriteCommandLine("Go Dormant");
                                choices.Add("Go Dormant");
                                Grammar grammar = new Grammar(new GrammarBuilder(choices));
                                speechRecognitionEngine.LoadGrammar(grammar);
                                this.WriteLine($"Language and Category: {languageMatched} {categoryMatched}");
                                return;
                            }
                        }
                    }
                }

            }

        }

        private void ListCommands()
        {
            Choices choices = new Choices();
            speechRecognitionEngine.UnloadAllGrammars();
            Dispatcher.Invoke(() =>
            {
                Commands.Text = "";
            });

            this.WriteCommandLine($"Quit Application");
            choices.Add($"quit application");
            AddApplicationsToKillToChoices(choices);
            //this.WriteCommandLine($"Launch <application name>");
            this.WriteCommandLine($"Toggle Microphone");
            choices.Add("Toggle Microphone");
            this.WriteCommandLine($"Restart Dragon");
            choices.Add("Restart Dragon");
            //this.WriteCommandLine($"With Intent Mode");
            //choices.Add("With Intent Mode");
            this.WriteCommandLine($"Short Phrase Mode");
            choices.Add("Short Phrase Mode");
            //this.WriteCommandLine($"long dictation mode");
            //choices.Add("Long Dictation Mode");
            this.WriteCommandLine($"Global IntelliSense");
            choices.Add("Global IntelliSense");
            this.WriteCommandLine($"Visual Basic IntelliSense");
            choices.Add("Visual Basic IntelliSense");
            this.WriteCommandLine($"C Sharp IntelliSense");
            choices.Add("C Sharp IntelliSense");
            this.WriteCommandLine($"JavaScript IntelliSense");
            choices.Add("JavaScript IntelliSense");
            this.WriteCommandLine($"Keyboard");
            choices.Add("Keyboard");
            this.WriteCommandLine("Go Dormant");
            choices.Add("Go Dormant");
            this.WriteCommandLine($"List Commands");
            choices.Add("List Commands");
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognitionEngine.LoadGrammarAsync(grammar);
            dispatcherTimer.Stop();
        }

        private void AddApplicationsToKillToChoices(Choices choices)
        {
            using (var db = new MyDatabase())
            {
                var applicationsToKill = db.ApplicationsToKill.Where(k => k.Display==true).OrderBy(k => k.CommandName).ToList();
                foreach (var application in applicationsToKill)
                {
                    choices.Add($"kill {application.CommandName}");
                    this.WriteCommandLine($"kill {application.CommandName}");
                }
            }
        }

        private void KillAllProcesses(string name)
        {
            var processName = GetProcessNameFromApplicationsToKill(name);
            if (processName.Length > 0)
            {
                foreach (var process in Process.GetProcessesByName(processName))
                {
                    try
                    {
                        process.Kill();
                    }
                    catch (Exception exception)
                    {
                        System.Windows.MessageBox.Show(exception.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Called when a final response is received;
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnDataShortPhraseResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            Dispatcher.Invoke((System.Action)(() =>
            {
                this.WriteLine("--- OnDataShortPhraseResponseReceivedHandler ---");

                // we got the final result, so it we can end the mic reco.  No need to do this
                // for dataReco, since we already called endAudio() on it as soon as we were done
                // sending all the data.
                this.WriteResponseResult(e);
                _startButton.IsEnabled = true;
                _radioGroup.IsEnabled = true;
            }));
        }

        /// <summary>
        /// Writes the response result.
        /// </summary>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void WriteResponseResult(SpeechResponseEventArgs e)
        {
            if (e.PhraseResponse.Results.Length == 0)
            {
                this.WriteLine("No phrase response is available.");
                //dispatcherTimer.Stop();
                lastResult = null;
            }
            else
            {
                if (!e.PhraseResponse.Results[0].LexicalForm.StartsWith("choose ") && e.PhraseResponse.Results[0].LexicalForm!="transfer to notepad")
                {
                    Dispatcher.Invoke(() => {
                        var position = Commands.Text.IndexOf("********* Final n-BEST Results *********");
                        if (position>0)
                        {
                            Commands.Text = Commands.Text.Substring(0, (position ));
                        }
                    });
                    lastResult = e;
                    this.WriteCommandLine("********* Final n-BEST Results *********");
                    for (int i = 0; i < e.PhraseResponse.Results.Length; i++)
                    {
                        this.WriteCommandLine(
                            "[Choose {0}] Confidence={1}, Text=\"{2}\"", 
                            i, 
                            e.PhraseResponse.Results[i].Confidence,
                            e.PhraseResponse.Results[i].DisplayText);
                    }
                if (e.PhraseResponse.Results[0].Confidence!=Confidence.Low)
                {
                    Dispatcher.Invoke(() =>
                    {
                            finalResult.Text = (e.PhraseResponse.Results[0].DisplayText);
                    });
                        //List<string> keys = new List<string>(new string[] { e.PhraseResponse.Results[0].DisplayText });
                        //SendKeysCustom(null, null, keys, currentProcess.ProcessName);
                    }
                }
            }
        }

        /// <summary>
        /// Called when a final response is received;
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnMicDictationResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            this.WriteLine("--- OnMicDictationResponseReceivedHandler ---");
            if (e.PhraseResponse.RecognitionStatus == RecognitionStatus.EndOfDictation ||
                e.PhraseResponse.RecognitionStatus == RecognitionStatus.DictationEndSilenceTimeout)
            { 
                if (e.PhraseResponse.Results[0].LexicalForm == "transfer to notepad" && e.PhraseResponse.RecognitionStatus == RecognitionStatus.RecognitionSuccess)
                {
                    List<string> keys = new List<string>(new string[] { Environment.NewLine + this.finalResult.Text + Environment.NewLine });
                    SendKeysCustom(null, "Untitled - Notepad", keys, "Notepad.exe");
                }
                Dispatcher.Invoke(
                    (System.Action)(() => 
                    {
                        // we got the final result, so it we can end the mic reco.  No need to do this
                        // for dataReco, since we already called endAudio() on it as soon as we were done
                        // sending all the data.
                        this.micClient.EndMicAndRecognition();
                        this._startButton.IsEnabled = true;
                        this._radioGroup.IsEnabled = true;
                    }));                
            }

            this.WriteResponseResult(e);
        }

        /// <summary>
        /// Called when a final response is received;
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnDataDictationResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            this.WriteLine("--- OnDataDictationResponseReceivedHandler ---");
            if (e.PhraseResponse.RecognitionStatus == RecognitionStatus.EndOfDictation ||
                e.PhraseResponse.RecognitionStatus == RecognitionStatus.DictationEndSilenceTimeout)
            {
                Dispatcher.Invoke(
                    (System.Action)(() => 
                    {
                        _startButton.IsEnabled = true;
                        _radioGroup.IsEnabled = true;

                        // we got the final result, so it we can end the mic reco.  No need to do this
                        // for dataReco, since we already called endAudio() on it as soon as we were done
                        // sending all the data.
                    }));
            }

            this.WriteResponseResult(e);
        }

        /// <summary>
        /// Called when a final response is received and its intent is parsed
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechIntentEventArgs"/> instance containing the event data.</param>
        private void OnIntentHandler(object sender, SpeechIntentEventArgs e)
        {
            this.WriteLine("--- Intent received by OnIntentHandler() ---");

            LuisResult luisResult = new LuisResult();
            JsonConvert.DefaultSettings = () => new JsonSerializerSettings()
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver(),
                Formatting = Newtonsoft.Json.Formatting.Indented,
                NullValueHandling = NullValueHandling.Ignore,
            };

            var _Data = JsonConvert.DeserializeObject<MyLUIS>(e.Payload);
            //luisResult = JsonConvert.DeserializeObject<LuisResult>(e.Payload);

            if (_Data.intents.Count()>0)
            {
                this.WriteLine($"Intent: {_Data.intents[0].intent}");
                if (_Data.entities.Count()>1)
                {
                    this.WriteLine($"Entities:{_Data.entities[0].entity} {_Data.entities[0].type} {_Data.entities[1].entity} {_Data.entities[1].type}  query: {_Data.query}");
                }
            }
            this.WriteLine();
        }



        public class MyLUIS
        {
            public string query { get; set; }
            public lIntent[] intents { get; set; }
            public lEntity[] entities { get; set; }
        }

        public class lIntent
        {
            public string intent { get; set; }
            public float score { get; set; }
        }

        public class lEntity
        {
            public string entity { get; set; }
            public string type { get; set; }
            public int startIndex { get; set; }
            public int endIndex { get; set; }
            public float score { get; set; }
        }

        /// <summary>
        /// Called when a partial response is received.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="PartialSpeechResponseEventArgs"/> instance containing the event data.</param>
        private void OnPartialResponseReceivedHandler(object sender, PartialSpeechResponseEventArgs e)
        {
            //this.WriteLine("--- Partial result received by OnPartialResponseReceivedHandler() ---");
            //this.WriteLine("{0}", e.PartialResult);
            Dispatcher.Invoke(() =>
            {
                Result.Text = (e.PartialResult);
            });
            //this.WriteLine();
        }

        /// <summary>
        /// Called when an error is received.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="SpeechErrorEventArgs"/> instance containing the event data.</param>
        private void OnConversationErrorHandler(object sender, SpeechErrorEventArgs e)
        {
           Dispatcher.Invoke(() =>
           {
               _startButton.IsEnabled = true;
               _radioGroup.IsEnabled = true;
           });

            this.WriteLine("--- Error received by OnConversationErrorHandler() ---");
            this.WriteLine("Error code: {0}", e.SpeechErrorCode.ToString());
            this.WriteLine("Error text: {0}", e.SpeechErrorText);
            this.WriteLine();
        }

        /// <summary>
        /// Called when the microphone status has changed.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="MicrophoneEventArgs"/> instance containing the event data.</param>
        private void OnMicrophoneStatus(object sender, MicrophoneEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                WriteLine("--- Microphone status change received by OnMicrophoneStatus() ---");
                WriteLine("********* Microphone status: {0} *********", e.Recording);
                if (e.Recording)
                {
                    WriteLine("Please start speaking.");
                }

                WriteLine();
            });
        }

        /// <summary>
        /// Writes the line.
        /// </summary>
        private void WriteLine()
        {
            this.WriteLine(string.Empty);
        }

        /// <summary>
        /// Writes the line.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <param name="args">The arguments.</param>
        private void WriteLine(string format, params object[] args)
        {
            var formattedStr = "";
            if (format.Contains("{") || format.Contains("}"))
            {
                formattedStr = format;
            }
            else
            {
                formattedStr = string.Format(format, args);
            }
            Trace.WriteLine(formattedStr);
            Dispatcher.Invoke(() =>
            {
                _logText.Text += (formattedStr + "\n");
                _logText.ScrollToEnd();
            });
        }

        private void WriteCommandLine(string format, params object[] args)
        {
            var formattedStr = string.Format(format, args);
            Trace.WriteLine(formattedStr);
            Dispatcher.Invoke(() =>
            {
                Commands.Text += (formattedStr + "\n");
                Commands.ScrollToEnd();
            });
        }

        /// <summary>
        /// Gets the subscription key from isolated storage.
        /// </summary>
        /// <returns>The subscription key.</returns>
        private string GetSubscriptionKeyFromIsolatedStorage()
        {
            string subscriptionKey = null;

            using (IsolatedStorageFile isoStore = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Assembly, null, null))
            {
                try
                {
                    using (var iStream = new IsolatedStorageFileStream(IsolatedStorageSubscriptionKeyFileName, FileMode.Open, isoStore))
                    {
                        using (var reader = new StreamReader(iStream))
                        {
                            subscriptionKey = reader.ReadLine();
                        }
                    }
                }
                catch (FileNotFoundException)
                {
                    subscriptionKey = null;
                }
            }

            if (string.IsNullOrEmpty(subscriptionKey))
            {
                subscriptionKey = DefaultSubscriptionKeyPromptMessage;
            }

            return subscriptionKey;
        }

        /// <summary>
        /// Handles the Click event of the subscription key save button.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void SaveKey_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveSubscriptionKeyToIsolatedStorage(this.SubscriptionKey);
                System.Windows.MessageBox.Show("Subscription key is saved in your disk.\nYou do not need to paste the key next time.", "Subscription Key");
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show(
                    "Fail to save subscription key. Error message: " + exception.Message,
                    "Subscription Key", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Handles the Click event of the DeleteKey control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void DeleteKey_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.SubscriptionKey = DefaultSubscriptionKeyPromptMessage;
                SaveSubscriptionKeyToIsolatedStorage(string.Empty);
                System.Windows.MessageBox.Show("Subscription key is deleted from your disk.", "Subscription Key");
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show(
                    "Fail to delete subscription key. Error message: " + exception.Message,
                    "Subscription Key", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Helper function for INotifyPropertyChanged interface 
        /// </summary>
        /// <typeparam name="T">Property type</typeparam>
        /// <param name="caller">Property name</param>
        private void OnPropertyChanged<T>([CallerMemberName]string caller = null)
        {
            var handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(caller));
            }
        }

        /// <summary>
        /// Handles the Click event of the RadioButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void RadioButton_Click(object sender, RoutedEventArgs e)
        {
            // Reset everything
            ResetEverything();
        }

        private void ResetEverything()
        {
            if (this.micClient != null)
            {
                this.micClient.EndMicAndRecognition();
                this.micClient.Dispose();
                this.micClient = null;
            }

            if (this.dataClient != null)
            {
                this.dataClient.Dispose();
                this.dataClient = null;
            }

            this._logText.Text = string.Empty;
            this._startButton.IsEnabled = true;
            this._radioGroup.IsEnabled = true;
        }

        public string GetCommandLineFromLauncherName(string name)
        {
            using (var db=new MyDatabase())
            {
                var launcher = db.tblLaunchers.Where(l => l.Name.ToLower() == name).FirstOrDefault();

                if (launcher!= null )
                {
                    return launcher.CommandLine; 
                }
                return "";
            }
        }

        public string GetProcessNameFromApplicationsToKill(string name)
        {
            using (var db = new MyDatabase())
            {
                var applicationToKill = db.ApplicationsToKill.Where(k => k.CommandName.ToLower()==name).FirstOrDefault();
                if (applicationToKill!= null )
                {
                    return applicationToKill.ProcessName;
                }
                return "";
            }
        }

        public IEnumerable<CustomIntelliSense> GetCustomIntelliSenses(string language, string category)
        {
            using (var db=new MyDatabase())
            {
                var languageId = db.tblLanguages.Where(l => l.Language == language).FirstOrDefault()?.ID;
                var categoryId = db.tblCategories.Where(c => c.Category == category).FirstOrDefault()?.MenuNumber;
                IEnumerable<CustomIntelliSense> customIntelliSenses = db.tblCustomIntelliSenses.Where(c => c.Language_ID == languageId && c.Category_ID == categoryId);
                if (customIntelliSenses!= null )
                {
                    return customIntelliSenses.ToList();
                }
                return null;
            }
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            if (this._startButton.IsEnabled == true)
            {
                ButtonAutomationPeer peer = new ButtonAutomationPeer(this._startButton);
                IInvokeProvider invokeProvider = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                invokeProvider.Invoke();
            }
        }
        void ActivateApp(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            //Activate the first application refined with this name
            if (processes.Count()>0)
            {
                SetForegroundWindow(processes[0].MainWindowHandle);
            }
        }
        void ActivateApp(IntPtr windowHandle)
        {
            SetForegroundWindow(windowHandle);
        }
    }
}
