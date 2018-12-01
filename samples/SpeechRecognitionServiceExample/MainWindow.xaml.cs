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
    using System.Globalization;

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
        //Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        private SpeechRecognitionEngine  speechRecognitionEngine= new SpeechRecognitionEngine();
        private SpeechRecognizer speechRecognizer;
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
        private bool filterInProgress;
        private string lastCommand;
        private SpeechRecognizedEventArgs lastRecognition;
        private bool launcherCategoryMatched;
        private string lastLauncherCategory;
        private bool completed;
        

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
        {
            if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                System.Windows.Application.Current.Shutdown();
                return;
            }
            StartWindowsSpeechRecognition();
            if (speechRecognizer!= null )
            {
                speechRecognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognitionEngine_SpeechRecognized);
                this.InitializeComponent();
                this.Initialize();
                return;
            }
            System.Windows.Application.Current.Shutdown();
            return;
        }
        private void LoadGrammarLauncher(bool showCommands= true )
        {
            Choices choices = new Choices();
            using (var db = new MyDatabase())
            {
                List<tblCategory> categories = null;
                categories = db.tblCategories.Where(c => c.Category_Type == "Launch Applications").OrderBy(c => c.Category).ToList();
                var computerId = db.Computers.Where(c => c.ComputerName == Environment.MachineName).FirstOrDefault()?.ID;
                foreach (var category in categories)
                {
                    var count = db.tblLaunchers.Where(l => l.Menu == category.MenuNumber && (l.ComputerID== null  ||  l.ComputerID==computerId)).Count();
                    choices.Add($"Launcher {category.Category}");
                    if (showCommands==true)
                    {
                        this.WriteCommandLine($"Launcher {category.Category} ({count})");
                    }
                }
                if (showCommands==true)
                {
                    this.WriteLine($"Launcher grammars loaded...");
                    choices.Add("Stop Launcher");
                    this.WriteCommandLine("Stop Launcher");
                    this.WriteCommandLine("Go Dormant");
                    choices.Add("Go Dormant");
                    this.WriteCommandLine("List Commands");
                    choices.Add("List Commands");
                }
            }
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognizer.LoadGrammarAsync(grammar);
        }
        private void LoadGrammarCustomIntellisense(string specificLanguage,bool showCommands=true,bool useEngine=false)
        {
            Choices choices = new Choices();
            using (var db = new MyDatabase())
            {
                List<tblLanguage> languages = null;
                if (specificLanguage != null)
                {
                    languages = db.tblLanguages.Where(l => l.Language == specificLanguage).OrderBy(l => l.Language).ToList();
                }
                else
                {
                    languages = db.tblLanguages.Where(l => l.tblCustomIntelliSenses.Count > 0 && l.Active == true).OrderBy(l => l.Language).ToList();
                }
                var computerId = db.Computers.Where(c => c.ComputerName == Environment.MachineName).FirstOrDefault()?.ID;
                foreach (var language in languages)
                {
                    List<tblCategory> categories1 = db.tblCategories.OrderBy(c => c.Category).Where(c => c.tblCustomIntelliSenses.Count > 0 && c.Category_Type == "IntelliSense Command" && c.tblCustomIntelliSenses.Where(s => s.Language_ID == language.ID && (s.ComputerID == null || s.ComputerID == computerId)).Count() > 0).ToList();
                    foreach (var category in categories1)
                    {
                        var tempLanguage = language.Language;
                        if (tempLanguage == "Not Applicable")
                        {
                            tempLanguage = "Intellisense";
                        }
                        var count = db.tblCustomIntelliSenses.Where(s => s.Category_ID == category.MenuNumber && s.Language_ID == language.ID && (s.ComputerID == null || s.ComputerID == computerId)).Count();
                        if (category.Category=="jQuery")
                        {
                            choices.Add($"{tempLanguage} {"Jay Query"}");
                        }
                        else
                        {
                            choices.Add($"{tempLanguage} {category.Category}");
                        }
                        if (showCommands==true)
                        {
                            this.WriteCommandLine($"{tempLanguage} {category.Category} ({count})");
                        }
                    }
                }
                if (showCommands==true)
                {
                    this.WriteLine($"IntelliSense grammars loaded...");
                    this.WriteCommandLine("Stop IntelliSense");
                    this.WriteCommandLine("Go Dormant");
                    choices.Add("Go Dormant");
                    choices.Add("Stop IntelliSense");
                    BuildPhoneticAlphabetGrammars(useEngine);
                }
            }
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            if (useEngine==true)
            {
                speechRecognitionEngine.LoadGrammarAsync(grammar);
            }
            else
            {
                speechRecognizer.LoadGrammarAsync(grammar);
            }
        }
  
        private void LoadGrammarKeyboard(bool showCommands= true, bool useEngine= false)
        {
            //dispatcherTimer.Stop();
            List<string> phoneticAlphabet = new List<string> { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet", "Kilo", "Lima", "Mike", "November", "Oscar", "Papa", "Qubec", "Romeo", "Sierra", "Tango", "Uniform", "Victor", "Whiskey", "X-ray", "Yankee", "Zulu" };
            Choices choices = new Choices();
            if (showCommands==true)
            {
                this.WriteCommandLine($"Keyboard Commands:");
                this.WriteCommandLine("Press Alpha-Zulu");
                this.WriteCommandLine("Press 1-9");
                this.WriteCommandLine("Press Control Alpha-Zulu");
                this.WriteCommandLine("Press Alt Alpha-Zulu");
                this.WriteCommandLine("Press Shift Alpha-Zulu");
                this.WriteCommandLine("Press Zero");
            }
            foreach (var item in phoneticAlphabet)
            {
                choices.Add($"Press {item}");
            }
            choices.Add("Press Zero");
            for (int i = 1; i < 10; i++)
            {
                choices.Add($"Press {i}");
            }
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
            if (showCommands==true)
            {

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
                AddChoiceAndWriteCommandline(choices, "Press Semicolon");
                AddChoiceAndWriteCommandline(choices, "Press Colon");
                AddChoiceAndWriteCommandline(choices, "Press Percent");
                AddChoiceAndWriteCommandline(choices, "Press Ampersand");
                AddChoiceAndWriteCommandline(choices, "Press Dollar");
                AddChoiceAndWriteCommandline(choices, "Press Exclamation Mark");
                AddChoiceAndWriteCommandline(choices, "Press Double Quote");
                AddChoiceAndWriteCommandline(choices, "Press Pound");
                AddChoiceAndWriteCommandline(choices, "Press Asterix");
                AddChoiceAndWriteCommandline(choices, "Press Open Bracket");
                AddChoiceAndWriteCommandline(choices, "Press Close Bracket");
                AddChoiceAndWriteCommandline(choices, "Press Plus");
                AddChoiceAndWriteCommandline(choices, "Press Stop");
                AddChoiceAndWriteCommandline(choices, "Press Equal");
                AddChoiceAndWriteCommandline(choices, "Press Apostrophe");
            }
            else
            {
                choices.Add("Press Down");
                choices.Add( "Press Up");
                choices.Add( "Press Left");
                choices.Add( "Press Right");
                choices.Add( "Press Enter");
                choices.Add( "Press Backspace");
                choices.Add( "Press Alt Tab");
                choices.Add( "Press Tab");
                choices.Add( "Press Shift Tab");
                choices.Add( "Press Escape");
                choices.Add( "Press Delete");
                choices.Add( "Press Space");
                choices.Add( "Press Alt Space");
                choices.Add( "Press Home");
                choices.Add( "Press Control Home");
                choices.Add( "Press Control Shift Home");
                choices.Add( "Press End");
                choices.Add( "Press Control End");
                choices.Add( "Press Control Shift End");
                choices.Add( "Press Page Down");
                choices.Add( "Press Page Up");
                choices.Add( "Press Hyphen");
                choices.Add( "Press _");
                choices.Add( "Press Semicolon");
                choices.Add( "Press Colon");
                choices.Add( "Press Percent");
                choices.Add( "Press Ampersand");
                choices.Add( "Press Dollar");
                choices.Add( "Press Exclamation Mark");
                choices.Add( "Press Double Quote");
                choices.Add( "Press Pound");
                choices.Add( "Press Asterix");
                choices.Add( "Press Open Bracket");
                choices.Add( "Press Close Bracket");
                choices.Add( "Press Plus");
                choices.Add( "Press Equal");
                choices.Add( "Press Apostrophe");
            }

            choices.Add("perform replacement");
            this.WriteCommandLine("Perform Replacement");
            if (useEngine==true)
            {
                BuildReplaceLettersGrammars(speechRecognitionEngine, "Replace");
                BuildReplaceLettersGrammars(speechRecognitionEngine, "With this");
            }
            else
            {
                BuildReplaceLettersGrammars(speechRecognizer, "Replace");
                BuildReplaceLettersGrammars(speechRecognizer, "With this");
            }
            

            choices.Add("Transfer to Application");
            this.WriteCommandLine("Transfer to Application");

            if (showCommands==true)
            {
                this.WriteCommandLine("Press Function 1-12");
            }
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

            if (showCommands==true)
            {
                BuildPhoneticAlphabetGrammars(useEngine);
                choices.Add("Stop Keyboard");
                this.WriteCommandLine("Stop Keyboard");
                this.WriteCommandLine("Go Dormant");
                choices.Add("Go Dormant");
            }
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            if (useEngine==true)
            {
                speechRecognizer.LoadGrammarAsync(grammarDirections);
                speechRecognizer.LoadGrammarAsync(grammar);
            }
            else
            {
                speechRecognizer.LoadGrammarAsync(grammarDirections);
                speechRecognizer.LoadGrammarAsync(grammar);
            }

            SetUpSymbolGrammarCommands(useEngine);

            UpdateCurrentProcess();
        }


        private void BuildReplaceLettersGrammars(SpeechRecognizer speechRecognizer,string replaceType)
        {
            Choices phoneticAlphabet = new Choices(new string[] { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet","Kilo","Lima","Mike","November","Oscar","Papa","Qubec","Romeo","Sierra","Tango","Uniform","Victor","Whiskey","X-ray","Yankee","Zulu" });
            GrammarBuilder grammarBuilder2 = new GrammarBuilder();
            grammarBuilder2.Append(phoneticAlphabet);
            grammarBuilder2.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder3 = new GrammarBuilder();
            grammarBuilder3.Append(phoneticAlphabet);
            grammarBuilder3.Append(phoneticAlphabet);
            grammarBuilder3.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder4 = new GrammarBuilder();
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder5 = new GrammarBuilder();
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder6 = new GrammarBuilder();
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder7 = new GrammarBuilder();
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            Choices phoneticAlphabet2to7 = new Choices(new GrammarBuilder[] { grammarBuilder2, grammarBuilder3, grammarBuilder4,grammarBuilder5,grammarBuilder6,grammarBuilder7 });
            Grammar grammarPhoneticAlphabets = new Grammar((GrammarBuilder)phoneticAlphabet2to7);
            grammarPhoneticAlphabets.Name = "Phonetic Alphabet";
            speechRecognizer.LoadGrammarAsync(grammarPhoneticAlphabets);

            Choices choices = new Choices("Lower");
            GrammarBuilder grammarBuilderLower2 = new GrammarBuilder();
            grammarBuilderLower2.Append(choices);
            grammarBuilderLower2.Append(phoneticAlphabet);
            grammarBuilderLower2.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower3 = new GrammarBuilder();
            grammarBuilderLower3.Append(choices);
            grammarBuilderLower3.Append(phoneticAlphabet);
            grammarBuilderLower3.Append(phoneticAlphabet);
            grammarBuilderLower3.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower4 = new GrammarBuilder();
            grammarBuilderLower4.Append(choices);
            grammarBuilderLower4.Append(phoneticAlphabet);
            grammarBuilderLower4.Append(phoneticAlphabet);
            grammarBuilderLower4.Append(phoneticAlphabet);
            grammarBuilderLower4.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower5 = new GrammarBuilder();
            grammarBuilderLower5.Append(choices);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower6 = new GrammarBuilder();
            grammarBuilderLower6.Append(choices);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower7 = new GrammarBuilder();
            grammarBuilderLower7.Append(choices);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            Choices phoneticAlphabetLower2to7 = new Choices(new GrammarBuilder[] { grammarBuilderLower2, grammarBuilderLower3, grammarBuilderLower4, grammarBuilderLower5, grammarBuilderLower6, grammarBuilderLower7 });
            Grammar grammarPhoneticAlphabetsLower = new Grammar((GrammarBuilder)phoneticAlphabetLower2to7);
            grammarPhoneticAlphabetsLower.Name = "Phonetic Alphabet Lower";
            speechRecognizer.LoadGrammarAsync(grammarPhoneticAlphabetsLower);
            this.WriteCommandLine("<Alpha-Zulu 2-7>");
        }

        private void BuildReplaceLettersGrammars(SpeechRecognitionEngine speechRecognitionEngine,string replaceText)
        {
            Choices phoneticAlphabet = new Choices(new string[] { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet", "Kilo", "Lima", "Mike", "November", "Oscar", "Papa", "Qubec", "Romeo", "Sierra", "Tango", "Uniform", "Victor", "Whiskey", "X-ray", "Yankee", "Zulu", "Upper Alpha", "Upper Bravo", "Upper Charlie", "Upper Delta", "Upper Echo", "Upper Foxtrot", "Upper Golf", "Upper Hotel", "Upper India", "Upper Juliet", "Upper Kilo", "Upper Lima", "Upper Mike", "Upper November", "Upper Oscar", "Upper Papa", "Upper Qubec", "Upper Romeo", "Upper Sierra", "Upper Tango", "Upper Uniform", "Upper Victor", "Upper Whiskey", "Upper X-ray", "Upper Yankee", "Upper Zulu" });
            GrammarBuilder grammarBuilder1 = new GrammarBuilder();
            grammarBuilder1.Append(replaceText);
            grammarBuilder1.Append(phoneticAlphabet);

            GrammarBuilder grammarBuilder2 = new GrammarBuilder();
            grammarBuilder2.Append(replaceText);
            grammarBuilder2.Append(phoneticAlphabet);
            grammarBuilder2.Append(phoneticAlphabet);

            GrammarBuilder grammarBuilder3 = new GrammarBuilder();
            grammarBuilder3.Append(replaceText);
            grammarBuilder3.Append(phoneticAlphabet);
            grammarBuilder3.Append(phoneticAlphabet);
            grammarBuilder3.Append(phoneticAlphabet);

            GrammarBuilder grammarBuilder4 = new GrammarBuilder();
            grammarBuilder4.Append(replaceText);
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);

            GrammarBuilder grammarBuilder5 = new GrammarBuilder();
            grammarBuilder5.Append(replaceText);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder6 = new GrammarBuilder();
            grammarBuilder6.Append(replaceText);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder7 = new GrammarBuilder();
            grammarBuilder7.Append(replaceText);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            Choices phoneticAlphabet1to7 = new Choices(new GrammarBuilder[] { grammarBuilder1, grammarBuilder2, grammarBuilder3, grammarBuilder4, grammarBuilder5, grammarBuilder6, grammarBuilder7 });
            Grammar grammarPhoneticAlphabets = new Grammar((GrammarBuilder)phoneticAlphabet1to7);
            grammarPhoneticAlphabets.Name = "Replace Letters";
            speechRecognitionEngine.LoadGrammarAsync(grammarPhoneticAlphabets);
            this.WriteCommandLine($"<{replaceText} Alpha-Zulu 1-7>");
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
            try
            {
                SpeechUI.SendTextFeedback(e.Result, $"Recognised: {e.Result.Text}", true);
            }
            catch (Exception)
            {
                //This will fail if were using the engine
            }
            this.WriteLine($"Recognised: {e.Result.Text}");
            this.WriteLine($"Confidence: {e.Result.Confidence}");
            if (e.Result.Text.ToLower() == "keyboard" && e.Result.Confidence > 0.5)
            {
                this.WriteLine($"Keyboard mode...");
                isKeyboard = true;
                UnloadGrammarsAndClearCommands();
                LoadGrammarKeyboard();
            }
            else if (e.Result.Grammar.Name=="Mouse Command" && e.Result.Confidence>0.4)
            {
                PerformMouseCommand(e);
            }
            else if (e.Result.Grammar.Name == "Symbols" && e.Result.Confidence > 0.3)
            {
                PerformanceSymbolsCommand(e);
            }
            else if (e.Result.Grammar.Name.ToLower() == "camel case" && e.Result.Confidence > 0.3 )
            {
                CamelCaseCommand(e);
            }
            else if (e.Result.Grammar.Name.ToLower() == "variable" && e.Result.Confidence > 0.3)
            {
                VariableCommand(e);
            }
            else if (e.Result.Text.ToLower().StartsWith("replace") && e.Result.Grammar.Name == "Replace Letters" && e.Result.Confidence > 0.8)
            {
                SetupReplaceText(e, "Replace");
            }
            else if (e.Result.Text.ToLower().StartsWith("with this") && e.Result.Grammar.Name == "Replace Letters" && e.Result.Confidence > 0.8)
            {
                SetupReplaceText(e, "With this");
            }
            else if (e.Result.Text.ToLower() == "perform replacement" && e.Result.Confidence > 0.8)
            {
                PerformReplacement();
            }
            else if (e.Result.Text.ToLower() == "windows speech recognition")
            {
                StartWindowsSpeechRecognition();
            }
            else if (e.Result.Text.ToLower() == "launch applications" && e.Result.Confidence > 0.5 || e.Result.Text.ToLower() == "go back" && launcherCategoryMatched == true)
            {
                this.WriteLine($"Launch Applications Mode… ");
                UnloadGrammarsAndClearCommands();
                LoadGrammarLauncher();
            }
            else if (e.Result.Text.ToLower().StartsWith("filter ") && e.Result.Confidence > 0.5 || (e.Result.Grammar.Name.Contains("Phonetic Alphabet") && isKeyboard == false))
            {
                filterInProgress = true;
                var filterBy = e.Result.Text.ToLower().Replace("filter ", "").Replace("lower", "");
                if (e.Result.Grammar.Name.Contains("Phonetic Alphabet"))
                {
                    filterBy = Get1stLetterFromPhoneticAlphabet(e, filterBy).ToLower();
                }
                if (launcherCategoryMatched == true)
                {
                    ShowLauncherOptions(lastLauncherCategory, filterBy);
                }
                else
                {
                    ListCommandsInLanguageAndCategory($"{languageMatched} {categoryMatched}", filterBy);
                }
            }
            else if (e.Result.Text.ToLower().StartsWith("launcher ") && e.Result.Confidence > 0.5)
            {
                var result = e.Result.Text.ToLower();
                result = result.Replace("launcher ", "");
                ShowLauncherOptions(result, null);
            }
            else if ((e.Result.Text.ToLower().StartsWith("press ") && e.Result.Confidence > 0.5) || (e.Result.Grammar.Name.Contains("Phonetic Alphabet")) && isKeyboard == true)
            {
                UpdateCurrentProcess();
                ProcessKeyboardCommand(e);
            }
            else if (e.Result.Text.ToLower() == "edit list" && e.Result.Confidence > 0.7)
            {
                EditList();
            }
            else if (e.Result.Text.ToLower() == "toggle microphone" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                this.WriteLine("*************Keys sent to toggle the microphone*****************");
                List<string> keys = new List<string>(new string[] { "{ADD}" });
                SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            }
            else if (e.Result.Text.ToLower() == "restart dragon" && e.Result.Confidence > 0.97)
            {
                isKeyboard = false;
                RestartDragon();
            }
            else if (e.Result.Text.ToLower().StartsWith("kill") && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                var name = e.Result.Text.ToLower();
                name = name.Replace("kill", "").Trim();
                KillAllProcesses(name);
            }
            else if (e.Result.Text.ToLower() == "quit application" && e.Result.Confidence > 0.5)
            {
                QuitApplication();
            }
            else if ((e.Result.Text.ToLower() == "start emergency speech" && e.Result.Confidence > 0.84) || (e.Result.Text.ToLower() == "list commands" && e.Result.Confidence > 0.84))
            {
                launcherCategoryMatched = false;
                UnloadGrammarsAndClearCommands();
                EmergencySpeech();
            }
            else if (e.Result.Text.ToLower() == "short phrase mode" && e.Result.Confidence > 0.5)
            {
                isKeyboard = false;
                SetDragonToSleep();
                StartShortPhraseMode();
            }
            else if (e.Result.Text.ToLower().StartsWith("choose "))
            {
                var choiceNumber = Int32.Parse(e.Result.Text.Substring(7));
                isKeyboard = false;
                if (lastResult != null && choiceNumber <= (lastResult.PhraseResponse.Results.Length - 1))
                {
                    Dispatcher.Invoke(() => { finalResult.Text = lastResult.PhraseResponse.Results[choiceNumber].DisplayText; });
                }
            }
            else if (e.Result.Text.ToLower() == "transfer as paragraph" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                ToggleTransferAsParagraph();
            }
            else if (e.Result.Text.ToLower() == "remove spaces" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                ToggleRemoveSpaces();
            }
            else if (e.Result.Text.ToLower() == "remove punctuation" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                ToggleRemovePunctuation();
            }
            else if (e.Result.Text.ToLower() == "camel case" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                ToggleCamelCase();
            }
            else if (e.Result.Text.ToLower() == "variable" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                ToggleVariable();
            }
            else if (e.Result.Text.ToLower() == "transfer to application" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                TransferToApplication();
            }
            else if ((e.Result.Text.ToLower() == "grammar mode" && e.Result.Confidence > 0.965) || (e.Result.Text.ToLower() == "stop launcher" && e.Result.Confidence > 0.96))
            {
                speechRecognitionEngine.UnloadAllGrammars();
                speechRecognitionEngine.RecognizeAsyncStop();
                ListMainMenuCommands();
                launcherCategoryMatched = false;
                isKeyboard = false;
            }
            else if (e.Result.Text.ToLower() == "long dictation mode" && e.Result.Confidence > 0.965)
            {
                isKeyboard = false;
                StartLongDictationMode();
            }
            else if (e.Result.Text.ToLower() == "with intent mode" && e.Result.Confidence > 0.97)
            {
                StartWithIntentMode();
            }
            else if ((e.Result.Text.ToLower() == "visual basic intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower() == "go back" && languageMatched == "Visual Basic" && e.Result.Confidence > 0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                UnloadGrammarsAndClearCommands();
                LoadGrammarCustomIntellisense("Visual Basic");
                languageAndCategoryAlreadyMatched = false;
            }
            else if ((e.Result.Text.ToLower() == "c sharp intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower() == "go back" && languageMatched.ToLower() == "c sharp" && e.Result.Confidence > 0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                UnloadGrammarsAndClearCommands();
                LoadGrammarCustomIntellisense("C Sharp");
                languageAndCategoryAlreadyMatched = false;
            }
            else if ((e.Result.Text.ToLower() == "javascript intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower() == "go back" && languageMatched.ToLower() == "javascript" && e.Result.Confidence > 0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                UnloadGrammarsAndClearCommands();
                LoadGrammarCustomIntellisense("JavaScript");
                languageAndCategoryAlreadyMatched = false;
            }
            else if ((e.Result.Text.ToLower() == "global intellisense" && e.Result.Confidence > 0.5) || (e.Result.Text.ToLower() == "go back" && languageMatched.ToLower() != "visual basic" && e.Result.Confidence > 0.5))
            {
                isKeyboard = false;
                this.WriteLine($"Now Loading Global IntelliSense mode...");
                UnloadGrammarsAndClearCommands();
                LoadGrammarCustomIntellisense(null);
                languageAndCategoryAlreadyMatched = false;
            }
            else if (e.Result.Text.ToLower() == "go dormant" && e.Result.Confidence > 0.9)
            {
                isKeyboard = false;
                GoDormant();
            }
            else if (e.Result.Text == "Stop IntelliSense" || e.Result.Text == "Stop Keyboard" || e.Result.Text == "Stop Launcher")
            {
                isKeyboard = false;

                //speechRecognizer.RecognizeAsyncCancel();
                ListMainMenuCommands();
                //speechRecognizer.SetInputToDefaultAudioDevice();
                //speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
                languageAndCategoryAlreadyMatched = false;
            }
            else if (languageAndCategoryAlreadyMatched == true && e.Result.Confidence > 0.6)
            {
                isKeyboard = false;
                filterInProgress = false;
                PerformGlobalIntelliSense(e);
            }
            else if (launcherCategoryMatched == true && e.Result.Confidence > 0.9)
            {
                LaunchApplication(e.Result.Text);
                ListMainMenuCommands();
                filterInProgress = false;
            }
            else
            {
                isKeyboard = false;
                ListCommandsInLanguageAndCategory(e.Result.Text, null);
            }

            if (e.Result.Text.ToLower()!="go dormant")
            {
                lastCommand = e.Result.Text.ToLower();
                lastRecognition = e;
            }
        }

        private void PerformMouseCommand(SpeechRecognizedEventArgs e)
        {
            Win32.POINT p = new Win32.POINT();
            p.x = 100;
            p.y = 100;
            var horizontalCoordinate = e.Result.Words[1].Text;
            if (horizontalCoordinate=="Zero")
            {
                p.x = 5;
            }
            else if (horizontalCoordinate=="Alpha")
            {
                p.x = 50;
            }
            else if (horizontalCoordinate == "Bravo")
            {
                p.x = 100;
            }
            else if (horizontalCoordinate=="Charlie")
            {
                p.x = 150;
            }
            else if (horizontalCoordinate=="Delta")
            {
                p.x = 200;
            }
            else if (horizontalCoordinate=="Echo")
            {
                p.x = 250;
            }
            else if (horizontalCoordinate=="FoxTrot")
            {
                p.x = 300;
            }
            else if (horizontalCoordinate=="Golf")
            {
                p.x = 350;
            }
            else if (horizontalCoordinate=="Hotel")
            {
                p.x = 400;
            }
            else if (horizontalCoordinate == "India")
            {
                p.x = 450;
            }
            else if (horizontalCoordinate == "Juliet")
            {
                p.x = 500;
            }
            else if (horizontalCoordinate == "Kilo")
            {
                p.x = 550;
            }
            else if (horizontalCoordinate == "Lima")
            {
                p.x = 600;
            }
            else if (horizontalCoordinate == "Mike")
            {
                p.x = 650;
            }
            else if (horizontalCoordinate == "November")
            {
                p.x = 700;
            }
            else if (horizontalCoordinate == "Oscar")
            {
                p.x = 750;
            }
            else if (horizontalCoordinate == "Papa")
            {
                p.x = 800;
            }
            else if (horizontalCoordinate == "Qubec")
            {
                p.x = 850;
            }
            else if (horizontalCoordinate == "Romeo")
            {
                p.x = 900;
            }
            else if (horizontalCoordinate == "Sierra")
            {
                p.x = 950;
            }
            else if (horizontalCoordinate == "Tango")
            {
                p.x = 1000;
            }
            else if (horizontalCoordinate == "Uniform")
            {
                p.x = 1050;
            }
            else if (horizontalCoordinate == "Victor")
            {
                p.x = 1100;
            }
            else if (horizontalCoordinate == "Whiskey")
            {
                p.x = 1150;
            }
            else if (horizontalCoordinate == "X-ray")
            {
                p.x = 1200;
            }
            else if (horizontalCoordinate == "Yankee")
            {
                p.x = 1250;
            }
            else if (horizontalCoordinate == "Zulu")
            {
                p.x = 1300;
            }
            else if (horizontalCoordinate == "1")
            {
                p.x = 1350;
            }
            else if (horizontalCoordinate == "2")
            {
                p.x = 1400;
            }
            else if (horizontalCoordinate == "3")
            {
                p.x = 1450;
            }
            else if (horizontalCoordinate == "4")
            {
                p.x = 1500;
            }
            else if (horizontalCoordinate == "5")
            {
                p.x = 1550;
            }
            else if (horizontalCoordinate == "6")
            {
                p.x = 1600;
            }
            else if (horizontalCoordinate == "7")
            {
                p.x = 1650;
            }
            var verticalCoordinate = e.Result.Words[2].Text;
            if (verticalCoordinate=="Zero")
            {
                p.y = 5;
            }
            else if (verticalCoordinate=="Alpha")
            {
                p.y = 50;
            }
            else if (verticalCoordinate=="Bravo")
            {
                p.y = 100;
            }
            else if (verticalCoordinate == "Charlie")
            {
                p.y = 150;
            }
            else if (verticalCoordinate == "Delta")
            {
                p.y = 200;
            }
            else if (verticalCoordinate == "Echo")
            {
                p.y = 250;
            }
            else if (verticalCoordinate == "FoxTrot")
            {
                p.y = 300;
            }
            else if (verticalCoordinate == "Golf")
            {
                p.y = 350;
            }
            else if (verticalCoordinate == "Hotel")
            {
                p.y = 400;
            }
            else if (verticalCoordinate == "India")
            {
                p.y = 450;
            }
            else if (verticalCoordinate == "Juliet")
            {
                p.y = 500;
            }
            else if (verticalCoordinate == "Kilo")
            {
                p.y = 550;
            }
            else if (verticalCoordinate == "Lima")
            {
                p.y = 600;
            }
            else if (verticalCoordinate == "Mike")
            {
                p.y = 650;
            }
            else if (verticalCoordinate == "November")
            {
                p.y = 700;
            }
            else if (verticalCoordinate == "Oscar")
            {
                p.y = 750;
            }
            else if (verticalCoordinate == "Papa")
            {
                p.y = 800;
            }
            else if (verticalCoordinate == "Qubec")
            {
                p.y = 850;
            }
            else if (verticalCoordinate == "Romeo")
            {
                p.y = 900;
            }
            else if (verticalCoordinate == "Sierra")
            {
                p.y = 950;
            }
            else if (verticalCoordinate == "Tango")
            {
                p.y = 1000;
            }
            else if (verticalCoordinate == "Uniform")
            {
                p.y = 1050;
            }
            else if (verticalCoordinate == "Victor")
            {
                p.y = 1100;
            }
            else if (verticalCoordinate == "Whiskey")
            {
                p.y = 1150;
            }
            else if (verticalCoordinate == "X-ray")
            {
                p.y = 1200;
            }
            else if (verticalCoordinate == "Yankee")
            {
                p.y = 1250;
            }
            else if (verticalCoordinate == "Zulu")
            {
                p.y = 1300;
            }
            else if (verticalCoordinate == "1")
            {
                p.y = 1350;
            }
            else if (verticalCoordinate == "2")
            {
                p.y = 1400;
            }
            else if (verticalCoordinate == "3")
            {
                p.y = 1450;
            }
            else if (verticalCoordinate == "4")
            {
                p.y = 1500;
            }
            else if (verticalCoordinate == "5")
            {
                p.y = 1550;
            }
            else if (verticalCoordinate == "6")
            {
                p.y = 1600;
            }
            else if (verticalCoordinate == "7")
            {
                p.y = 1650;
            }
            var screen = e.Result.Words[0].Text;
            if (screen=="Screen" || screen=="Touch")
            {
                p.x = p.x + 1680;
            }

            Win32.SetCursorPos(p.x, p.y);
            SpeechUI.SendTextFeedback(e.Result, $" {e.Result.Text} H{p.x} V{p.y}",false);
            if (screen=="Click" || screen=="Touch")
            {
                Win32.mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)p.x, (uint)p.y, 0, 0);
            }
        }

        private void VariableCommand(SpeechRecognizedEventArgs e)
        {
            var value = "";
            foreach (var word in e.Result.Words.Where(w => w.Text.ToLower()!="variable"))
            {
                value = value + word.Text.Substring(0, 1).ToUpper() + word.Text.Substring(1).ToLower();
            }
            List<string> keys = new List<string> { value };
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
        }


        private void CamelCaseCommand(SpeechRecognizedEventArgs e)
        {
            var value = "";
            var counter = 0;
            foreach (var word in e.Result.Words.Where(w => w.Text.ToLower()!="camel" && w.Text.ToLower()!="case"))
            {
                counter++;
                if (counter==1)
                {
                    value = word.Text.ToLower();
                }
                else
                {
                    value = value + word.Text.Substring(0, 1).ToUpper() + word.Text.Substring(1).ToLower();
                }
            }
            List<string> keys = new List<string> { value };
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
        }

        private void PerformanceSymbolsCommand(SpeechRecognizedEventArgs e)
        {
            List<string> keys= new List<string>();
            var text = e.Result.Text.ToLower();
            this.WriteLine("Recognised:" + e.Result.Text);
            if (text.Contains("square brackets"))
            {
                keys.Add("[]");
            }
            else if (text.Contains("curly brackets"))
            {
                keys.Add("{{}");
                keys.Add("{}}");
            }
            else if (text.Contains("brackets"))
            {
                keys.Add("(");
                keys.Add(")");
            }
            else if (text.Contains("apostrophes"))
            {
                keys.Add("''");
            }
            else if (text.Contains("quotes"))
            {
                keys.Add("\"");
                keys.Add("\"");
            }
            else if (text.Contains("at signs"))
            {
                keys.Add("@@");
            }
            else if (text.Contains("chevrons"))
            {
                keys.Add("<>");
            }
            else if (text.Contains("equals"))
            {
                keys.Add("==");
            }
            else if (text.Contains("not equal"))
            {
                keys.Add("!=");
            }
            else if (text.Contains("plus"))
            {
                keys.Add("++");
            }
            else if (text.Contains("dollar"))
            {
                keys.Add("$$");
            }
            else if (text.Contains("hash"))
            {
                keys.Add("##");
            }
            else if (text.Contains("question marks"))
            {
                keys.Add("??");
            }
            else if (text.Contains("pipes"))
            {
                keys.Add("||");
            }
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            if (text.EndsWith("in"))
            {
                List<string> keysLeft = new List<string> { "{Left}" };
                SendKeysCustom(null, null, keysLeft, currentProcess.ProcessName);
            }
        }

        private void ToggleVariable()
        {
            if (this.Variable.IsChecked == false)
            {
                Dispatcher.Invoke(() => { Variable.IsChecked = true; });
                Dispatcher.Invoke(() => { CamelCase.IsChecked = false; });
            }
            else
            {
                Dispatcher.Invoke(() => { Variable.IsChecked = false; });
            }
        }

        private void ToggleCamelCase()
        {
            if (this.CamelCase.IsChecked == false)
            {
                Dispatcher.Invoke(() => { CamelCase.IsChecked = true; });
                Dispatcher.Invoke(() => { Variable.IsChecked = false; });
            }
            else
            {
                Dispatcher.Invoke(() => { CamelCase.IsChecked = false; });
            }
        }

        private void PerformReplacement()
        {
            if (this.ReplaceText.Text!= null  && this.ReplaceText.Text.Length>0 && this.ReplaceWith.Text!= null  && this.ReplaceWith.Text.Length>0)
            {
                var newFinal = this.finalResult.Text;
                newFinal = newFinal.Replace(this.ReplaceText.Text, this.ReplaceWith.Text);
                Dispatcher.Invoke(() =>
                {
                    finalResult.Text = newFinal;
                });
            }
        }

        private void SetupReplaceText(SpeechRecognizedEventArgs e,string replaceText)
        {
            var result = Get1stLetterFromPhoneticAlphabet(e, "");
            if (replaceText=="Replace")
            {
                Dispatcher.Invoke(() =>
                {
                    ReplaceText.Text = result;
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    ReplaceWith.Text = result;
                });
            }

            
        }

        private void StartWindowsSpeechRecognition()
        {
            //speechRecognitionEngine.RecognizeAsyncCancel();
            //speechRecognitionEngine.UnloadAllGrammars();
            //speechRecognitionEngine.Dispose();
            try
            {
                speechRecognizer = new SpeechRecognizer();
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show($"Error loading Windows Speech Recognition {exception.Message}","Error",MessageBoxButton.OK,MessageBoxImage.Error );
                return;
            }
            //Wildcard wildcard = new Wildcard();
            //Grammar grammar = wildcard.CreatePasswordGrammar();
            //speechRecognizer.LoadGrammarAsync(grammar);
            // Create and load a sample grammar.  
            //Grammar testGrammar =
            //  new Grammar(new GrammarBuilder("testing testing"));
            //testGrammar.Name = "Test Grammar";
            //recognizer.LoadGrammar(testGrammar);

            // Attach event handlers for recognition events.  
            //speechRecognizer.SpeechRecognized +=
            //      new EventHandler<SpeechRecognizedEventArgs>(
            //        SpeechRecognitionEngine_SpeechRecognized);
            //recognizer.SpeechRecognized +=
            //  new EventHandler<SpeechRecognizedEventArgs>(
            //    SpeechRecognizedHandler);
            //speechRecognizer.EmulateRecognizeCompleted +=
            //      new EventHandler<EmulateRecognizeCompletedEventArgs>(
            //        EmulateRecognizeCompletedHandler);

            //completed = false;
            isKeyboard = true;
            //        Start asynchronous emulated recognition.   
            //         This matches the grammar and generates a SpeechRecognized event.  
            //speechRecognizer.EmulateRecognizeAsync("testing testing");
            //         Check to see if recognizer is loaded, wait if it is not loaded.
            //System.Windows.MessageBox.Show($"The current state is: {speechRecognizer.State}");
            //if (speechRecognizer.State != RecognizerState.Listening)
            //    {
            //        Thread.Sleep(5000);
            //        // Put recognizer in lisHer job page of onto pressed her job stop listening wake up stopped listening wake up tening state.  
            //        speechRecognizer.EmulateRecognizeAsync("Start listening");
            //    }
            //    else
            //    {
            //        speechRecognizer.EmulateRecognizeAsync("Start listening");

            //}
            
        }

        private void BuildPhoneticAlphabetGrammars(bool useEngine= false)
        {
            Choices phoneticAlphabet = new Choices(new string[] { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet", "Kilo", "Lima", "Mike", "November", "Oscar", "Papa", "Qubec", "Romeo", "Sierra", "Tango", "Uniform", "Victor", "Whiskey", "X-ray", "Yankee", "Zulu" });
            GrammarBuilder grammarBuilder2 = new GrammarBuilder();
            grammarBuilder2.Append(phoneticAlphabet);
            grammarBuilder2.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder3 = new GrammarBuilder();
            grammarBuilder3.Append(phoneticAlphabet);
            grammarBuilder3.Append(phoneticAlphabet);
            grammarBuilder3.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder4 = new GrammarBuilder();
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            grammarBuilder4.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder5 = new GrammarBuilder();
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            grammarBuilder5.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder6 = new GrammarBuilder();
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            grammarBuilder6.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilder7 = new GrammarBuilder();
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            grammarBuilder7.Append(phoneticAlphabet);
            Choices phoneticAlphabet2to7 = new Choices(new GrammarBuilder[] { grammarBuilder2, grammarBuilder3, grammarBuilder4, grammarBuilder5, grammarBuilder6, grammarBuilder7 });
            Grammar grammarPhoneticAlphabets = new Grammar((GrammarBuilder)phoneticAlphabet2to7);
            grammarPhoneticAlphabets.Name = "Phonetic Alphabet";
            if (useEngine==true)
            {
                speechRecognitionEngine.LoadGrammarAsync(grammarPhoneticAlphabets);
            }
            else
            {
                speechRecognizer.LoadGrammarAsync(grammarPhoneticAlphabets);
            }

            Choices choices = new Choices("Lower");
            GrammarBuilder grammarBuilderLower2 = new GrammarBuilder();
            grammarBuilderLower2.Append(choices);
            grammarBuilderLower2.Append(phoneticAlphabet);
            grammarBuilderLower2.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower3 = new GrammarBuilder();
            grammarBuilderLower3.Append(choices);
            grammarBuilderLower3.Append(phoneticAlphabet);
            grammarBuilderLower3.Append(phoneticAlphabet);
            grammarBuilderLower3.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower4 = new GrammarBuilder();
            grammarBuilderLower4.Append(choices);
            grammarBuilderLower4.Append(phoneticAlphabet);
            grammarBuilderLower4.Append(phoneticAlphabet);
            grammarBuilderLower4.Append(phoneticAlphabet);
            grammarBuilderLower4.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower5 = new GrammarBuilder();
            grammarBuilderLower5.Append(choices);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            grammarBuilderLower5.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower6 = new GrammarBuilder();
            grammarBuilderLower6.Append(choices);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            grammarBuilderLower6.Append(phoneticAlphabet);
            GrammarBuilder grammarBuilderLower7 = new GrammarBuilder();
            grammarBuilderLower7.Append(choices);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            grammarBuilderLower7.Append(phoneticAlphabet);
            Choices phoneticAlphabetLower2to7 = new Choices(new GrammarBuilder[] { grammarBuilderLower2, grammarBuilderLower3, grammarBuilderLower4, grammarBuilderLower5, grammarBuilderLower6, grammarBuilderLower7 });
            Grammar grammarPhoneticAlphabetsLower = new Grammar((GrammarBuilder)phoneticAlphabetLower2to7);
            grammarPhoneticAlphabetsLower.Name = "Phonetic Alphabet Lower";
            if (useEngine == true)
            {
                speechRecognitionEngine.LoadGrammarAsync(grammarPhoneticAlphabetsLower);
            }
            else
            {
                speechRecognizer.LoadGrammarAsync(grammarPhoneticAlphabetsLower);
            }
        }

        private void EmulateRecognizeCompletedHandler(object sender, EmulateRecognizeCompletedEventArgs e)
        {
            completed = true;

        }

        private void SpeechRecognizedHandler(object sender, SpeechRecognizedEventArgs e)
        {

            //recognizer.Dispose();
            completed = false;
        }
        private void GoDormant()
        {
            //this._mainWindow.Topmost = false;
            //this._mainWindow.Hide();
            UnloadGrammarsAndClearCommands();
            Choices choices = new Choices();
            choices.Add("start emergency speech");
            this.WriteCommandLine("Start Emergency Speech");
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognizer.LoadGrammarAsync(grammar);
            this.WriteLine("****Pressing the / Key on the number pad****");
            List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            launcherCategoryMatched = false;
        }

        private void StartWithIntentMode()
        {
            isKeyboard = false;
            //speechRecognizer.RecognizeAsyncStop();
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

        private void StartLongDictationMode()
        {
            speechRecognizer.EmulateRecognize("Stop listening");            
            speechRecognizer.UnloadAllGrammars();

            
            speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognitionEngine_SpeechRecognized);
            Dispatcher.Invoke(() =>
            {
                Commands.Text = "";
            });
            //speechRecognizer.RecognizeAsyncStop();
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
            LoadGrammarKeyboard(false,false);
            Choices choices = SetupDictationGrammarCommands();
            for (int i = 0; i < 6; i++)
            {
                choices.Add($"Choose {i}");
                this.WriteCommandLine($"Choose {i}");
            }

            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognitionEngine.LoadGrammarAsync(grammar);
            speechRecognitionEngine.SetInputToDefaultAudioDevice();
            speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void TransferToApplication()
        {
            var resultText = "";
            if (this.TransferAsParagraph.IsChecked == true)
            {
                resultText = "{Enter}" + this.finalResult.Text + "{Enter}";
            }
            else
            {
                resultText =this.finalResult.Text;
            }
            List<string> keys = new List<string>(new string[] { resultText });
            SendKeysCustom(null, null, keys, currentProcess.ProcessName, null);
        }
        private void ToggleRemoveSpaces()
        {
            if (this.RemoveSpaces.IsChecked==false)
            {
                Dispatcher.Invoke(() => { RemoveSpaces.IsChecked = true; });
            }
            else
            {
                Dispatcher.Invoke(() => { RemoveSpaces.IsChecked = false; });
            }
        }

        private void ToggleRemovePunctuation()
        {
            if (this.RemovePunctuation.IsChecked == false)
            {
                Dispatcher.Invoke(() => { RemovePunctuation.IsChecked = true; });
            }
            else
            {
                Dispatcher.Invoke(() => { RemovePunctuation.IsChecked = false; });
            }
        }
        private void ToggleTransferAsParagraph()
        {
            if (this.TransferAsParagraph.IsChecked == false)
            {
                Dispatcher.Invoke(() => { TransferAsParagraph.IsChecked = true; });
            }
            else
            {
                Dispatcher.Invoke(() => { TransferAsParagraph.IsChecked = false; });
            }
        }

        private void StartShortPhraseMode()
        {
            speechRecognizer.UnloadAllGrammars();
            //speechRecognizer.RecognizeAsyncStop();

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
            dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 0, 1);
            dispatcherTimer.Start();
            Dispatcher.Invoke(() =>
            {
                Commands.Text = "";
            });
            Choices choices = SetupDictationGrammarCommands();
            for (int i = 0; i < 6; i++)
            {
                choices.Add($"Choose {i}");
                this.WriteCommandLine($"Choose {i}");
            }

            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognizer.LoadGrammarAsync(grammar);
            //speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
        }
        private void SetUpSymbolGrammarCommands(bool useEngine)
        {
            Choices choices = new Choices();
            choices.Add("Square Brackets");
            choices.Add("Brackets");
            choices.Add("Curly Brackets");
            choices.Add("Single Quotes");
            choices.Add("Apostrophes");
            choices.Add("Quotes");
            choices.Add("At Signs");
            choices.Add("Chevrons");
            choices.Add("Equals");
            choices.Add("Not Equal");
            choices.Add("Plus");
            choices.Add("Dollar");
            choices.Add("Hash");
            choices.Add("Pipes");
            choices.Add("Ampersands");

            Choices choicesInOut = new Choices("In","Out");

            GrammarBuilder grammarBuilder = new GrammarBuilder();
            grammarBuilder.Append(choices);
            grammarBuilder.Append(choicesInOut);
            Grammar grammarSymbols = new Grammar((GrammarBuilder)grammarBuilder);
            grammarSymbols.Name = "Symbols";
            if (useEngine==true)
            {
                speechRecognitionEngine.LoadGrammarAsync(grammarSymbols);
            }
            else
            {
                speechRecognizer.LoadGrammarAsync(grammarSymbols);
            }
        }

        private void CreateDictationGrammar(string initialPhrase, string grammarName)
        {
            GrammarBuilder grammarBuilder= new GrammarBuilder();
            grammarBuilder.Append(new Choices(initialPhrase));
            grammarBuilder.AppendDictation();

            Grammar grammar = new Grammar((GrammarBuilder)grammarBuilder);
            grammar.Name = grammarName;
            speechRecognizer.LoadGrammarAsync(grammar);
        }


        private Choices SetupDictationGrammarCommands()
        {
            this.WriteCommandLine("Grammar Mode");
            this.WriteCommandLine("Transfer to Application");
            this.WriteCommandLine("Transfer as Paragraph");
            this.WriteCommandLine("Remove Spaces");
            this.WriteCommandLine("Remove Punctuation");
            this.WriteCommandLine("Camel Case");
            this.WriteCommandLine("Variable");

            Choices choices = new Choices();
            choices.Add("Grammar Mode");
            choices.Add("Transfer to Application");
            choices.Add("Transfer as Paragraph");
            choices.Add("Remove Spaces");
            choices.Add("Remove Punctuation");
            choices.Add("Camel Case");
            choices.Add("Variable");
            return choices;
        }

        private void EmergencySpeech()
        {
            if (lastCommand == null)
            {
                ListMainMenuCommands();
            }
            else if (lastCommand.ToLower() == "global intellisense")
            {
                LoadGrammarCustomIntellisense(null);
            }
            else if (lastCommand.ToLower() == "visual basic intellisense")
            {
                LoadGrammarCustomIntellisense("Visual Basic");
            }
            else if (lastCommand.ToLower() == "c sharp intellisense")
            {
                LoadGrammarCustomIntellisense("C Sharp");
            }
            else if (lastCommand.ToLower() == "javascript intellisense")
            {
                LoadGrammarCustomIntellisense("JavaScript");
            }
            else if (lastCommand.ToLower().StartsWith("press "))
            {
                this.WriteLine($"Keyboard mode...");
                isKeyboard = true;
                LoadGrammarKeyboard();
            }
            else if (languageAndCategoryAlreadyMatched == true)
            {
                ListCommandsInLanguageAndCategory(languageMatched + " " + categoryMatched, null);
            }
            else
            {
                ListMainMenuCommands();
            }
            //this._mainWindow.IsVisible=true;


            //this._mainWindow.Topmost = true;
            this._mainWindow.Show();
            //SetForegroundWindow(currentProcess.Handle);
            this.WriteLine("****Pressing the / Key on the number pad****");
            List<string> keys = new List<string>(new string[] { "{DIVIDE}","%{Tab}" });
            UpdateCurrentProcess();
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
        }

        private void QuitApplication()
        {
            List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
            try
            {
                System.Windows.Application.Current.Shutdown();
            }
            catch (Exception exception)
            {
                this.WriteLine(exception.Message);
            }
        }

        private void RestartDragon()
        {
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
                    SendKeysCustom(null, null, keysKB, currentProcess.ProcessName);

                }
            }
            catch (Exception exception)
            {
                Dispatcher.Invoke(() =>
                {
                    _logText.Text += $"An error has occurred: {exception.Message}";
                });
            }
        }

        private void EditList()
        {
            var arguments = "";
            if (languageAndCategoryAlreadyMatched == true)
            {
                arguments = $"\"C:\\Msoffice\\Access\\DragonScripts.accdb\" /cmd \"{languageMatched}|{categoryMatched}\"";
            }
            else
            {
                arguments = $"\"C:\\Msoffice\\Access\\DragonScripts.accdb\" /cmd \"Launcher|{lastLauncherCategory}\"";
            }

            var commandline = $"\"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\MSACCESS.EXE\"";

            try
            {
                Process.Start(commandline, arguments);
                this.WriteLine($"Launching {commandline} with arguments {arguments}");
            }
            catch (Exception exception)
            {
                this.WriteLine(exception.Message);
            }
        }

        private void ProcessKeyboardCommand(SpeechRecognizedEventArgs e)
        {
            var value = e.Result.Text;
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
            value = value.Replace("Semicolon", ";");
            value = value.Replace("Control", "^");
            value = value.Replace("Alt Space", "% ");
            value = value.Replace("Alt", "%");
            value = value.Replace("Escape", "{Esc}");
            value = value.Replace("Zero", "0");
            value = value.Replace("Stop", ".");
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
            value = value.Replace("Colon", ":");
            value = value.Replace("Ampersand", "&");
            value = value.Replace("Dollar", "$");
            value = value.Replace("Exclamation Mark", "!");
            value = value.Replace("Double Quote", "\"");
            value = value.Replace("Pound", "£");
            value = value.Replace("Asterix", "*");
            value = value.Replace("Apostrophe", "'");
            value = value.Replace("Equal", "=");
            value = value.Replace("Open Bracket", "(");
            value = value.Replace("Close Bracket", ")");

            


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
            if (value.Contains("{Up}") && IsNumber(value.Substring(value.IndexOf("}") + 1)))
            {
                value = "{Up " + value.Substring(value.IndexOf("}") + 1) + "}";
            }
            if (value.Contains("{Down}") && IsNumber(value.Substring(value.IndexOf("}") + 1)))
            {
                value = "{Down " + value.Substring(value.IndexOf("}") + 1) + "}";
            }
            if (value.Contains("{Left}") && IsNumber(value.Substring(value.IndexOf("}") + 1)))
            {
                value = "{Left " + value.Substring(value.IndexOf("}") + 1) + "}";
            }
            if (value.Contains("{Right}") && IsNumber(value.Substring(value.IndexOf("}") + 1)))
            {
                value = "{Right " + value.Substring(value.IndexOf("}") + 1) + "}";
            }
            value = value.Replace("Percent", "{%}");
            value = value.Replace("Plus", "{+}");
            if (e.Result.Grammar.Name.Contains("Phonetic Alphabet"))
            {
                value = Get1stLetterFromPhoneticAlphabet(e, value);
            }

            this.WriteLine($"*****Sending Keys: {value.Replace("{", "").Replace("}", "").ToString()}*******");

            List<string> keys = new List<string>(new string[] { value });
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
        }

        private static string Get1stLetterFromPhoneticAlphabet(SpeechRecognizedEventArgs e, string value)
        {
            if (e.Result.Grammar.Name == "Phonetic Alphabet")
            {
                value = "";
                foreach (var word in e.Result.Words)
                {
                    value = value + word.Text.Substring(0, 1);
                }
            }
            else if (e.Result.Grammar.Name == "Phonetic Alphabet Lower")
            {
                value = "";
                foreach (var word in e.Result.Words)
                {
                    if (word.Text != "Lower")
                    {
                        value = value + word.Text.ToLower().Substring(0, 1);
                    }
                }
            }
            else if (e.Result.Grammar.Name=="Replace Letters")
            {
                value = "";
                var upper = false;
                foreach (var word in e.Result.Words)
                {
                    if (word.Text=="Upper")
                    {
                        upper = true;
                    }
                    else if (word.Text=="Replace" || word.Text=="With" || word.Text=="this")
                    {
                        //Do nothing
                    }
                    else 
                    {
                        if (upper==true)
                        {
                            value = value + word.Text.ToUpper().Substring(0, 1);
                            upper = false;
                        }
                        else
                        {
                            value = value + word.Text.ToLower().Substring(0, 1);
                        }
                    }
                }
            }
            return value;
        }

        private void UnloadGrammarsAndClearCommands()
        {
            speechRecognizer.UnloadAllGrammars();
            Dispatcher.Invoke(() =>
            {
                Commands.Text = "";
            });
        }

        private void ShowLauncherOptions(string result, string filterBy,bool useEngine= false)
        {
            speechRecognizer.UnloadAllGrammars();
            Choices choices = new Choices(); Dispatcher.Invoke(() =>
            {
                Commands.Text = "";
            });
            using (var db = new MyDatabase())
            {
                var categoryId = db.tblCategories.Where(c => c.Category.ToLower() == result).FirstOrDefault().MenuNumber;
                var computerId = db.Computers.Where(c => c.ComputerName == Environment.MachineName).FirstOrDefault()?.ID;
                List<tblLauncher> launchItems = null;
                if (filterBy != null)
                {
                    launchItems = db.tblLaunchers.Where(l => l.Menu == categoryId && (l.ComputerID == null || l.ComputerID == computerId) && l.Name.ToLower().Contains(filterBy)).OrderBy(l => l.Name).ToList();
                }
                else
                {
                    launchItems = db.tblLaunchers.Where(l => l.Menu == categoryId && (l.ComputerID == null || l.ComputerID == computerId)).OrderBy(l => l.Name).ToList();
                }
                foreach (var launchItem in launchItems)
                {
                    this.WriteCommandLine($"{launchItem.Name}");
                    choices.Add(launchItem.Name);
                }
            }
            this.WriteCommandLine("");
            this.WriteCommandLine("Edit List");
            choices.Add("Edit List");
            choices.Add("Go Back");
            this.WriteCommandLine("Go Back");
            this.WriteCommandLine("Go Dormant");
            choices.Add("Go Dormant");
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognizer.LoadGrammar(grammar);
            launcherCategoryMatched = true;
            lastLauncherCategory = result;
            BuildPhoneticAlphabetGrammars(useEngine);
            this.WriteLine($"Launcher set up for: {result}");
        }

        private void CreateSpellingDictationGrammar()
        {
            DictationGrammar spellingDictationGrammar = new DictationGrammar("grammar:dictation#spelling");
            spellingDictationGrammar.Name = "spelling dictation";
            spellingDictationGrammar.Enabled = true;
            speechRecognizer.LoadGrammar(spellingDictationGrammar);
            this.WriteCommandLine("<abc..>");
        }

        private void CreateDictationGrammar()
        {
            //Create the filter dictation Grammar
            DictationGrammar dictationGrammar = new DictationGrammar("grammar:dictation");
            dictationGrammar.Name = "Filter Dictation";
            dictationGrammar.Enabled = true;
            speechRecognizer.LoadGrammar(dictationGrammar);
            dictationGrammar.SetDictationContext("Filter", null);
            this.WriteCommandLine("Filter <dictation>");
        }

        private void LaunchApplication(string result)
        {
            var commandline = GetCommandLineFromLauncherName(result);
            try
            {
                Process.Start(commandline);
                this.WriteLine("Launching " + commandline);
                launcherCategoryMatched = false;
            }
            catch (Exception exception)
            {
                this.WriteLine(exception.Message);
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
            if (speechRecognizer!= null )
            {
                speechRecognizer.Dispose();
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

            ListMainMenuCommands();
            //speechRecognizer.SetInputToDefaultAudioDevice();
            //speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
            UpdateCurrentProcess();
            List<string> keys = new List<string>(new string[] { "{DIVIDE}" });
            SendKeysCustom(null, null, keys, currentProcess.ProcessName);
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

        private void ListenForSpeech(object sender, RoutedEventArgs e=null)
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
                this.WriteLine($"--- OnMicShortPhraseResponseReceivedHandler {e.PhraseResponse.RecognitionStatus.ToString()}---");

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
                        if (categoryMatched=="Access" || categoryMatched=="Blocks")
                        {
                            keys = new List<string>(new string[] { "{Enter 2}" });
                            SendKeysCustom("", "", keys, currentProcess.ProcessName);
                        }
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

        private void ListCommandsInLanguageAndCategory(string phrase,string filterBy,bool useEngine= false)
        {
            phrase = phrase.ToLower().Replace("intellisense", "not applicable");
            phrase = phrase.Replace("jay query", "jquery");
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
                        var computerId = db.Computers.Where(c => c.ComputerName == Environment.MachineName).FirstOrDefault()?.ID;
                        foreach (var category in categories)
                        {
                            var temporary = phrase.Replace(language.Language.ToLower(), "").Trim();
                            if (category.Category.ToLower().EndsWith( temporary) )
                            {
                                categoryMatched = category.Category;
                                languageAndCategoryAlreadyMatched = true;
                                List<CustomIntelliSense> commands = null;
                                if (filterBy!= null )
                                {
                                    commands = db.tblCustomIntelliSenses.Where(i => i.Language_ID == language.ID && i.Category_ID == category.MenuNumber && (i.ComputerID == null || i.ComputerID == computerId) && i.Display_Value.ToLower().Contains(filterBy)).OrderBy(s => s.Display_Value).ToList();
                                }
                                else
                                {
                                    commands = db.tblCustomIntelliSenses.Where(i => i.Language_ID == language.ID && i.Category_ID == category.MenuNumber && (i.ComputerID== null || i.ComputerID==computerId)  ).OrderBy(s => s.Display_Value).ToList();
                                }
                                speechRecognizer.UnloadAllGrammars();
                                Choices choices = new Choices();
                                Dispatcher.Invoke(() =>
                                {
                                    Commands.Text = "";
                                });
                                foreach (var command in commands)
                                {
                                    if (category.Category=="Passwords")
                                    {
                                        this.WriteCommandLine($"{command.Display_Value}");
                                    }
                                    else if (command.Display_Value.Length + command.SendKeys_Value.Length >59)
                                    {
                                        this.WriteCommandLine($"{command.Display_Value}      ({command.SendKeys_Value.Substring(0,59 - command.Display_Value.Length)})");
                                    }
                                    else
                                    {
                                        this.WriteCommandLine($"{command.Display_Value}      ({command.SendKeys_Value})");
                                    }
                                    //Load a new grammar here
                                    choices.Add(command.Display_Value);
                                }
                                this.WriteCommandLine("");
                                choices.Add("Stop IntelliSense");
                                this.WriteCommandLine("Stop IntelliSense");
                                choices.Add("Global IntelliSense");
                                this.WriteCommandLine("Global IntelliSense");
                                this.WriteCommandLine("Edit List");
                                choices.Add("Edit List");
                                choices.Add("Go Back");
                                this.WriteCommandLine("Go Back");
                                this.WriteCommandLine("Go Dormant");
                                choices.Add("Go Dormant");
                                Grammar grammar = new Grammar(new GrammarBuilder(choices));
                                speechRecognizer.LoadGrammar(grammar);
                                this.WriteLine($"Language and Category: {languageMatched} {categoryMatched}");
                                BuildPhoneticAlphabetGrammars(useEngine);
                                return;
                            }
                        }
                    }
                }

            }

        }

        private void ListMainMenuCommands()
        {
            if (this.micClient!= null )
            {
                this.micClient.EndMicAndRecognition();
            }
            Choices choices = new Choices();
            speechRecognizer.UnloadAllGrammars();
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
            //this.WriteCommandLine($"Short Phrase Mode");
            //choices.Add("Short Phrase Mode");
            this.WriteCommandLine($"Long Dictation Mode");
            choices.Add("Long Dictation Mode");
            //this.WriteCommandLine($"Windows Speech Recognition");
            //choices.Add("Windows Speech Recognition");
            CreateDictationGrammar("camel case", "camel case");
            CreateDictationGrammar("variable", "variable");
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
            this.WriteCommandLine("Launch Applications");
            choices.Add("Launch Applications");
            this.WriteCommandLine($"List Commands");
            choices.Add("List Commands");
            Grammar grammar = new Grammar(new GrammarBuilder(choices));
            speechRecognizer.LoadGrammarAsync(grammar);
            LoadGrammarCustomIntellisense(null , false);
            LoadGrammarLauncher(false);
            LoadGrammarKeyboard(false);
            LoadGrammarMouseCommands();

            dispatcherTimer.Stop();
        }

        private void LoadGrammarMouseCommands()
        {
            List<string> screenCoordinates = new List<string> { "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet", "Kilo", "Lima", "Mike", "November", "Oscar", "Papa", "Qubec", "Romeo", "Sierra", "Tango", "Uniform", "Victor", "Whiskey", "X-ray", "Yankee", "Zulu","1","2","3","4","5","6","7","Zero" };
            Choices choices = new Choices();
            List<string> monitorNames = new List<string> { "Arrow", "Screen","Click","Touch" };
            foreach (var item in screenCoordinates)
            {
                foreach (var monitorName in monitorNames)
                {
                    foreach (var item2 in screenCoordinates)
                    {
                        if (item2=="Uniform")
                        {
                            break;
                        }
                        choices.Add($"{monitorName} {item} {item2}");
                    }
                }
            }
            GrammarBuilder grammarBuilder = new GrammarBuilder(choices);
            Grammar grammar = new Grammar((GrammarBuilder)grammarBuilder);
            grammar.Name = "Mouse Command";
            speechRecognizer.LoadGrammarAsync(grammar);
        }

        private void AddApplicationsToKillToChoices(Choices choices)
        {
            using (var db = new MyDatabase())
            {
                List<ApplicationsToKill> applicationsToKill = null;
                try
                {
                    applicationsToKill = db.ApplicationsToKill.Where(k => k.Display == true).OrderBy(k => k.CommandName).ToList();
                }
                catch (Exception  exception )
                {
                    System.Windows.MessageBox.Show($"There is a problem connecting to the database. {exception.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                }
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
                //lastResult = null;
            }
            else
            {
                if (!e.PhraseResponse.Results[0].LexicalForm.StartsWith("choose ") && e.PhraseResponse.Results[0].LexicalForm!="transfer to application" && !e.PhraseResponse.Results[0].LexicalForm.StartsWith("replace") && !e.PhraseResponse.Results[0].LexicalForm.StartsWith("with this") && e.PhraseResponse.Results[0].LexicalForm!="perform replacement"  && e.PhraseResponse.Results[0].LexicalForm!="phone replacement" && e.PhraseResponse.Results[0].LexicalForm!="grammar mode")
                {
                    if (e.PhraseResponse.Results.Length>1)
                    {
                        Dispatcher.Invoke(() => {
                            var position = Commands.Text.IndexOf("********* Final n-BEST Results *********");
                            if (position>0)
                            {
                                Commands.Text = Commands.Text.Substring(0, (position ));
                            }
                        });
                        lastResult = e;
                        for (int i = 0; i < e.PhraseResponse.Results.Length; i++)
                        {
                            if (i==0)
                            {
                                this.WriteCommandLine("********* Final n-BEST Results *********");
                            }
                            this.WriteCommandLine($"{i} {e.PhraseResponse.Results[i].Confidence.ToString().Substring(0,1)} ({e.PhraseResponse.Results[i].DisplayText})");
                        }
                    }
                    if (e.PhraseResponse.Results[0].Confidence!=Confidence.Low )
                    {
                        Dispatcher.Invoke(() =>
                        {
                            var text = "";
                            if (this.RemovePunctuation.IsChecked==true)
                            {
                                text = e.PhraseResponse.Results[0].LexicalForm;
                            }
                            else if (this.CamelCase.IsChecked==true)
                            {
                                TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
                                text = e.PhraseResponse.Results[0].LexicalForm.ToLower();
                                text = textInfo.ToTitleCase(text);
                                text = text.Replace(" ", "");
                                text = text.Substring(0, 1).ToLower() + text.Substring(1);
                            }
                            else if (this.Variable.IsChecked==true)
                            {
                                TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
                                text = e.PhraseResponse.Results[0].LexicalForm.ToLower();
                                text = textInfo.ToTitleCase(text);
                                text = text.Replace(" ", "");
                            }
                            else
                            {
                                text = e.PhraseResponse.Results[0].DisplayText;
                            }
                            if (this.RemoveSpaces.IsChecked==true)
                            {
                                text = text.Replace(" ", "");
                            }
                            finalResult.Text = (text);
                        });
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
                if (e.PhraseResponse.Results.Count()>0 &&  e.PhraseResponse.Results[0].LexicalForm == "transfer to application" && e.PhraseResponse.RecognitionStatus == RecognitionStatus.RecognitionSuccess)
                {
                    List<string> keys;
                    if (this.TransferAsParagraph.IsChecked==true)
                    {
                        keys = new List<string>(new string[] { "{Enter}" + finalResult.Text + "{Enter}" });
                    }
                    else
                    {
                        keys = new List<string>(new string[] {  finalResult.Text  });
                    }
                    SendKeysCustom(null,  null , keys,currentProcess.ProcessName, null );
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
            var resultLexical = "";
            if (e.PhraseResponse.Results.Count()>0)
            {
                resultLexical = e.PhraseResponse.Results[0].LexicalForm;
            }
            if ((resultLexical.Contains("transfer") && resultLexical.Contains("application")) || (resultLexical.Contains("transfer") && resultLexical.Contains("paragraph")) || resultLexical.Contains("mode") || resultLexical.Contains("grammer") || resultLexical.Contains("remove punctuation") || resultLexical.Contains("remove spaces") || resultLexical.Contains("camel case") || resultLexical.Contains("camelcase") || resultLexical.Contains("variable"))
            {
                //Do not change the final result
            }
            else
            {
                this.WriteResponseResult(e);
            }
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
            var formattedStr = "";
            if (format.Contains("{") || format.Contains("}"))
            {
                formattedStr = format;
            }
            else
            {
                formattedStr = string.Format(format, args);
            }
            formattedStr = formattedStr.Replace("/n", "");
            formattedStr = formattedStr.Replace("/r", "");
            formattedStr = formattedStr.Replace(Environment.NewLine, "");

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
                //ButtonAutomationPeer peer = new ButtonAutomationPeer(this._startButton);
                //IInvokeProvider invokeProvider = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                //invokeProvider.Invoke();
                ListenForSpeech(sender);
            }
        }
        void ActivateApp(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            //Activate the first application refined with this name
            if (processes.Count()>0)
            {
                var currentWindowHandle = GetForegroundWindow();
                if (processes[0].Handle != currentWindowHandle)
                {
                    SetForegroundWindow(processes[0].MainWindowHandle);
                }
            }
        }
        void ActivateApp(IntPtr windowHandle)
        {
            var currentWindowHandle = GetForegroundWindow();
            if (windowHandle!=currentWindowHandle)
            {
                SetForegroundWindow(windowHandle);
            }
        }
    }
}
