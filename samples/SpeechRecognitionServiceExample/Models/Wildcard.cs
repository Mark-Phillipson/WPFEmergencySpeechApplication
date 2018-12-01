using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Speech.Recognition;

namespace SpeechToTextWPFSample.Models
{
    public class Wildcard
    {


        public Grammar CreatePasswordGrammar()
        {
            GrammarBuilder wildcardBuilder = new GrammarBuilder();
            wildcardBuilder.AppendWildcard();
            SemanticResultKey passwordKey =
              new SemanticResultKey("Password", wildcardBuilder);

            GrammarBuilder passwordBuilder =
              new GrammarBuilder("My Password is");
            passwordBuilder.Append(passwordKey);

            Grammar passwordGrammar = new Grammar(passwordBuilder);
            passwordGrammar.Name = "Password input";

            passwordGrammar.SpeechRecognized +=
              new EventHandler<SpeechRecognizedEventArgs>(
                PasswordInputHandler);

            return passwordGrammar;
        }

        // Handle the SpeechRecognized event for the password grammar.
        private void PasswordInputHandler(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result == null) return;

            RecognitionResult result = e.Result;
            SemanticValue semantics = e.Result.Semantics;

            if (semantics.ContainsKey("Password"))
            {
                RecognizedAudio passwordAudio =
                  result.GetAudioForWordRange(
                    result.Words[3], result.Words[result.Words.Count - 1]);

                if (IsValidPassword(passwordAudio))
                {
                    Console.WriteLine("Password accepted.");

                    // Add code to handle a valid password here.
                }
                else
                {
                    Console.WriteLine("Invalid password.");

                    // Add code to handle an invalid password here.
                }
            }
        }

        // Validate the password input. 
        private bool IsValidPassword(RecognizedAudio passwordAudio)
        {
            Console.WriteLine("Validating password.");

            // Add password validation code here.

            return false;
        }
    }
}
