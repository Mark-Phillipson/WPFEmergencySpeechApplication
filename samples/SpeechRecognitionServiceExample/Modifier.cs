using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeechToTextWPFSample
{
     class Modifier
    {
        public  string ModifierName { get; set; }

        public string SendKeysCode { get; set; }

        public List<Modifier> GetModifiers()
        {
            List<Modifier> modifiers= new List<Modifier>();
            Modifier modifier = new Modifier { ModifierName = "Control", SendKeysCode = "^" };
            modifiers.Add(modifier);
            modifier = new Modifier { ModifierName = "Alt", SendKeysCode = "%" };
            modifiers.Add(modifier);
            modifier = new Modifier { ModifierName = "Shift", SendKeysCode = "+" };
            modifiers.Add(modifier);
            return modifiers;
        }
}
}
