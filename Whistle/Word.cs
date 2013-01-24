using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Whistle
{
    class Word
    {
        public Word() { }
        public String Text { get; set; }
        public string AttachedText { get; set; }
        public bool IsShellCommand { get; set; }
    }
}
