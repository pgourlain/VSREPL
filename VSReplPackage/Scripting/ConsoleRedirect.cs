using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSReplPackage.Scripting
{
    class ConsoleRedirect : IDisposable
    {
        TextWriter oldWriter;
        TextWriter newWriter;
        public ConsoleRedirect()
        {
            newWriter = new TextWriterToVsTools();
            oldWriter = Console.Out;
            Console.SetOut(newWriter);
        }
        public void Dispose()
        {
            Console.SetOut(oldWriter);
            newWriter.Dispose();
        }
    }


    class TextWriterToVsTools : TextWriter
    {
        public override Encoding Encoding
        {
            get
            {
                return Encoding.Default;
            }
        }

        public override void Write(char value)
        {
            VSTools.LogDebug(new string (value, 1));
        }

        public override void Write(char[] buffer, int index, int count)
        {
            VSTools.LogDebug(new string(buffer, index, count));
        }

        public override void Write(string value)
        {
            VSTools.LogDebug(value);
        }
    }
}
