using System;
using System.Runtime.Serialization;
namespace VisualFiParser
{
    [Serializable]
    class ArrayBoundaryException : Exception
    {

        public ArrayBoundaryException() : base()
        {

        }
        public ArrayBoundaryException(String message) : base(message)
        {
            this.HelpLink = "www.google.com";
        }
        public ArrayBoundaryException(String message, int k) : base(message)
        {
            Console.Out.WriteLine(message);
            Console.Out.WriteLine("Controllare la riga {0}", k);
        }
        public ArrayBoundaryException(String message, String filename, int k) : base(message)
        {
            Console.Out.WriteLine(message);
            Console.Out.WriteLine("Controllare la riga {0} nel file:", k);
            Console.Out.WriteLine(filename);

        }
        public ArrayBoundaryException(String message, Exception innerException) : base(message, innerException)
        { }
        public ArrayBoundaryException(SerializationInfo info, StreamingContext context)
        : base(info, context)
        { }

    }
}
