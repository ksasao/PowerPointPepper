using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointPepper
{
    public enum Command { Start, Speech }
    public class Message
    {
        public Command Command { get; set; }
        public string Data { get; set; }
    }
}
