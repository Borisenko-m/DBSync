using System;
using System.Collections.Generic;
using System.Text;

namespace Excel_Interop
{
    class Comment
    {
        public Comment(User user,
                       string text,
                       byte[] data,
                       DateTime time)
        {
            User = user;
            Text = text;
            Data = data;
            Time = time;
        }
        User User { get; }
        string Text { get; }
        byte[] Data { get; }
        DateTime Time { get; }

    }
}
