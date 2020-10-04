using System;
using System.Collections.Generic;
using System.Text;

namespace Excel_Interop
{
    class Contact
    {
        IEnumerable<Telephone> Telephones { get; }
        IEnumerable<Email> Emails { get; }


        class Telephone
        {

        }
        class Email
        {

        }
    }
}
