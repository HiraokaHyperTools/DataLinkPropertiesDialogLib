using DataLinkPropertiesDialogLib;
using PromptDataLinkProperties.Properties;
using System;
using System.Collections.Generic;
using System.Text;

namespace PromptDataLinkProperties
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length >= 1 && args[0] == "new")
            {
                Console.Write(DataLinkProperties.PromptNew());
            }
            else if (args.Length >= 2 && args[0] == "edit")
            {
                Console.Write(DataLinkProperties.PromptEdit(args[1]));
            }
            else
            {
                Console.Error.Write(Resources.Usage);
                Environment.ExitCode = 1;
            }
        }
    }
}
