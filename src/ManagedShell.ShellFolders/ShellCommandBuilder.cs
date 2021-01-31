using System.Collections.Generic;
using ManagedShell.ShellFolders.Enums;

namespace ManagedShell.ShellFolders
{
    public class ShellCommandBuilder
    {
        public List<ShellCommand> Commands = new List<ShellCommand>();

        public void AddCommand(ShellCommand command)
        {
            Commands.Add(command);
        }

        public void AddSeparator()
        {
            Commands.Add(new ShellCommand {Flags = MFT.SEPARATOR, Label = string.Empty, UID = 0});
        }

        public void AddShellNewMenu()
        {
            Commands.Add(new ShellNewCommand());
        }
    }
}
