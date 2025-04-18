using Register;


if (args.Length > 0)
{
    try
    {
        switch (args[0].ToLower())
        {
            case "/u":
                Registration.UnregisterAddIn();
                break;

            default:
                PrintMessage("Invalid argument. Use /u to unregister the add-in.", MessageSeverity.Warning);
                break;
        }
    }
    catch (Exception ex)
    {
        PrintMessage($"An error occurred: {ex.Message}", MessageSeverity.Error);  
    }

}
else
{
    try
    {
        Registration.RegisterAddIn();
    }
    catch (Exception ex)
    {
        PrintMessage($"An error occurred: {ex.Message}", MessageSeverity.Error);
    }
}


static void PrintMessage(string message, MessageSeverity severity)
    {
        Console.ForegroundColor = severity switch
        {
            MessageSeverity.Warning => ConsoleColor.Yellow,
            MessageSeverity.Error => ConsoleColor.Red,
            MessageSeverity.Success => ConsoleColor.Green,
            _ => ConsoleColor.Cyan,
        };
        Console.WriteLine(message);
        Console.ResetColor();
    }
