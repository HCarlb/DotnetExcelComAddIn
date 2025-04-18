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

            case "/?":
                Console.WriteLine("Usage: Register.exe [/u] [/h]");
                Console.WriteLine("  'Register.exe' to register addin");
                Console.WriteLine("  'Register.exe /u' to unregister addin");
                break;

            default:
                PrintMessage("Invalid argument. Use /u to unregister the add-in.", MessageSeverity.Warning);
                break;
        }
    }
    catch (UnauthorizedAccessException ex)
    {
        PrintMessage($"Access denied: {ex.Message}", MessageSeverity.Error);
        PrintMessage("Please run the application as an administrator.", MessageSeverity.Warning);
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
    catch(FileNotFoundException ex)
    {
        PrintMessage($"File not found: {ex.Message}", MessageSeverity.Error);
    }
    catch (UnauthorizedAccessException ex)
    {
        PrintMessage($"Access denied: {ex.Message}", MessageSeverity.Error);
        PrintMessage("Please run the application as an administrator.", MessageSeverity.Warning);
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
