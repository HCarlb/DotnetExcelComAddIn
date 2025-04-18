using Register;
using COMContract;

var noArgMessage = "No arguments provided. Use /r to register or /u to unregister.";

Console.WriteLine();
PrintMessage($"################## Start of Registration of {ContractGuids.ProgId} ######################", MessageSeverity.Info);

if (args.Length > 0)
{
    try
    {
        switch (args[0].ToLower())
        {
            case "/r":
                Register.Registration.RegisterAddIn();
                break;

            case "/u":
                Register.Registration.UnregisterAddIn();
                break;

            default:
                PrintMessage(noArgMessage,MessageSeverity.Warning);
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
    Console.WriteLine(noArgMessage);
    //Registrations.RegisterAddIn();
}

PrintMessage($"################## End of Registration of {ContractGuids.ProgId} ######################", MessageSeverity.Info);
Console.WriteLine();


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
