namespace ReportGenerator.App;

static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length == 2 && args[0] == "--create-sample-template")
        {
            ReportGenerator.PowerPoint.SampleTemplateFactory.Create(args[1]);
            return;
        }

        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
    }
}
