using System;

static class SubMain
{

    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // Creating an object
        Manager oManager = new Manager();

        //  Start Message Loop
        System.Windows.Forms.Application.Run();
    }

}
