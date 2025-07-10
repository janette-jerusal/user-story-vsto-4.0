partial class ThisAddIn
{
    private void InternalStartup()
    {
        this.Startup += new System.EventHandler(ThisAddIn_Startup);
        this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
    }
}
