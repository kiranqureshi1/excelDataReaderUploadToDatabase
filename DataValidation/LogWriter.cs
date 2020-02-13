
using System;
using System.IO;
using System.Reflection;


public class LogWriter
{
    private string m_exePath = string.Empty;
    private string logMessage;
    private string fileName;

    public LogWriter(string logMessage, string fileName)
    {
        this.logMessage = logMessage;
        this.fileName = fileName;
        //LogWrite(logMessage);
    }
    public void LogWrite()
    {
        m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        try
        {
            using (StreamWriter w = File.AppendText(m_exePath + "\\" + "ErrorLog.txt"))
            {
                Log(logMessage, w, fileName);
            }
        }
        catch (Exception ex)
        {
        }
    }

    //Writes the actual log message in this format shown below in this method.
    public void Log(string logMessage, TextWriter txtWriter, string fileName)
    {
        try
        {
            txtWriter.Write("\r\nLog Entry : ");
            txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
            DateTime.Now.ToLongDateString());
            //txtWriter.WriteLine("  :");
            txtWriter.WriteLine("  :{0}", logMessage);
            txtWriter.WriteLine("  :{0}", fileName);
            txtWriter.WriteLine();
            txtWriter.WriteLine("-------------------------------");
        }
        catch (Exception ex)
        {
        }
    }
}
