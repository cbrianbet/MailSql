using System;
using Microsoft.Exchange.WebServices.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Windows.Forms;

namespace MailSql
{
    class Program
    {
        static void Main(string[] args)
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            ExchangeService _service;

            try
            {
                Console.WriteLine("Registering Exchange connection");

                _service = new ExchangeService
                {
                    Credentials = new WebCredentials("cbrianbet@outlook.com", "kaka10139")
                };
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return;
            }
            _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            // Prepare seperate class for writing email to the database
            try
            {
                Write2DB db = new Write2DB();

                Console.WriteLine("Reading mail");

                foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.Inbox, new ItemView(5)))
                {
                    db.Save(email, "Inbox");

                }

                foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.SentItems, new ItemView(5)))
                {
                    db.Save(email, "Sent Items");

                }

                Console.WriteLine("Success.\nPress any key to exit...");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error has occured: \n" + e.Message);
                Console.ReadKey();
            }
        }
    }
}
class Write2DB
{
    SqlConnection conn = null;

    public Write2DB()
    {
        Console.WriteLine("Connecting to SQL Server");
        try
        {
            NewMethod();
            conn.Open();
            Console.WriteLine("Database connected");
        }
        catch (System.Data.SqlClient.SqlException e)
        {
             Console.WriteLine(e);
        }
    }

    private void NewMethod()
    {
        conn = new SqlConnection("Data Source = (LocalDB)\\MSSQLLocalDB; Initial Catalog = chat; Persist Security Info = True; User ID = brian; Password =root");
    }
    
    public void Save(EmailMessage email, String type)
    {
        email.Load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));

        SqlCommand cmd = new SqlCommand("dbo.usp_servicedesk_savemail", conn)
        {
            CommandType = System.Data.CommandType.StoredProcedure,
            CommandTimeout = 1500
        };

        string recipients = "";

        foreach (EmailAddress emailAddress in email.CcRecipients)
        {
            recipients += ";" + emailAddress.Address.ToString();
        }
        cmd.Parameters.AddWithValue("@message_id", email.InternetMessageId);
        cmd.Parameters.AddWithValue("@from", email.From.Address);
        cmd.Parameters.AddWithValue("@body", email.TextBody.ToString());
        cmd.Parameters.AddWithValue("@cc", recipients);
        if (email.Subject.Length > 0)
        {
            cmd.Parameters.AddWithValue("@subject", email.Subject);
        }
        else
        {
            cmd.Parameters.AddWithValue("@subject", "No SUBJECT");
        }
        cmd.Parameters.AddWithValue("@received_time", email.DateTimeReceived.ToUniversalTime().ToString());
        cmd.Parameters.AddWithValue("@folder", type);

        recipients = "";
        foreach (EmailAddress emailAddress in email.ToRecipients)
        {
            recipients += ";" + emailAddress.Address.ToString();
        }
        cmd.Parameters.AddWithValue("@to", recipients);

        cmd.ExecuteNonQuery();
    }
    ~Write2DB()
    {
        Console.WriteLine("Disconnecting from SQLServer");
    }
}
