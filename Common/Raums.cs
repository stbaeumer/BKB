using Microsoft.Data.SqlClient;

public class Raums : List<Raum>
{
    public int Anzahl { get; internal set; }

    public Raums()
    {
    }

    public Raums(int periode)
    {
        using (SqlConnection odbcConnection = new SqlConnection(Global.ConnectionStringUntis))
        {
            try
            {
                string queryString = @"SELECT Room.ROOM_ID, 
                                                Room.Name,  
                                                Room.Longname,
                                                Room.Capacity
                                                FROM Room
                                                WHERE (((Room.SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Room.SCHOOL_ID)=177659) AND  ((Room.TERM_ID)=" + periode + "))";

                SqlCommand odbcCommand = new SqlCommand(queryString, odbcConnection);
                odbcConnection.Open();
                SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                while (sqlDataReader.Read())
                {
                    Raum raum = new Raum()
                    {
                        IdUntis = sqlDataReader.GetInt32(0),
                        Raumnummer = Global.SafeGetString(sqlDataReader, 1)
                    };

                    this.Add(raum);
                }

                sqlDataReader.Close();
            }
            catch (Microsoft.Data.SqlClient.SqlException ex)
            {
                Console.WriteLine("Netzwerkbezogener oder instanzspezifischer Fehler beim Herstellen einer Verbindung mit SQL Server. Der Server wurde nicht gefunden, oder auf ihn kann nicht zugegriffen werden.");
                Console.ReadKey();
            }
            finally
            {
                odbcConnection.Close();
                Global.ZeileSchreiben("Räume", this.Count().ToString(), ConsoleColor.Green, ConsoleColor.White);
            }
        }
    }
}