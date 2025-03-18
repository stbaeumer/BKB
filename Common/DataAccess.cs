using System.Globalization;
using MySqlConnector;

namespace Common;

public class DataAccess
{
    private readonly string _connetionString;

    public DataAccess()
    {
        var server = Global.MariaIp;
        var database = Global.MariaDb;
        var user = Global.MariaUser;
        var port = Global.MariaPort;
        var password = Global.MariaPw;

        _connetionString = $"Server={server};Database={database};User ID={user};Password={password};Port={port};";
    }

    internal DateTime GetSchildVersion()
    {
        using (var connection = new MySqlConnection(_connetionString))
        {
            try
            {
                DateTime schildVersion;
                const string query = @"SELECT Schild_Verwaltung.Version FROM Schild_Verwaltung;";
                connection.Open();

                using (var command = new MySqlCommand(query, connection))
                {
                    schildVersion = Convert.ToDateTime(command.ExecuteScalar()?.ToString());
                }

                Global.ZeileSchreiben( "Installierte SchILD-Version", schildVersion.ToString(CultureInfo.InvariantCulture), ConsoleColor.White, ConsoleColor.Black);
                return schildVersion;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Kein zulässiges Kennwort"))
                {
                    Console.WriteLine(ex.Message + " Sie müssen das Kennwort eingeben, das herstellerseitig gesetzt wurde. Datenbankkennwort eingeben!");
                }
                else if (ex.Message.Contains("Parameter"))
                {
                    Console.WriteLine(
                        "Fehler beim Lesen der Datenbank. Dieser Fehler könnte durch eine syntaktisch falsche Anfrage an die Datenbank ausgelöst worden sein: " + ex.Message);
                }
                else
                {
                    Console.WriteLine(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }
        }

        return new DateTime();
    }

    internal List<Student> GetStudents(Klasse klasse)
    {
        var students = new List<Student>();

        using var connection = new MySqlConnection(_connetionString);
        try
        {
            //string queryString = @"SELECT Schueler.ID, Schueler.Klasse, Schueler.FotoVorhanden, LENGTH(SchuelerFotos.Foto) AS Laenge, Schueler.SchulnrEigner FROM Schueler LEFT JOIN SchuelerFotos ON Schueler.ID = SchuelerFotos.Schueler_ID WHERE Schueler.Geloescht = '-' AND Schueler.Status = 2 AND Schueler.AktSchuljahr = " + aktSj + @" ORDER BY Schueler.Klasse, Schueler.Name, Schueler.Vorname;";
            var queryString = """
                              SELECT 
                                  Schueler.ID, 
                                  Klassen.Klasse,     
                                  Schueler.FotoVorhanden, 
                                  LENGTH(SchuelerFotos.FotoBase64) AS Laenge,
                                  Schueler.Vorname,
                                  Schueler.Name
                              FROM 
                                  Schueler
                                                                      LEFT JOIN 
                                                                          SchuelerFotos ON Schueler.ID = SchuelerFotos.Schueler_ID
                                                                      LEFT JOIN 
                                                                          SchuelerLernabschnittsdaten ON Schueler.ID = SchuelerLernabschnittsdaten.Schueler_ID
                                                                      LEFT JOIN 
                                                                          Klassen ON SchuelerLernabschnittsdaten.Klassen_ID = Klassen.ID
                                                                      WHERE 
                                                                          Schueler.Geloescht = '-' 
                                                                          AND Schueler.Status = 2
                                                                          AND Klassen.Klasse = '
                              """ + klasse.Name + @"'
                                        ORDER BY 
                                            Klassen.Klasse, Schueler.Name, Schueler.Vorname;
                                        ";

            MySqlCommand mySqlCommand = new MySqlCommand(queryString, connection);
            connection.Open();
            MySqlDataReader mySqlDataReader = mySqlCommand.ExecuteReader();

            while (mySqlDataReader.Read())
            {
                var student = new Student
                {
                    IdSchildInt = mySqlDataReader.GetInt32(0),
                    KlasseString = Global.SafeGetString(mySqlDataReader, 1),
                    FotoVorhanden = Global.SafeGetString(mySqlDataReader, 2) == "-" ? false : true,
                    FotoBinary = mySqlDataReader.GetValue(3) == DBNull.Value ? false : true,
                    Vorname = Global.SafeGetString(mySqlDataReader, 4),
                    Nachname = Global.SafeGetString(mySqlDataReader, 5)
                };

                switch (student.FotoBinary)
                {
                    case true when !student.FotoVorhanden:
                    {
                        using var updateConnection = new MySqlConnection(_connetionString);
                        mySqlCommand = new MySqlCommand();
                        mySqlCommand = updateConnection.CreateCommand();
                        mySqlCommand.CommandText = $"UPDATE Schueler SET FotoVorhanden = '+' WHERE ID = @id";
                        mySqlCommand.Parameters.AddWithValue("@id", student.IdSchild);
                        updateConnection.Open();
                        mySqlCommand.ExecuteNonQuery();
                        updateConnection.Close();
                        student.FotoVorhanden = true;

                        break;
                    }
                    case false when student.FotoVorhanden:
                    {
                        using var updateConnection = new MySqlConnection(_connetionString);
                        mySqlCommand = new MySqlCommand();
                        mySqlCommand = updateConnection.CreateCommand();
                        mySqlCommand.CommandText = $"UPDATE Schueler SET FotoVorhanden = '-' WHERE ID = @id";
                        mySqlCommand.Parameters.AddWithValue("@id", student.IdSchild);
                        updateConnection.Open();
                        mySqlCommand.ExecuteNonQuery();
                        updateConnection.Close();
                        student.FotoVorhanden = true;

                        break;
                    }
                }

                students.Add(student);
            }

            mySqlDataReader.Close();
            Global.ZeileSchreiben( "SuS der Klasse " + klasse + " in der SchILD-DB", students.Count().ToString(),
                ConsoleColor.Yellow, ConsoleColor.White);
            return students;
        }

        catch (Exception ex)
        {
            if (ex.Message.Contains("Parameter"))
            {
                Console.WriteLine(
                    "Fehler beim Lesen der Datenbank. Dieser Fehler könnte durch eine syntaktisch falsche Anfrage an die Datenbank ausgelöst worden sein: " +
                    ex.Message);
            }

            Console.WriteLine(ex.Message);
        }
        finally
        {
            connection.Close();
        }

        return null!;
    }

    public string SafeGetString(MySqlDataReader reader, int colIndex)
    {
        return !reader.IsDBNull(colIndex) ? reader.GetString(colIndex) : string.Empty;
    }

    internal void InsertImage(Student student)
    {
        using var connection = new MySqlConnection(_connetionString);
        try
        {
            connection.Open();
            var mySqlCommand = new MySqlCommand();
            mySqlCommand.Parameters.AddWithValue("@id", student.IdSchildInt);
            mySqlCommand.Parameters.AddWithValue("@foto", student.Foto);
            mySqlCommand = connection.CreateCommand();
            mySqlCommand.CommandText = "UPDATE Schueler SET FotoVorhanden = '+' WHERE ID = " + student.IdSchildInt;
            mySqlCommand.ExecuteNonQuery();
            student.FotoVorhanden = true;
            mySqlCommand = new MySqlCommand();
            mySqlCommand = connection.CreateCommand();
            mySqlCommand.CommandText = "INSERT INTO SchuelerFotos (Schueler_ID, FotoBase64) VALUES (" +
                                       student.IdSchildInt + ", '" + student.Foto + "')";
            mySqlCommand.ExecuteNonQuery();
            student.FotoBinary = true;
            Console.Write("neu angelegt.");
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("Parameter"))
            {
                Console.WriteLine(
                    "Fehler beim Lesen der Datenbank. Dieser Fehler könnte durch eine syntaktisch falsche Anfrage an die Datenbank ausgelöst worden sein: " +
                    ex.Message);
            }

            Console.WriteLine(ex.Message);
        }
        finally
        {
            connection.Close();
        }
    }

    internal void DeleteImage(Student student)
    {
        using var connection = new MySqlConnection(_connetionString);
        try
        {
            connection.Open();
            var mySqlCommand = connection.CreateCommand();
            mySqlCommand.CommandText = $"DELETE FROM SchuelerFotos WHERE SchuelerFotos.Schueler_ID = @id";
            mySqlCommand.Parameters.AddWithValue("@id", student.IdSchild);
            mySqlCommand.ExecuteNonQuery();
            mySqlCommand = new MySqlCommand();
            mySqlCommand = connection.CreateCommand();
            mySqlCommand.CommandText = "UPDATE Schueler SET FotoVorhanden = '-' WHERE ID = @id";
            mySqlCommand.Parameters.AddWithValue("@id", student.IdSchild);
            mySqlCommand.ExecuteNonQuery();
            student.FotoVorhanden = false;
            Console.Write("Foto gelöscht ...");
        }

        catch (Exception ex)
        {
            if (ex.Message.Contains("Parameter"))
            {
                Console.WriteLine(
                    "Fehler beim Lesen der Datenbank. Dieser Fehler könnte durch eine syntaktisch falsche Anfrage an die Datenbank ausgelöst worden sein: " +
                    ex.Message);
            }

            Console.WriteLine(ex.Message);
        }
        finally
        {
            connection.Close();
        }
    }
}