using AutocadCommandDWG;
using System.Data.SQLite;

internal static class CommandHelpers
{

    public static Block GetBlockInfo(SQLiteConnection connection, string id)
    {
        string selectQuery = "SELECT * FROM Parts WHERE ID = @Id;";

        using (SQLiteCommand command = new SQLiteCommand(selectQuery, connection))
        {
            command.Parameters.AddWithValue("@Id", id);

            using (SQLiteDataReader reader = command.ExecuteReader())
            {
                if (reader.Read())
                {
                    Block block = new Block
                    {
                        Id = reader.GetString(reader.GetOrdinal("ID")),
                        Weight = reader.GetDouble(reader.GetOrdinal("Weight")),
                        Diameter = reader.GetInt32(reader.GetOrdinal("Diameter")),
                        Fullname = reader.GetString(reader.GetOrdinal("FullNameTemplate"))
                    };

                    return block;
                }
            }
        }

        return null;
    }
}