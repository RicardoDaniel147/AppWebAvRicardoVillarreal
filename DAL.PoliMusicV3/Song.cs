using System.Data;
using System.Data.SqlClient;

namespace DAL.PoliMusicV3
{
    public class Song
    {
        private readonly string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public DataTable Read()
        {
            DataTable dtSongs = new DataTable();

            using (SqlConnection con = new SqlConnection(strConnectionString))
            using (SqlCommand cmd = new SqlCommand(
                "SELECT [ID_SONG], [SONG_NAME], [SONG_PATH], [PLAYS] " +
                "FROM [TBL_SONG] ORDER BY [SONG_NAME]", con))
            using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
            {
                sda.Fill(dtSongs);
            }

            return dtSongs;
        }
    }
}
