using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
{
    /// <summary>
    /// Business layer for the songs. It only validates the information and
    /// delegates the database work to the DAL layer.
    /// </summary>
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public int Save(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Save(song);
        }

        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }

        public DataTable ReadById(int idSong)
        {
            return new DAL.Song(strConnectionString).ReadById(idSong);
        }

        public int Update(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Update(song);
        }

        public int Delete(int idSong)
        {
            return new DAL.Song(strConnectionString).Delete(idSong);
        }

        public int IncreasePlays(int idSong)
        {
            return new DAL.Song(strConnectionString).IncreasePlays(idSong);
        }

        public int GetPlays(int idSong)
        {
            return new DAL.Song(strConnectionString).GetPlays(idSong);
        }

        /// <summary>
        /// Increment the play counter and return the new value, so the web page
        /// can refresh the "plays" column without a full postback.
        /// </summary>
        public int RegisterPlay(int idSong)
        {
            if (idSong <= 0)
            {
                return 0;
            }
            int affectedRows = new DAL.Song(strConnectionString).IncreasePlays(idSong);
            if (affectedRows == 0)
            {
                return 0;
            }
            return new DAL.Song(strConnectionString).GetPlays(idSong);
        }

        public string ValidationsRequired(Entities.Song song)
        {
            if (string.IsNullOrEmpty(song.Name))
            {
                return "Song name is required!";
            }
            if (string.IsNullOrEmpty(song.Path))
            {
                return "Song path is required!";
            }
            if (song.Plays < 0)
            {
                return "Plays can not be a negative value!";
            }
            return string.Empty;
        }
    }
}
