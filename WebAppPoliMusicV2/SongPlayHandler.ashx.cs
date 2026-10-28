using System;
using System.Configuration;
using System.Web;

namespace WebAppPoliMusicV2
{
    /// <summary>
    /// Increment the PLAYS column of a song without a full page postback, so the
    /// audio player keeps playing while the counter is refreshed.
    /// Usage: SongPlayHandler.ashx?idSong=1
    /// </summary>
    public class SongPlayHandler : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            context.Response.Cache.SetCacheability(HttpCacheability.NoCache);

            if (context.User == null || !context.User.Identity.IsAuthenticated)
            {
                context.Response.StatusCode = 401;
                context.Response.Write("0");
                return;
            }

            int idSong;
            if (!int.TryParse(context.Request["idSong"], out idSong) || idSong <= 0)
            {
                context.Response.StatusCode = 400;
                context.Response.Write("0");
                return;
            }

            string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
            int plays = new BLL.Song(strConnString).RegisterPlay(idSong);
            context.Response.Write(plays.ToString());
        }

        public bool IsReusable
        {
            get { return false; }
        }
    }
}
