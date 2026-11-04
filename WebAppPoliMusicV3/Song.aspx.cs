using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        private readonly string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                LoadLoggedUser();
                LoadSongs();
            }
        }

        private void LoadLoggedUser()
        {
            EntityLayer.PoliMusicV3.User user = Session[EntityLayer.PoliMusicV3.Util.Constants.USER] as EntityLayer.PoliMusicV3.User;
            string userName = user == null ? string.Empty : user.UserName;

            if (string.IsNullOrEmpty(userName) && Page.User != null && Page.User.Identity.IsAuthenticated)
            {
                userName = Page.User.Identity.Name;
            }

            if (string.IsNullOrEmpty(userName))
            {
                Response.Redirect("~/Login.aspx", true);
                return;
            }

            lblUserName.Text = userName;
        }

        private void LoadSongs()
        {
            try
            {
                DataTable dtSongs = new BLL.PoliMusicV3.Song(strConnString).Read();
                rptSongs.DataSource = dtSongs;
                rptSongs.DataBind();
                lblSongCount.Text = dtSongs.Rows.Count == 1 ? "1 canción" : dtSongs.Rows.Count + " canciones";
                pnlEmpty.Visible = dtSongs.Rows.Count == 0;
                pnlMessage.Visible = false;
            }
            catch (Exception ex)
            {
                pnlEmpty.Visible = false;
                pnlMessage.Visible = true;
                lblMessage.Text = "No fue posible cargar las canciones: " + Server.HtmlEncode(ex.Message);
            }
        }
    }
}