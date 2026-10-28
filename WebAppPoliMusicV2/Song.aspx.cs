using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                LoadLoggedUser();
                LoadSongs();
            }
        }

        /// <summary>
        /// Show the name and the photo of the authenticated user in the top bar.
        /// </summary>
        private void LoadLoggedUser()
        {
            Entities.User user = Session[Util.Constants.USER] as Entities.User;
            string userName = string.Empty;
            string userPhoto = "~/userImages/default.jpg";

            if (user != null)
            {
                userName = user.UserName;
                if (!string.IsNullOrEmpty(user.Photo))
                {
                    userPhoto = user.Photo;
                }
            }
            else if (Page.User != null && Page.User.Identity.IsAuthenticated)
            {
                // Fallback: the session expired but the forms authentication ticket is still valid.
                userName = Page.User.Identity.Name;
            }

            if (string.IsNullOrEmpty(userName))
            {
                Response.Redirect("~/Login.aspx", true);
                return;
            }

            lblUserName.Text = Server.HtmlEncode(userName);
            lblSidebarUser.Text = Server.HtmlEncode(userName);
            imgUserPhoto.ImageUrl = userPhoto;
        }

        /// <summary>
        /// Load the song list from TBL_SONG through the BLL layer.
        /// </summary>
        private void LoadSongs()
        {
            try
            {
                DataTable dtSongs = new BLL.Song(strConnString).Read();
                rptSongs.DataSource = dtSongs;
                rptSongs.DataBind();

                int totalSongs = dtSongs.Rows.Count;
                lblSongCount.Text = totalSongs == 1 ? "1 song" : totalSongs + " songs";
                pnlEmpty.Visible = totalSongs == 0;
                pnlMessage.Visible = false;
            }
            catch (Exception ex)
            {
                rptSongs.DataSource = new DataTable();
                rptSongs.DataBind();
                lblSongCount.Text = "0 songs";
                pnlEmpty.Visible = false;
                pnlMessage.Visible = true;
                lblMessage.Text = "The song list could not be loaded: " + Server.HtmlEncode(ex.Message);
            }
        }
    }
}
