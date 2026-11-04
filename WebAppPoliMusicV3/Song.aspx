<%@ Page Title="Canciones" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <style type="text/css">
        .songs-page {
            --songs-background: #100d18;
            --songs-panel: #1d172b;
            --songs-panel-light: #28203a;
            --songs-purple: #a970ff;
            --songs-purple-dark: #7136b5;
            --songs-text: #f5efff;
            --songs-muted: #b9accb;
            background: linear-gradient(135deg, #100d18 0%, #211335 100%);
            border-radius: 24px;
            color: var(--songs-text);
            margin: 20px 0 35px;
            overflow: hidden;
            padding: 34px;
        }

        .songs-header {
            align-items: flex-end;
            display: flex;
            justify-content: space-between;
            margin-bottom: 28px;
        }

        .songs-kicker {
            color: var(--songs-purple);
            font-size: .78rem;
            font-weight: 700;
            letter-spacing: 2px;
            text-transform: uppercase;
        }

        .songs-title {
            color: var(--songs-text);
            font-size: 2.45rem;
            font-weight: 700;
            margin: 5px 0 8px;
        }

        .songs-subtitle, .songs-count {
            color: var(--songs-muted);
        }

        .songs-count {
            background: rgba(169, 112, 255, .13);
            border: 1px solid rgba(169, 112, 255, .35);
            border-radius: 30px;
            padding: 10px 17px;
        }

        .song-list {
            display: grid;
            gap: 12px;
        }

        .song-card {
            align-items: center;
            background: rgba(40, 32, 58, .92);
            border: 1px solid rgba(255, 255, 255, .07);
            border-radius: 16px;
            display: grid;
            gap: 18px;
            grid-template-columns: 48px minmax(150px, 1fr) 90px minmax(150px, 1fr) minmax(220px, 1.4fr);
            padding: 15px 18px;
            transition: border-color .2s ease, transform .2s ease;
        }

        .song-card:hover {
            border-color: var(--songs-purple);
            transform: translateY(-2px);
        }

        .song-number {
            align-items: center;
            background: linear-gradient(135deg, var(--songs-purple-dark), #a970ff);
            border-radius: 12px;
            display: flex;
            font-weight: 700;
            height: 42px;
            justify-content: center;
            width: 42px;
        }

        .song-name {
            color: var(--songs-text);
            font-size: 1.05rem;
            font-weight: 700;
        }

        .song-path, .song-plays {
            color: var(--songs-muted);
            font-size: .86rem;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .song-path-label {
            color: var(--songs-purple);
            display: block;
            font-size: .7rem;
            letter-spacing: 1px;
            margin-bottom: 3px;
            text-transform: uppercase;
        }

        .song-player {
            min-width: 0;
            width: 100%;
        }

        .song-player audio {
            height: 38px;
            width: 100%;
        }

        .songs-message {
            background: rgba(169, 112, 255, .12);
            border: 1px solid rgba(169, 112, 255, .35);
            border-radius: 14px;
            color: var(--songs-muted);
            padding: 22px;
        }

        @media (max-width: 800px) {
            .songs-page { padding: 24px 16px; }
            .songs-header { align-items: flex-start; flex-direction: column; gap: 16px; }
            .songs-title { font-size: 2rem; }
            .song-card { grid-template-columns: 42px 1fr; }
            .song-path, .song-plays, .song-player { grid-column: 2; }
        }
    </style>

    <section class="songs-page">
        <div class="songs-header">
            <div>
                <span class="songs-kicker">PoliMusic library</span>
                <h1 class="songs-title">Hola, <asp:Label ID="lblUserName" runat="server" /> </h1>
                <p class="songs-subtitle">Descubre y reproduce tus canciones favoritas.</p>
            </div>
            <asp:Label ID="lblSongCount" runat="server" CssClass="songs-count" />
        </div>

        <asp:Panel ID="pnlMessage" runat="server" CssClass="songs-message" Visible="false">
            <asp:Label ID="lblMessage" runat="server" />
        </asp:Panel>

        <asp:Panel ID="pnlEmpty" runat="server" CssClass="songs-message" Visible="false">
            Aún no hay canciones disponibles en la biblioteca.
        </asp:Panel>

        <div class="song-list">
            <asp:Repeater ID="rptSongs" runat="server">
                <ItemTemplate>
                    <article class="song-card">
                        <div class="song-number"><%# Container.ItemIndex + 1 %></div>
                        <div class="song-name"><%# Server.HtmlEncode(Eval("SONG_NAME").ToString()) %></div>
                        <div class="song-path">
                            <span class="song-path-label">Ruta</span>
                            <%# Server.HtmlEncode(Eval("SONG_PATH").ToString()) %>
                        </div>
                        <div class="song-plays">
                            <span class="song-path-label">Reproducciones</span>
                            <%# Eval("PLAYS") %>
                        </div>
                        <div class="song-player">
                            <audio controls preload="none">
                                <source src='<%# ResolveUrl(Eval("SONG_PATH").ToString()) %>' type="audio/mpeg" />
                                Tu navegador no soporta el reproductor de audio.
                            </audio>
                        </div>
                    </article>
                </ItemTemplate>
            </asp:Repeater>
        </div>
    </section>
</asp:Content>
