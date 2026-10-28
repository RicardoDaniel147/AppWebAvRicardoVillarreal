<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>PoliMusic | Song list</title>
    <style type="text/css">
        :root {
            --bg: #0b0b0f;
            --bg-elev: #121218;
            --panel: #16161d;
            --panel-2: #1e1e27;
            --line: #26262f;
            --text: #f5f5f7;
            --muted: #9a9aa8;
            --purple: #8b5cf6;
            --purple-2: #a78bfa;
            --purple-deep: #5b21b6;
        }

        * { box-sizing: border-box; }

        html, body { height: 100%; }

        body {
            margin: 0;
            background: var(--bg);
            color: var(--text);
            font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
            -webkit-font-smoothing: antialiased;
        }

        /* ---------- Layout ---------- */
        .app {
            display: flex;
            min-height: 100vh;
            padding-bottom: 96px;
        }

        .sidebar {
            width: 246px;
            flex: 0 0 246px;
            background: #08080b;
            border-right: 1px solid var(--line);
            padding: 22px 16px;
            position: sticky;
            top: 0;
            height: calc(100vh - 92px);
            display: flex;
            flex-direction: column;
            gap: 26px;
        }

        .brand {
            display: flex;
            align-items: center;
            gap: 11px;
            padding: 0 6px;
        }

        .brand-mark {
            width: 38px;
            height: 38px;
            border-radius: 11px;
            display: grid;
            place-items: center;
            color: #fff;
            background: linear-gradient(135deg, var(--purple-2), var(--purple-deep));
            box-shadow: 0 6px 18px rgba(139, 92, 246, .35);
        }

        .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
        .brand-text strong { font-size: 16px; letter-spacing: .2px; }
        .brand-text span { font-size: 10.5px; color: var(--muted); letter-spacing: .12em; text-transform: uppercase; }

        .nav { display: flex; flex-direction: column; gap: 4px; }

        .nav-title {
            font-size: 10.5px;
            letter-spacing: .14em;
            text-transform: uppercase;
            color: #6f6f7d;
            padding: 0 10px 6px;
        }

        .nav-item {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 10px;
            border-radius: 9px;
            color: #c9c9d4;
            font-size: 13.5px;
            text-decoration: none;
            transition: background .15s ease, color .15s ease;
        }

        .nav-item:hover { background: #14141b; color: #fff; }

        .nav-item.active {
            background: linear-gradient(90deg, rgba(139, 92, 246, .22), rgba(139, 92, 246, .04));
            color: #fff;
            box-shadow: inset 2px 0 0 var(--purple);
        }

        .nav-item svg { flex: 0 0 auto; opacity: .9; }

        .sidebar-card {
            margin-top: auto;
            background: linear-gradient(150deg, #221a33, #14141a);
            border: 1px solid #2c2440;
            border-radius: 13px;
            padding: 15px;
            font-size: 12px;
            color: var(--muted);
            line-height: 1.5;
        }

        .sidebar-card b { color: #e9e2ff; display: block; margin-bottom: 4px; font-size: 12.5px; }

        .content { flex: 1 1 auto; min-width: 0; display: flex; flex-direction: column; }

        /* ---------- Top bar ---------- */
        .topbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 14px 32px;
            background: rgba(11, 11, 15, .88);
            border-bottom: 1px solid var(--line);
            position: sticky;
            top: 0;
            z-index: 10;
            backdrop-filter: blur(12px);
        }

        .user-chip { display: flex; align-items: center; gap: 12px; }

        .avatar {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            object-fit: cover;
            background: var(--panel-2);
            border: 2px solid rgba(139, 92, 246, .55);
        }

        .user-text { display: flex; flex-direction: column; line-height: 1.25; }
        .user-hello { font-size: 11px; text-transform: uppercase; letter-spacing: .12em; color: var(--muted); }
        .user-name { font-size: 15px; font-weight: 600; color: #fff; }

        .btn-logout {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 9px 18px;
            border-radius: 999px;
            border: 1px solid rgba(167, 139, 250, .45);
            color: #d9cffb;
            font-size: 13px;
            font-weight: 600;
            text-decoration: none;
            transition: background .18s ease, color .18s ease, border-color .18s ease, transform .18s ease;
        }

        .btn-logout:hover {
            background: linear-gradient(135deg, var(--purple), var(--purple-deep));
            border-color: transparent;
            color: #fff;
            transform: translateY(-1px);
        }

        /* ---------- Hero ---------- */
        .hero {
            position: relative;
            padding: 40px 32px 34px;
            background:
                radial-gradient(1000px 380px at 8% -30%, rgba(139, 92, 246, .34), transparent 62%),
                linear-gradient(160deg, #23222c 0%, #17171d 52%, #101015 100%);
            border-bottom: 1px solid var(--line);
            overflow: hidden;
        }

        .hero::after {
            content: "";
            position: absolute;
            right: -90px;
            top: -110px;
            width: 320px;
            height: 320px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(167, 139, 250, .22), transparent 68%);
            pointer-events: none;
        }

        .hero-tag {
            display: inline-block;
            font-size: 10.5px;
            letter-spacing: .16em;
            text-transform: uppercase;
            color: #cbb8ff;
            background: rgba(139, 92, 246, .16);
            border: 1px solid rgba(139, 92, 246, .3);
            padding: 5px 12px;
            border-radius: 999px;
            margin-bottom: 16px;
        }

        .hero h1 {
            margin: 0 0 8px;
            font-size: 46px;
            line-height: 1.05;
            letter-spacing: -.5px;
            font-weight: 800;
        }

        .hero p { margin: 0; color: var(--muted); font-size: 13.5px; }

        .hero-actions { margin-top: 24px; display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }

        .btn-play-all {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            border: 0;
            cursor: pointer;
            padding: 12px 26px;
            border-radius: 999px;
            font-size: 14px;
            font-weight: 700;
            color: #fff;
            background: linear-gradient(135deg, var(--purple), var(--purple-deep));
            box-shadow: 0 10px 24px rgba(109, 40, 217, .38);
            transition: transform .16s ease, box-shadow .16s ease;
        }

        .btn-play-all:hover { transform: translateY(-2px) scale(1.01); box-shadow: 0 14px 30px rgba(109, 40, 217, .5); }

        .hero-meta { font-size: 12px; color: var(--muted); letter-spacing: .04em; }

        /* ---------- Toolbar ---------- */
        .toolbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 22px 32px 10px;
        }

        .toolbar h2 {
            margin: 0;
            font-size: 17px;
            font-weight: 700;
            letter-spacing: .2px;
        }

        .search {
            position: relative;
            display: flex;
            align-items: center;
        }

        .search svg { position: absolute; left: 12px; color: var(--muted); pointer-events: none; }

        .search input {
            width: 260px;
            max-width: 46vw;
            padding: 10px 14px 10px 36px;
            border-radius: 999px;
            border: 1px solid var(--line);
            background: var(--panel);
            color: var(--text);
            font-size: 13px;
            outline: none;
            transition: border-color .16s ease, background .16s ease;
        }

        .search input:focus { border-color: rgba(139, 92, 246, .6); background: var(--panel-2); }
        .search input::placeholder { color: #6f6f7d; }

        /* ---------- Song list ---------- */
        .list-wrap { padding: 6px 24px 40px; flex: 1 1 auto; }

        .list-head,
        .song-row {
            display: grid;
            grid-template-columns: 40px minmax(190px, 1.7fr) 112px minmax(110px, 1fr) 74px 62px 48px;
            gap: 14px;
            align-items: center;
        }

        .list-head {
            padding: 10px 14px;
            font-size: 10.5px;
            letter-spacing: .13em;
            text-transform: uppercase;
            color: #71717f;
            border-bottom: 1px solid var(--line);
            margin-bottom: 6px;
        }

        .list-head .c-plays, .list-head .c-time { text-align: right; }
        .list-head .c-action { text-align: center; }

        .song-row {
            padding: 9px 14px;
            border-radius: 10px;
            border: 1px solid transparent;
            cursor: pointer;
            transition: background .15s ease, border-color .15s ease;
        }

        .song-row:hover { background: var(--panel-2); }

        .song-row.current {
            background: linear-gradient(90deg, rgba(139, 92, 246, .18), rgba(139, 92, 246, .03));
            border-color: rgba(139, 92, 246, .28);
        }

        .c-index { position: relative; display: flex; align-items: center; justify-content: center; height: 26px; }

        .row-number { font-size: 13px; color: var(--muted); font-variant-numeric: tabular-nums; }

        .eq { display: none; gap: 2.5px; align-items: flex-end; height: 15px; }
        .eq i {
            display: block;
            width: 3px;
            border-radius: 2px;
            background: var(--purple-2);
            animation: bounce 900ms ease-in-out infinite;
        }
        .eq i:nth-child(1) { height: 7px; animation-delay: -200ms; }
        .eq i:nth-child(2) { height: 14px; animation-delay: -600ms; }
        .eq i:nth-child(3) { height: 9px; animation-delay: -400ms; }

        @keyframes bounce {
            0%, 100% { transform: scaleY(.45); }
            50% { transform: scaleY(1); }
        }

        .song-row.is-playing .row-number { display: none; }
        .song-row.is-playing .eq { display: flex; }

        .c-title { display: flex; align-items: center; gap: 13px; min-width: 0; }

        .cover {
            width: 42px;
            height: 42px;
            flex: 0 0 42px;
            border-radius: 7px;
            display: grid;
            place-items: center;
            color: rgba(255, 255, 255, .92);
            background: linear-gradient(140deg, #4b3a6b, #241f33 65%, #1a1a22);
            border: 1px solid rgba(167, 139, 250, .22);
        }

        .song-row.current .cover {
            background: linear-gradient(140deg, var(--purple), var(--purple-deep));
            border-color: transparent;
        }

        .title-text { display: flex; flex-direction: column; min-width: 0; line-height: 1.3; }

        .song-name {
            font-size: 14px;
            font-weight: 600;
            color: #f2f2f6;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .song-row.current .song-name { color: #d8cbff; }

        .song-sub {
            font-size: 11.5px;
            color: var(--muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 999px;
            font-size: 10.5px;
            letter-spacing: .05em;
            text-transform: uppercase;
            color: #cbb8ff;
            background: rgba(139, 92, 246, .13);
            border: 1px solid rgba(139, 92, 246, .26);
            white-space: nowrap;
        }

        .c-path {
            font-family: Consolas, "Courier New", monospace;
            font-size: 11.5px;
            color: #8b8b99;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .c-plays { text-align: right; }

        .plays-value {
            display: inline-block;
            min-width: 26px;
            padding: 3px 9px;
            border-radius: 7px;
            font-size: 12px;
            color: #d5d5df;
            background: #1a1a22;
            border: 1px solid #2a2a34;
            font-variant-numeric: tabular-nums;
        }

        .c-time { text-align: right; }

        .time-value {
            font-size: 12px;
            color: var(--muted);
            font-variant-numeric: tabular-nums;
        }

        .c-action { text-align: center; }

        .row-play {
            width: 36px;
            height: 36px;
            border-radius: 50%;
            border: 0;
            cursor: pointer;
            display: grid;
            place-items: center;
            color: #fff;
            background: linear-gradient(135deg, var(--purple), var(--purple-deep));
            box-shadow: 0 6px 14px rgba(91, 33, 182, .4);
            opacity: 0;
            transform: scale(.86);
            transition: opacity .16s ease, transform .16s ease;
        }

        .song-row:hover .row-play,
        .song-row.current .row-play { opacity: 1; transform: scale(1); }

        .row-play:hover { transform: scale(1.07) !important; }

        .row-play .i-pause { display: none; }
        .song-row.is-playing .row-play .i-pause { display: block; }
        .song-row.is-playing .row-play .i-play { display: none; }

        /* ---------- Empty states ---------- */
        .empty {
            margin: 26px 8px;
            padding: 34px;
            text-align: center;
            border: 1px dashed #303040;
            border-radius: 14px;
            color: var(--muted);
            font-size: 13.5px;
            background: #101016;
        }

        .empty strong { display: block; color: #e6e6ee; font-size: 15px; margin-bottom: 6px; }

        #noResults { display: none; }

        .alert {
            margin: 0 32px;
            padding: 11px 16px;
            border-radius: 10px;
            background: rgba(190, 24, 93, .13);
            border: 1px solid rgba(190, 24, 93, .4);
            color: #fbcfe8;
            font-size: 13px;
        }

        /* ---------- Player bar ---------- */
        .player {
            position: fixed;
            left: 0;
            right: 0;
            bottom: 0;
            height: 92px;
            z-index: 30;
            display: grid;
            grid-template-columns: minmax(180px, 1fr) minmax(300px, 620px) minmax(150px, 1fr);
            align-items: center;
            gap: 20px;
            padding: 0 22px;
            background: rgba(13, 13, 18, .97);
            border-top: 1px solid #2a2a36;
            box-shadow: 0 -12px 32px rgba(0, 0, 0, .5);
            backdrop-filter: blur(14px);
        }

        .np { display: flex; align-items: center; gap: 12px; min-width: 0; }

        .np-cover {
            width: 54px;
            height: 54px;
            flex: 0 0 54px;
            border-radius: 8px;
            display: grid;
            place-items: center;
            color: #fff;
            background: linear-gradient(140deg, #3a2c52, #1b1b23);
            border: 1px solid #33294a;
        }

        .np-cover.active {
            background: linear-gradient(140deg, var(--purple-2), var(--purple-deep));
            border-color: transparent;
            box-shadow: 0 0 20px rgba(139, 92, 246, .45);
        }

        .np-text { display: flex; flex-direction: column; min-width: 0; line-height: 1.35; }

        .np-title {
            font-size: 13.5px;
            font-weight: 600;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .np-genre {
            font-size: 11.5px;
            color: var(--muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .player-center { display: flex; flex-direction: column; align-items: center; gap: 7px; }

        .controls { display: flex; align-items: center; gap: 16px; }

        .ctrl {
            background: none;
            border: 0;
            padding: 0;
            cursor: pointer;
            color: #d3d3de;
            display: grid;
            place-items: center;
            transition: color .15s ease, transform .15s ease;
        }

        .ctrl:hover { color: #fff; transform: scale(1.08); }

        .ctrl-main {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            color: #fff;
            background: linear-gradient(135deg, var(--purple), var(--purple-deep));
            box-shadow: 0 8px 18px rgba(91, 33, 182, .45);
        }

        .ctrl-main:hover { color: #fff; transform: scale(1.06); }
        .ctrl-main .i-pause { display: none; }
        .ctrl-main.is-playing .i-pause { display: block; }
        .ctrl-main.is-playing .i-play { display: none; }

        .progress { display: flex; align-items: center; gap: 10px; width: 100%; }

        .time {
            font-size: 11px;
            color: var(--muted);
            min-width: 38px;
            text-align: center;
            font-variant-numeric: tabular-nums;
        }

        .track {
            position: relative;
            flex: 1 1 auto;
            height: 5px;
            border-radius: 999px;
            background: #2e2e3a;
            cursor: pointer;
        }

        .track .fill {
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 0;
            border-radius: 999px;
            background: linear-gradient(90deg, var(--purple-deep), var(--purple-2));
        }

        .track .knob {
            position: absolute;
            top: 50%;
            left: 0;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: #fff;
            transform: translate(-50%, -50%);
            opacity: 0;
            transition: opacity .15s ease;
            box-shadow: 0 0 8px rgba(167, 139, 250, .8);
        }

        .track:hover .knob { opacity: 1; }

        .player-right { display: flex; align-items: center; justify-content: flex-end; gap: 10px; color: var(--muted); }

        .volume {
            -webkit-appearance: none;
            appearance: none;
            width: 110px;
            height: 4px;
            border-radius: 999px;
            background: #2e2e3a;
            outline: none;
        }

        .volume::-webkit-slider-thumb {
            -webkit-appearance: none;
            appearance: none;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: #fff;
            cursor: pointer;
        }

        .volume::-moz-range-thumb {
            width: 12px;
            height: 12px;
            border: 0;
            border-radius: 50%;
            background: #fff;
            cursor: pointer;
        }

        /* ---------- Responsive ---------- */
        @media (max-width: 1080px) {
            .list-head, .song-row { grid-template-columns: 36px minmax(170px, 1.6fr) 100px minmax(90px, 1fr) 66px 56px 46px; gap: 10px; }
        }

        @media (max-width: 900px) {
            .sidebar { display: none; }
            .hero h1 { font-size: 34px; }
            .list-head .c-path, .song-row .c-path { display: none; }
            .list-head, .song-row { grid-template-columns: 34px minmax(150px, 2fr) 96px 64px 54px 44px; }
            .player { grid-template-columns: 1fr; height: 78px; padding: 8px 16px; }
            .player-right, .np { display: none; }
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="app">
            <!-- ============ Sidebar ============ -->
            <aside class="sidebar">
                <div class="brand">
                    <div class="brand-mark">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                            <path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z" />
                        </svg>
                    </div>
                    <div class="brand-text">
                        <strong>PoliMusic</strong>
                        <span>Streaming</span>
                    </div>
                </div>

                <nav class="nav">
                    <div class="nav-title">Menu</div>
                    <a class="nav-item active" href="Song.aspx">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M4 5h16v2H4zm0 6h16v2H4zm0 6h10v2H4z" /></svg>
                        Song list
                    </a>
                    <a class="nav-item" href="#songs">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M10 4a6 6 0 1 0 3.5 10.9l4.3 4.3 1.4-1.4-4.3-4.3A6 6 0 0 0 10 4zm0 2a4 4 0 1 1 0 8 4 4 0 0 1 0-8z" /></svg>
                        Search
                    </a>
                    <a class="nav-item" href="#songs">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M4 4h6v6H4zm10 0h6v6h-6zM4 14h6v6H4zm10 0h6v6h-6z" /></svg>
                        Your library
                    </a>
                </nav>

                <div class="sidebar-card">
                    <b>Now streaming</b>
                    <asp:Label ID="lblSidebarUser" runat="server" Text=""></asp:Label>'s session is active.
                    Enjoy the PoliMusic catalog.
                </div>
            </aside>

            <!-- ============ Main content ============ -->
            <div class="content">
                <header class="topbar">
                    <div class="user-chip">
                        <asp:Image ID="imgUserPhoto" runat="server" CssClass="avatar" AlternateText="User photo" />
                        <div class="user-text">
                            <span class="user-hello">Welcome</span>
                            <asp:Label ID="lblUserName" runat="server" CssClass="user-name"></asp:Label>
                        </div>
                    </div>
                    <a class="btn-logout" href="Logout.aspx" title="Close session">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M10 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h4v-2H6V5h4V3zm5.6 4.6L14.2 9l2 2H9v2h7.2l-2 2 1.4 1.4L20 12l-4.4-4.4z" /></svg>
                        Logout
                    </a>
                </header>

                <section class="hero">
                    <span class="hero-tag">Your library</span>
                    <h1>Song list</h1>
                    <p>Pick a track and press play. Every reproduction is stored in the database.</p>
                    <div class="hero-actions">
                        <button type="button" class="btn-play-all" id="btnPlayAll">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M8 5v14l11-7z" /></svg>
                            Play all
                        </button>
                        <span class="hero-meta">
                            <asp:Label ID="lblSongCount" runat="server" Text="0 songs"></asp:Label>
                        </span>
                    </div>
                </section>

                <asp:Panel ID="pnlMessage" runat="server" Visible="false" CssClass="alert">
                    <asp:Label ID="lblMessage" runat="server"></asp:Label>
                </asp:Panel>

                <div class="toolbar">
                    <h2 id="songs">Songs</h2>
                    <div class="search">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M10 4a6 6 0 1 0 3.5 10.9l4.3 4.3 1.4-1.4-4.3-4.3A6 6 0 0 0 10 4zm0 2a4 4 0 1 1 0 8 4 4 0 0 1 0-8z" /></svg>
                        <input type="text" id="searchInput" placeholder="Filter by name or genre" autocomplete="off" />
                    </div>
                </div>

                <div class="list-wrap">
                    <div class="list-head">
                        <span class="c-index">#</span>
                        <span class="c-title">Name</span>
                        <span class="c-genre">Genre</span>
                        <span class="c-path">Route</span>
                        <span class="c-plays">Plays</span>
                        <span class="c-time">Time</span>
                        <span class="c-action">Play</span>
                    </div>

                    <asp:Repeater ID="rptSongs" runat="server">
                        <ItemTemplate>
                            <div class="song-row"
                                data-id='<%# Eval("ID_SONG") %>'
                                data-src='<%# ResolveUrl(Convert.ToString(Eval("SONG_PATH"))) %>'
                                data-genre='<%# Server.HtmlEncode(Convert.ToString(Eval("GENRE"))) %>'
                                data-search='<%# Server.HtmlEncode((Convert.ToString(Eval("SONG_NAME")) + " " + Convert.ToString(Eval("GENRE"))).ToLower()) %>'>
                                <div class="c-index">
                                    <span class="row-number"><%# Container.ItemIndex + 1 %></span>
                                    <span class="eq"><i></i><i></i><i></i></span>
                                </div>
                                <div class="c-title">
                                    <div class="cover">
                                        <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z" /></svg>
                                    </div>
                                    <div class="title-text">
                                        <span class="song-name"><%# Server.HtmlEncode(Convert.ToString(Eval("SONG_NAME"))) %></span>
                                        <span class="song-sub">PoliMusic | <%# Server.HtmlEncode(Convert.ToString(Eval("GENRE"))) %></span>
                                    </div>
                                </div>
                                <div class="c-genre"><span class="badge"><%# Server.HtmlEncode(Convert.ToString(Eval("GENRE"))) %></span></div>
                                <div class="c-path" title='<%# Server.HtmlEncode(Convert.ToString(Eval("SONG_PATH"))) %>'><%# Server.HtmlEncode(Convert.ToString(Eval("SONG_PATH"))) %></div>
                                <div class="c-plays"><span class="plays-value"><%# Eval("PLAYS") %></span></div>
                                <div class="c-time"><span class="time-value">--:--</span></div>
                                <div class="c-action">
                                    <button type="button" class="row-play" title="Play song">
                                        <svg class="i-play" width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M8 5v14l11-7z" /></svg>
                                        <svg class="i-pause" width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M6 5h4v14H6zm8 0h4v14h-4z" /></svg>
                                    </button>
                                </div>
                            </div>
                        </ItemTemplate>
                    </asp:Repeater>

                    <div class="empty" id="noResults">
                        <strong>No matches</strong>
                        No song matches your filter. Try another name or genre.
                    </div>

                    <asp:Panel ID="pnlEmpty" runat="server" Visible="false" CssClass="empty">
                        <strong>There are no songs yet</strong>
                        The TBL_SONG table is empty. Add songs to start using the player.
                    </asp:Panel>
                </div>
            </div>
        </div>

        <!-- ============ Player bar ============ -->
        <div class="player">
            <div class="np">
                <div class="np-cover" id="npCover">
                    <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z" /></svg>
                </div>
                <div class="np-text">
                    <span class="np-title" id="npTitle">Nothing playing</span>
                    <span class="np-genre" id="npGenre">Select a song to start</span>
                </div>
            </div>

            <div class="player-center">
                <div class="controls">
                    <button type="button" class="ctrl" id="btnPrev" title="Previous">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M6 5h2v14H6zm12 0v14l-9-7z" /></svg>
                    </button>
                    <button type="button" class="ctrl ctrl-main" id="btnPlay" title="Play / Pause">
                        <svg class="i-play" width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M8 5v14l11-7z" /></svg>
                        <svg class="i-pause" width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M6 5h4v14H6zm8 0h4v14h-4z" /></svg>
                    </button>
                    <button type="button" class="ctrl" id="btnNext" title="Next">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M16 5h2v14h-2zM6 5l9 7-9 7z" /></svg>
                    </button>
                </div>
                <div class="progress">
                    <span class="time" id="curTime">0:00</span>
                    <div class="track" id="track">
                        <div class="fill" id="trackFill"></div>
                        <div class="knob" id="trackKnob"></div>
                    </div>
                    <span class="time" id="durTime">0:00</span>
                </div>
            </div>

            <div class="player-right">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M4 9v6h4l5 4V5L8 9H4zm12.5-.9v7.8a4.5 4.5 0 0 0 0-7.8z" /></svg>
                <input type="range" class="volume" id="volume" min="0" max="100" value="80" />
            </div>
        </div>

        <audio id="audioPlayer" preload="metadata"></audio>
    </form>

    <script type="text/javascript">
        (function () {
            "use strict";

            var audio = document.getElementById("audioPlayer");
            if (!audio) { return; }

            var rows = Array.prototype.slice.call(document.querySelectorAll(".song-row"));
            var el = {
                btnPlay: document.getElementById("btnPlay"),
                btnPrev: document.getElementById("btnPrev"),
                btnNext: document.getElementById("btnNext"),
                btnPlayAll: document.getElementById("btnPlayAll"),
                track: document.getElementById("track"),
                fill: document.getElementById("trackFill"),
                knob: document.getElementById("trackKnob"),
                curTime: document.getElementById("curTime"),
                durTime: document.getElementById("durTime"),
                npCover: document.getElementById("npCover"),
                npTitle: document.getElementById("npTitle"),
                npGenre: document.getElementById("npGenre"),
                volume: document.getElementById("volume"),
                search: document.getElementById("searchInput"),
                noResults: document.getElementById("noResults")
            };

            var currentIndex = -1;
            var countedId = null;
            var DEFAULT_TITLE = el.npTitle.textContent;
            var DEFAULT_GENRE = el.npGenre.textContent;

            function formatTime(seconds) {
                if (!isFinite(seconds) || seconds < 0) { return "--:--"; }
                var total = Math.floor(seconds);
                var m = Math.floor(total / 60);
                var s = total % 60;
                return m + ":" + (s < 10 ? "0" + s : s);
            }

            function rowTime(index) {
                if (index < 0) { return null; }
                return rows[index].querySelector(".time-value");
            }

            function setActiveRow(index) {
                for (var i = 0; i < rows.length; i++) {
                    rows[i].classList.remove("is-playing", "current");
                }
                if (index < 0) { return; }
                rows[index].classList.add("current");

                var name = rows[index].querySelector(".song-name").textContent;
                var genre = rows[index].getAttribute("data-genre") || "PoliMusic";
                el.npTitle.textContent = name;
                el.npGenre.textContent = genre;
                el.npCover.className = "np-cover active";
            }

            function syncPlayState() {
                var playing = !audio.paused && !audio.ended && audio.src !== "";
                if (playing) {
                    el.btnPlay.className = "ctrl ctrl-main is-playing";
                } else {
                    el.btnPlay.className = "ctrl ctrl-main";
                }
                for (var i = 0; i < rows.length; i++) {
                    if (i === currentIndex && playing) {
                        rows[i].classList.add("is-playing");
                    } else {
                        rows[i].classList.remove("is-playing");
                    }
                }
            }

            function registerPlay(row) {
                var id = row.getAttribute("data-id");
                if (!id || id === countedId) { return; }
                countedId = id;
                var xhr = new XMLHttpRequest();
                xhr.open("GET", "SongPlayHandler.ashx?idSong=" + encodeURIComponent(id), true);
                xhr.onreadystatechange = function () {
                    if (xhr.readyState !== 4) { return; }
                    if (xhr.status === 200) {
                        var plays = parseInt(xhr.responseText, 10);
                        var cell = row.querySelector(".plays-value");
                        if (!isNaN(plays) && cell) { cell.textContent = plays; }
                    }
                };
                xhr.send(null);
            }

            function playIndex(index) {
                if (index < 0 || index >= rows.length) { return; }
                if (index === currentIndex) {
                    if (audio.paused) { audio.play(); } else { audio.pause(); }
                    return;
                }
                currentIndex = index;
                setActiveRow(index);
                audio.src = rows[index].getAttribute("data-src");
                var cell = rowTime(index);
                if (cell) { cell.textContent = "--:--"; }
                el.fill.style.width = "0%";
                el.knob.style.left = "0%";
                el.curTime.textContent = "0:00";
                el.durTime.textContent = "0:00";
                registerPlay(rows[index]);
                var promise = audio.play();
                if (promise && typeof promise.catch === "function") {
                    promise.catch(function () { el.npGenre.textContent = "Press play to start"; });
                }
            }

            function playNext() {
                if (rows.length === 0) { return; }
                var next = currentIndex + 1;
                if (next >= rows.length) { next = 0; }
                playIndex(next);
            }

            function playPrevious() {
                if (rows.length === 0) { return; }
                if (audio.currentTime > 3) {
                    audio.currentTime = 0;
                    return;
                }
                var prev = currentIndex - 1;
                if (prev < 0) { prev = rows.length - 1; }
                playIndex(prev);
            }

            /* ---- Row clicks (event delegation) ---- */
            for (var i = 0; i < rows.length; i++) {
                (function (index) {
                    rows[index].addEventListener("click", function () { playIndex(index); });
                })(i);
            }

            /* ---- Player controls ---- */
            el.btnPlay.addEventListener("click", function () {
                if (currentIndex < 0) {
                    if (rows.length > 0) { playIndex(0); }
                    return;
                }
                if (audio.paused) { audio.play(); } else { audio.pause(); }
            });

            el.btnNext.addEventListener("click", playNext);
            el.btnPrev.addEventListener("click", playPrevious);

            if (el.btnPlayAll) {
                el.btnPlayAll.addEventListener("click", function () {
                    if (rows.length > 0) { playIndex(0); }
                });
            }

            /* ---- Audio events ---- */
            audio.addEventListener("play", syncPlayState);
            audio.addEventListener("pause", syncPlayState);
            audio.addEventListener("ended", function () { playNext(); });

            audio.addEventListener("loadedmetadata", function () {
                el.durTime.textContent = formatTime(audio.duration);
                var cell = rowTime(currentIndex);
                if (cell) { cell.textContent = formatTime(audio.duration); }
            });

            audio.addEventListener("timeupdate", function () {
                var pct = 0;
                if (audio.duration > 0) { pct = (audio.currentTime / audio.duration) * 100; }
                el.fill.style.width = pct + "%";
                el.knob.style.left = pct + "%";
                el.curTime.textContent = formatTime(audio.currentTime);
            });

            audio.addEventListener("error", function () {
                if (currentIndex < 0) { return; }
                el.npGenre.textContent = "Audio file not available";
            });

            /* ---- Seek ---- */
            el.track.addEventListener("click", function (event) {
                if (!audio.duration || currentIndex < 0) { return; }
                var rect = el.track.getBoundingClientRect();
                var ratio = (event.clientX - rect.left) / rect.width;
                if (ratio < 0) { ratio = 0; }
                if (ratio > 1) { ratio = 1; }
                audio.currentTime = ratio * audio.duration;
            });

            /* ---- Volume ---- */
            if (el.volume) {
                audio.volume = el.volume.value / 100;
                el.volume.addEventListener("input", function () {
                    audio.volume = el.volume.value / 100;
                });
            }

            /* ---- Client side filter ---- */
            if (el.search) {
                el.search.addEventListener("input", function () {
                    var query = el.search.value.toLowerCase().replace(/^\s+|\s+$/g, "");
                    var visible = 0;
                    for (var j = 0; j < rows.length; j++) {
                        var haystack = rows[j].getAttribute("data-search") || "";
                        var match = haystack.indexOf(query) !== -1;
                        rows[j].style.display = match ? "" : "none";
                        if (match) { visible++; }
                    }
                    if (el.noResults) {
                        el.noResults.style.display = (visible === 0 && rows.length > 0) ? "block" : "none";
                    }
                });
            }

            /* ---- Space bar = play / pause ---- */
            document.addEventListener("keydown", function (event) {
                var tag = (event.target && event.target.tagName) ? event.target.tagName.toLowerCase() : "";
                if (tag === "input" || tag === "textarea" || tag === "select") { return; }
                if (event.keyCode === 32) {
                    event.preventDefault();
                    if (currentIndex < 0) { playIndex(0); return; }
                    if (audio.paused) { audio.play(); } else { audio.pause(); }
                }
            });

            if (rows.length === 0) {
                el.npTitle.textContent = DEFAULT_TITLE;
                el.npGenre.textContent = DEFAULT_GENRE;
            }
        })();
    </script>
</body>
</html>
