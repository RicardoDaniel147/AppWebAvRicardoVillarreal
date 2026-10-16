    </main>

    <!-- ============================================================= PLAYER -->
    <footer class="pm-player">
        <div class="pm-player-now">
            <img id="pm-now-cover" alt="" src="<%=PlaceholderUrl("album", "PoliMusic")%>" />
            <span class="t">
                <strong id="pm-now-title">Sin reproduccion</strong>
                <span id="pm-now-artist">Selecciona una cancion</span>
            </span>
        </div>

        <div class="pm-player-center">
            <div class="pm-player-controls">
                <button type="button" class="pm-ctrl" id="pm-shuffle" title="Aleatorio">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M16 3h5v5" /><path d="M4 20 21 3" /><path d="M21 16v5h-5" />
                        <path d="M15 15l6 6" /><path d="M4 4l5 5" />
                    </svg>
                </button>

                <button type="button" class="pm-ctrl" id="pm-prev" title="Anterior">
                    <svg viewBox="0 0 24 24" fill="currentColor"><path d="M6 5h2v14H6zM20 5v14L10 12z" /></svg>
                </button>

                <button type="button" class="pm-ctrl pm-ctrl-play" id="pm-play" title="Reproducir / Pausar">
                    <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z" /></svg>
                </button>

                <button type="button" class="pm-ctrl" id="pm-next" title="Siguiente">
                    <svg viewBox="0 0 24 24" fill="currentColor"><path d="M16 5h2v14h-2zM4 5v14l10-7z" /></svg>
                </button>

                <button type="button" class="pm-ctrl" id="pm-repeat" title="Repetir">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M17 2l4 4-4 4" /><path d="M3 11V9a4 4 0 0 1 4-4h14" />
                        <path d="M7 22l-4-4 4-4" /><path d="M21 13v2a4 4 0 0 1-4 4H3" />
                    </svg>
                </button>
            </div>

            <div class="pm-progress">
                <span id="pm-cur">0:00</span>
                <div class="pm-bar" id="pm-bar"><div class="pm-bar-fill" id="pm-fill"></div></div>
                <span id="pm-dur">0:00</span>
            </div>
        </div>

        <div class="pm-player-right">
            <span class="pm-eq paused" id="pm-eq"><i></i><i></i><i></i><i></i></span>
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"
                 style="width:17px;height:17px;color:var(--text-faint)">
                <path d="M11 5 6 9H3v6h3l5 4z" /><path d="M16 9a4 4 0 0 1 0 6" />
            </svg>
            <input type="range" min="0" max="100" class="pm-vol" id="pm-vol" aria-label="Volumen" />
        </div>
    </footer>
</div>

<audio id="pm-audio" preload="metadata"></audio>

<script type="text/javascript">
    window.PM_BASE = "<%=JS(AppPath())%>";
    window.PM_TOKEN = "<%=JS(CsrfToken())%>";
</script>
<script src="<%=Asset("player.js")%>"></script>
</body>
</html>
