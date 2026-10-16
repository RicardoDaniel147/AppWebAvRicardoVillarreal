/* =====================================================================
   PoliMusic - Reproductor global
   Requiere que header.asp defina:  PM_BASE (ruta de la app) y PM_TOKEN (CSRF)
   ===================================================================== */
(function () {
    "use strict";

    var audio = document.getElementById("pm-audio");
    if (!audio) return;

    var BASE = (typeof window.PM_BASE === "string" ? window.PM_BASE : "");
    var TOKEN = (typeof window.PM_TOKEN === "string" ? window.PM_TOKEN : "");

    var el = {
        cover: document.getElementById("pm-now-cover"),
        title: document.getElementById("pm-now-title"),
        artist: document.getElementById("pm-now-artist"),
        play: document.getElementById("pm-play"),
        prev: document.getElementById("pm-prev"),
        next: document.getElementById("pm-next"),
        shuffle: document.getElementById("pm-shuffle"),
        repeat: document.getElementById("pm-repeat"),
        bar: document.getElementById("pm-bar"),
        fill: document.getElementById("pm-fill"),
        cur: document.getElementById("pm-cur"),
        dur: document.getElementById("pm-dur"),
        vol: document.getElementById("pm-vol"),
        eq: document.getElementById("pm-eq")
    };

    var state = { queue: [], index: -1, shuffle: false, repeat: false };

    var ICON_PLAY = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>';
    var ICON_PAUSE = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M6 5h4v14H6zM14 5h4v14h-4z"/></svg>';

    function fmt(seconds) {
        if (!isFinite(seconds) || seconds < 0) return "0:00";
        var m = Math.floor(seconds / 60);
        var s = Math.floor(seconds % 60);
        return m + ":" + (s < 10 ? "0" : "") + s;
    }

    function placeholder(kind, seed) {
        return BASE + "/assets/placeholder.asp?kind=" + encodeURIComponent(kind) +
               "&seed=" + encodeURIComponent(seed || "PoliMusic");
    }

    function toast(message, kind) {
        var host = document.getElementById("pm-toasts");
        if (!host) {
            host = document.createElement("div");
            host.id = "pm-toasts";
            host.style.cssText =
                "position:fixed;right:22px;bottom:calc(var(--player-h) + 18px);z-index:9999;" +
                "display:flex;flex-direction:column;gap:10px;pointer-events:none";
            document.body.appendChild(host);
        }
        var node = document.createElement("div");
        node.className = "pm-alert " + (kind === "err" ? "pm-alert-err" : "pm-alert-info");
        node.style.cssText = "margin:0;box-shadow:0 12px 34px rgba(0,0,0,.55);pointer-events:auto;max-width:330px";
        node.textContent = message;
        host.appendChild(node);
        setTimeout(function () {
            node.style.transition = "opacity .3s ease, transform .3s ease";
            node.style.opacity = "0";
            node.style.transform = "translateY(8px)";
            setTimeout(function () { node.remove(); }, 320);
        }, 3600);
    }

    function readRow(row) {
        return {
            id: parseInt(row.dataset.songId, 10) || 0,
            name: row.dataset.songName || "Sin titulo",
            artists: row.dataset.songArtists || "Artista desconocido",
            src: row.dataset.songSrc || "",
            cover: row.dataset.songCover || "",
            exists: row.dataset.songExists === "1"
        };
    }

    function rowsInScope(row) {
        var scope = row.closest("[data-playlist-scope]") || document;
        return Array.prototype.slice.call(scope.querySelectorAll(".pm-row[data-song-id]"));
    }

    function paint() {
        var track = state.index >= 0 ? state.queue[state.index] : null;
        if (el.title) el.title.textContent = track ? track.name : "Sin reproduccion";
        if (el.artist) el.artist.textContent = track ? track.artists : "Selecciona una cancion";
        if (el.cover) el.cover.src = track ? (track.cover || placeholder("album", track.name)) : placeholder("album", "PoliMusic");

        document.querySelectorAll(".pm-row.playing").forEach(function (r) { r.classList.remove("playing"); });
        if (track) {
            var active = document.querySelector('.pm-row[data-song-id="' + track.id + '"]');
            if (active) active.classList.add("playing");
        }
    }

    function setPlayingUi(isPlaying) {
        if (el.play) el.play.innerHTML = isPlaying ? ICON_PAUSE : ICON_PLAY;
        if (el.eq) el.eq.classList.toggle("paused", !isPlaying);
    }

    function registerPlay(id) {
        if (!id) return;
        fetch(BASE + "/play.asp?id=" + id, {
            method: "POST",
            headers: { "X-CSRF-Token": TOKEN, "X-Requested-With": "XMLHttpRequest" }
        }).catch(function () { /* best effort */ });
    }

    function load(index, autoplay) {
        if (index < 0 || index >= state.queue.length) return;
        state.index = index;
        var track = state.queue[index];

        audio.src = track.src || "";
        paint();

        if (!track.exists) {
            toast('El archivo de audio no esta disponible en el servidor: "' + track.name + '"', "err");
            setPlayingUi(false);
            return;
        }

        if (autoplay !== false) {
            audio.play().then(function () {
                setPlayingUi(true);
                registerPlay(track.id);
            }).catch(function () {
                setPlayingUi(false);
                toast("No se pudo iniciar la reproduccion. Revisa que el archivo exista.", "err");
            });
        }
    }

    function toggle() {
        if (state.index < 0) { PoliMusic.playFirst(); return; }
        if (audio.paused) {
            audio.play().then(function () {
                setPlayingUi(true);
                registerPlay(state.queue[state.index].id);
            }).catch(function () { toast("No se pudo reproducir el archivo.", "err"); });
        } else {
            audio.pause();
            setPlayingUi(false);
        }
    }

    function nextIndex() {
        if (state.queue.length === 0) return -1;
        if (state.shuffle) {
            var candidate = state.index;
            if (state.queue.length > 1) {
                while (candidate === state.index) candidate = Math.floor(Math.random() * state.queue.length);
            }
            return candidate;
        }
        return state.index + 1 < state.queue.length ? state.index + 1 : (state.repeat ? 0 : -1);
    }

    function prevIndex() {
        if (state.queue.length === 0) return -1;
        if (audio.currentTime > 3) return state.index;
        return state.index - 1 >= 0 ? state.index - 1 : state.queue.length - 1;
    }

    function playFrom(row) {
        state.queue = rowsInScope(row).map(readRow);
        var id = parseInt(row.dataset.songId, 10);
        var index = state.queue.findIndex(function (t) { return t.id === id; });
        load(index < 0 ? 0 : index, true);
    }

    document.addEventListener("click", function (event) {
        var row = event.target.closest(".pm-row[data-song-id]");
        if (!row) return;
        if (event.target.closest("a, button.pm-no-play, form, input, select")) return;
        event.preventDefault();
        playFrom(row);
    });

    document.addEventListener("keydown", function (event) {
        if (event.target.matches("input, textarea, select")) return;
        if (event.code === "Space") { event.preventDefault(); toggle(); }
        else if (event.code === "ArrowRight" && event.shiftKey) { load(nextIndex(), true); }
        else if (event.code === "ArrowLeft" && event.shiftKey) { load(prevIndex(), true); }
    });

    if (el.play) el.play.addEventListener("click", toggle);
    if (el.next) el.next.addEventListener("click", function () {
        var i = nextIndex();
        if (i < 0) { toast("Fin de la cola de reproduccion."); return; }
        load(i, true);
    });
    if (el.prev) el.prev.addEventListener("click", function () {
        var i = prevIndex();
        if (i >= 0) load(i, true);
    });

    if (el.shuffle) el.shuffle.addEventListener("click", function () {
        state.shuffle = !state.shuffle;
        el.shuffle.style.color = state.shuffle ? "var(--accent-2)" : "";
        toast(state.shuffle ? "Aleatorio activado." : "Aleatorio desactivado.");
    });

    if (el.repeat) el.repeat.addEventListener("click", function () {
        state.repeat = !state.repeat;
        el.repeat.style.color = state.repeat ? "var(--accent-2)" : "";
        toast(state.repeat ? "Repeticion activada." : "Repeticion desactivada.");
    });

    audio.addEventListener("timeupdate", function () {
        if (!audio.duration) return;
        if (el.fill) el.fill.style.width = ((audio.currentTime / audio.duration) * 100) + "%";
        if (el.cur) el.cur.textContent = fmt(audio.currentTime);
    });

    audio.addEventListener("loadedmetadata", function () {
        if (el.dur) el.dur.textContent = fmt(audio.duration);
    });

    audio.addEventListener("play", function () { setPlayingUi(true); });
    audio.addEventListener("pause", function () { setPlayingUi(false); });

    audio.addEventListener("ended", function () {
        var i = nextIndex();
        if (i < 0) { setPlayingUi(false); return; }
        load(i, true);
    });

    audio.addEventListener("error", function () {
        if (!audio.src) return;
        setPlayingUi(false);
        var track = state.index >= 0 ? state.queue[state.index] : null;
        toast("No se encontro el archivo de audio" + (track ? ': "' + track.name + '"' : "."), "err");
    });

    if (el.bar) {
        el.bar.addEventListener("click", function (event) {
            if (!audio.duration) return;
            var rect = el.bar.getBoundingClientRect();
            var ratio = (event.clientX - rect.left) / rect.width;
            audio.currentTime = Math.max(0, Math.min(1, ratio)) * audio.duration;
        });
    }

    if (el.vol) {
        el.vol.value = 80;
        audio.volume = 0.8;
        el.vol.addEventListener("input", function () {
            audio.volume = parseFloat(el.vol.value) / 100;
        });
    }

    window.PoliMusic = {
        play: function (songId) {
            var row = document.querySelector('.pm-row[data-song-id="' + songId + '"]');
            if (row) playFrom(row);
        },
        playFirst: function () {
            var first = document.querySelector(".pm-row[data-song-id]");
            if (first) playFrom(first);
            else toast("No hay canciones disponibles para reproducir.");
        },
        toggle: toggle,
        toast: toast
    };

    paint();
})();
