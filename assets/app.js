/* Mission Control V2 - vanilla JS */

const API_BASE = "/.netlify/functions/mission";

function $(sel, root = document) { return root.querySelector(sel); }
function $all(sel, root = document) { return Array.from(root.querySelectorAll(sel)); }

function qparam(name) { return new URLSearchParams(location.search).get(name); }

function clamp(n, min, max) { return Math.max(min, Math.min(max, n)); }

function now() { return Date.now(); }

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function formatTime(sec) {
  sec = Math.max(0, Math.floor(sec));
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return `${m}:${s.toString().padStart(2, "0")}`;
}

// Deterministic seeded RNG (xmur3 + sfc32)
function xmur3(str) {
  let h = 1779033703 ^ str.length;
  for (let i = 0; i < str.length; i++) {
    h = Math.imul(h ^ str.charCodeAt(i), 3432918353);
    h = (h << 13) | (h >>> 19);
  }
  return function() {
    h = Math.imul(h ^ (h >>> 16), 2246822507);
    h = Math.imul(h ^ (h >>> 13), 3266489909);
    h ^= h >>> 16;
    return h >>> 0;
  };
}
function sfc32(a, b, c, d) {
  return function() {
    a >>>= 0; b >>>= 0; c >>>= 0; d >>>= 0;
    let t = (a + b) | 0;
    a = b ^ (b >>> 9);
    b = (c + (c << 3)) | 0;
    c = (c << 21) | (c >>> 11);
    d = (d + 1) | 0;
    t = (t + d) | 0;
    c = (c + t) | 0;
    return (t >>> 0) / 4294967296;
  };
}
function rngFromSeed(seed) {
  const h = xmur3(seed);
  return sfc32(h(), h(), h(), h());
}
function pick(rng, arr) { return arr[Math.floor(rng() * arr.length)]; }
function shuffle(rng, arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

async function apiGet(code) {
  const res = await fetch(`${API_BASE}?code=${encodeURIComponent(code)}`, { cache: "no-store" });
  if (!res.ok) throw new Error(`GET ${res.status}`);
  const data = await res.json();
  return data.mission;
}

async function apiWrite(code, patch, expectedVersion = null) {
  const payload = { ...patch };
  if (expectedVersion !== null) payload.version = expectedVersion;
  const res = await fetch(`${API_BASE}?code=${encodeURIComponent(code)}`, {
    method: "PATCH",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const err = data?.error || `PATCH ${res.status}`;
    const e = new Error(err);
    e.status = res.status;
    e.data = data;
    throw e;
  }
  return data.mission;
}

// ---- Mission model helpers ----

const DEFAULT_DURATION_SEC = 8 * 60;
const ALLOWLIST_PASSWORDS = ["swsw","grfr","sksa","lchi","bigu","guest"];
const COLORS = ["Red","Blue","Green","Yellow","Orange","Purple","White","Gray"];
const PANEL_META = {
  wiring: { label: "Wiring", icon: "🧷", accent: "#60a5fa" },
  cosmic: { label: "Cosmic", icon: "✦", accent: "#f472b6" },
  thruster: { label: "Thruster", icon: "⟟", accent: "#fbbf24" },
  frequency: { label: "Frequency", icon: "⟒", accent: "#a78bfa" },
  life: { label: "Life", icon: "⧉", accent: "#4ade80" },
  keypad: { label: "Keypad", icon: "⌬", accent: "#fb7185" },
  aux: { label: "AUX", icon: "▣", accent: "#94a3b8" },
  power: { label: "Power", icon: "⏻", accent: "#6ee7b7" }
};

function makePanelId(rng) {
  // 5-digit numeric string
  return String(Math.floor(rng() * 90000) + 10000);
}

function computeTimer(mission) {
  const duration = mission?.snapshot?.durationSec ?? DEFAULT_DURATION_SEC;
  const startedAt = mission?.runtime?.startedAt ?? null;
  const penalties = mission?.runtime?.penaltySec ?? 0;
  const stoppedAt = mission?.runtime?.stoppedAt ?? null;

  if (!startedAt) return { state: "idle", remaining: duration, duration, penalties };
  const endRef = stoppedAt || now();
  const elapsed = (endRef - startedAt) / 1000;
  const remaining = duration - elapsed - penalties;
  if (stoppedAt) return { state: "stopped", remaining, duration, penalties };
  if (remaining <= 0) return { state: "expired", remaining: 0, duration, penalties };
  return { state: "running", remaining, duration, penalties };
}

function allRequiredPanelsComplete(mission) {
  const c = mission?.runtime?.completed || {};
  return !!(c.wiring && c.cosmic && c.thruster && c.frequency && c.life && c.keypad);
}

function missionFailed(mission) {
  return mission?.runtime?.status === "failed";
}
function missionWon(mission) {
  return mission?.runtime?.status === "won";
}

// ---- UI helpers ----

function setBanner(text, kind = "warn") {
  const el = $("#banner");
  if (!el) return;
  el.textContent = text || "";
  el.classList.remove("hidden");
  el.classList.toggle("bad", kind === "bad");
  el.classList.toggle("ok", kind === "ok");
  if (!text) el.classList.add("hidden");
}

function toast(msg) {
  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.classList.add("show"), 10);
  setTimeout(() => {
    t.classList.remove("show");
    setTimeout(() => t.remove(), 400);
  }, 2200);
}

function flash(kind="warn") {
  const overlay = document.createElement("div");
  overlay.className = `flash ${kind}`;
  document.body.appendChild(overlay);
  requestAnimationFrame(() => overlay.classList.add("show"));
  setTimeout(() => overlay.remove(), 220);
}

// ---- Page boot ----

document.addEventListener("DOMContentLoaded", () => {
  const role = document.body?.dataset?.role;
  if (role === "facilitator") bootFacilitator();
  if (role === "astronaut") bootAstronaut();
  if (role === "manuals") bootManuals();
});

// ---------------- Facilitator ----------------

async function bootFacilitator() {
  const codeInput = $("#missionCode");
  const genBtn = $("#genCode");
  const connectBtn = $("#connect");
  const pwModal = $("#pwModal");
  const pwInput = $("#pw");
  const pwBtn = $("#pwSubmit");
  const showSetupBtn = $("#showSetup");
  const hideSetupBtn = $("#hideSetup");
  const applyBtn = $("#applySetup");
  const lockBtn = $("#lockMission");
  const resetBtn = $("#resetMission");
  const regenBtn = $("#regenMission");
  const genLinksBtn = $("#generateLinks");
  const genManualsBtn = $("#generateManuals");
  const dlManualsBtn = $("#downloadManuals");

  const durationInput = $("#duration");
  const allowCustom = $("#allowCustom");

  let code = qparam("mission") || "";
  if (code) codeInput.value = code;

  function requirePassword() {
    const ok = sessionStorage.getItem("facilitator_authed") === "1";
    if (!ok) {
      pwModal.classList.remove("hidden");
      pwInput.value = "";
      pwInput.focus();
    } else {
      pwModal.classList.add("hidden");
    }
  }

  pwBtn.addEventListener("click", () => {
    const val = (pwInput.value || "").trim();
    if (!ALLOWLIST_PASSWORDS.includes(val)) {
      toast("Nope. Wrong password.");
      flash("bad");
      return;
    }
    sessionStorage.setItem("facilitator_authed", "1");
    pwModal.classList.add("hidden");
  });

  showSetupBtn.addEventListener("click", () => {
    sessionStorage.removeItem("facilitator_authed");
    requirePassword();
    $("#setup").classList.remove("hidden");
    $("#spectator").classList.add("hidden");
  });
  hideSetupBtn.addEventListener("click", () => {
    sessionStorage.removeItem("facilitator_authed");
    $("#setup").classList.add("hidden");
    $("#spectator").classList.remove("hidden");
    requirePassword();
  });

  genBtn.addEventListener("click", () => {
    const rng = rngFromSeed(String(now()));
    const parts = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
    let out = "";
    for (let i = 0; i < 6; i++) out += parts[Math.floor(rng() * parts.length)];
    codeInput.value = out;
  });

  connectBtn.addEventListener("click", async () => {
    code = (codeInput.value || "").trim();
    if (!code) return toast("Enter a mission code");
    history.replaceState({}, "", `?mission=${encodeURIComponent(code)}`);
    await ensureMissionExists(code);
    requirePassword();
    startFacilitatorPoll(code);
    toast("Connected");
  });

  async function ensureMissionExists(code) {
    let m = await apiGet(code);
    if (m) return m;
    // Create empty mission container
    m = await apiWrite(code, {
      code,
      createdAt: now(),
      locked: false,
      snapshot: null,
      runtime: {
        status: "idle",
        startedAt: null,
        stoppedAt: null,
        penaltySec: 0,
        strikes: 0,
        completed: {},
        panelState: {},
      },
      manuals: { generatedAt: 0, items: [] },
    });
    return m;
  }

  // Apply Setup => creates snapshot + scrambles order + panel IDs + answers
  applyBtn.addEventListener("click", async () => {
    code = (codeInput.value || "").trim();
    if (!code) return toast("Enter mission code first");
    let m = await apiGet(code);
    if (!m) m = await ensureMissionExists(code);
    if (m.locked) return toast("Mission is locked");

    const durationSec = clamp(parseInt(durationInput.value || "480", 10) * 60, 60, 60 * 60);
    const seed = `${code}::snapshot::${now()}`;
    const rng = rngFromSeed(seed);

    const panelIds = {
      wiring: makePanelId(rng),
      cosmic: makePanelId(rng),
      thruster: makePanelId(rng),
      frequency: makePanelId(rng),
      life: makePanelId(rng),
      keypad: makePanelId(rng),
      aux: makePanelId(rng),
      power: makePanelId(rng),
    };

    const order = shuffle(rng, ["wiring","cosmic","thruster","frequency","life","keypad","aux","power"]);

    // Wiring generation
    const wireCount = Math.floor(rng() * 4) + 3; // 3..6
    const palette = shuffle(rng, COLORS);
    const wires = [];
    const usedCounts = {};
    while (wires.length < wireCount) {
      const c = pick(rng, palette);
      usedCounts[c] = (usedCounts[c] || 0) + 1;
      if (usedCounts[c] > 2) continue;
      wires.push(c);
    }
    // Choose how many correct cuts: 0..wireCount
    const correctCutCount = Math.floor(rng() * (wireCount + 1));
    let correctOrder = [];
    if (correctCutCount > 0) {
      const idxs = shuffle(rng, Array.from({length: wireCount}, (_, i) => i + 1)).slice(0, correctCutCount);
      // order of cutting is another shuffle
      correctOrder = shuffle(rng, idxs);
    }

    // Cosmic: 4 symbols, shuffled display, correct order separate
    const SYMBOLS = ["☉","☾","✶","✷","✹","✦","✧","✩","✪","✫","⌬","⌖","⟡","⟁","⟠","⟟"];
    const symbols = shuffle(rng, SYMBOLS).slice(0, 4);
    const cosmicCorrect = shuffle(rng, symbols.slice());

    const thrusterTarget = Math.floor(rng() * 100) + 1;
    const freqTargets = [
      Math.floor(rng() * 10),
      Math.floor(rng() * 10),
      Math.floor(rng() * 10),
    ];

    const lifeChecks = {
      oxygen: rng() < 0.6,
      coolant: rng() < 0.6,
      scrubbers: rng() < 0.6,
    };
    const lifeMode = pick(rng, ["NOMINAL","ECO","EMERGENCY"]);

    const keypadCode = String(Math.floor(rng() * 9000) + 1000);

    const snapshot = {
      seed,
      createdAt: now(),
      durationSec,
      panelIds,
      panelOrder: order,
      answers: {
        wiring: { wireCount, wires, correctOrder },
        cosmic: { symbols, correct: cosmicCorrect },
        thruster: { target: thrusterTarget },
        frequency: { target: freqTargets },
        life: { checks: lifeChecks, mode: lifeMode },
        keypad: { code: keypadCode },
      },
      overrides: {
        customInstructions: {},
      },
      manualsOutOfDate: true,
      allowCustom: !!allowCustom.checked,
    };

    // reset runtime to pre-mayday
    const runtime = {
      status: "idle",
      startedAt: null,
      stoppedAt: null,
      penaltySec: 0,
      strikes: 0,
      completed: {},
      panelState: {},
      powerOnline: false,
    };

    const updated = await apiWrite(code, { snapshot, runtime, locked: false });
    markApplied(true);
    toast("Setup applied");
    return updated;
  });

  lockBtn.addEventListener("click", async () => {
    code = (codeInput.value || "").trim();
    if (!code) return;
    const m = await apiGet(code);
    if (!m?.snapshot) return toast("Apply Setup first");
    await apiWrite(code, { locked: true });
    toast("Mission locked");
  });

  resetBtn.addEventListener("click", async () => {
    code = (codeInput.value || "").trim();
    if (!code) return;
    const m = await apiGet(code);
    if (!m) return;
    if (!m.snapshot) return;
    // Restart mission: keep snapshot, clear runtime
    const runtime = {
      status: "idle",
      startedAt: null,
      stoppedAt: null,
      penaltySec: 0,
      strikes: 0,
      completed: {},
      panelState: {},
      powerOnline: false,
    };
    await apiWrite(code, { runtime });
    toast("Restarted (same mission)");
  });

  regenBtn.addEventListener("click", async () => {
    // Regenerate mission: just click Apply Setup again (new snapshot)
    applyBtn.click();
  });

  genLinksBtn.addEventListener("click", () => {
    code = (codeInput.value || "").trim();
    if (!code) return;
    const base = location.origin;
    $("#linkAstronaut").value = `${base}/astronaut?mission=${encodeURIComponent(code)}`;
    $("#linkManuals").value = `${base}/manuals?mission=${encodeURIComponent(code)}`;
    toast("Links generated");
  });

  genManualsBtn.addEventListener("click", async () => {
    code = (codeInput.value || "").trim();
    if (!code) return;
    const m = await apiGet(code);
    if (!m?.snapshot) return toast("Apply Setup first");

    const manuals = generateManuals(m);
    const updated = await apiWrite(code, { manuals, snapshot: { ...m.snapshot, manualsOutOfDate: false } }, m.version);
    const url = `/manuals?mission=${encodeURIComponent(code)}`;
    window.open(url, "_blank", "noopener,noreferrer");
    toast("Manuals generated");
    return updated;
  });

  dlManualsBtn.addEventListener("click", () => {
    // HTML manuals are on-page; this downloads a simple HTML file snapshot
    code = (codeInput.value || "").trim();
    if (!code) return;
    window.open(`/manuals?mission=${encodeURIComponent(code)}&download=1`, "_blank", "noopener,noreferrer");
  });

  function markApplied(on) {
    const el = $("#applied");
    if (!el) return;
    el.textContent = on ? "Specs applied" : "Specs not applied";
    el.classList.toggle("ok", on);
  }

  async function startFacilitatorPoll(code) {
    let lastVersion = null;
    while (true) {
      try {
        const m = await apiGet(code);
        if (m && m.version !== lastVersion) {
          lastVersion = m.version;
          renderFacilitator(m);
        } else if (m) {
          // still update timer display
          renderFacilitatorTimer(m);
        }
      } catch (e) {
        console.error(e);
      }
      await sleep(1000);
    }
  }

  function renderFacilitator(mission) {
    // applied indicator
    markApplied(!!mission?.snapshot);

    const room = $("#roomCode");
    if (room) room.textContent = mission.code || "";

    // out-of-date manuals
    const mod = $("#manualsStatus");
    if (mod) {
      const ood = mission?.snapshot?.manualsOutOfDate;
      mod.textContent = ood ? "Manuals out of date" : "Manuals current";
      mod.classList.toggle("warn", !!ood);
    }

    renderFacilitatorTimer(mission);

    const completed = mission?.runtime?.completed || {};
    const cards = $("#statusCards");
    if (cards) {
      cards.innerHTML = "";
      for (const key of ["wiring","cosmic","thruster","frequency","life","keypad"]) {
        const meta = PANEL_META[key];
        const done = !!completed[key];
        const div = document.createElement("div");
        div.className = `status ${done ? "ok" : ""}`;
        div.innerHTML = `<div><strong>${meta.icon} ${meta.label}</strong><div class="muted">Panel ID: ${mission?.snapshot?.panelIds?.[key] || "—"}</div></div><div>${done ? "COMPLETE" : "PENDING"}</div>`;
        cards.appendChild(div);
      }
    }

    const strikesEl = $("#strikes");
    if (strikesEl) strikesEl.textContent = String(mission?.runtime?.strikes || 0);

    const penEl = $("#penalties");
    if (penEl) penEl.textContent = `${mission?.runtime?.penaltySec || 0}s`;

    const st = $("#missionState");
    if (st) st.textContent = mission?.runtime?.status || "idle";
  }

  function renderFacilitatorTimer(mission) {
    const t = computeTimer(mission);
    const el = $("#timer");
    if (el) el.textContent = formatTime(t.remaining);
  }

  // initial
  requirePassword();
}

// ---------------- Astronaut ----------------

async function bootAstronaut() {
  const code = (qparam("mission") || "").trim();
  if (!code) {
    setBanner("Missing mission code", "bad");
    return;
  }
  $("#roomCode").textContent = code;

  let mission = await apiGet(code);
  if (!mission?.snapshot) {
    setBanner("Mission not set up yet. Tell the facilitator to Apply Setup.", "warn");
  }

  // wire UI and handlers
  $("#mayday").addEventListener("click", async () => {
    mission = await apiGet(code);
    if (missionFailed(mission) || missionWon(mission)) {
      toast("Mission ended. Restart if you want to run it again.");
      return;
    }
    if (!mission?.snapshot) return;
    if (mission?.runtime?.startedAt) return;
    // start timer
    mission = await apiWrite(code, {
      runtime: { ...mission.runtime, startedAt: now(), status: "running" },
    }, mission.version);
    toast("MAYDAY received. Timer started.");
  });

  $("#restart").addEventListener("click", async () => {
    mission = await apiGet(code);
    if (!mission?.snapshot) return;
    const runtime = {
      status: "idle",
      startedAt: null,
      stoppedAt: null,
      penaltySec: 0,
      strikes: 0,
      completed: {},
      panelState: {},
      powerOnline: false,
    };
    mission = await apiWrite(code, { runtime }, mission.version);
    setBanner("");
    toast("Restarted. Hit MAYDAY to begin.");
  });

  // poll + render
  let lastVersion = null;
  while (true) {
    try {
      mission = await apiGet(code);
      if (mission) {
        if (mission.version !== lastVersion) {
          lastVersion = mission.version;
          renderAstronaut(mission, code);
        } else {
          renderAstronautTimer(mission);
        }
        // handle timer expiry
        const t = computeTimer(mission);
        if (t.state === "expired" && mission?.runtime?.status !== "failed") {
          mission = await apiWrite(code, { runtime: { ...mission.runtime, status: "failed", stoppedAt: now() } }, mission.version).catch(()=>mission);
        }
      }
    } catch (e) {
      console.error(e);
    }
    await sleep(300);
  }
}

function renderAstronaut(mission, code) {
  const t = computeTimer(mission);
  renderAstronautTimer(mission);

  if (missionFailed(mission)) {
    setBanner("MISSION FAILED", "bad");
  } else if (missionWon(mission)) {
    setBanner("MISSION ACCOMPLISHED", "ok");
  } else {
    setBanner("");
  }

  // strikes + penalties UI
  $("#strikes").textContent = String(mission?.runtime?.strikes || 0);
  $("#penalties").textContent = `${mission?.runtime?.penaltySec || 0}s`;

  // build panels in shuffled order
  const grid = $("#panelGrid");
  grid.innerHTML = "";
  const order = mission?.snapshot?.panelOrder || [];
  for (const key of order) {
    grid.appendChild(renderPanelCard(key, mission, code));
  }
}

function renderAstronautTimer(mission) {
  const t = computeTimer(mission);
  $("#timer").textContent = formatTime(t.remaining);
  $("#timer").classList.toggle("warn", t.remaining < 60 && t.state === "running");
}

function renderPanelCard(key, mission, code) {
  const meta = PANEL_META[key];
  const id = mission?.snapshot?.panelIds?.[key] || "—";
  const card = document.createElement("div");
  card.className = "card";
  card.style.setProperty("--accent", meta?.accent || "#6ee7b7");

  const header = document.createElement("div");
  header.className = "cardHeader";
  header.innerHTML = `<div class="title">${meta?.icon || ""} Panel ID ${id}</div><div class="sub">${key === "aux" ? "AUXILIARY SYSTEM MONITOR" : ""}</div>`;
  card.appendChild(header);

  const body = document.createElement("div");
  body.className = "cardBody";
  card.appendChild(body);

  const done = !!mission?.runtime?.completed?.[key];
  if (done) card.classList.add("locked");

  if (key === "wiring") body.appendChild(renderWiring(mission, code, done));
  else if (key === "cosmic") body.appendChild(renderCosmic(mission, code, done));
  else if (key === "thruster") body.appendChild(renderThruster(mission, code, done));
  else if (key === "frequency") body.appendChild(renderFrequency(mission, code, done));
  else if (key === "life") body.appendChild(renderLife(mission, code, done));
  else if (key === "keypad") body.appendChild(renderKeypad(mission, code, done));
  else if (key === "aux") body.appendChild(renderAux(mission));
  else if (key === "power") body.appendChild(renderPower(mission, code));

  return card;
}

async function applyPenaltyOrFail(code, mission, opts) {
  // opts: { immediateFail?: boolean, reason?: string }
  if (missionFailed(mission) || missionWon(mission)) return mission;

  if (opts.immediateFail) {
    flash("bad");
    setBanner("MISSION FAILED", "bad");
    return apiWrite(code, { runtime: { ...mission.runtime, status: "failed", stoppedAt: now() } }, mission.version);
  }

  const strikes = (mission.runtime.strikes || 0) + 1;
  const penaltySec = (mission.runtime.penaltySec || 0) + 10;
  flash("warn");
  return apiWrite(code, { runtime: { ...mission.runtime, strikes, penaltySec } }, mission.version);
}

function renderWiring(mission, code, locked) {
  const wrap = document.createElement("div");
  const a = mission.snapshot.answers.wiring;
  const state = mission.runtime.panelState.wiring || { cut: [], bypassed: false };

  const list = document.createElement("div");
  list.className = "wires";
  for (let i = 0; i < a.wireCount; i++) {
    const idx = i + 1;
    const cut = state.cut.includes(idx);
    const btn = document.createElement("button");
    btn.className = `chip ${cut ? "ok" : ""}`;
    btn.textContent = `Wire ${idx}: ${a.wires[i]}`;
    btn.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission) || state.bypassed;
    btn.addEventListener("click", async () => {
      let m = await apiGet(code);
      const ans = m.snapshot.answers.wiring;
      const st = m.runtime.panelState.wiring || { cut: [], bypassed: false };

      // Clicking a wire equals cutting it.
      if (st.bypassed) return;
      if (st.cut.includes(idx)) return;

      // Determine if cut is valid sequence
      const nextExpected = ans.correctOrder[st.cut.length] || null;
      const hasCutsRequired = (ans.correctOrder || []).length > 0;

      // If no cuts required, any cut is instant fail
      if (!hasCutsRequired) {
        m = await applyPenaltyOrFail(code, m, { immediateFail: true });
        return;
      }

      // Wrong wire or wrong order => immediate failure
      if (nextExpected !== idx) {
        m = await applyPenaltyOrFail(code, m, { immediateFail: true });
        return;
      }

      st.cut = st.cut.concat([idx]);
      // If completed sequence, mark panel complete
      const completedNow = st.cut.length === ans.correctOrder.length;
      const completed = { ...m.runtime.completed, wiring: completedNow ? true : false };
      const runtime = {
        ...m.runtime,
        completed,
        panelState: { ...m.runtime.panelState, wiring: st },
      };
      m = await apiWrite(code, { runtime }, m.version);
      if (completedNow) {
        flash("ok");
        toast("Wiring panel complete");
      }
    });
    list.appendChild(btn);
  }

  const bypass = document.createElement("button");
  bypass.className = "btn";
  bypass.textContent = "BYPASS";
  bypass.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
  bypass.addEventListener("click", async () => {
    let m = await apiGet(code);
    const ans = m.snapshot.answers.wiring;
    const st = m.runtime.panelState.wiring || { cut: [], bypassed: false };

    const cutsRequired = (ans.correctOrder || []).length;
    if (cutsRequired > 0) {
      // Pressing BYPASS when cuts required => immediate failure
      m = await applyPenaltyOrFail(code, m, { immediateFail: true });
      return;
    }

    st.bypassed = true;
    const runtime = {
      ...m.runtime,
      completed: { ...m.runtime.completed, wiring: true },
      panelState: { ...m.runtime.panelState, wiring: st },
    };
    m = await apiWrite(code, { runtime }, m.version);
    flash("ok");
    toast("Wiring bypassed successfully");
  });

  const hint = document.createElement("div");
  hint.className = "muted";
  hint.textContent = "Cut wires by selecting them. Any wiring error fails the mission.";

  wrap.appendChild(hint);
  wrap.appendChild(list);
  wrap.appendChild(bypass);
  return wrap;
}

function renderCosmic(mission, code, locked) {
  const wrap = document.createElement("div");
  const ans = mission.snapshot.answers.cosmic;
  const st = mission.runtime.panelState.cosmic || { picked: [] };

  const row = document.createElement("div");
  row.className = "row";

  for (const sym of ans.symbols) {
    const btn = document.createElement("button");
    btn.className = "chip";
    btn.textContent = sym;
    btn.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
    btn.addEventListener("click", async () => {
      let m = await apiGet(code);
      const a = m.snapshot.answers.cosmic;
      const s = m.runtime.panelState.cosmic || { picked: [] };
      const next = a.correct[s.picked.length];
      if (sym !== next) {
        m = await applyPenaltyOrFail(code, m, { immediateFail: false });
        // if strikes hit 3 => fail
        const strikes = (m.runtime.strikes || 0);
        if (strikes >= 3) {
          m = await apiWrite(code, { runtime: { ...m.runtime, status: "failed", stoppedAt: now() } }, m.version);
        }
        return;
      }
      s.picked = s.picked.concat([sym]);
      const completeNow = s.picked.length === a.correct.length;
      const runtime = {
        ...m.runtime,
        panelState: { ...m.runtime.panelState, cosmic: s },
        completed: { ...m.runtime.completed, cosmic: completeNow ? true : false },
      };
      m = await apiWrite(code, { runtime }, m.version);
      if (completeNow) { flash("ok"); toast("Cosmic array complete"); }
    });
    row.appendChild(btn);
  }

  const reset = document.createElement("button");
  reset.className = "btn";
  reset.textContent = "Reset Sequence";
  reset.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
  reset.addEventListener("click", async () => {
    let m = await apiGet(code);
    const runtime = {
      ...m.runtime,
      panelState: { ...m.runtime.panelState, cosmic: { picked: [] } },
    };
    m = await apiWrite(code, { runtime }, m.version);
    toast("Selection cleared");
  });

  wrap.appendChild(document.createElement("div")).className="muted";
  wrap.firstChild.textContent = "Select symbols in the correct order. Mistakes cost time.";
  wrap.appendChild(row);
  wrap.appendChild(reset);
  return wrap;
}

function renderThruster(mission, code, locked) {
  const wrap = document.createElement("div");
  const ans = mission.snapshot.answers.thruster;
  const st = mission.runtime.panelState.thruster || { value: 50 };

  const label = document.createElement("div");
  label.className = "muted";
  label.textContent = `Set thruster output and lock it in.`;

  const input = document.createElement("input");
  input.type = "range";
  input.min = "1";
  input.max = "100";
  input.value = String(st.value ?? 50);
  input.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);

  const readout = document.createElement("div");
  readout.className = "mono";
  readout.textContent = `${input.value}%`;
  input.addEventListener("input", () => { readout.textContent = `${input.value}%`; });

  const btn = document.createElement("button");
  btn.className = "btn";
  btn.textContent = "Lock In";
  btn.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
  btn.addEventListener("click", async () => {
    let m = await apiGet(code);
    const target = m.snapshot.answers.thruster.target;
    const val = parseInt(input.value, 10);
    if (val !== target) {
      m = await applyPenaltyOrFail(code, m, { immediateFail: false });
      if ((m.runtime.strikes || 0) >= 3) {
        m = await apiWrite(code, { runtime: { ...m.runtime, status: "failed", stoppedAt: now() } }, m.version);
      }
      toast("Incorrect");
      return;
    }
    const runtime = {
      ...m.runtime,
      completed: { ...m.runtime.completed, thruster: true },
      panelState: { ...m.runtime.panelState, thruster: { value: val } },
    };
    m = await apiWrite(code, { runtime }, m.version);
    flash("ok");
    toast("Thruster locked");
  });

  wrap.appendChild(label);
  wrap.appendChild(readout);
  wrap.appendChild(input);
  wrap.appendChild(btn);
  return wrap;
}

function renderFrequency(mission, code, locked) {
  const wrap = document.createElement("div");
  const ans = mission.snapshot.answers.frequency;
  const st = mission.runtime.panelState.frequency || { a: 0, b: 0, c: 0 };

  const label = document.createElement("div");
  label.className = "muted";
  label.textContent = "Align all three channels, then press Align.";

  const row = document.createElement("div");
  row.className = "row";

  const inputs = ["A","B","C"].map((ch, i) => {
    const box = document.createElement("div");
    box.className = "field";
    const lab = document.createElement("div");
    lab.className = "muted";
    lab.textContent = `Channel ${ch}`;
    const inp = document.createElement("input");
    inp.type = "number";
    inp.min = "0";
    inp.max = "9";
    inp.value = String(st[["a","b","c"][i]] ?? 0);
    inp.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
    box.appendChild(lab);
    box.appendChild(inp);
    row.appendChild(box);
    return inp;
  });

  const btn = document.createElement("button");
  btn.className = "btn";
  btn.textContent = "Align";
  btn.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
  btn.addEventListener("click", async () => {
    let m = await apiGet(code);
    const target = m.snapshot.answers.frequency.target;
    const vals = inputs.map(i => parseInt(i.value || "0", 10));
    if (vals.some((v, i) => v !== target[i])) {
      m = await applyPenaltyOrFail(code, m, { immediateFail: false });
      if ((m.runtime.strikes || 0) >= 3) {
        m = await apiWrite(code, { runtime: { ...m.runtime, status: "failed", stoppedAt: now() } }, m.version);
      }
      toast("Incorrect");
      return;
    }
    const runtime = {
      ...m.runtime,
      completed: { ...m.runtime.completed, frequency: true },
      panelState: { ...m.runtime.panelState, frequency: { a: vals[0], b: vals[1], c: vals[2] } },
    };
    m = await apiWrite(code, { runtime }, m.version);
    flash("ok");
    toast("Frequency aligned");
  });

  wrap.appendChild(label);
  wrap.appendChild(row);
  wrap.appendChild(btn);
  return wrap;
}

function renderLife(mission, code, locked) {
  const wrap = document.createElement("div");
  const ans = mission.snapshot.answers.life;
  const st = mission.runtime.panelState.life || { oxygen: false, coolant: false, scrubbers: false, mode: "NOMINAL" };

  const label = document.createElement("div");
  label.className = "muted";
  label.textContent = "Set system toggles + mode, then Apply State.";

  const grid = document.createElement("div");
  grid.className = "row";

  const makeToggle = (key, text) => {
    const lab = document.createElement("label");
    lab.className = "toggle";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!st[key];
    cb.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
    const span = document.createElement("span");
    span.textContent = text;
    lab.appendChild(cb);
    lab.appendChild(span);
    return { lab, cb };
  };

  const t1 = makeToggle("oxygen", "Oxygen Routing");
  const t2 = makeToggle("coolant", "Coolant Flow");
  const t3 = makeToggle("scrubbers", "CO₂ Scrubbers");
  grid.appendChild(t1.lab);
  grid.appendChild(t2.lab);
  grid.appendChild(t3.lab);

  const mode = document.createElement("select");
  for (const m of ["NOMINAL","ECO","EMERGENCY"]) {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    if ((st.mode || "NOMINAL") === m) opt.selected = true;
    mode.appendChild(opt);
  }
  mode.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);

  const btn = document.createElement("button");
  btn.className = "btn";
  btn.textContent = "Apply State";
  btn.disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission);
  btn.addEventListener("click", async () => {
    let m = await apiGet(code);
    const a = m.snapshot.answers.life;
    const vals = {
      oxygen: !!t1.cb.checked,
      coolant: !!t2.cb.checked,
      scrubbers: !!t3.cb.checked,
      mode: mode.value,
    };
    const ok = vals.oxygen === a.checks.oxygen && vals.coolant === a.checks.coolant && vals.scrubbers === a.checks.scrubbers && vals.mode === a.mode;
    if (!ok) {
      m = await applyPenaltyOrFail(code, m, { immediateFail: false });
      if ((m.runtime.strikes || 0) >= 3) {
        m = await apiWrite(code, { runtime: { ...m.runtime, status: "failed", stoppedAt: now() } }, m.version);
      }
      toast("Incorrect");
      return;
    }
    const runtime = {
      ...m.runtime,
      completed: { ...m.runtime.completed, life: true },
      panelState: { ...m.runtime.panelState, life: vals },
    };
    m = await apiWrite(code, { runtime }, m.version);
    flash("ok");
    toast("Life support stable");
  });

  wrap.appendChild(label);
  wrap.appendChild(grid);
  wrap.appendChild(mode);
  wrap.appendChild(btn);
  return wrap;
}

function renderKeypad(mission, code, locked) {
  const wrap = document.createElement("div");
  const label = document.createElement("div");
  label.className = "muted";
  label.textContent = "Final Override is locked until all other panels complete.";

  const input = document.createElement("input");
  input.className = "mono";
  input.inputMode = "numeric";
  input.placeholder = "Enter 4-digit pin";

  const btn = document.createElement("button");
  btn.className = "btn";
  btn.textContent = "Submit";

  const prereq = ["wiring","cosmic","thruster","frequency","life"].every(k => !!mission?.runtime?.completed?.[k]);
  const disabled = locked || !mission.runtime.startedAt || missionFailed(mission) || missionWon(mission) || !prereq;
  input.disabled = disabled;
  btn.disabled = disabled;

  btn.addEventListener("click", async () => {
    let m = await apiGet(code);
    const prereq2 = ["wiring","cosmic","thruster","frequency","life"].every(k => !!m?.runtime?.completed?.[k]);
    if (!prereq2) return;
    const expected = m.snapshot.answers.keypad.code;
    const val = (input.value || "").trim();
    if (val !== expected) {
      m = await applyPenaltyOrFail(code, m, { immediateFail: false });
      if ((m.runtime.strikes || 0) >= 3) {
        m = await apiWrite(code, { runtime: { ...m.runtime, status: "failed", stoppedAt: now() } }, m.version);
      }
      toast("Incorrect");
      return;
    }
    const runtime = {
      ...m.runtime,
      completed: { ...m.runtime.completed, keypad: true },
    };
    m = await apiWrite(code, { runtime }, m.version);
    flash("ok");
    toast("Override accepted");
  });

  wrap.appendChild(label);
  wrap.appendChild(input);
  wrap.appendChild(btn);
  return wrap;
}

function renderAux(mission) {
  const wrap = document.createElement("div");
  const lines = document.createElement("div");
  lines.className = "mono";
  const rng = rngFromSeed(`${mission?.snapshot?.seed || "aux"}::${Math.floor(now()/1000)}`);
  const temp = (rng()*14 + 18).toFixed(1);
  const rad = (rng()*0.8).toFixed(3);
  const flux = (rng()*99).toFixed(2);
  lines.innerHTML = `TEMP: ${temp} C\nRADIATION: ${rad} mSv\nFLUX: ${flux} kHz\nUPLINK: STABLE`;
  wrap.appendChild(lines);
  return wrap;
}

function renderPower(mission, code) {
  const wrap = document.createElement("div");
  const label = document.createElement("div");
  label.className = "muted";
  label.textContent = "Flip ONLINE only after all panels complete.";

  const btn = document.createElement("button");
  btn.className = "btn";
  const online = !!mission?.runtime?.powerOnline;
  btn.textContent = online ? "ONLINE" : "OFFLINE";

  const canFlip = mission.runtime.startedAt && !missionFailed(mission) && !missionWon(mission) && allRequiredPanelsComplete(mission);
  btn.disabled = !canFlip;

  btn.addEventListener("click", async () => {
    let m = await apiGet(code);
    if (!allRequiredPanelsComplete(m)) return;
    const runtime = {
      ...m.runtime,
      powerOnline: true,
      stoppedAt: now(),
      status: "won",
    };
    m = await apiWrite(code, { runtime }, m.version);
    flash("ok");
    toast("System Power: ONLINE");
  });

  wrap.appendChild(label);
  wrap.appendChild(btn);
  return wrap;
}

// ---------------- Manuals ----------------

async function bootManuals() {
  const code = (qparam("mission") || "").trim();
  if (!code) {
    setBanner("Missing mission code", "bad");
    return;
  }
  $("#roomCode").textContent = code;

  let mission = await apiGet(code);
  if (!mission?.snapshot) {
    setBanner("Manuals unavailable until facilitator applies setup.", "warn");
    return;
  }

  if (mission?.snapshot?.manualsOutOfDate) {
    setBanner("Manuals are OUT OF DATE. Ask facilitator to Generate Manuals.", "bad");
  }

  const isDownload = qparam("download") === "1";
  if (isDownload) {
    // render and trigger download with simple html wrapper
    renderManuals(mission);
    document.title = `Manuals ${code}`;
    // Let browser's Save feature handle.
    return;
  }

  renderManuals(mission);

  // poll for regeneration
  let lastVersion = mission.version;
  while (true) {
    await sleep(1500);
    const m = await apiGet(code).catch(() => null);
    if (m && m.version !== lastVersion) {
      lastVersion = m.version;
      mission = m;
      renderManuals(mission);
    }
  }
}

function generateManuals(mission) {
  const code = mission.code;
  const seed = mission.snapshot.seed;
  const rng = rngFromSeed(`${seed}::manuals`);

  const panels = ["wiring","cosmic","thruster","frequency","life","keypad"];
  const items = panels.map((key, i) => {
    const id = mission.snapshot.panelIds[key];
    const header = `${PANEL_META[key].label} Manual #${i+1}`; // header doesn't include panel name? spec says generic headers only, not panel names.
    // We'll enforce generic type-only. Use Wiring Manual #1 style.
    const typeHeader = `${key.charAt(0).toUpperCase()+key.slice(1)} Manual #${i+1}`;
    const gateQCount = Math.floor(rng() * 3) + 1;
    const gates = Array.from({length: gateQCount}, (_, g) => {
      const grng = rngFromSeed(`${seed}::gate::${key}::${g}`);
      // simple puzzles
      const q = pick(grng, [
        "What is the third letter of the word ORBIT?",
        "How many sides does a hexagon have?",
        "Type the last digit of 3.1415.",
        "What comes after G in the alphabet?",
        "What is 2+5?",
      ]);
      const a = (
        q.includes("third letter") ? "B" :
        q.includes("hexagon") ? "6" :
        q.includes("3.1415") ? "5" :
        q.includes("after G") ? "H" :
        "7"
      );
      return { q, a };
    });

    const optionCount = Math.floor(rng() * 3) + 2; // 2..4
    const correct = buildCorrectInstruction(mission, key);
    const decoys = buildDecoys(mission, key, correct, optionCount - 1);
    const options = shuffle(rngFromSeed(`${seed}::opts::${key}`), [correct, ...decoys]).map((text, idx) => ({ id: `${key}::${idx}`, text, isCorrect: text === correct }));

    const custom = mission.snapshot.overrides?.customInstructions?.[key] || "";
    if (mission.snapshot.allowCustom && custom) {
      // Replace the correct option with custom instruction (still conditional)
      const correctIndex = options.findIndex(o => o.isCorrect);
      if (correctIndex >= 0) options[correctIndex] = { ...options[correctIndex], text: custom, isCorrect: true };
    }

    return {
      key,
      panelId: id,
      header: `${PANEL_META[key].label} Manual #${i+1}`,
      // Enforce "Wiring Manual #1" style (no panel names beyond type)
      title: `${PANEL_META[key].label} Manual #${i+1}`.replace("Cosmic","Symbol Array").replace("Life","Life Support").replace("Keypad","Final Override").replace("Thruster","Thruster").replace("Frequency","Frequency").replace("Wiring","Wiring"),
      mapText: `Applies to Panel ID ${id}.`,
      gates,
      options,
    };
  });

  return {
    generatedAt: now(),
    items,
  };
}

function buildCorrectInstruction(mission, key) {
  const seed = mission.snapshot.seed;
  const rng = rngFromSeed(`${seed}::instr::${key}`);
  // Conditional, not direct imperative.
  if (key === "wiring") {
    const ans = mission.snapshot.answers.wiring;
    if ((ans.correctOrder || []).length === 0) {
      return "If the panel’s indicator is steady and no cut sequence is listed, then engage the BYPASS control.";
    }
    const first = ans.correctOrder[0];
    const last = ans.correctOrder[ans.correctOrder.length - 1];
    return `If the first wire is ${ans.wires[first-1].toLowerCase()} and the last correct wire is numbered ${last}, then the cut sequence begins with wire ${first}.`;
  }
  if (key === "cosmic") {
    const ans = mission.snapshot.answers.cosmic;
    return `If the smallest symbol appears to the left of a star-like glyph, then select the symbols in the order the astronaut reads them aloud, starting with ${ans.correct[0]}.`;
  }
  if (key === "thruster") {
    const target = mission.snapshot.answers.thruster.target;
    return `If the requested output is an odd value, then lock in at ${target} percent; otherwise, lock in at the nearest even value.`;
  }
  if (key === "frequency") {
    const t = mission.snapshot.answers.frequency.target;
    return `If Channel A equals ${t[0]}, then the remaining channels follow as B=${t[1]} and C=${t[2]} before alignment.`;
  }
  if (key === "life") {
    const a = mission.snapshot.answers.life;
    const mode = a.mode;
    return `If two or more life toggles are active, then set the mode to ${mode} before applying state.`;
  }
  if (key === "keypad") {
    return "If all other panels report COMPLETE, then the final override accepts a four-digit pattern that the specialist validates before entry.";
  }
  return "If conditions match, proceed.";
}

function buildDecoys(mission, key, correct, count) {
  const seed = mission.snapshot.seed;
  const out = [];
  for (let i = 0; i < 12 && out.length < count; i++) {
    const rng = rngFromSeed(`${seed}::decoy::${key}::${i}`);
    let text = correct;
    // Mutate numbers/symbol references while staying plausible
    if (key === "wiring") {
      const ans = mission.snapshot.answers.wiring;
      if ((ans.correctOrder || []).length === 0) {
        text = "If the panel’s indicator is steady and no cut sequence is listed, then do not engage BYPASS until the astronaut confirms an alert tone.";
      } else {
        const wrongFirst = clamp(ans.correctOrder[0] + (rng() < 0.5 ? 1 : -1), 1, ans.wireCount);
        text = `If the first wire is ${ans.wires[wrongFirst-1].toLowerCase()} and the last correct wire is numbered ${ans.correctOrder.length}, then the cut sequence begins with wire ${wrongFirst}.`;
      }
    } else if (key === "cosmic") {
      const ans = mission.snapshot.answers.cosmic;
      const alt = pick(rng, ans.symbols);
      text = `If the crescent icon appears adjacent to a star, then begin selection with ${alt} and continue clockwise.`;
    } else if (key === "thruster") {
      const target = mission.snapshot.answers.thruster.target;
      const alt = clamp(target + (rng() < 0.5 ? 5 : -5), 1, 100);
      text = `If the requested output is an odd value, then lock in at ${alt} percent; otherwise, lock in at the nearest even value.`;
    } else if (key === "frequency") {
      const t = mission.snapshot.answers.frequency.target;
      const a = clamp(t[0] + 1, 0, 9);
      text = `If Channel A equals ${a}, then the remaining channels follow as B=${t[1]} and C=${t[2]} before alignment.`;
    } else if (key === "life") {
      const a = mission.snapshot.answers.life;
      const alt = pick(rng, ["NOMINAL","ECO","EMERGENCY"].filter(x => x !== a.mode));
      text = `If two or more life toggles are active, then set the mode to ${alt} before applying state.`;
    } else if (key === "keypad") {
      text = "If any panel is still pending, then the override will accept a four-digit pattern only after a full reset sequence.";
    }

    if (text !== correct && !out.includes(text)) out.push(text);
  }
  return out.slice(0, count);
}

function renderManuals(mission) {
  const root = $("#manualRoot");
  root.innerHTML = "";

  const manuals = mission.manuals?.items?.length ? mission.manuals : generateManuals(mission);

  for (const item of manuals.items) {
    const card = document.createElement("div");
    card.className = "card manual";
    const h = document.createElement("div");
    h.className = "cardHeader";
    h.innerHTML = `<div class="title">${item.title}</div><div class="sub">${item.mapText}</div>`;
    card.appendChild(h);

    const body = document.createElement("div");
    body.className = "cardBody";

    const gateWrap = document.createElement("div");
    gateWrap.className = "manual";

    const gateTitle = document.createElement("div");
    gateTitle.className = "muted";
    gateTitle.textContent = `Gate questions (${item.gates.length})`;
    gateWrap.appendChild(gateTitle);

    const inputs = [];
    item.gates.forEach((g, idx) => {
      const row = document.createElement("div");
      row.className = "field";
      row.innerHTML = `<div class="muted">Q${idx+1}: ${g.q}</div>`;
      const inp = document.createElement("input");
      inp.placeholder = "answer";
      row.appendChild(inp);
      gateWrap.appendChild(row);
      inputs.push({ inp, ans: g.a });
    });

    const unlock = document.createElement("button");
    unlock.className = "btn";
    unlock.textContent = "Unlock";

    const optionsWrap = document.createElement("div");
    optionsWrap.className = "manual hidden";

    unlock.addEventListener("click", () => {
      const ok = inputs.every(({inp, ans}) => (inp.value || "").trim().toUpperCase() === String(ans).toUpperCase());
      if (!ok) {
        flash("warn");
        toast("Incorrect gate answer");
        return;
      }
      optionsWrap.classList.remove("hidden");
      unlock.disabled = true;
      inputs.forEach(({inp}) => inp.disabled = true);
      flash("ok");
    });

    gateWrap.appendChild(unlock);

    const optTitle = document.createElement("div");
    optTitle.className = "muted";
    optTitle.textContent = "Instruction options (exactly one is correct)";
    optionsWrap.appendChild(optTitle);

    item.options.forEach((o, idx) => {
      const opt = document.createElement("div");
      opt.className = "option";
      opt.innerHTML = `<strong>Option ${idx+1}</strong><div>${o.text}</div>`;
      optionsWrap.appendChild(opt);
    });

    body.appendChild(gateWrap);
    body.appendChild(optionsWrap);
    card.appendChild(body);
    root.appendChild(card);
  }
}

