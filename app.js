const player = document.getElementById("player");
const sinks = document.getElementById("sinks");
const status = document.getElementById("status");
const refreshBtn = document.getElementById("refresh");

const ASTRO_HINT = /astro|a50|mixamp/i;
const GAME_HINT = /game|stereo|auricular|headphone|earphone/i;
const VOICE_HINT = /voice|hands-?free|communications|chat/i;

function scoreAstroLabel(label) {
  const text = label || "";
  if (!ASTRO_HINT.test(text)) return 0;
  if (VOICE_HINT.test(text) && !GAME_HINT.test(text)) return 1;
  if (GAME_HINT.test(text)) return 3;
  return 2;
}

function setStatus(text, kind) {
  status.textContent = text;
  status.className = `status${kind ? ` ${kind}` : ""}`;
}

function canSelectSink() {
  return typeof player.setSinkId === "function";
}

async function unlockOutputs() {
  if (!navigator.mediaDevices?.enumerateDevices) return;
  try {
    await navigator.mediaDevices.getUserMedia({ audio: true });
  } catch {
    // Permission may be denied; labels might stay generic.
  }
}

async function listSinks() {
  if (!navigator.mediaDevices?.enumerateDevices) {
    sinks.innerHTML = "";
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "Este navegador no lista salidas de audio";
    sinks.append(option);
    sinks.disabled = true;
    setStatus(
      "No se pueden listar dispositivos aquí. En la PC los ASTRO A50 tienen que aparecer por la base USB como “ASTRO A50 Game”.",
      "warn"
    );
    return;
  }

  const devices = await navigator.mediaDevices.enumerateDevices();
  const outputs = devices.filter((device) => device.kind === "audiooutput");

  sinks.innerHTML = "";
  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = "Salida predeterminada del sistema";
  sinks.append(defaultOption);

  let bestAstro = null;
  let bestScore = 0;
  for (const device of outputs) {
    const option = document.createElement("option");
    option.value = device.deviceId;
    const label = device.label || `Salida ${device.deviceId.slice(0, 8)}`;
    const score = scoreAstroLabel(label);
    option.textContent = score ? `${label} (ASTRO A50)` : label;
    if (score > bestScore) {
      bestScore = score;
      bestAstro = option;
    }
    sinks.append(option);
  }

  sinks.disabled = !canSelectSink();

  if (bestAstro) {
    sinks.value = bestAstro.value;
    await applySink();
    if (bestScore === 1) {
      setStatus(
        "Encontré los A50, pero como Voice/Hands-Free. Cambiá en Windows a “ASTRO A50 Game” para escuchar la canción.",
        "warn"
      );
    } else {
      setStatus(
        "Encontré los ASTRO A50 como salida de audio. Ya los seleccioné: dale play.",
        "ok"
      );
    }
    return;
  }

  if (!outputs.length) {
    setStatus(
      "No hay salidas extra. Los A50 están en Bluetooth, pero la PC no los ve como parlante USB.",
      "warn"
    );
    return;
  }

  const named = outputs.some((device) => device.label);
  if (!named) {
    setStatus(
      "Hay salidas, pero el navegador oculta los nombres. Dale a Actualizar lista y aceptá el micrófono. Igual: los A50 tienen que figurar como “ASTRO A50 Game” por la base USB.",
      "warn"
    );
    return;
  }

  setStatus(
    "Los ASTRO A50 no están en esta lista. Si solo los ves en Bluetooth, desconectá eso y usá la estación base por USB.",
    "warn"
  );
}

async function applySink() {
  if (!canSelectSink()) {
    setStatus(
      "Este navegador no deja cambiar la salida. En Windows elegí “ASTRO A50 Game” en Configuración → Sonido.",
      "warn"
    );
    return;
  }

  const id = sinks.value;
  try {
    await player.setSinkId(id);
    const selected = sinks.selectedOptions[0]?.textContent || "";
    if (scoreAstroLabel(selected) === 1) {
      setStatus(
        "Esa salida es Voice/Hands-Free. Pasate a “ASTRO A50 Game” o no se va a oír bien la canción.",
        "warn"
      );
    } else if (scoreAstroLabel(selected) > 1) {
      setStatus("Sonido enviado a los ASTRO A50. Dale play para escuchar.", "ok");
    }
  } catch (error) {
    setStatus(
      `No pude usar esa salida (${error.message}). Si son los A50, Windows tiene que listar “ASTRO A50 Game” por USB.`,
      "warn"
    );
  }
}

refreshBtn.addEventListener("click", async () => {
  await unlockOutputs();
  await listSinks();
});

sinks.addEventListener("change", applySink);

if (navigator.mediaDevices?.addEventListener) {
  navigator.mediaDevices.addEventListener("devicechange", listSinks);
}

listSinks();
