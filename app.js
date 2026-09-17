const player = document.getElementById("player");
const sinks = document.getElementById("sinks");
const status = document.getElementById("status");
const refreshBtn = document.getElementById("refresh");

const looksLikeA50 = (label) => /a50/i.test(label || "");

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
      "No se pueden listar dispositivos aquí. Reproducí la canción en el A50 abriendo esta página en el teléfono.",
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

  let foundA50 = false;
  for (const device of outputs) {
    const option = document.createElement("option");
    option.value = device.deviceId;
    option.textContent = device.label || `Salida ${device.deviceId.slice(0, 8)}`;
    if (looksLikeA50(device.label)) {
      foundA50 = true;
      option.textContent += " (A50)";
    }
    sinks.append(option);
  }

  sinks.disabled = !canSelectSink();

  if (foundA50) {
    const a50 = [...sinks.options].find((opt) => looksLikeA50(opt.textContent));
    if (a50) {
      sinks.value = a50.value;
      await applySink();
    }
    setStatus(
      "El A50 aparece como salida de audio. Ya lo seleccioné: dale play.",
      "ok"
    );
    return;
  }

  if (!outputs.length) {
    setStatus(
      "No hay salidas extra. El A50 está en Bluetooth, pero no figura como parlante.",
      "warn"
    );
    return;
  }

  const named = outputs.some((device) => device.label);
  if (!named) {
    setStatus(
      "Hay salidas, pero el navegador oculta los nombres hasta que des permiso de micrófono. El A50 igual suele no listarse: es teléfono, no auricular.",
      "warn"
    );
    return;
  }

  setStatus(
    "El A50 no está en esta lista. Puede estar conectado en Bluetooth y aun así no servir para escuchar desde acá.",
    "warn"
  );
}

async function applySink() {
  if (!canSelectSink()) {
    setStatus(
      "Este navegador no deja cambiar la salida. Elegí el A50 en Configuración → Sonido de Windows, o abrí la página en el teléfono.",
      "warn"
    );
    return;
  }

  const id = sinks.value;
  try {
    await player.setSinkId(id);
    if (looksLikeA50(sinks.selectedOptions[0]?.textContent)) {
      setStatus("Sonido enviado al A50. Dale play para escuchar.", "ok");
    }
  } catch (error) {
    setStatus(
      `No pude usar esa salida (${error.message}). Si es el A50, el sistema no lo expone como parlante.`,
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
