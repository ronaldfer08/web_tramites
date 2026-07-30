(function () {
  const FILE_FIELDS = [
    { name: "doc_identidad", label: "dni_o_ce" },
    { name: "antecedentes", label: "antecedentes" },
    { name: "recibo_servicios", label: "recibo_servicios" },
    { name: "movimiento_migratorio", label: "movimiento_migratorio" },
    { name: "estudios", label: "estudios" },
    { name: "trabajo", label: "trabajo" },
    { name: "retenciones", label: "retenciones" },
    { name: "pareja_matrimonio", label: "pareja_matrimonio" },
    { name: "dni_hijos", label: "dni_hijos" },
    { name: "estudios_hijo", label: "estudios_hijo" },
  ];

  const form = document.getElementById("form-ingreso");
  const statusEl = document.getElementById("upload-status");
  const submitBtn = document.getElementById("btn-enviar");

  if (!form || !window.supabase) {
    return;
  }

  const { url, anonKey, bucket } = window.SUPABASE_CONFIG || {};

  function normalizeSupabaseUrl(rawUrl) {
    return String(rawUrl || "")
      .trim()
      .replace(/\/rest\/v1\/?$/i, "")
      .replace(/\/+$/, "");
  }

  const projectUrl = normalizeSupabaseUrl(url);
  const configReady =
    projectUrl &&
    anonKey &&
    !projectUrl.includes("PEGAR_AQUI") &&
    !anonKey.includes("PEGAR_AQUI");

  const client = configReady
    ? window.supabase.createClient(projectUrl, anonKey)
    : null;

  function setStatus(message, type) {
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.className = `upload-status alert alert-${type}`;
    statusEl.hidden = false;
  }

  function clearStatus() {
    if (!statusEl) return;
    statusEl.hidden = true;
    statusEl.textContent = "";
    statusEl.className = "upload-status";
  }

  function extensionFromFile(file) {
    const parts = file.name.split(".");
    const ext = parts.length > 1 ? parts.pop().toLowerCase() : "";
    if (["pdf", "jpg", "jpeg", "png", "webp"].includes(ext)) {
      return ext === "jpeg" ? "jpg" : ext;
    }
    return "bin";
  }

  function collectFiles(formElement) {
    return FILE_FIELDS.map((field) => {
      const input = formElement.elements.namedItem(field.name);
      const file = input && input.files && input.files[0] ? input.files[0] : null;
      return { ...field, file };
    }).filter((item) => item.file);
  }

  async function uploadDocuments(dni, nombres, files) {
    const uploaded = [];

    for (let i = 0; i < files.length; i += 1) {
      const item = files[i];
      const ext = extensionFromFile(item.file);
      const path = `${dni}/${item.label}.${ext}`;

      setStatus(
        `Subiendo ${i + 1} de ${files.length}: ${item.label}...`,
        "info"
      );

      let { error } = await client.storage.from(bucket).upload(path, item.file, {
        cacheControl: "3600",
        upsert: false,
        contentType: item.file.type || undefined,
      });

      // Si ya existe, sobrescribe (requiere política UPDATE)
      if (error && /already exists/i.test(error.message)) {
        ({ error } = await client.storage.from(bucket).upload(path, item.file, {
          cacheControl: "3600",
          upsert: true,
          contentType: item.file.type || undefined,
        }));
      }

      if (error) {
        throw new Error(`${item.label}: ${error.message}`);
      }

      uploaded.push(path);
    }

    const metaPath = `${dni}/datos.json`;
    const meta = {
      dni,
      nombres_completos: nombres,
      enviado_en: new Date().toISOString(),
      archivos: uploaded,
    };

    const metaBlob = new Blob([JSON.stringify(meta, null, 2)], {
      type: "application/json",
    });

    let { error: metaError } = await client.storage
      .from(bucket)
      .upload(metaPath, metaBlob, {
        cacheControl: "3600",
        upsert: false,
        contentType: "application/json",
      });

    if (metaError && /already exists/i.test(metaError.message)) {
      ({ error: metaError } = await client.storage
        .from(bucket)
        .upload(metaPath, metaBlob, {
          cacheControl: "3600",
          upsert: true,
          contentType: "application/json",
        }));
    }

    if (metaError) {
      throw new Error(`No se pudo guardar datos.json: ${metaError.message}`);
    }

    return uploaded;
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearStatus();

    if (!client) {
      setStatus(
        "Falta configurar Supabase en js/config.js (URL y anon key).",
        "danger"
      );
      return;
    }

    const dni = String(form.dni.value || "").trim();
    const nombres = String(form.nombres_completos.value || "").trim();

    if (!/^\d{8}$/.test(dni)) {
      setStatus("El DNI debe tener exactamente 8 dígitos.", "warning");
      return;
    }

    const files = collectFiles(form);
    if (!files.length) {
      setStatus("Debes adjuntar al menos un documento.", "warning");
      return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = "Enviando...";

    try {
      const uploaded = await uploadDocuments(dni, nombres, files);
      setStatus(
        `Listo. Se guardaron ${uploaded.length} archivo(s) en la carpeta ${dni}.`,
        "success"
      );
      form.reset();
    } catch (error) {
      setStatus(
        `Error al subir: ${error.message || "intenta nuevamente."}`,
        "danger"
      );
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = "📤 Enviar Mis Documentos";
    }
  });
})();
