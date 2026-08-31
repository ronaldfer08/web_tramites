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
  const convertedFiles = new Map();
  let pendingConversions = 0;

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
    return parts.length > 1 ? parts.pop().toLowerCase() : "";
  }

  function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error(`No se pudo leer ${file.name}.`));
      reader.readAsDataURL(file);
    });
  }

  function loadImage(dataUrl, fileName) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () =>
        reject(new Error(`No se pudo procesar la imagen ${fileName}.`));
      image.src = dataUrl;
    });
  }

  async function convertImageToPdf(file) {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      throw new Error(
        "No se pudo cargar el conversor de imágenes a PDF. Recarga la página."
      );
    }

    const extension = extensionFromFile(file);
    const imageFormat =
      file.type === "image/png" || extension === "png" ? "PNG" : "JPEG";
    const dataUrl = await readFileAsDataUrl(file);
    const image = await loadImage(dataUrl, file.name);
    const orientation =
      image.naturalWidth > image.naturalHeight ? "landscape" : "portrait";
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
      orientation,
      unit: "pt",
      format: "a4",
      compress: true,
    });

    const margin = 24;
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const availableWidth = pageWidth - margin * 2;
    const availableHeight = pageHeight - margin * 2;
    const scale = Math.min(
      availableWidth / image.naturalWidth,
      availableHeight / image.naturalHeight
    );
    const width = image.naturalWidth * scale;
    const height = image.naturalHeight * scale;
    const x = (pageWidth - width) / 2;
    const y = (pageHeight - height) / 2;

    pdf.addImage(dataUrl, imageFormat, x, y, width, height, undefined, "MEDIUM");

    const baseName = file.name.replace(/\.[^.]+$/, "") || "documento";
    return new File([pdf.output("blob")], `${baseName}.pdf`, {
      type: "application/pdf",
      lastModified: Date.now(),
    });
  }

  function nextPaint() {
    return new Promise((resolve) => {
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    });
  }

  async function convertDocxToPdf(file) {
    if (!window.docx || typeof window.docx.renderAsync !== "function") {
      throw new Error(
        "No se pudo cargar el lector de Word. Recarga la página."
      );
    }

    if (!window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) {
      throw new Error(
        "No se pudo cargar el conversor de Word a PDF. Recarga la página."
      );
    }

    const renderContainer = document.createElement("div");
    renderContainer.setAttribute("aria-hidden", "true");
    Object.assign(renderContainer.style, {
      position: "fixed",
      left: "-100000px",
      top: "0",
      width: "900px",
      background: "#ffffff",
      pointerEvents: "none",
      zIndex: "-1",
    });
    document.body.appendChild(renderContainer);

    try {
      const data = await file.arrayBuffer();
      await window.docx.renderAsync(data, renderContainer, null, {
        breakPages: true,
        ignoreLastRenderedPageBreak: false,
        useBase64URL: true,
      });

      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
      await nextPaint();

      let pages = Array.from(renderContainer.querySelectorAll("section.docx"));
      if (!pages.length) {
        const wrapper = renderContainer.querySelector(".docx-wrapper");
        if (wrapper) pages = [wrapper];
      }
      if (!pages.length) {
        throw new Error("El documento Word no contiene páginas procesables.");
      }

      const { jsPDF } = window.jspdf;
      let pdf = null;

      for (let index = 0; index < pages.length; index += 1) {
        const page = pages[index];
        const canvas = await window.html2canvas(page, {
          scale: Math.min(window.devicePixelRatio || 1, 2),
          useCORS: true,
          backgroundColor: "#ffffff",
          logging: false,
        });
        const orientation =
          canvas.width > canvas.height ? "landscape" : "portrait";

        if (!pdf) {
          pdf = new jsPDF({
            orientation,
            unit: "pt",
            format: "a4",
            compress: true,
          });
        } else {
          pdf.addPage("a4", orientation);
        }

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const scale = Math.min(
          pageWidth / canvas.width,
          pageHeight / canvas.height
        );
        const width = canvas.width * scale;
        const height = canvas.height * scale;
        const x = (pageWidth - width) / 2;
        const y = (pageHeight - height) / 2;

        pdf.addImage(
          canvas.toDataURL("image/jpeg", 0.92),
          "JPEG",
          x,
          y,
          width,
          height,
          undefined,
          "MEDIUM"
        );
      }

      const baseName = file.name.replace(/\.[^.]+$/, "") || "documento";
      return new File([pdf.output("blob")], `${baseName}.pdf`, {
        type: "application/pdf",
        lastModified: Date.now(),
      });
    } finally {
      renderContainer.remove();
    }
  }

  async function prepareFileAsPdf(file) {
    const extension = extensionFromFile(file);

    if (file.type === "application/pdf" || extension === "pdf") {
      return file.type === "application/pdf"
        ? file
        : new File([file], file.name, {
            type: "application/pdf",
            lastModified: file.lastModified,
          });
    }

    if (
      ["image/jpeg", "image/png"].includes(file.type) ||
      ["jpg", "jpeg", "png"].includes(extension)
    ) {
      return convertImageToPdf(file);
    }

    if (extension === "docx") {
      return convertDocxToPdf(file);
    }

    throw new Error(
      `${file.name}: formato no permitido. Usa PDF, JPG, JPEG, PNG o DOCX.`
    );
  }

  function isImageFile(file) {
    const extension = extensionFromFile(file);
    return (
      ["image/jpeg", "image/png"].includes(file.type) ||
      ["jpg", "jpeg", "png"].includes(extension)
    );
  }

  function isDocxFile(file) {
    return extensionFromFile(file) === "docx";
  }

  function collectFiles(formElement) {
    return FILE_FIELDS.map((field) => {
      const input = formElement.elements.namedItem(field.name);
      const selectedFile =
        input && input.files && input.files[0] ? input.files[0] : null;
      const file = convertedFiles.get(field.name) || selectedFile;
      return { ...field, file };
    }).filter((item) => item.file);
  }

  FILE_FIELDS.forEach((field) => {
    const input = form.elements.namedItem(field.name);
    if (!input) return;

    input.addEventListener("change", async () => {
      convertedFiles.delete(field.name);
      const selectedFile = input.files && input.files[0];
      if (
        !selectedFile ||
        (!isImageFile(selectedFile) && !isDocxFile(selectedFile))
      ) {
        return;
      }

      pendingConversions += 1;
      submitBtn.disabled = true;
      setStatus(`Convirtiendo ${field.label} a PDF...`, "info");

      try {
        const pdfFile = isDocxFile(selectedFile)
          ? await convertDocxToPdf(selectedFile)
          : await convertImageToPdf(selectedFile);

        // El archivo convertido se guarda aparte para funcionar también en
        // navegadores móviles que no permiten modificar input.files.
        const currentFile = input.files && input.files[0];
        const selectionIsCurrent =
          currentFile &&
          currentFile.name === selectedFile.name &&
          currentFile.size === selectedFile.size &&
          currentFile.lastModified === selectedFile.lastModified;

        if (!selectionIsCurrent) return;

        convertedFiles.set(field.name, pdfFile);
        setStatus(`${field.label} se convirtió correctamente a PDF.`, "success");
      } catch (error) {
        convertedFiles.delete(field.name);
        input.value = "";
        setStatus(
          `No se pudo convertir ${field.label}: ${error.message}`,
          "danger"
        );
      } finally {
        pendingConversions -= 1;
        if (pendingConversions === 0) {
          submitBtn.disabled = false;
        }
      }
    });
  });

  function sanitizeStorageName(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[\/\\?%*:|"<>]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function buildFolderName(apellidosNombres) {
    return sanitizeStorageName(apellidosNombres);
  }

  async function uploadDocuments(apellidosNombres, files) {
    const uploaded = [];
    const folderName = buildFolderName(apellidosNombres);

    for (let i = 0; i < files.length; i += 1) {
      const item = files[i];

      setStatus(
        `Procesando ${i + 1} de ${files.length}: ${item.label}...`,
        "info"
      );

      const pdfFile = await prepareFileAsPdf(item.file);
      const storedFileName = `${item.label}.pdf`;
      const path = `${folderName}/${storedFileName}`;

      if (
        pdfFile.type !== "application/pdf" ||
        extensionFromFile(pdfFile) !== "pdf"
      ) {
        throw new Error(
          `${item.label}: la conversión no produjo un archivo PDF válido.`
        );
      }

      setStatus(
        `Subiendo ${i + 1} de ${files.length}: ${storedFileName}...`,
        "info"
      );

      let { error } = await client.storage.from(bucket).upload(path, pdfFile, {
        cacheControl: "3600",
        upsert: false,
        contentType: "application/pdf",
      });

      // Si ya existe, sobrescribe (requiere política UPDATE)
      if (error && /already exists/i.test(error.message)) {
        ({ error } = await client.storage.from(bucket).upload(path, pdfFile, {
          cacheControl: "3600",
          upsert: true,
          contentType: "application/pdf",
        }));
      }

      if (error) {
        throw new Error(`${item.label}: ${error.message}`);
      }

      uploaded.push(path);
    }

    const metaPath = `${folderName}/datos.json`;
    const meta = {
      apellidos_y_nombres: apellidosNombres,
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

    return { uploaded, folderName };
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

    const apellidosNombres = String(form.apellidos_nombres.value || "").trim();

    if (pendingConversions > 0) {
      setStatus(
        "Espera a que terminen de convertirse los documentos a PDF.",
        "warning"
      );
      return;
    }

    if (!apellidosNombres) {
      setStatus("Debes ingresar los apellidos y nombres.", "warning");
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
      const { uploaded, folderName } = await uploadDocuments(
        apellidosNombres,
        files
      );
      setStatus(
        `Listo. Se guardaron ${uploaded.length} archivo(s) en la carpeta ${folderName}.`,
        "success"
      );
      form.reset();
      convertedFiles.clear();
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
