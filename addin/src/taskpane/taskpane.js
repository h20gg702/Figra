// ========= その他の既存機能群（ファイル保存、テンプレート等）=========

// Add-in version (update this when making breaking changes)
const ADDIN_VERSION = "1.0.0";

// Grouped chart types constant (used throughout the file)
const GROUPED_CHART_TYPES = ["bar_grouped", "bar_grouped_error", "bar_grouped_error_dot", "box_grouped", "box_grouped_dot", "violin_grouped", "violin_grouped_dot", "line_grouped", "line_grouped_error", "line_grouped_error_raw"];

// ========= Registration System =========
const REGISTRATION_KEY = "figra_registered";
const GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSczQzQZmWDl0JBtTyy0mv-5h5FsRmSyFDEIXfWXet_vs6WvKQ/formResponse";

// Google Form entry IDs (extracted from form)
const FORM_FIELDS = {
  email: "entry.650578732",
  country: "entry.838311149",
  jobTitle: "entry.16463827",
  affiliation: "entry.810594215",
  agreement: "entry.438747924"
};

function isRegistered() {
  try {
    return localStorage.getItem(REGISTRATION_KEY) === "true";
  } catch (e) {
    console.warn("localStorage not available:", e);
    return true; // If localStorage fails, skip registration
  }
}

function setRegistered() {
  try {
    localStorage.setItem(REGISTRATION_KEY, "true");
  } catch (e) {
    console.warn("Could not save registration status:", e);
  }
}

function showRegistrationOverlay() {
  const overlay = document.getElementById("registrationOverlay");
  if (overlay) {
    overlay.classList.remove("hidden");
  }
}

function hideRegistrationOverlay() {
  const overlay = document.getElementById("registrationOverlay");
  if (overlay) {
    overlay.classList.add("hidden");
  }
}

async function submitRegistration() {
  const email = document.getElementById("regEmail")?.value?.trim();
  const jobTitle = document.getElementById("regJobTitle")?.value?.trim();
  const affiliation = document.getElementById("regAffiliation")?.value?.trim();
  const country = document.getElementById("regCountry")?.value;
  const termsAccepted = document.getElementById("regTerms")?.checked;
  const errorEl = document.getElementById("regError");
  const submitBtn = document.getElementById("regSubmitBtn");

  // Validation
  if (!email || !jobTitle || !affiliation || !country) {
    errorEl.textContent = "Please fill in all required fields.";
    errorEl.style.display = "block";
    return;
  }

  if (!termsAccepted) {
    errorEl.textContent = "Please accept the Terms of Service.";
    errorEl.style.display = "block";
    return;
  }

  // Email validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    errorEl.textContent = "Please enter a valid email address.";
    errorEl.style.display = "block";
    return;
  }

  errorEl.style.display = "none";
  submitBtn.disabled = true;
  submitBtn.textContent = "Submitting...";

  try {
    // Submit to Google Forms using hidden iframe (more reliable in Office Add-in)
    const params = new URLSearchParams();
    params.append(FORM_FIELDS.email, email);
    params.append(FORM_FIELDS.jobTitle, jobTitle);
    params.append(FORM_FIELDS.affiliation, affiliation);
    params.append(FORM_FIELDS.country, country);
    params.append(FORM_FIELDS.agreement, termsAccepted ? "I have read and agree to the Terms of Use and Privacy Policy." : "No");

    // Create hidden iframe for form submission
    const iframe = document.createElement("iframe");
    iframe.name = "hidden_iframe";
    iframe.style.display = "none";
    document.body.appendChild(iframe);

    // Create and submit form
    const form = document.createElement("form");
    form.method = "POST";
    form.action = GOOGLE_FORM_URL;
    form.target = "hidden_iframe";

    for (const [key, value] of params) {
      const input = document.createElement("input");
      input.type = "hidden";
      input.name = key;
      input.value = value;
      form.appendChild(input);
    }

    document.body.appendChild(form);
    form.submit();

    // Cleanup after a delay
    setTimeout(() => {
      form.remove();
      iframe.remove();
    }, 2000);

    // Mark as registered and hide overlay
    setRegistered();
    hideRegistrationOverlay();

    console.log("✅ Registration submitted successfully");

  } catch (error) {
    console.error("Registration error:", error);
    // Even if submission fails, allow user to proceed (mark as registered)
    // This prevents blocking users due to network issues
    setRegistered();
    hideRegistrationOverlay();
  }
}

function initRegistration() {
  if (isRegistered()) {
    hideRegistrationOverlay();
  } else {
    showRegistrationOverlay();

    // Attach submit handler
    const submitBtn = document.getElementById("regSubmitBtn");
    if (submitBtn) {
      submitBtn.addEventListener("click", submitRegistration);
    }

    // Allow Enter key to submit
    const inputs = document.querySelectorAll("#registrationOverlay input, #registrationOverlay select");
    inputs.forEach(input => {
      input.addEventListener("keypress", (e) => {
        if (e.key === "Enter") {
          submitRegistration();
        }
      });
    });
  }
}

// Initialize registration when DOM is ready
document.addEventListener("DOMContentLoaded", () => {
  initRegistration();
});

// ========= PNG Metadata Utilities =========
// Lightweight PNG chunk reader/writer for embedding metadata

// CRC32 calculation for PNG chunks
function crc32(data) {
  let crc = -1;
  for (let i = 0; i < data.length; i++) {
    crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ data[i]) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

// CRC32 lookup table
const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c;
  }
  return table;
})();

// Extract iTXt chunks from PNG
async function extractPngMetadata(pngBlob) {
  console.log("🔍 extractPngMetadata: Starting, blob size:", pngBlob.size);
  const arrayBuffer = await pngBlob.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const decoder = new TextDecoder();

  console.log("🔍 extractPngMetadata: ArrayBuffer size:", data.length);

  // Check PNG signature
  const pngSignature = [137, 80, 78, 71, 13, 10, 26, 10];
  for (let i = 0; i < 8; i++) {
    if (data[i] !== pngSignature[i]) {
      console.error("🔍 extractPngMetadata: Invalid PNG signature at byte", i, "expected", pngSignature[i], "got", data[i]);
      throw new Error('Not a valid PNG file');
    }
  }

  console.log("🔍 extractPngMetadata: PNG signature verified");

  let offset = 8; // Skip PNG signature
  const view = new DataView(data.buffer);
  let chunkCount = 0;
  let foundMetadata = false;

  while (offset < data.length) {
    // Read chunk length (4 bytes, big-endian)
    const length = view.getUint32(offset);
    offset += 4;

    // Read chunk type (4 bytes)
    const typeBytes = data.slice(offset, offset + 4);
    const type = decoder.decode(typeBytes);
    offset += 4;

    chunkCount++;
    console.log(`🔍 extractPngMetadata: Chunk ${chunkCount}: type="${type}", length=${length}`);

    // Read chunk data
    const chunkData = data.slice(offset, offset + length);
    offset += length;

    // Skip CRC (4 bytes)
    offset += 4;

    // Process iTXt chunks
    if (type === 'iTXt') {
      console.log("🔍 extractPngMetadata: Found iTXt chunk, size:", length);
      const text = decoder.decode(chunkData);
      // iTXt format: keyword\0compressionFlag\0compressionMethod\0languageTag\0translatedKeyword\0text
      const parts = text.split('\0');

      const keyword = parts[0];
      console.log("🔍 extractPngMetadata: iTXt keyword:", keyword);
      console.log("🔍 extractPngMetadata: iTXt parts count:", parts.length);

      if (keyword === 'OPEN_FIGURE_META') {
        console.log("🔍 extractPngMetadata: Found OPEN_FIGURE_META!");
        const jsonText = parts[5]; // keyword, flag, method, lang, trans, text
        console.log("🔍 extractPngMetadata: JSON text length:", jsonText?.length || 0);
        console.log("🔍 extractPngMetadata: JSON preview:", jsonText?.substring(0, 200));

        try {
          const parsed = JSON.parse(jsonText);
          console.log("🔍 extractPngMetadata: JSON parsed successfully");
          foundMetadata = true;
          return parsed;
        } catch (e) {
          console.error('🔍 extractPngMetadata: JSON parse error:', e);
          console.error('🔍 extractPngMetadata: Failed JSON text:', jsonText?.substring(0, 500));
          throw new Error('Invalid metadata JSON: ' + e.message);
        }
      }
    }

    // Stop at IEND chunk
    if (type === 'IEND') {
      console.log("🔍 extractPngMetadata: Reached IEND chunk");
      break;
    }
  }

  console.log(`🔍 extractPngMetadata: Processed ${chunkCount} chunks, found metadata: ${foundMetadata}`);
  return null;
}

// Add iTXt chunk to PNG (UTF-8 support, cross-platform compatible)
async function embedPngMetadata(pngBlob, metadataObj) {
  const arrayBuffer = await pngBlob.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const encoder = new TextEncoder();

  // Find IHDR chunk (first chunk after signature)
  const view = new DataView(data.buffer);
  let offset = 8; // Skip PNG signature
  const ihdrLen = view.getUint32(offset);
  const ihdrTotal = 4 + 4 + ihdrLen + 4; // length + type + data + CRC

  // We'll insert iTXt chunk right after IHDR
  const signature = data.slice(0, 8);
  const ihdrChunk = data.slice(8, 8 + ihdrTotal);
  const rest = data.slice(8 + ihdrTotal);

  // Create iTXt chunk with all metadata as single JSON
  const keyword = "OPEN_FIGURE_META";
  const json = JSON.stringify(metadataObj);

  // iTXt format: keyword\0compressionFlag\0compressionMethod\0languageTag\0translatedKeyword\0text
  const textField =
    keyword + "\0" +      // keyword
    "\0" +                // compressionFlag = 0 (uncompressed)
    "\0" +                // compressionMethod = 0
    "\0" +                // languageTag (empty)
    "\0" +                // translatedKeyword (empty)
    json;                 // actual JSON text

  const textBytes = encoder.encode(textField);
  const chunkType = encoder.encode("iTXt");
  const length = textBytes.length;

  // Build chunk: length(4) + type(4) + data(length) + CRC(4)
  const chunk = new Uint8Array(4 + 4 + length + 4);
  const chunkView = new DataView(chunk.buffer);

  // Length (big-endian)
  chunkView.setUint32(0, length);

  // Type
  chunk.set(chunkType, 4);

  // Data
  chunk.set(textBytes, 8);

  // CRC over type + data
  const crcData = chunk.slice(4, 8 + length);
  const crc = crc32(crcData);
  chunkView.setUint32(8 + length, crc);

  // Assemble new PNG: signature + IHDR + iTXt + rest
  const newData = new Uint8Array(signature.length + ihdrChunk.length + chunk.length + rest.length);
  let pos = 0;
  newData.set(signature, pos); pos += signature.length;
  newData.set(ihdrChunk, pos); pos += ihdrChunk.length;
  newData.set(chunk, pos); pos += chunk.length;
  newData.set(rest, pos);

  return new Blob([newData], { type: 'image/png' });
}

// ========= Custom Filename Dialog =========
// Office Add-ins don't support prompt(), so we use custom HTML dialog
function showFilenameDialog(defaultFilename) {
  return new Promise((resolve) => {
    const dialog = document.getElementById("filenameDialog");
    const input = document.getElementById("filenameInput");
    const saveBtn = document.getElementById("filenameSaveBtn");
    const cancelBtn = document.getElementById("filenameCancelBtn");

    // Set default value
    input.value = defaultFilename;

    // Show dialog
    dialog.classList.add("show");
    input.focus();
    input.select();

    // Handle save
    const handleSave = () => {
      let filename = input.value.trim() || defaultFilename;
      if (!filename.toLowerCase().endsWith('.png')) {
        filename += '.png';
      }
      cleanup();
      resolve(filename);
    };

    // Handle cancel
    const handleCancel = () => {
      cleanup();
      resolve(null);
    };

    // Handle Enter key
    const handleKeyPress = (e) => {
      if (e.key === 'Enter') {
        handleSave();
      } else if (e.key === 'Escape') {
        handleCancel();
      }
    };

    // Cleanup function
    const cleanup = () => {
      dialog.classList.remove("show");
      saveBtn.removeEventListener("click", handleSave);
      cancelBtn.removeEventListener("click", handleCancel);
      input.removeEventListener("keypress", handleKeyPress);
    };

    // Add event listeners
    saveBtn.addEventListener("click", handleSave);
    cancelBtn.addEventListener("click", handleCancel);
    input.addEventListener("keypress", handleKeyPress);
  });
}

// ========= Debug Helper =========
function debugLog(message) {
  console.log(message);
  // UI display suppressed - debug messages only in console
  // const status = document.getElementById("status");
  // if (status) {
  //   const currentText = status.textContent || "";
  //   status.textContent = currentText + "\n[DEBUG] " + message;
  // }
}

// ========= Save Figure with Metadata =========
// Uses ONLY standard browser APIs: Blob, URL.createObjectURL, <a download>
// NO File System Access API or prompt() - works on both Mac and Windows
async function saveFigureWithMetadata() {
  try {
    setStatus("Preparing figure with metadata...");
    debugLog("Step 1: Starting save process");

    // 1. Get current PNG from canvas/preview
    const canvas = document.getElementById("plot");
    if (!canvas) {
      setStatus("❌ No figure to save. Please preview first.");
      return;
    }
    debugLog("Step 2: Canvas found");

    // Convert canvas to blob
    const dataUrl = canvas.toDataURL("image/png");
    debugLog("Step 3: Canvas converted to dataURL (length: " + dataUrl.length + ")");

    const response = await fetch(dataUrl);
    const pngBlob = await response.blob();
    debugLog("Step 4: PNG blob created (" + pngBlob.size + " bytes)");

    // 2. Collect metadata as single JSON object
    const metadata = {
      version: 1,
      addinVersion: ADDIN_VERSION,
      generator: "Figra",
      created_utc: new Date().toISOString(),

      data: {
        headers: window.lastProcessedData ? window.lastProcessedData[0] : [],
        rows: window.lastProcessedData ? window.lastProcessedData.slice(1) : []
      },

      chart: {
        chartType: document.getElementById("chartType")?.value || "",
        xColumn: document.getElementById("xColumn")?.value || "",
        yColumn: document.getElementById("yColumn")?.value || "",
        groupColumn: document.getElementById("groupColumn")?.value || "",
        errorColumn: document.getElementById("errorColumn")?.value || "",
        bins: document.getElementById("bins")?.value || "20"
      },

      settings: collectCurrentSettings()
    };
    debugLog("Step 5: Metadata collected");

    // 2b. Retrieve statistical results if statistics are enabled AND chart type supports it
    const STATS_SUPPORTED_CHART_TYPES = [
      "box", "box_dot", "violin", "violin_dot", "bar_error_dot",
      "bar_grouped_error_dot", "box_grouped", "box_grouped_dot",
      "violin_grouped", "violin_grouped_dot",
      "line_grouped_error_raw"
    ];

    const chartType = document.getElementById("chartType")?.value || "";
    const addStatistics = document.getElementById("addStatistics")?.checked;
    const isChartTypeSupported = STATS_SUPPORTED_CHART_TYPES.includes(chartType);

    if (addStatistics && isChartTypeSupported) {
      debugLog("Step 5b: Retrieving statistical results for " + chartType + "...");
      try {
        let statResults = null;

        // Get statistical results based on chart type
        if (chartType === "line_grouped_error_raw") {
          statResults = await getLineStatisticalResultsText();
        } else if (chartType === "bar_grouped_error_dot" || chartType === "box_grouped" || chartType === "box_grouped_dot" || chartType === "violin_grouped" || chartType === "violin_grouped_dot") {
          statResults = await getGroupedBarStatisticalResultsText();
        } else {
          // For other supported chart types, try the general statistical results
          statResults = await getStatisticalResultsText();
        }

        // Only add to metadata if we got valid results
        if (statResults &&
            !statResults.includes("No statistical") &&
            !statResults.includes("ERROR") &&
            !statResults.includes("Statistics disabled") &&
            statResults.trim() !== "") {
          metadata.statisticalResults = statResults;
          debugLog("Step 5b: ✅ Statistical results added (" + statResults.length + " chars)");
        } else {
          metadata.statisticalResults = null;
          debugLog("Step 5b: ⚠️ No valid statistical results: " + (statResults ? statResults.substring(0, 50) : "null"));
        }
      } catch (e) {
        console.error("Error retrieving statistical results for metadata:", e);
        metadata.statisticalResults = null;
        debugLog("Step 5b: ❌ Error: " + e.message);
      }
    } else {
      metadata.statisticalResults = null;
      if (!addStatistics) {
        debugLog("Step 5b: Statistics disabled");
      } else if (!isChartTypeSupported) {
        debugLog("Step 5b: Chart type '" + chartType + "' does not support statistics");
      }
    }

    // 3. Embed metadata into PNG iTXt chunk
    const pngWithMetadata = await embedPngMetadata(pngBlob, metadata);
    debugLog("Step 6: Metadata embedded (" + pngWithMetadata.size + " bytes)");

    // 4. Show custom filename dialog
    const title = document.getElementById("titleText")?.value?.trim() || "plot";
    const defaultFilename = title.replace(/[^\w\-]+/g, "_") + ".png";
    debugLog("Step 7: Showing filename dialog (default: " + defaultFilename + ")");

    // Show dialog and wait for user input
    const filename = await showFilenameDialog(defaultFilename);
    if (!filename) {
      setStatus("⚠️ Save cancelled");
      debugLog("Step 8: User cancelled");
      return;
    }
    debugLog("Step 8: Filename confirmed: " + filename);

    // 5. Show image in dialog for right-click save (works on Mac!)
    debugLog("Step 9: Displaying image for right-click save...");

    const blobUrl = URL.createObjectURL(pngWithMetadata);
    debugLog("Step 10: Blob URL created");

    // Set image source
    const imgElement = document.getElementById("saveImagePreview");
    const filenameElement = document.getElementById("suggestedFilename");
    const dialog = document.getElementById("saveImageDialog");
    const closeBtn = document.getElementById("closeSaveImageDialog");

    imgElement.src = blobUrl;
    filenameElement.textContent = filename;

    // Make image draggable to desktop/Finder
    imgElement.draggable = true;
    imgElement.ondragstart = (e) => {
      debugLog("Drag started - trying to save as: " + filename);
      try {
        e.dataTransfer.effectAllowed = "copy";
        e.dataTransfer.setData("DownloadURL", `image/png:${filename}:${blobUrl}`);
        debugLog("DataTransfer set successfully");
      } catch (err) {
        debugLog("DataTransfer error: " + err.message);
      }
    };

    // Show dialog
    dialog.style.display = "block";
    debugLog("Step 11: ✅ Image displayed!");

    // Handle close
    const handleClose = () => {
      dialog.style.display = "none";
      URL.revokeObjectURL(blobUrl);
      closeBtn.removeEventListener("click", handleClose);
      imgElement.ondragstart = null;
      debugLog("Step 12: Dialog closed, blob URL revoked");
    };

    closeBtn.addEventListener("click", handleClose);

    setStatus(`✅ Image ready to save!\n\nTRY THESE OPTIONS:\n\n1. DRAG the image to your Desktop or Finder\n\n2. Right-click → "Save Image As..."\n\nSave as: ${filename}`);

  } catch (e) {
    console.error('Save figure error:', e);
    debugLog("ERROR: " + e.message);
    setStatus("❌ Error saving figure: " + (e?.message || e));
  }
}

// ========= Download Metadata as JSON =========
// Backup solution: Save metadata separately (always works on Mac)
async function downloadMetadataJson() {
  try {
    setStatus("Preparing metadata...");

    // Collect metadata
    const metadata = {
      version: 1,
      generator: "Figra",
      created_utc: new Date().toISOString(),

      data: {
        headers: window.lastProcessedData ? window.lastProcessedData[0] : [],
        rows: window.lastProcessedData ? window.lastProcessedData.slice(1) : []
      },

      chart: {
        chartType: document.getElementById("chartType")?.value || "",
        xColumn: document.getElementById("xColumn")?.value || "",
        yColumn: document.getElementById("yColumn")?.value || "",
        groupColumn: document.getElementById("groupColumn")?.value || "",
        errorColumn: document.getElementById("errorColumn")?.value || "",
        bins: document.getElementById("bins")?.value || "20"
      },

      settings: collectCurrentSettings()
    };

    // Retrieve statistical results if applicable
    const STATS_SUPPORTED_CHART_TYPES = [
      "box", "box_dot", "violin", "violin_dot", "bar_error_dot",
      "bar_grouped_error_dot", "box_grouped", "box_grouped_dot",
      "violin_grouped", "violin_grouped_dot",
      "line_grouped_error_raw"
    ];

    const chartType = document.getElementById("chartType")?.value || "";
    const addStatistics = document.getElementById("addStatistics")?.checked;
    const isChartTypeSupported = STATS_SUPPORTED_CHART_TYPES.includes(chartType);

    if (addStatistics && isChartTypeSupported) {
      try {
        let statResults = null;

        if (chartType === "line_grouped_error_raw") {
          statResults = await getLineStatisticalResultsText();
        } else if (chartType === "bar_grouped_error_dot" || chartType === "box_grouped" || chartType === "box_grouped_dot" || chartType === "violin_grouped" || chartType === "violin_grouped_dot") {
          statResults = await getGroupedBarStatisticalResultsText();
        } else {
          statResults = await getStatisticalResultsText();
        }

        if (statResults &&
            !statResults.includes("No statistical") &&
            !statResults.includes("ERROR") &&
            !statResults.includes("Statistics disabled") &&
            statResults.trim() !== "") {
          metadata.statisticalResults = statResults;
        } else {
          metadata.statisticalResults = null;
        }
      } catch (e) {
        console.error("Error retrieving statistical results for metadata:", e);
        metadata.statisticalResults = null;
      }
    } else {
      metadata.statisticalResults = null;
    }

    // Convert to JSON
    const jsonString = JSON.stringify(metadata, null, 2);
    const blob = new Blob([jsonString], { type: 'application/json' });

    // Create filename
    const title = document.getElementById("titleText")?.value?.trim() || "plot";
    const filename = title.replace(/[^\w\-]+/g, "_") + "_metadata.json";

    // Download using <a download> (works on Mac)
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();

    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);

    setStatus(`✅ Metadata saved: ${filename}\nSave the Excel figure separately, then use both files together!`);

  } catch (e) {
    console.error('Download metadata error:', e);
    setStatus("❌ Error downloading metadata: " + (e?.message || e));
  }
}

// ========= Load from Figure =========
// Uses standard <input type="file"> - works on both Mac and Windows
// Accepts both PNG (with embedded metadata) and JSON (separate metadata)
async function loadFromFigure() {
  try {
    setStatus("Select a PNG or JSON file...");

    // Create file input
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,.json";

    input.onchange = async (e) => {
      try {
        const file = e.target.files[0];
        if (!file) return;

        setStatus(`Loading metadata from ${file.name}...`);

        let metadata = null;

        // Check file type
        if (file.name.endsWith('.json')) {
          // Load from JSON file
          const text = await file.text();
          metadata = JSON.parse(text);
          setStatus(`Metadata loaded from JSON file`);
        } else {
          // Extract metadata from PNG iTXt chunk
          console.log("🔍 DEBUG: Loading PNG file, size:", file.size, "bytes");
          console.log("🔍 DEBUG: Platform:", navigator.platform, "UserAgent:", navigator.userAgent);

          try {
            metadata = await extractPngMetadata(file);
            console.log("🔍 DEBUG: Metadata extracted successfully");
            console.log("🔍 DEBUG: Metadata keys:", metadata ? Object.keys(metadata) : "null");
            console.log("🔍 DEBUG: Data headers:", metadata?.data?.headers);
            console.log("🔍 DEBUG: Data rows count:", metadata?.data?.rows?.length);
            console.log("🔍 DEBUG: Has statisticalResults:", !!metadata?.statisticalResults);
            console.log("🔍 DEBUG: StatisticalResults length:", metadata?.statisticalResults?.length || 0);
          } catch (e) {
            console.error("🔍 DEBUG: Error extracting metadata:", e);
            setStatus("❌ Error extracting metadata from PNG: " + e.message);
            return;
          }

          // Check if this is an Figra PNG
          if (!metadata || !metadata.data) {
            console.error("🔍 DEBUG: Invalid metadata - missing data field");
            setStatus("❌ This PNG was not created by Figra (no metadata found) or saved by un-supported environment like Mac.");
            return;
          }
        }

        // Check version compatibility
        const savedVersion = metadata.addinVersion || "Unknown";
        let versionMessage = "";

        if (!metadata.addinVersion) {
          versionMessage = `⚠️ Version info: This figure was created with an older version of Figra (before version tracking).\nCurrent version: ${ADDIN_VERSION}\n\nThe figure may not reproduce exactly. Proceed with caution.\n\n`;
          console.warn("⚠️ No version info in metadata - created with older add-in");
        } else if (savedVersion === ADDIN_VERSION) {
          versionMessage = `✅ Version match: Figure created with same version (${ADDIN_VERSION})\n\n`;
          console.log(`✅ Version match: ${savedVersion}`);
        } else {
          versionMessage = `⚠️ Version mismatch!\nFigure version: ${savedVersion}\nCurrent version: ${ADDIN_VERSION}\n\nSettings and appearance may differ. Consider recreating the figure with the current version.\n\n`;
          console.warn(`⚠️ Version mismatch: Figure=${savedVersion}, Current=${ADDIN_VERSION}`);
        }

        // Restore data
        window.lastProcessedData = [metadata.data.headers, ...metadata.data.rows];

        // Restore chart configuration
        if (metadata.chart.chartType) {
          document.getElementById("chartType").value = metadata.chart.chartType;
        }

        // Populate column dropdowns
        populateColumnSelectors(metadata.data.headers);

        if (metadata.chart.xColumn) document.getElementById("xColumn").value = metadata.chart.xColumn;
        if (metadata.chart.yColumn) document.getElementById("yColumn").value = metadata.chart.yColumn;
        if (metadata.chart.groupColumn) document.getElementById("groupColumn").value = metadata.chart.groupColumn;
        if (metadata.chart.errorColumn) document.getElementById("errorColumn").value = metadata.chart.errorColumn;
        if (metadata.chart.bins) document.getElementById("bins").value = metadata.chart.bins;

        // Restore all settings
        applySettings(metadata.settings);

        // Display statistical results if available
        if (metadata.statisticalResults) {
          console.log("📊 Statistical results found in metadata:");
          console.log(metadata.statisticalResults);
          // Note: Statistical results are stored in metadata but need to be regenerated
          // by clicking Preview to populate R global variables
        }

        // Insert data and statistical results into a single Excel sheet
        try {
          await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;

            // Create unique sheet name based on file name
            const baseSheetName = file.name.replace(/\.(png|json)$/i, '').substring(0, 20);
            let sheetName = baseSheetName;

            // Check if sheet exists and create unique name
            sheets.load("items/name");
            await context.sync();

            let counter = 1;
            while (sheets.items.some(s => s.name === sheetName)) {
              sheetName = `${baseSheetName}_${counter}`;
              counter++;
            }

            // Create new sheet
            const newSheet = sheets.add(sheetName);
            newSheet.activate();

            // Prepare data (headers + rows)
            const dataToInsert = [metadata.data.headers, ...metadata.data.rows];
            const numRows = dataToInsert.length;
            const numCols = metadata.data.headers.length;

            // Insert data into sheet
            const dataRange = newSheet.getRangeByIndexes(0, 0, numRows, numCols);
            dataRange.values = dataToInsert;

            // Format header row
            const headerRange = newSheet.getRangeByIndexes(0, 0, 1, numCols);
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = "#4472C4";
            headerRange.format.font.color = "white";

            // Auto-fit columns
            dataRange.format.autofitColumns();

            await context.sync();

            console.log(`Data inserted into new sheet: ${sheetName}`);

            // Add statistical results to the right side of the data if available
            const hasStats = metadata.statisticalResults &&
                            !metadata.statisticalResults.includes("No statistical") &&
                            !metadata.statisticalResults.includes("ERROR") &&
                            !metadata.statisticalResults.includes("Statistics disabled");

            if (hasStats) {
              // Place stats starting from the column after data (with 1 empty column gap)
              const statsStartCol = numCols + 1;

              // Add header for statistical section
              const statsHeaderRange = newSheet.getRangeByIndexes(0, statsStartCol, 1, 1);
              statsHeaderRange.values = [["STATISTICAL ANALYSIS RESULTS"]];
              statsHeaderRange.format.font.bold = true;
              statsHeaderRange.format.fill.color = "#4472C4";
              statsHeaderRange.format.font.color = "white";
              statsHeaderRange.format.font.size = 12;

              // Parse and format the results
              const formattedResults = metadata.statisticalResults.replace(/\\n/g, '\n');
              const allLines = formattedResults.split('\n');
              const lines = allLines.filter(line => line.trim().length > 0);

              // Insert results line by line (starting from row 1, below header)
              for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                const cellRange = newSheet.getRangeByIndexes(1 + i, statsStartCol, 1, 1);
                cellRange.values = [[line]];

                // Format headers FIRST (with priority over significance colors)
                if (line.startsWith('===') || line.startsWith('Category') ||
                    line.startsWith('Test:') || line.startsWith('Normality') ||
                    line.startsWith('Post-hoc')) {
                  cellRange.format.font.bold = true;
                  cellRange.format.fill.color = "#E7E6E6";
                }
                // Only apply significance colors to NON-header lines
                else if (line.includes('***') || line.includes('**') || line.includes('*')) {
                  cellRange.format.fill.color = "#C6EFCE";
                }
                // Highlight non-significant
                else if (line.includes('(ns)') || line.includes('ns ')) {
                  cellRange.format.fill.color = "#FFC7CE";
                }
              }

              // Auto-fit the stats column
              const statsRange = newSheet.getRangeByIndexes(0, statsStartCol, lines.length + 1, 1);
              statsRange.format.autofitColumns();

              console.log(`Statistical results added to sheet: ${sheetName}`);
            }

            await context.sync();

            // Select the data range (NOT including stats) so it's ready for loading
            dataRange.select();
            await context.sync();
          });
        } catch (err) {
          console.error('Error inserting data to Excel:', err);
          // Don't fail the whole load operation if sheet creation fails
        }

        // Auto-load the data into UI so it's ready for preview
        try {
          console.log("Auto-loading data into UI...");
          await loadHeadersFromSelection();
          console.log("✅ Data auto-loaded into UI");
        } catch (err) {
          console.error('Error auto-loading data:', err);
          // Don't fail the whole operation if auto-load fails
        }

        // Show success message with version info
        const createdDate = new Date(metadata.created_utc).toLocaleString();
        const hasStats = metadata.statisticalResults &&
                        !metadata.statisticalResults.includes("No statistical") &&
                        !metadata.statisticalResults.includes("ERROR");
        const statMsg = hasStats ? " (with statistical results)" : "";
        const rowCount = metadata.data.rows.length;
        const colCount = metadata.data.headers.length;
        setStatus(`${versionMessage}✅ Loaded from ${file.name} (created: ${createdDate})\n📋 Data loaded and ready: ${rowCount} rows × ${colCount} columns${statMsg}\n🎯 Click Preview to verify reproducibility`);

        console.log("Metadata loaded:", metadata);

      } catch (err) {
        console.error('Load figure error:', err);
        setStatus("❌ Error loading figure: " + (err?.message || err));
      }
    };

    input.click();

  } catch (e) {
    console.error('Load figure error:', e);
    setStatus("❌ Error loading figure: " + (e?.message || e));
  }
}

// Helper: Populate column selectors with headers
function populateColumnSelectors(headers) {
  const selectors = ["xColumn", "yColumn", "groupColumn", "errorColumn"];

  for (const selectorId of selectors) {
    const selector = document.getElementById(selectorId);
    if (!selector) continue;

    selector.innerHTML = "";
    headers.forEach(header => {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header;
      selector.appendChild(option);
    });
  }
}

// Helper: Apply settings to UI
function applySettings(settings) {
  const el = (id) => document.getElementById(id);

  // Helper to set value if element exists
  const setValue = (id, value) => {
    const elem = el(id);
    if (elem && value !== undefined) {
      if (elem.type === 'checkbox') {
        elem.checked = (value === 'true' || value === true);
      } else {
        elem.value = value;
      }
    }
  };

  // Apply all settings
  for (const [key, value] of Object.entries(settings)) {
    // Skip special cases that need custom handling
    if (key === 'customOrderGroup' || key === 'customOrderCategory' ||
        key === 'customComparisons' || key === 'customPositions') {
      continue;
    }
    setValue(key, value);
  }

  // Handle custom data ordering
  if (settings.customOrderGroup && typeof settings.customOrderGroup === 'string') {
    const groups = settings.customOrderGroup.split(',').map(g => g.trim()).filter(g => g);
    if (groups.length > 0 && typeof window !== 'undefined') {
      window.detectedGroups = groups;
      console.log('📋 Restored custom group order:', groups);
    }
  }

  if (settings.customOrderCategory && typeof settings.customOrderCategory === 'string') {
    const categories = settings.customOrderCategory.split(',').map(c => c.trim()).filter(c => c);
    if (categories.length > 0 && typeof window !== 'undefined') {
      window.detectedCategories = categories;
      console.log('📋 Restored custom category order:', categories);
    }
  }

  // Handle custom comparisons (stored as JSON string)
  if (settings.customComparisons && typeof settings.customComparisons === 'string') {
    try {
      const comparisons = JSON.parse(settings.customComparisons);
      console.log('📋 Restored custom comparisons:', comparisons);
      // Note: Checkboxes will be populated when data is loaded
    } catch (e) {
      console.error('Error parsing customComparisons:', e);
    }
  }

  // Handle custom positions (stored as JSON string)
  if (settings.customPositions && typeof settings.customPositions === 'string') {
    try {
      const positions = JSON.parse(settings.customPositions);
      console.log('📋 Restored custom positions:', positions);
      // Note: Position sliders will be populated when data is loaded
    } catch (e) {
      console.error('Error parsing customPositions:', e);
    }
  }

  // Trigger visibility updates
  if (typeof handleChartTypeChange === 'function') {
    handleChartTypeChange();
  }
  if (typeof updateStatisticsAvailability === 'function') {
    updateStatisticsAvailability();
  }
  if (typeof updateVbracketVisibility === 'function') {
    updateVbracketVisibility();
  }
  if (typeof updateStatsDetailedControlsVisibility === 'function') {
    updateStatsDetailedControlsVisibility();
  }
  if (typeof updateGroupColorVisibility === 'function') {
    updateGroupColorVisibility();
  }
}

// 先頭あたりに追加：共通セーブ
async function saveFileWithPicker({ data, suggestedName = "plot.png", mime = "image/png", pickerTypes = [{ description: "PNG Image", accept: { "image/png": [".png"] } }] }) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: mime });

  // File System Access API が使える場合（一番わかりやすい）
  if ("showSaveFilePicker" in window) {
    const handle = await window.showSaveFilePicker({ suggestedName, types: pickerTypes });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();

    // Full path unavailable. Show file name and "using picker"
    setStatus(`Save completed: ${handle.name} (using picker; full path unavailable by specification)`);
    return { mode: "picker", fileName: handle.name };
  }

  // フォールバック（ダウンロード扱い）
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = suggestedName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  // To default download folder (typically ~/Downloads on Mac)
  setStatus(`Save completed: ${suggestedName} to browser download folder (typically ~/Downloads)`);
  return { mode: "download", fileName: suggestedName };
}

async function tplSave(){
  const name = (el("tplName")?.value || "").trim();
  if(!name){ setStatus("Please enter template name"); return; }
  const data = uiOpts();
  try{
    // ストレージにも保存（従来の便利機能）
    await OfficeRuntime.storage.setItem(`sato:tpl:${name}`, JSON.stringify(data));

    // さらに "保存先を選べる" 書き出し（任意）
    await saveFileWithPicker({
      data: JSON.stringify({kind:"sato.tpl.v1", name, opts:data}, null, 2),
      suggestedName: `${name}.json`,
      mime: 'application/json'
    });

    // ドロップダウンへ反映
    const tsel = el("template");
    if([...tsel.options].every(o=>o.value!==name)){ tsel.add(new Option(name, name)); }
    setStatus(`Template saved: "${name}"`);
  }catch(e){
    console.error(e); setStatus("Template save error: " + (e?.message || e));
  }
}

// === 追加 : キャンバスPNGを書き出す（保存ダイアログ; Macはダウンロードにフォールバック） ===
async function exportCanvasPng() {
  try {
    const cv = document.getElementById("plot");
    if (!cv) throw new Error("Canvas not found (id=plot)");

    const dataUrl = cv.toDataURL("image/png");
    const title = (document.getElementById("titleText")?.value || "").trim();
    const yName = (document.getElementById("yColumn")?.value || "plot").trim();
    const base  = (title || `Histogram_of_${yName}`).replace(/[^\w\-]+/g, "_");

    const res = await fetch(dataUrl);
    const blob = await res.blob();

    const info = await saveToDirectoryEachTime({
      data: blob,
      fileName: base + ".png",
      mime: "image/png"
    });

    console.log("Save info:", info);
    setStatus(`Save mode: ${info.mode}, File name: ${info.fileName}`);
  } catch (e) {
    console.error(e);
    setStatus("PNG export error: " + (e?.message || e));
  }
}

// === 毎回フォルダ選択で保存 ===
// 保存：毎回フォルダ選択 → 失敗時はファイルピッカー → さらに失敗時はダウンロード＋クリップボード案内
async function saveToDirectoryEachTime({ data, fileName, mime = "image/png" }) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: mime });
  const isMac = () =>
    (navigator.platform || "").toLowerCase().includes("mac") ||
    (navigator.userAgent || "").toLowerCase().includes("mac");
  const pasteHint = () => (isMac() ? "⌘V" : "Ctrl+V");
  const status = (msg) => (typeof setStatus === "function" ? setStatus(msg) : console.log(msg));

  // A. ディレクトリピッカー（毎回フォルダを選べる）
  if ("showDirectoryPicker" in window) {
    try {
      const dirHandle = await window.showDirectoryPicker();          // ← 毎回選ぶ
      const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      status(`Save completed: ${fileName} (directory picker used)`);
      return { mode: "directoryPicker", fileName };
    } catch (e) {
      // User cancellation, etc. falls through here → fallback to next method
      console.warn("showDirectoryPicker failed; fallback to file picker:", e);
    }
  }

  // B. File picker (specify location and name each time)
  if ("showSaveFilePicker" in window) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: fileName,
        types: [{ description: "PNG Image", accept: { "image/png": [".png"] } }]
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      status(`Save completed: ${handle.name} (file picker used)`);
      return { mode: "filePicker", fileName: handle.name };
    } catch (e) {
      console.warn("showSaveFilePicker failed; fallback to download:", e);
    }
  }

  // C. Final fallback: Download (may be ignored on Mac WKWebView, etc.)
  try {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (e) {
    console.warn("download link fallback failed:", e);
  }

  // Copy to clipboard if possible, and show message
  let clipboard = false;
  try {
    if (navigator.clipboard && window.ClipboardItem) {
      await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]);
      clipboard = true;
    }
  } catch (e) {
    console.warn("clipboard write failed:", e);
  }

  const msg =
    `Unfortunately, due to macOS Office add-in limitations, this file may not be saved automatically. ` +
    `I've copied the image to your clipboard — please press ${pasteHint()} to paste it.`;

  status(`Download fallback: ${fileName}\n${msg}`);
  return { mode: "download", fileName, clipboard };
}

// 追加：.sato をファイルに保存
async function saveSato() {
  try {
    await Excel.run(async (ctx) => {
      const ws = await ensureSatoSheet(ctx); // ← 下に定義
      const cv = document.getElementById("plot");
      const dataUrl = cv.toDataURL("image/png");

      // --- ペイロード（JSON） ---
      const { values } = await getSelectedValues();
      const spec = {
        chart: document.getElementById("chartType").value,
        x: document.getElementById("xColumn").value || "",
        y: document.getElementById("yColumn").value || "",
        opts: uiOpts(),
        app: "sato.webr",
        version: "0.3.0",
      };
      const payload = {
        kind: "sato.v1",
        spec,
        dataCsv: toCsv(values),
        previewDataUrl: dataUrl,
        savedAt: new Date().toISOString(),
        id: Date.now().toString(), // → 一意キー
      };
      const json = JSON.stringify(payload);

      // --- 追記先の行を決める（上書き回避） ---
      const used = ws.getUsedRangeOrNullObject();
      used.load(["rowCount", "columnCount", "address"]);
      await ctx.sync();
      const nextRow = used.isNullObject ? 1 : (used.rowCount + 1);

      // JSON はA列へ、見出し行が欲しければ最初の1行目に入れておく
      const jsonCell = ws.getRangeByIndexes(nextRow - 1, 0, 1, 1);
      jsonCell.values = [[json]];
      jsonCell.format.columnWidth = 420;     // 読みやすく
      jsonCell.format.wrapText = true;
      jsonCell.format.horizontalAlignment = "Left";
      jsonCell.format.verticalAlignment = "Top";

      // --- 画像を比率固定で貼り付け（C列に配置） ---
      const shape = ws.shapes.addImage(dataUrl.split(",")[1]); // base64 部分
      const W0 = cv.width, H0 = cv.height;              // キャンバスの実ピクセル
      const targetW = 480;                               // 任意：見やすい幅
      const targetH = Math.round(targetW * H0 / W0);     // 比率で計算
      shape.width  = targetW;
      shape.height = targetH;

      // 配置：C列の nextRow 行の左上へ
      const place = ws.getRangeByIndexes(nextRow - 1, 2, 1, 1); // C列
      place.load(["left", "top", "rowHeight"]);
      await ctx.sync();

      // 行の高さが足りなければ伸ばす
      const needH = Math.max(place.rowHeight, targetH + 8);
      ws.getRangeByIndexes(nextRow - 1, 0, 1, 1).rowHeight = needH;

      // 位置
      shape.left = place.left;
      shape.top  = place.top;

      setStatus(`.sato appended to workbook (row: ${nextRow} / sheet: SATO_DATA)`);
    });
  } catch (e) {
    console.error(e);
    setStatus(".sato save error: " + (e?.message || e));
  }
}

async function ensureSatoSheet(ctx) {
  const name = "SATO_DATA";
  const sheets = ctx.workbook.worksheets;
  const ws = sheets.getItemOrNullObject(name);
  await ctx.sync();
  if (ws.isNullObject) {
    const created = sheets.add(name);
    created.visibility = Excel.SheetVisibility.visible; // Keep visible
    await ctx.sync();
    return created;
  }
  return ws;
}

// 選択セルから .sato JSON を読み込んでUI に反映 → 再描画
async function loadSatoFromSelection() {
  try {
    await Excel.run(async (ctx) => {
      const rng = ctx.workbook.getSelectedRange();
      rng.load(["text", "values", "rowCount", "columnCount"]);
      await ctx.sync();

      // 1セル目の非空テキストを拾う
      let raw = "";
      const T = rng.text || [];
      outer: for (let r = 0; r < T.length; r++) {
        for (let c = 0; c < (T[r] || []).length; c++) {
          const t = (T[r][c] || "").trim();
          if (t) { raw = t; break outer; }
        }
      }
      if (!raw) { setStatus("Selected cell is empty."); return; }

      // ---- エスケープ状況に応じて復元 ----
      let candidate = raw;

      // Case A: When cell contains JSON string full of escaped quotes → parse once to extract the string
      // Example: "{\"kind\":\"sato.v1\", ... }"
      if (!candidate.trim().startsWith("{") && candidate.includes('\\"')) {
        try { candidate = JSON.parse(candidate); } catch {}
      }
      // Case B: If double-stringified, parse once more
      if (typeof candidate === "string" && candidate.startsWith('"') && candidate.endsWith('"')) {
        try { candidate = JSON.parse(candidate); } catch {}
      }

      if (typeof candidate !== "string" || !candidate.trim().startsWith("{")) {
        setStatus("Selected cell content is not JSON."); return;
      }

      let sato;
      try { sato = JSON.parse(candidate); }
      catch (e) { setStatus("JSON parse failed: " + e.message); return; }

      if (!sato || sato.kind !== "sato.v1") {
        setStatus("This cell is not .sato format."); return;
      }

      // --- Apply to UI ---
      const spec = sato.spec || {};
      const o = spec.opts || {};
      const set = (id, v) => { if (el(id) != null && v != null) el(id).value = v; };

      // Column names
      set("xColumn", spec.x);
      set("yColumn", spec.y);

      // Options
      set("titleText", o.title);
      set("xLabel",    o.xlab);
      set("yLabel",    o.ylab);
      set("ggtheme",   o.theme);
      set("fontFam",   o.family);

      set("fillColor",   o.fill || o.fillPicker);
      set("fillPicker",  o.fill || o.fillPicker);
      set("fillAlpha",   o.fillAlpha);
      set("strokeColor", o.stroke || o.strokePicker);
      set("strokePicker",o.stroke || o.strokePicker);

      set("lineWidth",       o.lineWidth);
      set("titleSize",       o.titleSize);
      // Handle both old format (axisTitleSize/axisTextSize) and new format (4 separate values)
      set("xAxisTitleSize",  o.xAxisTitleSize || o.axisTitleSize);
      set("yAxisTitleSize",  o.yAxisTitleSize || o.axisTitleSize);
      set("xAxisTextSize",   o.xAxisTextSize || o.axisTextSize);
      set("yAxisTextSize",   o.yAxisTextSize || o.axisTextSize);
      set("legendTextSize",  o.legendTextSize);
      set("titleWeight",     o.titleWeight);
      set("axisTitleWeight", o.axisTitleWeight || o.axisWeight);
      set("axisTextWeight",  o.axisTextWeight || o.axisWeight);
      set("dataOrder",       o.dataOrder || "default");
      set("customOrder",     o.customOrder || "");
      set("numGroups",       o.numGroups || 2);
      for (let i = 0; i < 6; i++) {
        set(`groupColor${i+1}`, (o.groupColors && o.groupColors[i]) || ["#4C78A8", "#E15759", "#57C4AD", "#E9C46A", "#F76C6C", "#A8DADC"][i]);
      }
      if (o.xMin != null) set("xMin", o.xMin);
      if (o.xMax != null) set("xMax", o.xMax);
      if (o.yMin != null) set("yMin", o.yMin);
      if (o.yMax != null) set("yMax", o.yMax);
      if (o.bins != null)  set("bins", o.bins);

      // Fast redraw (with current selection data)
      await previewPlotWithDebug();
      setStatus(".sato loaded from selected cell");
    });
  } catch (e) {
    console.error(e);
    setStatus(".sato load error from selected cell: " + (e?.message || e));
  }
}

// UPDATED: R 側に渡す代表フォント名（簡素化版）
function primaryFont(name = "") {
  console.log(`DEBUG: Legacy primaryFont called with: "${name}"`);
  // Use the simplified function
  const result = validateFont(name);
  console.log(`DEBUG: primaryFont returning: "${result}"`);
  return result;
}

// UPDATED: SVG内の font-family を強制（簡素化版）
function forceSvgFont(svgText, familyCss) {
  console.log(`DEBUG: Legacy forceSvgFont called with font: "${familyCss}"`);
  // Extract primary font name from CSS stack and use the simplified function
  const primaryFontName = familyCss.split(',')[0].replace(/['"]/g, '').trim();
  const result = forceSvgFontAggressive(svgText, primaryFontName);
  console.log(`DEBUG: forceSvgFont delegated to forceSvgFontAggressive`);
  return result;
}

// UPDATED: 任意の名前をブラウザで解決しやすいフォント・スタックに展開（簡素化版）
function fontStack(name, { strict = false } = {}) {
  const n = validateFont(name || "Arial");
  if (strict || n === "Arial" || n === "Times New Roman") {
    // 単一ファミリで返す（フォールバック禁止）
    return n.includes(" ") ? `'${n}'` : n;
  }
  // 通常は従来どおりスタックを返す
  return buildCompleteFontStack(n);
}


// 図形が選択されていても落ちないアンカー座標
async function getSafeAnchorLeftTop() {
  let anchor = { left: 0, top: 0 };
  await Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const ws = wb.worksheets.getActiveWorksheet();

    try {
      const rng = wb.getSelectedRange();
      rng.load(["left", "top"]); await ctx.sync();
      anchor = { left: rng.left, top: rng.top };
      return;
    } catch (_) {}

    try {
      const ac = wb.getActiveCell();
      ac.load(["left", "top"]); await ctx.sync();
      anchor = { left: ac.left, top: ac.top };
      return;
    } catch (_) {}

    const a1 = ws.getRange("A1");
    a1.load(["left", "top"]); await ctx.sync();
    anchor = { left: a1.left, top: a1.top };
  });
  return anchor;
}

// 新規追加：SATO_DATA!B2 の JSON を読み込んで UI に反映
async function loadSatoFromSheet() {
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem("SATO_DATA");
      const cell = ws.getRange("B2");
      cell.load("values");
      await ctx.sync();

      const raw = String(cell.values?.[0]?.[0] || "");
      if (!raw) { setStatus(".sato not found in SATO_DATA"); return; }

      const p = JSON.parse(raw);
      const set = (id, v)=>{ if(el(id)!=null && v!=null) el(id).value = v; };

      set("chartType", p.chart);
      set("xColumn",   p.x);
      set("yColumn",   p.y);

      const o = p.opts || {};
      set("titleText", o.title);
      set("xLabel",    o.xlab);
      set("yLabel",    o.ylab);
      set("ggtheme",   o.theme);
      set("fontFam",   o.family);
      set("fillColor", o.fill);
      set("fillPicker",o.fill || o.fillPicker);
      set("fillAlpha", o.fillAlpha);
      set("strokeColor", o.stroke);
      set("strokePicker",o.stroke || o.strokePicker);
      set("lineWidth",    o.lineWidth);
      set("titleSize",    o.titleSize);
      // Handle both old format (axisTitleSize/axisTextSize) and new format (4 separate values)
      set("xAxisTitleSize",  o.xAxisTitleSize || o.axisTitleSize);
      set("yAxisTitleSize",  o.yAxisTitleSize || o.axisTitleSize);
      set("xAxisTextSize",   o.xAxisTextSize || o.axisTextSize);
      set("yAxisTextSize",   o.yAxisTextSize || o.axisTextSize);
      set("legendTextSize", o.legendTextSize);
      set("titleWeight",  o.titleWeight);
      set("axisTitleWeight", o.axisTitleWeight || o.axisWeight);
      set("axisTextWeight",  o.axisTextWeight || o.axisWeight);
      set("dataOrder",    o.dataOrder || "default");
      set("customOrder",  o.customOrder || "");
      set("numGroups",    o.numGroups || 2);
      for (let i = 0; i < 6; i++) {
        set(`groupColor${i+1}`, (o.groupColors && o.groupColors[i]) || ["#4C78A8", "#E15759", "#57C4AD", "#E9C46A", "#F76C6C", "#A8DADC"][i]);
      }
      if (o.xMin!=null) set("xMin", o.xMin);
      if (o.xMax!=null) set("xMax", o.xMax);
      if (o.yMin!=null) set("yMin", o.yMin);
      if (o.yMax!=null) set("yMax", o.yMax);
      if (o.bins!=null) set("bins", o.bins);

      setStatus(".sato loaded from SATO_DATA");
    });
  } catch (e) {
    console.error(e);
    setStatus(".sato load error: " + (e?.message || e));
  }
}

function getExportSettings() {
  const width  = parseFloat(document.getElementById("expWidth")?.value)  || 6;
  const height = parseFloat(document.getElementById("expHeight")?.value) || 4;
  const units  = (document.getElementById("expUnits")?.value || "in");
  const dpi    = parseInt(document.getElementById("expDpi")?.value, 10)  || 300;
  return { width, height, units, dpi };
}

// まだどこかで呼ばれている場合の互換スタブ（Excelには触れずRの dat をCSV化して返す）
async function getSelectedRangeAsCSV() {
  await ensureDat(); await initWebR();
  const r = `
    tmp <- tempfile(fileext = ".csv")
    utils::write.csv(dat, file = tmp, row.names = FALSE)
    paste(readLines(tmp, warn = FALSE), collapse = "\\n")
  `;
  const res = await webR.evalR(r);
  const js  = await res.toJs();
  return Array.isArray(js?.values) ? js.values[0] : String(js);
}

// ========= .wrplot.json のペイロードを組み立て =========
function buildWrplotBundle({ csvText, rCode, params, titleBase }) {
  return {
    version: "0.1.0",
    guid: crypto.randomUUID ? crypto.randomUUID() :
          ((Math.random()+"").slice(2)+"-"+Date.now()),
    title: titleBase || "Plot",
    data_csv: csvText,
    plot_r: rCode,
    params_json: params || {},
    export: getExportSettings(),
    created_at: new Date().toISOString()
  };
}

// ========= PNG と .wrplot を同時に書き出すメイン関数 =========
async function exportPngAndWrplot() {
  await initWebR()
  try {
    // 1) Excel 選択範囲 → CSV
    const csvText = await getSelectedRangeAsCSV();

    // 2) Rコード（あなたのUIに既存のrCodeがあればそれを使ってください）
    //    MVP: ヒストグラムの例
    const rCode =
`hasGg <- requireNamespace("ggplot2", quietly = TRUE)
if (hasGg) {
  library(ggplot2)
  p <- ggplot(dat, aes(x = dat[[1]])) + geom_histogram(bins = 20) + theme_minimal()
  print(p)
} else {
  hist(dat[[1]], breaks = 20, main = "Histogram")
}`;

    // 3) webR で PNG 生成（寸法/DPI 反映）
    const b64 = await runRtoPng(csvText, rCode, {});
    const dataUrl = "data:image/png;base64," + b64;

    // 4) ファイル名ベース
    const title = (document.getElementById("titleText")?.value || "").trim();
    const base  = (title || "Histogram").replace(/[^\w\-]+/g, "_");
    const pngName = base + ".png";
    const wrpName = base + ".wrplot.json";

    // 5) PNG 保存（毎回フォルダ選択 → フォールバック）
    const pngBlob = await (await fetch(dataUrl)).blob();
    const infoPng = await saveToDirectoryEachTime({ data: pngBlob, fileName: pngName, mime: "image/png" });

    // 6) .wrplot.json 保存（同じ方式）
    const bundle = buildWrplotBundle({ csvText, rCode, params: {}, titleBase: base });
    const jsonBlob = new Blob([JSON.stringify(bundle, null, 2)], { type: "application/json" });
    const infoJson = await saveToDirectoryEachTime({ data: jsonBlob, fileName: wrpName, mime: "application/json" });

    setStatus(`Export completed: ${infoPng.fileName} / ${infoJson.fileName}`);
  } catch (e) {
    console.error(e);
    setStatus("PNG + .wrplot export error: " + (e?.message || e));
  }
}

async function copyPngToClipboard() {
  try {
    // Adjust this to your PNG source:
    // Example: Use dataURL generated directly by webR or get dataURL from <canvas id="plot">
    const dataUrl = document.getElementById("plot").toDataURL("image/png");
    const blob = await (await fetch(dataUrl)).blob();

    if (navigator.clipboard && window.ClipboardItem) {
      await navigator.clipboard.write([ new ClipboardItem({ [blob.type]: blob }) ]);
      setStatus("PNG copied to clipboard. Press ⌘V in PowerPoint to paste.");
    } else {
      setStatus("Clipboard API not supported in this environment. Please use 'Insert to Sheet'.");
    }
  } catch (e) {
    console.error(e);
    setStatus("Clipboard copy failed: " + (e?.message || e));
  }
}

async function insertIntoSheet() {
  try {
    const dataUrl = document.getElementById("plot").toDataURL("image/png");
    const base64 = dataUrl.split(",")[1];
    await new Promise((res, rej) =>
      Office.context.document.setSelectedDataAsync(
        dataUrl, // ← ここは dataURL を渡す（prefixあり）
        { coercionType: Office.CoercionType.Image },
        r => r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error)
      )
    );
    setStatus("Inserted to sheet. Right-click → 'Save as Picture...' to save to any location.");
  } catch (e) {
    console.error(e);
    setStatus("Sheet insert failed: " + (e?.message || e));
  }
}

// ========= Export R Code Functions =========

// Download R code as .R file
function downloadRFile(rCode, filename = 'generated_plot.R') {
  try {
    // Strip any accidental wrapping quotes
    let cleanCode = rCode;
    if (cleanCode.startsWith('"') && cleanCode.endsWith('"')) {
      cleanCode = cleanCode.slice(1, -1);
    }
    if (cleanCode.startsWith("'") && cleanCode.endsWith("'")) {
      cleanCode = cleanCode.slice(1, -1);
    }

    const blob = new Blob([cleanCode], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';

    document.body.appendChild(a);
    a.click();

    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);

    return true;
  } catch (e) {
    console.error('Download failed:', e);
    return false;
  }
}

// Write R code to selected Excel cell
async function writeRCodeToCell(rCode) {
  try {
    // Replace all double quotes with single quotes to prevent Excel CSV-style escaping
    let cleanCode = rCode.replace(/"/g, "'");

    console.log('📝 Writing R code to Excel - first 500 chars:', cleanCode.substring(0, 500));
    console.log('📝 Check labs line in code:', cleanCode.match(/labs\(title.*?\)/)?.[0]);

    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("address");
      await context.sync();

      // Write R code to selected cell
      // Force text format to prevent Excel from mangling quotes
      selectedRange.numberFormat = [["@"]];  // @ = text format in Excel
      selectedRange.values = [[cleanCode]];

      await context.sync();
      setStatus(`✅ R code written to cell: ${selectedRange.address}`);
    });
    return true;
  } catch (e) {
    console.error('Write to cell failed:', e);
    setStatus("❌ Error: Please select a cell first, then click Export R Code");
    return false;
  }
}

// Generate subset R code from Excel data (data frame + basic ggplot)
function generateSubsetRCodeFromData(chartType, opts) {
  try {
    if (!window.lastProcessedData || window.lastProcessedData.length === 0) {
      return "# No data loaded. Please load data first.";
    }

    const data = window.lastProcessedData;
    const hasHeader = data[0].some(cell => typeof cell === 'string' && isNaN(cell));

    // Create data frame code
    let dataFrameCode = "# Create data frame\ndat <- data.frame(\n";

    // Store original headers for labels
    const originalHeaders = hasHeader ? data[0] : [];

    if (hasHeader) {
      const headers = data[0];
      const dataRows = data.slice(1);

      headers.forEach((header, colIdx) => {
        // Use simple safe column names (col1, col2, col3...)
        const colName = `col${colIdx + 1}`;

        const colValues = dataRows.map(row => row[colIdx]);
        const isNumeric = colValues.every(v => v === null || v === '' || !isNaN(v));

        if (isNumeric) {
          const values = colValues.map(v => v === null || v === '' ? 'NA' : v).join(', ');
          dataFrameCode += `  ${colName} = c(${values})`;
        } else {
          const values = colValues.map(v => v === null || v === '' ? 'NA' : `'${String(v).replace(/'/g, "\\'")}'`).join(', ');
          dataFrameCode += `  ${colName} = c(${values})`;
        }

        if (colIdx < headers.length - 1) dataFrameCode += ',\n';
      });
    } else {
      // No header - use V1, V2, V3... as column names
      const numCols = data[0].length;
      for (let colIdx = 0; colIdx < numCols; colIdx++) {
        const colValues = data.map(row => row[colIdx]);
        const isNumeric = colValues.every(v => v === null || v === '' || !isNaN(v));

        if (isNumeric) {
          const values = colValues.map(v => v === null || v === '' ? 'NA' : v).join(', ');
          dataFrameCode += `  V${colIdx + 1} = c(${values})`;
        } else {
          const values = colValues.map(v => v === null || v === '' ? 'NA' : `'${String(v).replace(/'/g, "\\'")}'`).join(', ');
          dataFrameCode += `  V${colIdx + 1} = c(${values})`;
        }

        if (colIdx < numCols - 1) dataFrameCode += ',\n';
      }
    }

    dataFrameCode += "\n)\n\n";

    // Create simple ggplot code based on chart type
    let ggplotCode = "# Create plot\nlibrary(ggplot2)\n\n";

    const {title, xlab, ylab, fillColor, strokeColor, fillAlpha,
           barWidth, dodgeWidth, lineWidth, groupColIndex, xColIndex, yColIndex, themeName, rColorVector} = opts;

    // Get additional settings from UI
    const fontFamily = document.getElementById("fontFam")?.value || "Arial";
    const titleSize = document.getElementById("titleSize")?.value || "24";
    const xAxisTitleSize = document.getElementById("xAxisTitleSize")?.value || "20";
    const yAxisTitleSize = document.getElementById("yAxisTitleSize")?.value || "20";
    const xAxisTextSize = document.getElementById("xAxisTextSize")?.value || "18";
    const yAxisTextSize = document.getElementById("yAxisTextSize")?.value || "18";
    const legendTextSize = document.getElementById("legendTextSize")?.value || "16";
    const titleWeight = document.getElementById("titleWeight")?.value || "bold";
    const axisTitleWeight = document.getElementById("axisTitleWeight")?.value || "plain";
    const axisTextWeight = document.getElementById("axisTextWeight")?.value || "plain";

    // Escape single quotes for R strings (use single quotes in R code)
    const escapeForRLabel = (str) => {
      if (!str) return '';
      return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    };

    // Get original column names for labels (escaped)
    const xLabelOrig = originalHeaders[0] ? escapeForRLabel(originalHeaders[0]) : (xlab || 'X');
    const yLabelOrig = originalHeaders[1] ? escapeForRLabel(originalHeaders[1]) : (ylab || 'Y');

    switch(chartType) {
      case 'histogram':
        ggplotCode += `p <- ggplot(dat, aes(x = col1)) +\n`;
        ggplotCode += `  geom_histogram(bins = 20, fill = '${fillColor || '#4C78A8'}', color = '${strokeColor || '#1f2937'}', alpha = ${fillAlpha || 1}) +\n`;
        ggplotCode += `  labs(title = '${escapeForRLabel(title || 'Histogram')}', x = '${xLabelOrig}', y = '${escapeForRLabel(ylab || 'Frequency')}') +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
        break;

      case 'bar':
      case 'bar_error':
        ggplotCode += `p <- ggplot(dat, aes(x = col1, y = col2)) +\n`;
        ggplotCode += `  geom_bar(stat = 'identity', fill = '${fillColor || '#4C78A8'}', color = '${strokeColor || '#1f2937'}', alpha = ${fillAlpha || 1}, width = ${barWidth || 0.8}, linewidth = ${lineWidth || 0.5}) +\n`;
        ggplotCode += `  labs(title = '${escapeForRLabel(title || 'Bar Plot')}', x = '${xLabelOrig}', y = '${yLabelOrig}') +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
        break;

      case 'bar_grouped':
      case 'bar_grouped_error_dot':
        // Parse group colors from rColorVector - convert to single quotes
        const groupColorsStr = (rColorVector || "c('#4C78A8', '#E15759', '#76B7B2', '#F28E2B')").replace(/"/g, "'");
        const groupCol = `col${groupColIndex || 1}`;
        const xCol = `col${xColIndex || 2}`;
        const yCol = `col${yColIndex || 3}`;
        ggplotCode += `p <- ggplot(dat, aes(x = ${xCol}, y = ${yCol}, fill = ${groupCol})) +\n`;
        ggplotCode += `  geom_bar(stat = 'identity', position = position_dodge(width = ${dodgeWidth || 0.9}), width = ${barWidth || 0.8}, color = '${strokeColor || '#1f2937'}', alpha = ${fillAlpha || 1}, linewidth = ${lineWidth || 0.5}) +\n`;
        ggplotCode += `  scale_fill_manual(values = ${groupColorsStr}) +\n`;
        ggplotCode += `  labs(title = '${escapeForRLabel(title || 'Grouped Bar Plot')}', x = '${escapeForRLabel(originalHeaders[xColIndex - 1] || xlab || 'Category')}', y = '${escapeForRLabel(originalHeaders[yColIndex - 1] || ylab || 'Value')}', fill = '${escapeForRLabel(originalHeaders[groupColIndex - 1] || 'Group')}') +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
        break;

      case 'bar_grouped_error':
        // Grouped bar chart with pre-calculated error bars (4 columns: Group, Category, Mean, Error)
        const groupColorsError = (rColorVector || "c('#4C78A8', '#E15759', '#76B7B2', '#F28E2B')").replace(/"/g, "'");
        const groupColE = `col${groupColIndex || 1}`;
        const xColE = `col${xColIndex || 2}`;
        const yColE = `col${yColIndex || 3}`;
        const errColE = `col${errorColIndex || 4}`;
        ggplotCode += `p <- ggplot(dat, aes(x = ${xColE}, y = ${yColE}, fill = ${groupColE})) +\n`;
        ggplotCode += `  geom_bar(stat = 'identity', position = position_dodge(width = ${dodgeWidth || 0.9}), width = ${barWidth || 0.8}, color = '${strokeColor || '#1f2937'}', alpha = ${fillAlpha || 1}, linewidth = ${lineWidth || 0.5}) +\n`;
        ggplotCode += `  geom_errorbar(aes(ymin = ${yColE} - ${errColE}, ymax = ${yColE} + ${errColE}), position = position_dodge(width = ${dodgeWidth || 0.9}), width = 0.25, linewidth = ${lineWidth * 0.8 || 0.4}) +\n`;
        ggplotCode += `  scale_fill_manual(values = ${groupColorsError}) +\n`;
        ggplotCode += `  labs(title = '${escapeForRLabel(title || 'Grouped Bar Plot with Error')}', x = '${escapeForRLabel(originalHeaders[xColIndex - 1] || xlab || 'Category')}', y = '${escapeForRLabel(originalHeaders[yColIndex - 1] || ylab || 'Mean')}', fill = '${escapeForRLabel(originalHeaders[groupColIndex - 1] || 'Group')}') +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
        break;

      case 'box_grouped':
      case 'box_grouped_dot':
      case 'violin_grouped':
      case 'violin_grouped_dot':
      case 'line_grouped':
      case 'line_grouped_error':
      case 'line_grouped_error_raw':
        const groupColorsLine = (rColorVector || "c('#4C78A8', '#E15759', '#76B7B2', '#F28E2B')").replace(/"/g, "'");
        const groupColLine = `col${groupColIndex || 1}`;
        const xColLine = `col${xColIndex || 2}`;
        const yColLine = `col${yColIndex || 3}`;
        ggplotCode += `p <- ggplot(dat, aes(x = ${xColLine}, y = ${yColLine}, color = ${groupColLine})) +\n`;
        ggplotCode += `  geom_point(size = 3) +\n`;
        ggplotCode += `  scale_color_manual(values = ${groupColorsLine}) +\n`;
        ggplotCode += `  labs(title = '${escapeForRLabel(title || 'Grouped Plot')}', x = '${escapeForRLabel(originalHeaders[xColIndex - 1] || xlab || 'X')}', y = '${escapeForRLabel(originalHeaders[yColIndex - 1] || ylab || 'Value')}', color = '${escapeForRLabel(originalHeaders[groupColIndex - 1] || 'Group')}') +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
        break;

      case 'box':
      case 'box_dot':
        ggplotCode += `p <- ggplot(dat, aes(x = col1, y = col2)) +\n`;
        ggplotCode += `  geom_boxplot(fill = '${fillColor || '#4C78A8'}', color = '${strokeColor || '#1f2937'}', alpha = ${fillAlpha || 1}) +\n`;
        ggplotCode += `  labs(title = '${escapeForRLabel(title || 'Box Plot')}', x = '${xLabelOrig}', y = '${yLabelOrig}') +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
        break;

      case 'dot':
        ggplotCode += `p <- ggplot(dat, aes(x = col1, y = col2)) +\n`;
        ggplotCode += `  geom_point(color = '${fillColor || '#4C78A8'}', size = 3, alpha = ${fillAlpha || 1}) +\n`;
        ggplotCode += `  labs(title = '${escapeForRLabel(title || 'Dot Plot')}', x = '${xLabelOrig}', y = '${yLabelOrig}') +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
        break;

      default:
        ggplotCode += `# Chart type: ${chartType}\n`;
        ggplotCode += `p <- ggplot(dat, aes(x = col1, y = col2)) +\n`;
        ggplotCode += `  geom_point() +\n`;
        ggplotCode += `  theme_${themeName || 'classic'}(base_size = ${xAxisTextSize}, base_family = '${fontFamily}')\n`;
    }

    // Add detailed theme customization for text sizes and appearance
    ggplotCode += `\n# Customize theme elements\n`;
    ggplotCode += `p <- p + theme(\n`;
    ggplotCode += `  plot.background = element_rect(fill = 'white', color = NA),\n`;
    ggplotCode += `  panel.background = element_rect(fill = 'white', color = NA),\n`;
    ggplotCode += `  panel.grid.major.y = element_line(color = '#e5e7eb', linewidth = 0.6),\n`;
    ggplotCode += `  panel.grid.minor = element_blank(),\n`;
    ggplotCode += `  axis.line = element_line(linewidth = 0.6),\n`;
    ggplotCode += `  axis.ticks = element_line(linewidth = 0.5),\n`;
    ggplotCode += `  plot.title = element_text(size = ${titleSize}, face = '${titleWeight}', family = '${fontFamily}', hjust = 0.5),\n`;
    ggplotCode += `  axis.title.x = element_text(size = ${xAxisTitleSize}, face = '${axisTitleWeight}', family = '${fontFamily}'),\n`;
    ggplotCode += `  axis.title.y = element_text(size = ${yAxisTitleSize}, face = '${axisTitleWeight}', family = '${fontFamily}'),\n`;
    ggplotCode += `  axis.text.x = element_text(size = ${xAxisTextSize}, family = '${fontFamily}', color = 'black', face = '${axisTextWeight}'),\n`;
    ggplotCode += `  axis.text.y = element_text(size = ${yAxisTextSize}, family = '${fontFamily}', color = 'black', face = '${axisTextWeight}'),\n`;
    ggplotCode += `  legend.text = element_text(size = ${legendTextSize}, family = '${fontFamily}'),\n`;
    ggplotCode += `  legend.title = element_text(size = ${legendTextSize}, family = '${fontFamily}')\n`;
    ggplotCode += `)\n`;

    // Add chart rotation (coord_flip) if needed
    const rotation = document.getElementById("rotation")?.value || "0";
    const tableStyleLabels = document.getElementById("tableStyleLabels")?.checked || false;

    if (rotation === "90") {
      ggplotCode += `\n# Rotate chart 90 degrees (horizontal bars)\n`;
      ggplotCode += `p <- p + coord_flip()\n`;

      // Add table-style Y-axis labels (left-aligned) for horizontal bars
      if (tableStyleLabels) {
        ggplotCode += `\n# Table-style Y-axis labels (left-aligned)\n`;
        ggplotCode += `# Remove default Y-axis text and title\n`;
        ggplotCode += `p <- p + theme(\n`;
        ggplotCode += `  axis.text.y = element_blank(),\n`;
        ggplotCode += `  axis.title.y = element_blank()\n`;
        ggplotCode += `)\n\n`;
        ggplotCode += `# Add left-aligned category labels at start of bars\n`;
        ggplotCode += `label_offset <- max(dat$col2, na.rm = TRUE) * 0.02  # 2% of max value\n`;
        ggplotCode += `p <- p + geom_text(\n`;
        ggplotCode += `  aes(x = col1, y = label_offset, label = col1),\n`;
        ggplotCode += `  hjust = 0,\n`;
        ggplotCode += `  size = ${yAxisTextSize} / 3,\n`;
        ggplotCode += `  family = '${fontFamily}'\n`;
        ggplotCode += `)\n\n`;
        ggplotCode += `# Adjust Y-axis to start from 0\n`;
        ggplotCode += `p <- p + scale_y_continuous(expand = expansion(mult = c(0, 0.05)))\n`;
      }
    }

    ggplotCode += "\nprint(p)\n";

    return dataFrameCode + ggplotCode;

  } catch (e) {
    console.error('Generate subset R code failed:', e);
    return `# Error generating subset R code: ${e.message}`;
  }
}

// Generate R code from current settings
function generateCurrentRCode() {
  try {
    // Get R code from the last preview
    if (window.lastRCode && window.lastRCode.length > 0) {
      return window.lastRCode;
    } else {
      return "# No R code generated yet.\n# Please click 'Preview' first to generate the chart and R code.";
    }
  } catch (e) {
    console.error('Generate R code failed:', e);
    return `# Error generating R code: ${e.message}`;
  }
}

// Generate educational R code with manual calculations (no custom functions)
// This creates readable ggplot2 code that beginners can learn from
function extractRelevantRCode() {
  try {
    if (!window.lastRCode || window.lastRCode.length === 0) {
      return "# No R code generated yet.\n# Please click 'Preview' first to generate the chart and R code.";
    }

    const fullCode = window.lastRCode;
    const chartType = window.lastChartType;

    console.log(`Generating educational R code for chart type: ${chartType}`);

    // Generate data frame from window.lastProcessedData (not from R code)
    let dataFrame = '';
    if (!window.lastProcessedData || window.lastProcessedData.length === 0) {
      console.error('No processed data available');
      dataFrame = '# ERROR: No data available\ndat <- data.frame()';
    } else {
      // Build data frame with proper formatting (using SINGLE quotes)
      const data = window.lastProcessedData;
      const headers = data[0];
      const dataRows = data.slice(1);

      // Store original headers for labels
      console.log('💾 Storing originalHeaders:', headers);
      if (headers[2]) {
        console.log('💾 Type check - headers[2]:', typeof headers[2], 'value:', headers[2]);
        console.log('💾 Char codes of headers[2]:', Array.from(headers[2]).map((c, i) => `[${i}]='${c}' (${c.charCodeAt(0)})`));
      }
      window.lastOriginalHeaders = headers;

      dataFrame = '# Create data frame from your Excel data\ndat <- data.frame(\n';

      headers.forEach((header, colIdx) => {
        // Use simple safe column names (col1, col2, col3...)
        const colName = `col${colIdx + 1}`;
        const colValues = dataRows.map(row => row[colIdx]);
        const isNumeric = colValues.every(v => v === null || v === '' || !isNaN(v));

        if (isNumeric) {
          // Numeric column
          const values = colValues.map(v => v === null || v === '' ? 'NA' : v).join(', ');
          dataFrame += `  ${colName} = c(${values})`;
        } else {
          // String column - USE SINGLE QUOTES
          const values = colValues.map(v => {
            if (v === null || v === '') return 'NA';
            // Escape single quotes in the string value
            const escaped = String(v).replace(/'/g, "\\'");
            return `'${escaped}'`;
          }).join(', ');
          dataFrame += `  ${colName} = c(${values})`;
        }

        if (colIdx < headers.length - 1) dataFrame += ',\n';
      });

      dataFrame += '\n)\n';
    }

    console.log('Generated data frame:', dataFrame.substring(0, 150) + '...');

    // Use stored settings from preview instead of reading from UI
    // This ensures the R code reproduces the exact preview figure
    if (!window.lastPlotSettings) {
      return "# No plot settings stored.\n# Please click 'Preview' first to generate the chart.";
    }

    const settings = window.lastPlotSettings;
    console.log('Using stored plot settings from preview');

    // Get group colors from stored settings
    const groupColors = settings.groupColors
      ? `c(${settings.groupColors.map(c => `'${c}'`).join(', ')})`
      : getGroupColorsForRCode();

    // Generate code based on chart type
    let rCode = generateEducationalRCode(chartType, dataFrame, settings, groupColors);

    console.log(`✅ Generated educational R code (${rCode.length} chars)`);
    console.log('🔍 Final R code labs line:', rCode.match(/labs\(title.*?\)/)?.[0]);

    // Replace any remaining double quotes with single quotes in R code
    // This prevents Excel from escaping them as ""
    rCode = rCode.replace(/"/g, "'");

    return rCode;

  } catch (e) {
    console.error('Generate educational R code failed:', e);
    console.error('Error details:', e.stack);
    return `# Error generating R code: ${e.message}`;
  }
}

// Helper function to get group colors in R format
function getGroupColorsForRCode() {
  const colorInputs = document.querySelectorAll('#groupColorControls input[type="color"]');
  if (colorInputs.length === 0) {
    return "c('#4C78A8', '#E15759', '#76B7B2', '#F28E2B')";
  }

  const colors = Array.from(colorInputs).map(input => `'${input.value}'`).join(', ');
  return `c(${colors})`;
}

// Generate educational R code with manual calculations
function generateEducationalRCode(chartType, dataFrame, settings, groupColors) {
  // Check if vbracket is needed (line_grouped with 3+ groups and statistics)
  const needsVbracket = chartType.startsWith('line_grouped') && settings.addStatistics && settings.numGroups >= 3;

  // Start with libraries - USE ONLY SINGLE QUOTES to avoid Excel double-quote escaping
  let code = `# ============================================
# Educational R Code - Generated by Office Add-in
# Chart Type: ${chartType}
# ============================================

# Load required libraries
library(ggplot2)
library(dplyr)
library(ggpubr)  # For statistical comparisons
${needsVbracket ? `library(vbracket)  # For custom legend with brackets` : ''}

`;

  // Add data frame
  code += dataFrame + '\n\n';

  // Generate code based on chart type
  if (chartType === 'bar_grouped_error_dot' || chartType === 'bar_grouped_error' || chartType === 'bar_grouped') {
    code += generateGroupedBarCode(chartType, settings, groupColors);
  } else if (chartType === 'box_grouped' || chartType === 'box_grouped_dot') {
    code += generateGroupedBoxCode(chartType, settings, groupColors);
  } else if (chartType === 'violin_grouped' || chartType === 'violin_grouped_dot') {
    code += generateGroupedViolinCode(chartType, settings, groupColors);
  } else if (chartType.startsWith('line_grouped')) {
    // Handle all line grouped types: line_grouped, line_grouped_error, line_grouped_error_raw
    code += generateLineGroupedCode(settings, groupColors);
  } else if (chartType === 'histogram' || chartType === 'box' || chartType === 'box_dot' ||
             chartType === 'dot' || chartType === 'bar' || chartType === 'bar_error' ||
             chartType === 'bar_error_dot' || chartType === 'violin' || chartType === 'violin_dot') {
    // Escape single quotes for R strings (use single quotes in R code)
    const escapeForRLabel = (str) => {
      if (!str) return '';
      return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    };

    const originalHeaders = window.lastOriginalHeaders || [];
    console.log('📋 originalHeaders:', originalHeaders);
    console.log('📋 yColIndex:', settings.yColIndex);
    console.log('📋 Raw Y header:', originalHeaders[settings.yColIndex - 1]);

    const rawYHeader = originalHeaders[settings.yColIndex - 1];
    console.log('🔍 Before escape - rawYHeader:', rawYHeader);
    console.log('🔍 rawYHeader char codes:', Array.from(String(rawYHeader)).map((c, i) => `[${i}]='${c}'(${c.charCodeAt(0)})`).join(' '));

    const xLabelOrig = originalHeaders[settings.xColIndex - 1] ? escapeForRLabel(originalHeaders[settings.xColIndex - 1]) : (settings.xLabel || 'X');
    const yLabelOrig = originalHeaders[settings.yColIndex - 1] ? escapeForRLabel(originalHeaders[settings.yColIndex - 1]) : (settings.yLabel || 'Y');

    console.log('📋 After escapeForRLabel - xLabelOrig:', xLabelOrig);
    console.log('📋 After escapeForRLabel - yLabelOrig:', yLabelOrig);
    console.log('🔍 yLabelOrig char codes:', Array.from(String(yLabelOrig)).map((c, i) => `[${i}]='${c}'(${c.charCodeAt(0)})`).join(' '));

    // Test: Build the labs line and show it
    const labsLine = `labs(title = '${escapeForRLabel(settings.title || 'Plot')}', x = '${xLabelOrig}', y = '${yLabelOrig}')`;
    console.log('🔍 Generated labs line:', labsLine);

    // Generate appropriate geom for chart type
    let geomCode = '';
    let additionalGeoms = '';

    if (chartType === 'bar' || chartType === 'bar_error') {
      geomCode = `geom_bar(stat = 'identity', fill = '${settings.fillColor || '#4C78A8'}', color = '${settings.strokeColor || '#1f2937'}', alpha = ${settings.fillAlpha || 0.9}, width = ${settings.barWidth || 0.4}, linewidth = ${settings.lineWidth || 0.7})`;
    } else if (chartType === 'bar_error_dot') {
      // For bar_error_dot: need stat_summary for bars + error bars, plus geom_point for dots
      const errorType = settings.errorBarType === 'SE' ? 'se' : (settings.errorBarType === 'CI95' ? 'ci95' : 'sd');

      geomCode = `stat_summary(fun = mean, geom = 'bar', fill = '${settings.fillColor || '#4C78A8'}', color = '${settings.strokeColor || '#1f2937'}', alpha = ${settings.fillAlpha || 0.9}, width = ${settings.barWidth || 0.4}, linewidth = ${settings.lineWidth || 0.7}) +
  stat_summary(fun.data = mean_${errorType}, geom = 'errorbar', width = 0.2, linewidth = ${settings.lineWidth || 0.7}) +
  geom_point(position = position_jitter(width = ${(settings.barWidth || 0.4) * 0.15}, height = 0), size = ${settings.dotSize || 3}, alpha = ${settings.dotAlpha || 1}, color = '${settings.dotColor || '#000000'}', shape = ${settings.dotShape || 16})`;
    } else if (chartType === 'histogram') {
      geomCode = `geom_histogram(bins = ${settings.bins || 20}, fill = '${settings.fillColor || '#4C78A8'}', color = '${settings.strokeColor || '#1f2937'}', alpha = ${settings.fillAlpha || 0.9})`;
    } else if (chartType === 'box' || chartType === 'box_dot') {
      geomCode = `geom_boxplot(fill = '${settings.fillColor || '#4C78A8'}', color = '${settings.strokeColor || '#1f2937'}', alpha = ${settings.fillAlpha || 0.9})`;
      if (chartType === 'box_dot') {
        additionalGeoms = `  geom_point(position = position_jitter(width = 0.15, height = 0), size = ${settings.dotSize || 3}, alpha = ${settings.dotAlpha || 0.6}, color = '${settings.dotColor || '#000000'}', shape = ${settings.dotShape || 16}) +\n`;
      }
    } else if (chartType === 'violin' || chartType === 'violin_dot') {
      geomCode = `geom_violin(fill = '${settings.fillColor || '#4C78A8'}', color = '${settings.strokeColor || '#1f2937'}', alpha = ${settings.fillAlpha || 0.9})`;
      if (chartType === 'violin_dot') {
        additionalGeoms = `  geom_point(position = position_jitter(width = 0.15, height = 0), size = ${settings.dotSize || 3}, alpha = ${settings.dotAlpha || 0.6}, color = '${settings.dotColor || '#000000'}', shape = ${settings.dotShape || 16}) +\n`;
      }
    } else {
      geomCode = `geom_point(color = '${settings.fillColor || '#4C78A8'}', size = 3, alpha = ${settings.fillAlpha || 0.9})`;
    }

    // Generate labels with plotmath support
    const titleLabel = convertToRPlotmath(settings.title || 'Plot');
    const xLabel = convertToRPlotmath(xLabelOrig);
    const yLabel = convertToRPlotmath(yLabelOrig);

    // Add factor ordering for single-group charts to preserve category order
    const xColName = `col${settings.xColIndex || 1}`;
    let actualCategories = [];

    if (window.lastProcessedData && window.lastProcessedData.length > 1) {
      const data = window.lastProcessedData;
      const catIdx = (settings.xColIndex || 1) - 1;  // Convert from 1-based to 0-based

      if (catIdx >= 0 && catIdx < data[0].length) {
        actualCategories = [...new Set(data.slice(1).map(row => row[catIdx]).filter(v => v))];
      }
    }

    // Set factor levels based on data order mode
    if (actualCategories.length > 0) {
      const dataOrder = settings.dataOrder || 'original';
      console.log('📊 [Plot] dataOrder mode:', dataOrder);
      console.log('📊 [Plot] actualCategories from data:', actualCategories);

      let finalCategories = [];

      if (dataOrder === 'custom') {
        // Custom order: use customOrderGroup (for single-group charts, X-axis values are stored here)
        const orderSource = settings.customOrderGroup || settings.customOrderCategory;
        console.log('📊 [Plot] Custom orderSource:', orderSource);

        let customCategories = [];
        if (orderSource) {
          if (typeof orderSource === 'string') {
            customCategories = orderSource.split(',').map(c => c.trim()).filter(c => c);
          } else if (Array.isArray(orderSource)) {
            customCategories = orderSource.map(c => String(c).trim());
          }
        }

        // Validate custom order
        const actualLower = actualCategories.map(c => String(c).toLowerCase().trim());
        const customLower = customCategories.map(c => String(c).toLowerCase().trim());
        const isValidCustomOrder = customCategories.length > 0 &&
                                   customLower.every(c => actualLower.includes(c));

        if (isValidCustomOrder) {
          finalCategories = customCategories;
          const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
          code += `# Set custom category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
          console.log('✅ [Plot] Using custom category order:', finalCategories);
        } else {
          // Fallback to original order if custom validation fails
          finalCategories = actualCategories;
          const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
          code += `# Set category order (original - custom validation failed)\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
          console.log('⚠️ [Plot] Custom validation failed, using original order:', finalCategories);
        }
      } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
        // Alphabetical order: sort categories
        finalCategories = [...actualCategories].sort((a, b) => String(a).localeCompare(String(b)));
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set alphabetical category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
        console.log('✅ [Plot] Using alphabetical order:', finalCategories);
      } else {
        // Original order: use order as values appear in data
        finalCategories = actualCategories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set original category order (as in data)\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
        console.log('✅ [Plot] Using original data order:', finalCategories);
      }
    }

    code += `# Create plot
p <- ggplot(dat, aes(x = col${settings.xColIndex || 1}, y = col${settings.yColIndex || 2})) +
  ${geomCode} +
${additionalGeoms}  labs(title = ${titleLabel}, x = ${xLabel}, y = ${yLabel}) +
  theme_${settings.themeName}(base_family = '${settings.fontFamily}') +
  theme(
    plot.background = element_rect(fill = 'white', color = NA),
    panel.background = element_rect(fill = 'white', color = NA),
    panel.grid.major.y = element_line(color = '#e5e7eb', linewidth = 0.6),
    panel.grid.minor = element_blank(),
    axis.line = element_line(linewidth = 0.6),
    axis.ticks = element_line(linewidth = 0.5),
    plot.title = element_text(size = ${settings.titleSize}, face = '${settings.titleWeight}', hjust = 0.5),
    axis.title.x = element_text(size = ${settings.xAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.title.y = element_text(size = ${settings.yAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.text.x = element_text(size = ${settings.xAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    axis.text.y = element_text(size = ${settings.yAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    legend.text = element_text(size = ${settings.legendTextSize}),
    legend.title = element_text(size = ${settings.legendTextSize})
  )
`;

    // Add statistical comparisons for single-group charts
    if (settings.addStatistics && (chartType === 'bar_error_dot' || chartType === 'box' ||
        chartType === 'box_dot' || chartType === 'violin_dot')) {
      code += generateSingleGroupStatisticalCode(settings, chartType);
    }

    // Add chart rotation if needed
    if (settings.rotation === 90 || settings.rotation === "90") {
      code += `\n# Rotate chart 90 degrees (horizontal bars)\np <- p + coord_flip()\n`;

      // Add table-style Y-axis labels if enabled
      if (settings.tableStyleLabels) {
        code += `\n# Table-style Y-axis labels (left-aligned)\n`;
        code += `# Remove default Y-axis text and title\n`;
        code += `p <- p + theme(\n`;
        code += `  axis.text.y = element_blank(),\n`;
        code += `  axis.title.y = element_blank()\n`;
        code += `)\n\n`;
        code += `# Add left-aligned category labels at start of bars\n`;
        code += `label_offset <- max(dat$col${settings.yColIndex || 2}, na.rm = TRUE) * 0.02  # 2% of max value\n`;
        code += `p <- p + geom_text(\n`;
        code += `  aes(x = col${settings.xColIndex || 1}, y = label_offset, label = col${settings.xColIndex || 1}),\n`;
        code += `  hjust = 0,\n`;
        code += `  size = ${settings.yAxisTextSize || 18} / 3,\n`;
        code += `  family = '${settings.fontFamily || 'Arial'}'\n`;
        code += `)\n\n`;
        code += `# Adjust Y-axis to start from 0\n`;
        code += `p <- p + scale_y_continuous(expand = expansion(mult = c(0, 0.05)))\n`;
      }
    }

    code += `\nprint(p)\n`;
  } else if (chartType === 'ic50_dose_response') {
    // IC50 Dose-Response curve
    const xColName = `col${settings.xColIndex || 1}`;
    const yColName = `col${settings.yColIndex || 2}`;
    const curveColor = settings.ic50CurveColor || '#2563eb';
    const curveWidth = settings.ic50CurveWidth || 1.5;
    const pointSize = settings.ic50PointSize || 3;
    const pointColor = settings.ic50PointColor || '#1f2937';
    const pointAlpha = settings.ic50PointAlpha ?? 1;
    const lineColor = settings.ic50LineColor || '#dc2626';
    const showIC50Value = settings.showIC50Value !== false;
    const showIC50Line = settings.showIC50Line !== false;
    const dataDisplay = settings.ic50DataDisplay || 'all_points';

    code += `# Prepare data for IC50 fitting
dat$conc <- as.numeric(dat$${xColName})
dat$response <- as.numeric(dat$${yColName})

# Remove missing values
dat <- dat[complete.cases(dat$conc, dat$response), ]

# 4-Parameter Logistic (4PL) model for IC50
# Model: response = Bottom + (Top - Bottom) / (1 + (conc/IC50)^Hill)

# Initial estimates
bottom_init <- min(dat$response, na.rm = TRUE)
top_init <- max(dat$response, na.rm = TRUE)
ic50_init <- median(dat$conc, na.rm = TRUE)
hill_init <- 1

# Fit the 4PL model
fit <- tryCatch({
  nls(
    response ~ Bottom + (Top - Bottom) / (1 + (conc / IC50)^Hill),
    data = dat,
    start = list(Bottom = bottom_init, Top = top_init, IC50 = ic50_init, Hill = hill_init),
    algorithm = 'port',
    lower = c(0, 0, min(dat$conc)/10, 0.1),
    upper = c(max(dat$response), max(dat$response)*1.5, max(dat$conc)*10, 10),
    control = list(maxiter = 500, warnOnly = TRUE)
  )
}, error = function(e) {
  # Fallback to 3PL (fixed Bottom = 0)
  tryCatch({
    nls(
      response ~ Top / (1 + (conc / IC50)^Hill),
      data = dat,
      start = list(Top = top_init, IC50 = ic50_init, Hill = hill_init),
      algorithm = 'port',
      lower = c(0, min(dat$conc)/10, 0.1),
      upper = c(max(dat$response)*1.5, max(dat$conc)*10, 10),
      control = list(maxiter = 500, warnOnly = TRUE)
    )
  }, error = function(e2) NULL)
})

# Extract IC50 value
ic50_value <- if (!is.null(fit)) coef(fit)['IC50'] else NA

# Generate curve data for plotting
curve_data <- data.frame(
  conc = 10^seq(log10(min(dat$conc)), log10(max(dat$conc)), length.out = 500)
)
if (!is.null(fit)) {
  curve_data$response <- predict(fit, newdata = curve_data)
}

${dataDisplay === 'mean_sd' || dataDisplay === 'mean_se' ? `# Calculate summary statistics for each concentration
summary_data <- aggregate(response ~ conc, data = dat, FUN = function(x) {
  c(mean = mean(x), sd = sd(x), se = sd(x)/sqrt(length(x)), n = length(x))
})
summary_data <- do.call(data.frame, summary_data)
colnames(summary_data) <- c('conc', 'mean', 'sd', 'se', 'n')
summary_data$error <- summary_data$${dataDisplay === 'mean_sd' ? 'sd' : 'se'}

# Create the plot with mean +/- ${dataDisplay === 'mean_sd' ? 'SD' : 'SE'}
p <- ggplot(summary_data, aes(x = conc, y = mean)) +
  geom_errorbar(aes(ymin = mean - error, ymax = mean + error),
                width = 0.1, linewidth = 0.5, color = '${pointColor}') +
  geom_point(size = ${pointSize}, color = '${pointColor}', alpha = ${pointAlpha}) +
` : `# Create the plot with all data points
p <- ggplot(dat, aes(x = conc, y = response)) +
  geom_point(size = ${pointSize}, color = '${pointColor}', alpha = ${pointAlpha}) +
`}  scale_x_log10() +
  labs(
    title = '${settings.title || 'Dose-Response Curve'}',
    x = '${settings.xLabel || 'Concentration'}',
    y = '${settings.yLabel || 'Response (%)'}'
  ) +
  theme_${settings.themeName}(base_family = '${settings.fontFamily}') +
  theme(
    plot.background = element_rect(fill = 'white', color = NA),
    panel.background = element_rect(fill = 'white', color = NA),
    plot.title = element_text(size = ${settings.titleSize}, hjust = 0.5),
    axis.title.x = element_text(size = ${settings.xAxisTitleSize}),
    axis.title.y = element_text(size = ${settings.yAxisTitleSize}),
    axis.text.x = element_text(size = ${settings.xAxisTextSize}, color = 'black'),
    axis.text.y = element_text(size = ${settings.yAxisTextSize}, color = 'black')
  )

# Add fitted curve
if (!is.null(fit)) {
  p <- p + geom_line(data = curve_data, aes(x = conc, y = response),
                     color = '${curveColor}', linewidth = ${curveWidth})
}

${showIC50Line ? `# Add IC50 reference lines
if (!is.na(ic50_value)) {
  y_at_ic50 <- predict(fit, newdata = data.frame(conc = ic50_value))
  p <- p +
    geom_vline(xintercept = ic50_value, linetype = 'dashed', color = '${lineColor}', linewidth = 0.8) +
    geom_hline(yintercept = y_at_ic50, linetype = 'dashed', color = 'gray50', linewidth = 0.5)
}
` : ''}
${showIC50Value ? `# Add IC50 annotation
if (!is.na(ic50_value)) {
  p <- p + annotate('text', x = max(dat$conc) * 0.7, y = max(dat$response) * 0.95,
                    label = sprintf('IC50 = %.3g', ic50_value), hjust = 1, vjust = 1,
                    size = 5, fontface = 'bold')
}
` : ''}
# Print IC50 result
cat('IC50 =', ic50_value, '\\n')

print(p)
`;
  } else {
    // Unknown chart type
    code += `# Chart type '${chartType}' is not yet supported for educational R code generation.
# Please use the complex R code instead, or request implementation for this chart type.

`;
  }

  return code;
}

// Helper function to convert formatted text to R plotmath expression
function convertToRPlotmath(text) {
  if (!text) return "''";

  // Check if text contains formatting markers
  const hasFormatting = text.includes('^') || text.includes('~') || text.includes('*') ||
                        text.includes('₂') || text.includes('₁₀') || text.includes('₀');

  if (!hasFormatting) {
    // No formatting - return simple string
    return `'${text.replace(/'/g, "\\'")}'`;
  }

  // Convert special Unicode subscripts to markup
  text = text.replace(/₂/g, '~2~');
  text = text.replace(/₁₀/g, '~10~');
  text = text.replace(/₀/g, '~0~');
  text = text.replace(/₁/g, '~1~');
  text = text.replace(/₃/g, '~3~');
  text = text.replace(/₄/g, '~4~');
  text = text.replace(/₅/g, '~5~');
  text = text.replace(/₆/g, '~6~');
  text = text.replace(/₇/g, '~7~');
  text = text.replace(/₈/g, '~8~');
  text = text.replace(/₉/g, '~9~');

  // Parse the text and build R plotmath expression
  let result = '';
  let current = '';
  let i = 0;

  while (i < text.length) {
    if (text[i] === '*' && text.indexOf('*', i + 1) > i) {
      // Italic found
      const endIdx = text.indexOf('*', i + 1);
      const italicText = text.substring(i + 1, endIdx);

      if (current) {
        // Add current text before italic
        result += (result ? '*' : '') + `'${current}'`;
        current = '';
      }
      // Add italic text
      result += (result ? '*' : '') + `italic('${italicText}')`;
      i = endIdx + 1;
    } else if (text[i] === '^' && text.indexOf('^', i + 1) > i) {
      // Superscript found
      const endIdx = text.indexOf('^', i + 1);
      const superText = text.substring(i + 1, endIdx);

      if (current) {
        // Add current text with superscript
        result += (result ? '*' : '') + `'${current}'^'${superText}'`;
        current = '';
      } else {
        // No preceding text, just superscript
        result += (result ? '*' : '') + `''^'${superText}'`;
      }
      i = endIdx + 1;
    } else if (text[i] === '~' && text.indexOf('~', i + 1) > i) {
      // Subscript found
      const endIdx = text.indexOf('~', i + 1);
      const subText = text.substring(i + 1, endIdx);

      if (current) {
        // Add current text with subscript
        result += (result ? '*' : '') + `'${current}'['${subText}']`;
        current = '';
      } else {
        // No preceding text, just subscript
        result += (result ? '*' : '') + `''['${subText}']`;
      }
      i = endIdx + 1;
    } else {
      current += text[i];
      i++;
    }
  }

  if (current) {
    result += (result ? '*' : '') + `'${current}'`;
  }

  return `expression(${result})`;
}

// Generate code for grouped bar charts with error bars and dots
function generateGroupedBarCode(chartType, settings, groupColors) {
  const hasErrorBars = chartType !== 'bar_grouped';
  const hasDots = chartType === 'bar_grouped_error_dot';

  // Use col1, col2, col3 naming to match data frame creation
  const groupColName = `col${settings.groupColIndex}`;
  const xColName = `col${settings.xColIndex}`;
  const yColName = `col${settings.yColIndex}`;
  const errorColName = hasErrorBars && settings.errorColIndex ? `col${settings.errorColIndex}` : null;

  let code = '';

  // Add factor ordering if custom order is set
  // Get actual unique values from data to validate and use correct values
  let actualCategories = [];
  let actualGroups = [];

  if (window.lastProcessedData && window.lastProcessedData.length > 1) {
    const data = window.lastProcessedData;
    const headers = data[0];
    const catIdx = settings.xColIndex - 1;  // Convert from 1-based to 0-based
    const grpIdx = settings.groupColIndex - 1;

    if (catIdx >= 0 && catIdx < headers.length) {
      actualCategories = [...new Set(data.slice(1).map(row => row[catIdx]).filter(v => v))];
    }
    if (grpIdx >= 0 && grpIdx < headers.length) {
      actualGroups = [...new Set(data.slice(1).map(row => row[grpIdx]).filter(v => v))];
    }
  }

  // Handle category ordering based on dataOrder mode
  if (actualCategories.length > 0) {
    let finalCategories = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let categories = [];
      if (settings.customOrderCategory) {
        if (typeof settings.customOrderCategory === 'string') {
          categories = settings.customOrderCategory.split(',').map(c => c.trim()).filter(c => c);
        } else if (Array.isArray(settings.customOrderCategory)) {
          categories = settings.customOrderCategory;
        }
      }

      // Validate: only use custom order if values match actual categories
      const isValidCategoryOrder = categories.length > 0 &&
                                   categories.every(c => actualCategories.includes(c));

      if (isValidCategoryOrder) {
        finalCategories = categories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set custom category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalCategories = actualCategories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set category order from data\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort categories alphabetically
      finalCategories = [...actualCategories].sort((a, b) => String(a).localeCompare(String(b)));
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set alphabetical category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalCategories = actualCategories;
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set category order from data (original order)\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    }
  }

  // Handle group ordering based on dataOrder mode
  if (actualGroups.length > 0) {
    let finalGroups = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let groups = [];
      if (settings.customOrderGroup) {
        if (typeof settings.customOrderGroup === 'string') {
          groups = settings.customOrderGroup.split(',').map(g => g.trim()).filter(g => g);
        } else if (Array.isArray(settings.customOrderGroup)) {
          groups = settings.customOrderGroup;
        }
      }

      // Validate: only use custom order if values match actual groups
      const isValidGroupOrder = groups.length > 0 &&
                                groups.every(g => actualGroups.includes(g));

      if (isValidGroupOrder) {
        finalGroups = groups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set custom group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalGroups = actualGroups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set group order from data\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort groups alphabetically
      finalGroups = [...actualGroups].sort((a, b) => String(a).localeCompare(String(b)));
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set alphabetical group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalGroups = actualGroups;
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set group order from data (original order)\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    }
  }

  // Check if we have pre-calculated error column (for bar_grouped_error)
  if (errorColName && chartType === 'bar_grouped_error') {
    // Data already has mean and error values - use directly
    code += `# ============ Prepare Data (Pre-calculated Values) ============

# Data already contains mean and error values
# Rename columns for easier reference
summary_data <- dat
colnames(summary_data) <- c('Group', 'Category', 'Mean', 'Error')

`;
  } else {
    // Calculate statistics from raw data
    code += `# ============ Calculate Summary Statistics ============

# Calculate mean and SD/SE for each group and category
summary_data <- dat %>%
  group_by(${groupColName}, ${xColName}) %>%
  summarise(
    mean_value = mean(${yColName}, na.rm = TRUE),
    sd_value = sd(${yColName}, na.rm = TRUE),
    se_value = sd(${yColName}, na.rm = TRUE) / sqrt(n()),
    n = n(),
    .groups = 'drop'
  )

# Rename columns for easier reference
colnames(summary_data) <- c('Group', 'Category', 'Mean', 'SD', 'SE', 'N')

# Add error values based on error type
summary_data$Error <- summary_data$${settings.errorBarType === 'SE' ? 'SE' : 'SD'}

`;
  }

  code += `# ============ Create Plot ============

# Create the base plot
p <- ggplot(summary_data, aes(x = Category, y = Mean, fill = Group)) +

  # Add bars
  geom_bar(
    stat = 'identity',
    position = position_dodge(width = ${settings.dodgeWidth}),
    width = ${settings.barWidth},
    color = '${settings.strokeColor}',
    alpha = ${settings.fillAlpha},
    linewidth = ${settings.lineWidth}
  ) +

`;

  if (hasErrorBars) {
    code += `  # Add error bars
  geom_errorbar(
    aes(ymin = Mean - Error, ymax = Mean + Error),
    position = position_dodge(width = ${settings.dodgeWidth}),
    width = 0.2,
    linewidth = ${settings.lineWidth}
  ) +

`;
  }

  if (hasDots) {
    code += `  # Add individual data points
  # Note: We need to match the dodge width exactly with the bars
  geom_point(
    data = dat,
    aes(x = ${xColName}, y = ${yColName}, fill = ${groupColName}),
    position = position_jitterdodge(
      dodge.width = ${settings.dodgeWidth},
      jitter.width = 0.15,
      jitter.height = 0
    ),
    size = ${settings.dotSize},
    alpha = ${settings.dotAlpha},
    color = '${settings.dotColor}',
    shape = ${settings.dotShape}
  ) +

`;
  }

  code += `  # Set colors
  scale_fill_manual(values = ${groupColors}, name = '${settings.selectedGroupColumn || 'Group'}') +

  # Apply theme
  theme_${settings.themeName}(base_family = '${settings.fontFamily}') +

  # Customize theme elements
  theme(
    plot.background = element_rect(fill = 'white', color = NA),
    panel.background = element_rect(fill = 'white', color = NA),
    panel.grid.major.y = element_line(color = '#e5e7eb', linewidth = 0.6),
    panel.grid.minor = element_blank(),
    axis.line = element_line(linewidth = 0.6),
    axis.ticks = element_line(linewidth = 0.5),
    plot.title = element_text(size = ${settings.titleSize}, face = '${settings.titleWeight}', hjust = 0.5),
    axis.title.x = element_text(size = ${settings.xAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.title.y = element_text(size = ${settings.yAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.text.x = element_text(size = ${settings.xAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    axis.text.y = element_text(size = ${settings.yAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    legend.text = element_text(size = ${settings.legendTextSize}),
    legend.title = element_text(size = ${settings.legendTextSize})
  )

# Add labels
`;

  if (settings.showTitle) {
    code += `p <- p + ggtitle(${convertToRPlotmath(settings.title)})\n`;
  }
  if (settings.showXLabel) {
    code += `p <- p + xlab(${convertToRPlotmath(settings.xlab)})\n`;
  }
  if (settings.showYLabel) {
    code += `p <- p + ylab(${convertToRPlotmath(settings.ylab)})\n`;
  }

  // Add statistical comparisons using helper function
  code += generateStatisticalComparisonCode(settings, groupColName, xColName);

  code += `\n# Display the plot
print(p)
`;

  return code;
}

// Helper function to generate statistical comparison code for grouped charts
function generateStatisticalComparisonCode(settings, groupColName, xColName) {
  if (!settings.addStatistics) return '';

  let code = `\n# ============ Statistical Comparisons ============
# Note: Statistical brackets from the preview are calculated using custom functions.
# For educational purposes, you can add comparisons manually using ggpubr:

library(ggpubr)
`;

  // Get actual unique groups and categories from data for better examples
  let uniqueGroups = [];
  let uniqueCategories = [];
  if (window.lastProcessedData && window.lastProcessedData.length > 1) {
    const data = window.lastProcessedData;
    const headers = data[0];
    const groupIdx = settings.groupColIndex - 1;
    const catIdx = settings.xColIndex - 1;

    if (groupIdx >= 0 && groupIdx < headers.length) {
      uniqueGroups = [...new Set(data.slice(1).map(row => row[groupIdx]).filter(v => v))];
    }
    if (catIdx >= 0 && catIdx < headers.length) {
      uniqueCategories = [...new Set(data.slice(1).map(row => row[catIdx]).filter(v => v))];
    }
  }

  // Parse custom comparisons if available (could be string or array)
  let customComps = [];
  if (settings.customComparisons) {
    if (typeof settings.customComparisons === 'string') {
      try {
        customComps = JSON.parse(settings.customComparisons);
      } catch (e) {
        console.warn('Failed to parse customComparisons:', e);
      }
    } else if (Array.isArray(settings.customComparisons)) {
      customComps = settings.customComparisons;
    }
  }

  // Generate actual comparisons based on mode
  if (settings.comparisonMode === 'custom' && customComps.length > 0) {
    // Parse custom comparisons from "Group1-Group2@Category" format
    const comparisonsByCategory = {};

    customComps.forEach(comp => {
      const parts = comp.split('@');
      if (parts.length === 2) {
        const [groupPair, category] = parts;
        const groups = groupPair.split('-');
        if (groups.length === 2) {
          if (!comparisonsByCategory[category]) {
            comparisonsByCategory[category] = [];
          }
          comparisonsByCategory[category].push(groups);
        }
      }
    });

    // Generate code - note: ggpubr's stat_compare_means applies to all categories
    // Custom per-category comparisons require more complex code not shown here
    const categories = Object.keys(comparisonsByCategory);
    if (categories.length > 0) {
      // Use the simplified approach for all categories
      code += generateCategoryByComparisonsCode(uniqueGroups, uniqueCategories, xColName, settings);
      code += `\n# Note: The add-in allows custom selection of which comparisons to show.
# The simplified ggpubr approach above shows all comparisons.
# For per-category control, you would need to manually compute tests and add annotations.
`;
    } else {
      // Fallback - generate for all categories
      code += generateCategoryByComparisonsCode(uniqueGroups, uniqueCategories, xColName, settings);
    }
  } else if (settings.comparisonMode === 'all') {
    // Generate comparisons for each category
    code += generateCategoryByComparisonsCode(uniqueGroups, uniqueCategories, xColName, settings);
  } else {
    // Significant only mode - still generate structure for each category
    code += generateCategoryByComparisonsCode(uniqueGroups, uniqueCategories, xColName, settings);
  }

  code += `\n# Modify the comparisons above as needed.

`;

  return code;
}

// Helper to generate category-by-category comparisons for grouped charts
function generateCategoryByComparisonsCode(groups, categories, xColName, settings) {
  if (groups.length < 2 || categories.length === 0) {
    return `# Please specify groups to compare
`;
  }

  // Get column names from settings - use colN format to match data frame
  const groupColName = `col${settings.groupColIndex}`;
  const yColName = `col${settings.yColIndex}`;

  let code = '';

  // Generate manual statistical test code for each category
  code += `
# ============ Perform Statistical Tests Manually ============
# This replicates the add-in's internal statistical analysis

# Initialize results storage
stat_results <- data.frame(
  category = character(),
  group1 = character(),
  group2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

`;

  // For each category, generate test code
  categories.forEach(category => {
    const safeName = category.replace(/[^a-zA-Z0-9]/g, '_');
    code += `# Test for ${category}
cat_data_${safeName} <- dat[dat$${xColName} == '${category}', ]
group1_data <- as.numeric(cat_data_${safeName}[cat_data_${safeName}$${groupColName} == '${groups[0]}', '${yColName}'])
group2_data <- as.numeric(cat_data_${safeName}[cat_data_${safeName}$${groupColName} == '${groups[1]}', '${yColName}'])

# Test for normality (Shapiro-Wilk test)
shapiro1_${safeName} <- shapiro.test(group1_data)
shapiro2_${safeName} <- shapiro.test(group2_data)
both_normal_${safeName} <- (shapiro1_${safeName}$p.value >= 0.05) && (shapiro2_${safeName}$p.value >= 0.05)

cat('  Normality test for ${category}:\\n')
cat('    ${groups[0]}: p =', shapiro1_${safeName}$p.value, if(shapiro1_${safeName}$p.value >= 0.05) '(normal)' else '(non-normal)', '\\n')
cat('    ${groups[1]}: p =', shapiro2_${safeName}$p.value, if(shapiro2_${safeName}$p.value >= 0.05) '(normal)' else '(non-normal)', '\\n')

# Test for equal variances (Levene test)
if (both_normal_${safeName}) {
  combined_data <- data.frame(
    values = c(group1_data, group2_data),
    group = factor(c(rep('${groups[0]}', length(group1_data)), rep('${groups[1]}', length(group2_data))))
  )
  group_means <- tapply(combined_data$values, combined_data$group, mean)
  abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
  levene_result <- anova(lm(abs_deviations ~ combined_data$group))
  levene_p_${safeName} <- levene_result[['Pr(>F)']][1]
  equal_var_${safeName} <- levene_p_${safeName} >= 0.05
  cat('  Variance test (Levene): p =', levene_p_${safeName}, if(equal_var_${safeName}) '(equal variances)' else '(unequal variances)', '\\n')
} else {
  equal_var_${safeName} <- TRUE  # Default for non-parametric tests
}

# Perform appropriate test
if (both_normal_${safeName}) {
  # Use t-test with appropriate var.equal parameter
  test_result_${safeName} <- t.test(group1_data, group2_data, var.equal = equal_var_${safeName})
  test_name_${safeName} <- if (equal_var_${safeName}) 'Student\\'s t-test' else 'Welch\\'s t-test'
} else {
  # Use Wilcoxon test for non-normal data
  test_result_${safeName} <- wilcox.test(group1_data, group2_data)
  test_name_${safeName} <- 'Wilcoxon test'
}

p_val_${safeName} <- test_result_${safeName}$p.value

# Determine significance symbol
sig_${safeName} <- if (p_val_${safeName} < 0.001) '***' else if (p_val_${safeName} < 0.01) '**' else if (p_val_${safeName} < 0.05) '*' else 'ns'

# Store result
stat_results <- rbind(stat_results, data.frame(
  category = '${category}',
  group1 = '${groups[0]}',
  group2 = '${groups[1]}',
  p_value = p_val_${safeName},
  significance = sig_${safeName}
))

cat('${category}:', test_name_${safeName}, 'p =', p_val_${safeName}, '(', sig_${safeName}, ')\\n')

`;
  });

  code += `
# Display results
print(stat_results)

# ============ Add Statistical Brackets using ggpubr ============

`;

  // Check if custom positions exist
  let customPositions = {};
  if (settings.customPositions) {
    if (typeof settings.customPositions === 'string') {
      try {
        customPositions = JSON.parse(settings.customPositions);
      } catch (e) {
        console.warn('Failed to parse customPositions:', e);
      }
    } else if (typeof settings.customPositions === 'object') {
      customPositions = settings.customPositions;
    }
  }

  // Generate bracket data frame for ggpubr
  const dodgeWidth = settings.dodgeWidth || 0.9;
  const categoryIndices = {};
  categories.forEach((cat, idx) => {
    categoryIndices[cat] = idx + 1; // R uses 1-based indexing
  });

  // Pre-calculate Y positions for all categories
  const yColIndex = settings.yColIndex - 1;
  const xColIndex = settings.xColIndex - 1;
  const allYValues = window.lastProcessedData
    .slice(1) // Skip header
    .map(row => parseFloat(row[yColIndex]))
    .filter(v => !isNaN(v));

  const yMin = Math.min(...allYValues);
  const yMax = Math.max(...allYValues);
  const yRange = yMax - yMin;

  // Build bracket positions for each category
  // Use exact positions from add-in if available
  const bracketData = [];

  // Check if we have bracket data from the add-in's calculation
  const hasAddInBrackets = window.lastBracketData && window.lastBracketData.length > 0;

  console.log("🔍 Educational R code - Checking bracket data:");
  console.log("  hasAddInBrackets:", hasAddInBrackets);
  console.log("  window.lastBracketData:", window.lastBracketData);
  console.log("  Number of categories:", categories.length);

  if (window.lastBracketData && window.lastBracketData.length > 0) {
    console.log("  First bracket object:", window.lastBracketData[0]);
    console.log("  First bracket y.position:", window.lastBracketData[0]["y.position"]);
    console.log("  All bracket y.positions:", window.lastBracketData.map(b => b["y.position"]));
  }

  categories.forEach((category, idx) => {
    const safeCategoryName = category.replace(/[^a-zA-Z0-9]/g, '_');
    const catIndex = categoryIndices[category];

    let yBracketPos;

    // Use exact Y position from add-in if available
    if (hasAddInBrackets && idx < window.lastBracketData.length) {
      yBracketPos = window.lastBracketData[idx]["y.position"];
      console.log(`📍 Using exact bracket position from add-in for ${category} (idx=${idx}): y=${yBracketPos}`);
    } else {
      console.log(`⚠️ No add-in bracket data for ${category} (idx=${idx}), calculating automatically`);
      // Fallback: Check for custom Y position
      const compKey1 = `${groups[0]}-${groups[1]}@${category}`;
      const compKey2 = `${groups[1]}-${groups[0]}@${category}`;
      const customYPos = customPositions[compKey1] || customPositions[compKey2];

      if (customYPos) {
        yBracketPos = customYPos;
      } else {
        // Calculate automatically from data
        const categoryYValues = window.lastProcessedData
          .slice(1)
          .filter(row => row[xColIndex] === category)
          .map(row => parseFloat(row[yColIndex]))
          .filter(v => !isNaN(v));

        const catYMax = categoryYValues.length > 0 ? Math.max(...categoryYValues) : yMax;
        yBracketPos = (catYMax + yRange * 0.1).toFixed(2);
      }
    }

    // Calculate xmin/xmax for this category's bracket
    const xmin = (catIndex - dodgeWidth/4).toFixed(3);
    const xmax = (catIndex + dodgeWidth/4).toFixed(3);

    bracketData.push({
      category: category,
      safeName: safeCategoryName,
      xmin: xmin,
      xmax: xmax,
      yPos: typeof yBracketPos === 'number' ? yBracketPos.toFixed(2) : yBracketPos
    });
  });

  // Generate R code to create bracket data frame
  const usingExactPositions = hasAddInBrackets;
  code += `# Create data frame for statistical brackets
${usingExactPositions ? '# Note: Y positions below are the exact values from the add-in preview\n' : ''}bracket_df <- data.frame(
  group1 = character(),
  group2 = character(),
  p.signif = character(),
  xmin = numeric(),
  xmax = numeric(),
  y.position = numeric(),
  stringsAsFactors = FALSE
)

`;

  // Add each bracket to the data frame
  bracketData.forEach(bracket => {
    code += `# Add bracket for ${bracket.category}
bracket_df <- rbind(bracket_df, data.frame(
  group1 = '${groups[0]}',
  group2 = '${groups[1]}',
  p.signif = sig_${bracket.safeName},
  xmin = ${bracket.xmin},
  xmax = ${bracket.xmax},
  y.position = ${bracket.yPos}
))

`;
  });

  code += `# Add all brackets to plot using ggpubr
p <- p + stat_pvalue_manual(
  bracket_df,
  label = 'p.signif',
  xmin = 'xmin',
  xmax = 'xmax',
  y.position = 'y.position',
  size = ${settings.statSymbolSize || 7},
  size.line = ${settings.statLineSize || 1.0},
  tip.length = ${settings.statTipLength || 0.04}
)
`;


  return code;
}

// Helper to generate fallback comparison code with actual group names
function generateFallbackComparisonCode(groups, categories) {
  if (groups.length >= 2 && categories.length >= 1) {
    const exampleComp = `c('${groups[0]}', '${groups[1]}')`;
    const exampleCategory = categories[0];
    return `# Example comparison (customize as needed):
comparisons <- list(${exampleComp})
p <- p + stat_compare_means(
  data = subset(dat, ${exampleCategory}),
  comparisons = comparisons,
  method = 't.test'
)
`;
  } else {
    return `# Example - replace with your actual comparisons:
comparisons <- list(c('Group1', 'Group2'))
p <- p + stat_compare_means(comparisons = comparisons, method = 't.test')
`;
  }
}

// Helper to generate all pairwise comparisons code
function generateAllComparisonsCode(groups, categories) {
  if (groups.length >= 2) {
    // Generate all pairwise combinations
    const allComps = [];
    for (let i = 0; i < groups.length; i++) {
      for (let j = i + 1; j < groups.length; j++) {
        allComps.push(`c('${groups[i]}', '${groups[j]}')`);
      }
    }
    const compList = allComps.join(', ');
    return `# All pairwise comparisons between groups:
comparisons <- list(${compList})
p <- p + stat_compare_means(comparisons = comparisons, method = 't.test')
`;
  } else {
    return `# Generate all pairwise comparisons
# Example for comparing all groups:
comparisons <- list(c('Group1', 'Group2'), c('Group1', 'Group3'), c('Group2', 'Group3'))
p <- p + stat_compare_means(comparisons = comparisons, method = 't.test')
`;
  }
}

// Helper to generate significant-only comparison code
function generateSignificantOnlyCode(groups, categories) {
  if (groups.length >= 2) {
    const exampleComp = `c('${groups[0]}', '${groups[1]}')`;
    return `# Significant comparisons only
# Run statistical tests first to identify significant pairs
# Example with actual group names:
comparisons <- list(${exampleComp})
p <- p + stat_compare_means(comparisons = comparisons, method = 't.test')
`;
  } else {
    return `# Significant comparisons only
# Run statistical tests first to identify significant pairs
# Example - replace with your actual significant comparisons:
comparisons <- list(c('Group1', 'Group2'))
p <- p + stat_compare_means(comparisons = comparisons, method = 't.test')
`;
  }
}

// Helper function to generate statistical comparison code for single-group charts
function generateSingleGroupStatisticalCode(settings, chartType) {
  if (!settings.addStatistics) return '';

  // Get unique categories from data
  let uniqueCategories = [];
  if (window.lastProcessedData && window.lastProcessedData.length > 1) {
    const data = window.lastProcessedData;
    const catIdx = settings.xColIndex - 1;
    if (catIdx >= 0) {
      uniqueCategories = [...new Set(data.slice(1).map(row => row[catIdx]).filter(v => v))];
    }
  }

  // Apply category order based on dataOrder mode (to match plot factor levels)
  const dataOrder = settings.dataOrder || 'original';
  console.log('📊 [Stats] dataOrder:', dataOrder);
  console.log('📊 [Stats] uniqueCategories from data:', uniqueCategories);
  console.log('📊 [Stats] postHocTest:', settings.postHocTest);

  if (dataOrder === 'custom') {
    // For single-group charts, the X-axis order may be stored in customOrderGroup
    const orderSource = settings.customOrderGroup || settings.customOrderCategory;
    console.log('📊 [Stats] Custom mode - orderSource:', orderSource);

    if (orderSource) {
      let customCategories = [];
      if (typeof orderSource === 'string') {
        customCategories = orderSource.split(',').map(c => c.trim()).filter(c => c);
      } else if (Array.isArray(orderSource)) {
        customCategories = orderSource.map(c => String(c).trim());
      }
      console.log('📊 [Stats] Parsed customCategories:', customCategories);

      // Use custom order if categories match (case-insensitive, type-coerced comparison)
      if (customCategories.length > 0) {
        // Convert all to lowercase strings for comparison
        const uniqueLower = uniqueCategories.map(c => String(c).toLowerCase().trim());
        const customLower = customCategories.map(c => String(c).toLowerCase().trim());

        // Check if all custom categories exist in data (regardless of order)
        const allValid = customLower.every(c => uniqueLower.includes(c));
        // Also check counts match to ensure no extra categories
        const countsMatch = customCategories.length === uniqueCategories.length;

        console.log('📊 [Stats] Validation - allValid:', allValid, 'countsMatch:', countsMatch);

        if (allValid && countsMatch) {
          // Use custom order, preserving original case from customCategories
          uniqueCategories = customCategories;
          console.log('✅ [Stats] Applied custom category order:', uniqueCategories);
        } else if (allValid) {
          // Custom has valid categories but different count - still use custom order
          // (might have subset or superset, but valid categories should still apply)
          uniqueCategories = customCategories;
          console.log('⚠️ [Stats] Applied custom order (count mismatch):', uniqueCategories);
        } else {
          console.log('❌ [Stats] Custom order validation failed.');
          console.log('   Custom (lowercase):', customLower);
          console.log('   Data (lowercase):', uniqueLower);
        }
      }
    }
  } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
    // Sort categories alphabetically
    uniqueCategories = [...uniqueCategories].sort((a, b) => String(a).localeCompare(String(b)));
    console.log('📊 [Stats] Applied alphabetical order:', uniqueCategories);
  } else {
    // Original order: use order as values appear in data (already done)
    console.log('📊 [Stats] Using original data order:', uniqueCategories);
  }

  if (uniqueCategories.length < 2) {
    return `\n# ============ Statistical Comparisons ============
# Need at least 2 categories for statistical comparisons

`;
  }

  const data = window.lastProcessedData;
  const headers = data && data.length > 0 ? data[0] : [];
  const xIdx = settings.xColIndex - 1;
  const yIdx = settings.yColIndex - 1;
  const xColName = `col${settings.xColIndex}`;
  const yColName = `col${settings.yColIndex}`;

  // Check test settings
  const isManualMode = settings.statisticalTestMode === 'manual';
  const isNonParametric = settings.dataType === 'nonparametric' ||
                          settings.statisticalTest === 'nonparametric' ||
                          settings.statisticalTest === 'kruskal';
  const isManualParametric = isManualMode && !isNonParametric;
  const postHocTest = settings.postHocTest || 'tukey';
  const isVsControl = postHocTest === 'dunnett' || postHocTest === 'steel';
  const controlGroup = settings.controlGroup || uniqueCategories[0];

  // Define pairs for ggpubr bracket generation (populated in each branch)
  let pairs = [];

  let code = `\n# ============ Statistical Comparisons ============
# This replicates the add-in's internal statistical analysis
# Data type: ${isNonParametric ? 'Non-parametric' : 'Parametric'}
# Post-hoc test: ${postHocTest}
${isVsControl ? `# Control group: ${controlGroup}` : '# Pairwise comparisons'}

`;

  // For 3+ groups with Steel/Dunnett, use vs-control approach
  if (uniqueCategories.length >= 3 && isVsControl) {
    // Generate pairs for vs-control comparisons
    uniqueCategories.filter(c => c !== controlGroup).forEach(trt => {
      pairs.push([controlGroup, trt]);
    });

    // Generate factor level order string for R
    const categoryOrder = uniqueCategories.map(c => `'${c}'`).join(', ');

    if (isNonParametric || postHocTest === 'steel') {
      // Steel test (non-parametric vs control)
      code += `# Install and load kSamples for Steel test
if (!requireNamespace('kSamples', quietly = TRUE)) {
  install.packages('kSamples')
}
library(kSamples)

# Steel test (non-parametric vs control)
# Control group: ${controlGroup}

# Set category order to match plot
dat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk):\\n')
all_normal <- TRUE
for (grp in unique(dat$${xColName})) {
  grp_data <- dat[dat$${xColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    is_non_normal <- sw_result$p.value < 0.05
    if (is_non_normal) all_normal <- FALSE
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(is_non_normal) '(non-normal)' else '(normal)', '\\n')
  } else {
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}
if (all_normal) {
  cat('\\nℹ️ NOTE: All groups appear normally distributed. Non-parametric test (Steel) was selected manually.\\n')
  cat('   Parametric test (Dunnett) may have more statistical power for normal data.\\n')
}
cat('\\n')

# Kruskal-Wallis overall test (non-parametric alternative to ANOVA) - proceeding as user selected
kw_result <- kruskal.test(${yColName} ~ ${xColName}, data = dat)
cat('Kruskal-Wallis test: p =', kw_result$p.value, '\\n\\n')

# Initialize results storage
stat_results <- data.frame(
  category1 = character(),
  category2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

# Get control group data
control_data <- dat[dat$${xColName} == '${controlGroup}', '${yColName}']

# Compare each treatment to control using Steel test
treatment_groups <- c(${uniqueCategories.filter(c => c !== controlGroup).map(c => `'${c}'`).join(', ')})

for (trt in treatment_groups) {
  trt_data <- dat[dat$${xColName} == trt, '${yColName}']

  # Steel test compares two samples - p-value is in st[2]
  steel_result <- Steel.test(list(control_data, trt_data))
  p_val <- steel_result$st[2]

  # Determine significance
  sig <- if (p_val < 0.001) '***' else if (p_val < 0.01) '**' else if (p_val < 0.05) '*' else 'ns'

  stat_results <- rbind(stat_results, data.frame(
    category1 = '${controlGroup}',
    category2 = trt,
    p_value = p_val,
    significance = sig
  ))

  cat('${controlGroup} vs', trt, ': Steel test p =', p_val, '(', sig, ')\\n')
}

# Display results
print(stat_results)

# Add significance brackets to plot using ggpubr
`;
    } else {
      // Dunnett test (parametric vs control)
      code += `# Install and load multcomp for Dunnett test
if (!requireNamespace('multcomp', quietly = TRUE)) {
  install.packages('multcomp')
}
library(multcomp)

# Dunnett test (parametric vs control)
# Control group: ${controlGroup}

# Set category order to match plot
dat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk):\\n')
any_non_normal <- FALSE
for (grp in unique(dat$${xColName})) {
  grp_data <- dat[dat$${xColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    is_non_normal <- sw_result$p.value < 0.05
    if (is_non_normal) any_non_normal <- TRUE
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(is_non_normal) '(non-normal)' else '(normal)', '\\n')
  } else {
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}
if (any_non_normal) {
  cat('\\n⚠️ WARNING: Some groups appear non-normal. Parametric test (Dunnett) was selected manually.\\n')
  cat('   Consider using non-parametric test (Steel) for more robust results.\\n')
}
cat('\\n')

# ANOVA overall test (parametric) - proceeding as user selected
aov_result <- aov(${yColName} ~ ${xColName}, data = dat)
anova_p <- summary(aov_result)[[1]][['Pr(>F)']][1]
cat('ANOVA: p =', anova_p, '\\n\\n')

# Initialize results storage
stat_results <- data.frame(
  category1 = character(),
  category2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

# Set control group as reference level
dat$${xColName} <- relevel(factor(dat$${xColName}), ref = '${controlGroup}')

# Run Dunnett test
dunnett_result <- glht(aov(${yColName} ~ ${xColName}, data = dat), linfct = mcp(${xColName} = 'Dunnett'))
dunnett_summary <- summary(dunnett_result)

# Extract results
for (i in seq_along(dunnett_summary$test$coefficients)) {
  comp_name <- names(dunnett_summary$test$coefficients)[i]
  p_val <- dunnett_summary$test$pvalues[i]

  # Parse comparison name (format: "Treatment - Control")
  grps <- strsplit(comp_name, ' - ')[[1]]
  trt <- trimws(grps[1])

  # Determine significance
  sig <- if (p_val < 0.001) '***' else if (p_val < 0.01) '**' else if (p_val < 0.05) '*' else 'ns'

  stat_results <- rbind(stat_results, data.frame(
    category1 = '${controlGroup}',
    category2 = trt,
    p_value = p_val,
    significance = sig
  ))

  cat('${controlGroup} vs', trt, ': Dunnett test p =', p_val, '(', sig, ')\\n')
}

# Display results
print(stat_results)

# Add significance brackets to plot using ggpubr
`;
    }
  } else if (uniqueCategories.length >= 3 && (postHocTest === 'bonferroni' || postHocTest === 'holm')) {
    // Bonferroni or Holm correction for 3+ groups (parametric pairwise with multiple testing correction)
    // Generate pairs for all pairwise comparisons
    for (let i = 0; i < uniqueCategories.length; i++) {
      for (let j = i + 1; j < uniqueCategories.length; j++) {
        pairs.push([uniqueCategories[i], uniqueCategories[j]]);
      }
    }

    const pAdjustMethod = postHocTest; // 'bonferroni' or 'holm'

    // Generate factor level order string for R
    const categoryOrder = uniqueCategories.map(c => `'${c}'`).join(', ');

    code += `# Parametric pairwise comparisons with ${postHocTest.charAt(0).toUpperCase() + postHocTest.slice(1)} correction
# For multiple testing correction across all pairwise comparisons

# Set category order to match plot
dat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk):\\n')
any_non_normal <- FALSE
for (grp in unique(dat$${xColName})) {
  grp_data <- dat[dat$${xColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    is_non_normal <- sw_result$p.value < 0.05
    if (is_non_normal) any_non_normal <- TRUE
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(is_non_normal) '(non-normal)' else '(normal)', '\\n')
  } else {
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}
if (any_non_normal) {
  cat('\\n\u26a0\ufe0f WARNING: Some groups appear non-normal. Parametric test with ${postHocTest.charAt(0).toUpperCase() + postHocTest.slice(1)} correction was selected.\\n')
  cat('   Consider using non-parametric test (Dunn) for more robust results.\\n')
}
cat('\\n')

# ANOVA overall test (parametric)
aov_result <- aov(${yColName} ~ ${xColName}, data = dat)
anova_p <- summary(aov_result)[[1]][['Pr(>F)']][1]
cat('ANOVA: p =', format(anova_p, digits = 4), '\\n\\n')

# Pairwise t-tests with ${postHocTest.charAt(0).toUpperCase() + postHocTest.slice(1)} correction
cat('Pairwise comparisons (${postHocTest.charAt(0).toUpperCase() + postHocTest.slice(1)} correction):\\n')
pairwise_result <- pairwise.t.test(dat$${yColName}, dat$${xColName},
                                    p.adjust.method = '${pAdjustMethod}')

# Initialize results storage
stat_results <- data.frame(
  category1 = character(),
  category2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

# Extract p-values from the matrix
p_matrix <- pairwise_result$p.value
for (i in 1:nrow(p_matrix)) {
  for (j in 1:ncol(p_matrix)) {
    if (!is.na(p_matrix[i, j])) {
      grp1 <- colnames(p_matrix)[j]
      grp2 <- rownames(p_matrix)[i]
      p_val <- p_matrix[i, j]

      # Determine significance
      sig <- if (p_val < 0.001) '***' else if (p_val < 0.01) '**' else if (p_val < 0.05) '*' else 'ns'

      stat_results <- rbind(stat_results, data.frame(
        category1 = grp1,
        category2 = grp2,
        p_value = p_val,
        significance = sig,
        stringsAsFactors = FALSE
      ))

      cat('  ', grp1, 'vs', grp2, ': p =', format(p_val, digits = 4), '(', sig, ')\\n')
    }
  }
}

cat('\\n')
print(stat_results)

# Add significance brackets to plot using ggpubr
`;
  } else if (uniqueCategories.length >= 3 && postHocTest === 'tukey') {
    // Tukey HSD for 3+ groups (ANOVA + TukeyHSD)
    // Generate pairs for all pairwise comparisons
    for (let i = 0; i < uniqueCategories.length; i++) {
      for (let j = i + 1; j < uniqueCategories.length; j++) {
        pairs.push([uniqueCategories[i], uniqueCategories[j]]);
      }
    }

    // Generate factor level order string for R
    const categoryOrder = uniqueCategories.map(c => `'${c}'`).join(', ');

    code += `# ANOVA with Tukey HSD post-hoc test
# For 3+ group comparisons

# Set category order to match plot
dat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk):\\n')
any_non_normal <- FALSE
for (grp in unique(dat$${xColName})) {
  grp_data <- dat[dat$${xColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    is_non_normal <- sw_result$p.value < 0.05
    if (is_non_normal) any_non_normal <- TRUE
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(is_non_normal) '(non-normal)' else '(normal)', '\\n')
  } else {
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}
if (any_non_normal) {
  cat('\\n⚠️ WARNING: Some groups appear non-normal. Parametric test (ANOVA/Tukey) was selected.\\n')
  cat('   Consider using non-parametric test (Kruskal-Wallis/Dunn) for more robust results.\\n')
}
cat('\\n')

# ANOVA overall test
aov_result <- aov(${yColName} ~ ${xColName}, data = dat)
anova_p <- summary(aov_result)[[1]][['Pr(>F)']][1]
cat('ANOVA: p =', format(anova_p, digits = 4), '\\n')

if (anova_p >= 0.05) {
  cat('ANOVA not significant (p >= 0.05). No post-hoc comparisons needed.\\n\\n')
}

# Tukey HSD post-hoc test
cat('\\nTukey HSD post-hoc comparisons:\\n')
tukey_result <- TukeyHSD(aov_result)

# Initialize results storage
stat_results <- data.frame(
  category1 = character(),
  category2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

# Extract results from Tukey HSD
tukey_df <- as.data.frame(tukey_result[['${xColName}']])
tukey_df[['comparison']] <- rownames(tukey_df)

for (i in 1:nrow(tukey_df)) {
  comp <- tukey_df[['comparison']][i]
  p_val <- tukey_df[['p adj']][i]

  # Parse comparison name (format: "Group2-Group1")
  grps <- strsplit(comp, '-')[[1]]
  grp1 <- trimws(grps[2])  # Second group (after -)
  grp2 <- trimws(grps[1])  # First group (before -)

  # Determine significance
  sig <- if (p_val < 0.001) '***' else if (p_val < 0.01) '**' else if (p_val < 0.05) '*' else 'ns'

  stat_results <- rbind(stat_results, data.frame(
    category1 = grp1,
    category2 = grp2,
    p_value = p_val,
    significance = sig,
    stringsAsFactors = FALSE
  ))

  cat('  ', grp1, 'vs', grp2, ': p =', format(p_val, digits = 4), '(', sig, ')\\n')
}

cat('\\n')
print(stat_results)

# Add significance brackets to plot using ggpubr
`;
  } else if (uniqueCategories.length >= 3 && postHocTest === 'dunn') {
    // Dunn test for 3+ groups (Kruskal-Wallis + Dunn)
    // Generate pairs for all pairwise comparisons
    for (let i = 0; i < uniqueCategories.length; i++) {
      for (let j = i + 1; j < uniqueCategories.length; j++) {
        pairs.push([uniqueCategories[i], uniqueCategories[j]]);
      }
    }

    // Generate factor level order string for R
    const categoryOrder = uniqueCategories.map(c => `'${c}'`).join(', ');

    code += `# Install and load dunn.test for Dunn test
if (!requireNamespace('dunn.test', quietly = TRUE)) {
  install.packages('dunn.test')
}
library(dunn.test)

# Kruskal-Wallis with Dunn post-hoc test (non-parametric)
# For 3+ group comparisons

# Set category order to match plot
dat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk):\\n')
all_normal <- TRUE
for (grp in unique(dat$${xColName})) {
  grp_data <- dat[dat$${xColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    is_non_normal <- sw_result$p.value < 0.05
    if (is_non_normal) all_normal <- FALSE
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(is_non_normal) '(non-normal)' else '(normal)', '\\n')
  } else {
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}
if (all_normal) {
  cat('\\nℹ️ NOTE: All groups appear normally distributed. Non-parametric test (Dunn) was selected.\\n')
  cat('   Parametric test (ANOVA/Tukey) may have more statistical power for normal data.\\n')
}
cat('\\n')

# Kruskal-Wallis overall test
kw_result <- kruskal.test(${yColName} ~ ${xColName}, data = dat)
cat('Kruskal-Wallis: p =', format(kw_result$p.value, digits = 4), '\\n')

if (kw_result$p.value >= 0.05) {
  cat('Kruskal-Wallis not significant (p >= 0.05). No post-hoc comparisons needed.\\n\\n')
}

# Dunn post-hoc test
cat('\\nDunn post-hoc comparisons:\\n')
dunn_result <- dunn.test(dat$${yColName}, dat$${xColName}, method = 'bonferroni', kw = FALSE)

# Initialize results storage
stat_results <- data.frame(
  category1 = character(),
  category2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

# Extract results from Dunn test
for (i in seq_along(dunn_result$comparisons)) {
  comp <- dunn_result$comparisons[i]
  p_val <- dunn_result$P.adjusted[i]

  # Parse comparison name (format: "Group1 - Group2")
  grps <- strsplit(comp, ' - ')[[1]]
  grp1 <- trimws(grps[1])
  grp2 <- trimws(grps[2])

  # Determine significance
  sig <- if (p_val < 0.001) '***' else if (p_val < 0.01) '**' else if (p_val < 0.05) '*' else 'ns'

  stat_results <- rbind(stat_results, data.frame(
    category1 = grp1,
    category2 = grp2,
    p_value = p_val,
    significance = sig,
    stringsAsFactors = FALSE
  ))

  cat('  ', grp1, 'vs', grp2, ': p =', format(p_val, digits = 4), '(', sig, ')\\n')
}

cat('\\n')
print(stat_results)

# Add significance brackets to plot using ggpubr
`;
  } else {
    // Original pairwise approach (for 2 groups only)
    // Generate factor level order string for R
    const categoryOrder = uniqueCategories.map(c => `'${c}'`).join(', ');

    // Add normality testing with warnings for manual mode
    if (isNonParametric) {
      code += `# Non-parametric pairwise comparisons (Wilcoxon test)
# Selected manually by user

# Set category order to match plot
dat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk):\\n')
all_normal <- TRUE
for (grp in unique(dat$${xColName})) {
  grp_data <- dat[dat$${xColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    is_non_normal <- sw_result$p.value < 0.05
    if (is_non_normal) all_normal <- FALSE
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(is_non_normal) '(non-normal)' else '(normal)', '\\n')
  } else {
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}
if (all_normal) {
  cat('\\nℹ️ NOTE: All groups appear normally distributed. Non-parametric test (Wilcoxon) was selected manually.\\n')
  cat('   Parametric test (t-test) may have more statistical power for normal data.\\n')
}
cat('\\n')

# Initialize results storage
stat_results <- data.frame(
  category1 = character(),
  category2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

`;
    } else {
      // Parametric or auto mode
      code += `# Parametric pairwise comparisons (t-test)
# Selected manually by user

# Set category order to match plot
dat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk):\\n')
any_non_normal <- FALSE
for (grp in unique(dat$${xColName})) {
  grp_data <- dat[dat$${xColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    is_non_normal <- sw_result$p.value < 0.05
    if (is_non_normal) any_non_normal <- TRUE
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(is_non_normal) '(non-normal)' else '(normal)', '\\n')
  } else {
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}
if (any_non_normal) {
  cat('\\n⚠️ WARNING: Some groups appear non-normal. Parametric test (t-test) was selected manually.\\n')
  cat('   Consider using non-parametric test (Wilcoxon) for more robust results.\\n')
}
cat('\\n')

# Initialize results storage
stat_results <- data.frame(
  category1 = character(),
  category2 = character(),
  p_value = numeric(),
  significance = character(),
  stringsAsFactors = FALSE
)

`;
    }

    // Generate pairwise comparisons based on comparison mode
    let pairs = [];

    if (settings.comparisonMode === 'custom' && settings.customComparisons) {
      // Parse custom comparisons
      let customComps = [];
      if (typeof settings.customComparisons === 'string') {
        try {
          customComps = JSON.parse(settings.customComparisons);
        } catch (e) {
          console.warn('Failed to parse customComparisons:', e);
        }
      } else if (Array.isArray(settings.customComparisons)) {
        customComps = settings.customComparisons;
      }

      // For single-group charts, custom comparisons are just category pairs (not Group@Category format)
      // Extract unique pairs
      customComps.forEach(comp => {
        // Handle both "Cat1-Cat2" and "Cat1-Cat2@X" formats
        const pairPart = comp.split('@')[0];
        const [cat1, cat2] = pairPart.split('-').map(c => c.trim());
        if (cat1 && cat2 && uniqueCategories.includes(cat1) && uniqueCategories.includes(cat2)) {
          // Avoid duplicates
          const alreadyExists = pairs.some(([p1, p2]) =>
            (p1 === cat1 && p2 === cat2) || (p1 === cat2 && p2 === cat1)
          );
          if (!alreadyExists) {
            pairs.push([cat1, cat2]);
          }
        }
      });

      // If no valid custom comparisons found, fall back to all pairs
      if (pairs.length === 0) {
        for (let i = 0; i < uniqueCategories.length; i++) {
          for (let j = i + 1; j < uniqueCategories.length; j++) {
            pairs.push([uniqueCategories[i], uniqueCategories[j]]);
          }
        }
      }
    } else {
      // For 'all' or 'significant' mode, generate all pairwise comparisons
      for (let i = 0; i < uniqueCategories.length; i++) {
        for (let j = i + 1; j < uniqueCategories.length; j++) {
          pairs.push([uniqueCategories[i], uniqueCategories[j]]);
        }
      }
    }

    // Generate test code for each pair
    pairs.forEach(([cat1, cat2]) => {
      const safeName1 = cat1.replace(/[^a-zA-Z0-9]/g, '_');
      const safeName2 = cat2.replace(/[^a-zA-Z0-9]/g, '_');
      const pairName = `${safeName1}_vs_${safeName2}`;

      if (isNonParametric) {
        // Non-parametric: always use Wilcoxon
        code += `# Test: ${cat1} vs ${cat2} (Non-parametric: Wilcoxon)
cat1_data <- dat[dat$${xColName} == '${cat1}', '${yColName}']
cat2_data <- dat[dat$${xColName} == '${cat2}', '${yColName}']

test_result_${pairName} <- wilcox.test(cat1_data, cat2_data)
test_name_${pairName} <- 'Wilcoxon test'
p_val_${pairName} <- test_result_${pairName}$p.value

# Determine significance symbol
sig_${pairName} <- if (p_val_${pairName} < 0.001) '***' else if (p_val_${pairName} < 0.01) '**' else if (p_val_${pairName} < 0.05) '*' else 'ns'

# Store result
stat_results <- rbind(stat_results, data.frame(
  category1 = '${cat1}',
  category2 = '${cat2}',
  p_value = p_val_${pairName},
  significance = sig_${pairName}
))

cat('${cat1} vs ${cat2}:', test_name_${pairName}, 'p =', p_val_${pairName}, '(', sig_${pairName}, ')\\n')

`;
      } else if (isManualParametric) {
        // Manual parametric mode: force t-test regardless of normality
        code += `# Test: ${cat1} vs ${cat2} (Parametric: t-test - manually selected)
cat1_data <- dat[dat$${xColName} == '${cat1}', '${yColName}']
cat2_data <- dat[dat$${xColName} == '${cat2}', '${yColName}']

# Test for equal variances (Levene test)
combined_data <- data.frame(
  values = c(cat1_data, cat2_data),
  group = factor(c(rep('${cat1}', length(cat1_data)), rep('${cat2}', length(cat2_data))))
)
group_means <- tapply(combined_data$values, combined_data$group, mean)
abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
levene_result <- anova(lm(abs_deviations ~ combined_data$group))
levene_p_${pairName} <- levene_result[['Pr(>F)']][1]
equal_var_${pairName} <- levene_p_${pairName} >= 0.05
cat('  Variance test (Levene) for ${cat1} vs ${cat2}: p =', levene_p_${pairName}, if(equal_var_${pairName}) '(equal variances)' else '(unequal variances)', '\\n')

# Use t-test (parametric - manually selected)
test_result_${pairName} <- t.test(cat1_data, cat2_data, var.equal = equal_var_${pairName})
test_name_${pairName} <- if (equal_var_${pairName}) 'Student\\'s t-test' else 'Welch\\'s t-test'
p_val_${pairName} <- test_result_${pairName}$p.value

# Determine significance symbol
sig_${pairName} <- if (p_val_${pairName} < 0.001) '***' else if (p_val_${pairName} < 0.01) '**' else if (p_val_${pairName} < 0.05) '*' else 'ns'

# Store result
stat_results <- rbind(stat_results, data.frame(
  category1 = '${cat1}',
  category2 = '${cat2}',
  p_value = p_val_${pairName},
  significance = sig_${pairName}
))

cat('${cat1} vs ${cat2}:', test_name_${pairName}, 'p =', p_val_${pairName}, '(', sig_${pairName}, ')\\n')

`;
      } else {
        // Auto-detect mode: test normality per pair and choose appropriate test
        code += `# Test: ${cat1} vs ${cat2} (Auto-detect mode)
cat1_data <- dat[dat$${xColName} == '${cat1}', '${yColName}']
cat2_data <- dat[dat$${xColName} == '${cat2}', '${yColName}']

# Test for normality (Shapiro-Wilk test)
shapiro1_${pairName} <- shapiro.test(cat1_data)
shapiro2_${pairName} <- shapiro.test(cat2_data)
both_normal_${pairName} <- (shapiro1_${pairName}$p.value >= 0.05) && (shapiro2_${pairName}$p.value >= 0.05)

cat('  Normality test for ${cat1} vs ${cat2}:\\n')
cat('    ${cat1}: p =', shapiro1_${pairName}$p.value, if(shapiro1_${pairName}$p.value >= 0.05) '(normal)' else '(non-normal)', '\\n')
cat('    ${cat2}: p =', shapiro2_${pairName}$p.value, if(shapiro2_${pairName}$p.value >= 0.05) '(normal)' else '(non-normal)', '\\n')

# Test for equal variances (Levene test)
if (both_normal_${pairName}) {
  combined_data <- data.frame(
    values = c(cat1_data, cat2_data),
    group = factor(c(rep('${cat1}', length(cat1_data)), rep('${cat2}', length(cat2_data))))
  )
  group_means <- tapply(combined_data$values, combined_data$group, mean)
  abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
  levene_result <- anova(lm(abs_deviations ~ combined_data$group))
  levene_p_${pairName} <- levene_result[['Pr(>F)']][1]
  equal_var_${pairName} <- levene_p_${pairName} >= 0.05
  cat('  Variance test (Levene): p =', levene_p_${pairName}, if(equal_var_${pairName}) '(equal variances)' else '(unequal variances)', '\\n')
} else {
  equal_var_${pairName} <- TRUE  # Default for non-parametric tests
}

# Perform appropriate test based on normality
if (both_normal_${pairName}) {
  # Use t-test with appropriate var.equal parameter
  test_result_${pairName} <- t.test(cat1_data, cat2_data, var.equal = equal_var_${pairName})
  test_name_${pairName} <- if (equal_var_${pairName}) 'Student\\'s t-test' else 'Welch\\'s t-test'
} else {
  # Use Wilcoxon test for non-normal data
  test_result_${pairName} <- wilcox.test(cat1_data, cat2_data)
  test_name_${pairName} <- 'Wilcoxon test'
}

p_val_${pairName} <- test_result_${pairName}$p.value

# Determine significance symbol
sig_${pairName} <- if (p_val_${pairName} < 0.001) '***' else if (p_val_${pairName} < 0.01) '**' else if (p_val_${pairName} < 0.05) '*' else 'ns'

# Store result
stat_results <- rbind(stat_results, data.frame(
  category1 = '${cat1}',
  category2 = '${cat2}',
  p_value = p_val_${pairName},
  significance = sig_${pairName}
))

cat('${cat1} vs ${cat2}:', test_name_${pairName}, 'p =', p_val_${pairName}, '(', sig_${pairName}, ')\\n')

`;
      }
    });

    code += `# Display results
print(stat_results)

# Add significance brackets to plot using ggpubr
`;
  }

  // Generate ggpubr comparison data frame
  const yColIndex = settings.yColIndex - 1;
  const xColIndex = settings.xColIndex - 1;
  const allYValues = window.lastProcessedData
    .slice(1)
    .map(row => parseFloat(row[yColIndex]))
    .filter(v => !isNaN(v));

  const yMax = allYValues.length > 0 ? Math.max(...allYValues) : 10;
  const yRange = yMax - Math.min(...allYValues);

  // Parse custom positions if available
  let customPositions = {};
  if (settings.customPositions) {
    if (typeof settings.customPositions === 'string') {
      try {
        customPositions = JSON.parse(settings.customPositions);
      } catch (e) {
        console.warn('Failed to parse customPositions:', e);
      }
    } else if (typeof settings.customPositions === 'object') {
      customPositions = settings.customPositions;
    }
  }

  // Build ggpubr comparison data frame
  // Check if we have bracket data from add-in
  const hasAddInBrackets = window.lastBracketData && window.lastBracketData.length > 0;

  // Determine if we use stat_results (vs-control or Bonferroni/Holm)
  const useStatResults = (uniqueCategories.length >= 3 && isVsControl) ||
                         (uniqueCategories.length >= 3 && (postHocTest === 'bonferroni' || postHocTest === 'holm' || postHocTest === 'tukey' || postHocTest === 'dunn'));

  // For tests that use stat_results (Steel/Dunnett, Bonferroni/Holm, Tukey, Dunn), build comparison_df from stat_results
  if (useStatResults) {
    // Calculate Y positions for each comparison
    const yPositions = pairs.map(([cat1, cat2], pairIdx) => {
      let yBracketPos;
      if (hasAddInBrackets && pairIdx < window.lastBracketData.length) {
        yBracketPos = window.lastBracketData[pairIdx]["y.position"];
        yBracketPos = typeof yBracketPos === 'number' ? yBracketPos.toFixed(2) : yBracketPos;
      } else {
        const compKey1 = `${cat1}-${cat2}`;
        const compKey2 = `${cat2}-${cat1}`;
        const customYPos = customPositions[compKey1] || customPositions[compKey2];
        if (customYPos !== undefined && customYPos !== null) {
          yBracketPos = Number(customYPos).toFixed(2);
        } else {
          const bracketLevel = pairIdx + 1;
          yBracketPos = (yMax + yRange * (0.05 + bracketLevel * 0.1)).toFixed(2);
        }
      }
      return yBracketPos;
    });

    const showAll = settings.comparisonMode === 'all' || settings.comparisonMode === 'custom';

    code += `# Build comparison_df from stat_results for ggpubr
# Y positions for brackets
y_positions <- c(${yPositions.join(', ')})

comparison_df <- data.frame(
  group1 = stat_results$category1,
  group2 = stat_results$category2,
  p.signif = stat_results$significance,
  y.position = y_positions[1:nrow(stat_results)],
  stringsAsFactors = FALSE
)

${showAll ? '' : `# Filter to only significant comparisons
comparison_df <- comparison_df[comparison_df$p.signif != 'ns', ]
`}
`;
  } else {
    // Original pairwise approach with individual variables
    code += `# Prepare comparison data for ggpubr
${hasAddInBrackets ? '# Note: Y positions below are the exact values from the add-in preview\n' : ''}comparison_df <- data.frame(
  group1 = character(),
  group2 = character(),
  p.signif = character(),
  y.position = numeric(),
  stringsAsFactors = FALSE
)

`;

    pairs.forEach(([cat1, cat2], pairIdx) => {
      const safeName1 = cat1.replace(/[^a-zA-Z0-9]/g, '_');
      const safeName2 = cat2.replace(/[^a-zA-Z0-9]/g, '_');
      const pairName = `${safeName1}_vs_${safeName2}`;

      // Calculate Y position: prioritize add-in data, then custom, then automatic
      let yBracketPos;

      // First priority: Use exact position from add-in if available
      if (hasAddInBrackets && pairIdx < window.lastBracketData.length) {
        yBracketPos = window.lastBracketData[pairIdx]["y.position"];
        console.log(`📍 Using exact bracket position from add-in for ${cat1} vs ${cat2}: y=${yBracketPos}`);
        yBracketPos = typeof yBracketPos === 'number' ? yBracketPos.toFixed(2) : yBracketPos;
      } else {
        // Second priority: Check for custom Y position
        const compKey1 = `${cat1}-${cat2}`;
        const compKey2 = `${cat2}-${cat1}`;
        const customYPos = customPositions[compKey1] || customPositions[compKey2];

        if (customYPos !== undefined && customYPos !== null) {
          yBracketPos = Number(customYPos).toFixed(2);
        } else {
          // Last resort: Calculate automatically (stack multiple brackets)
          const bracketLevel = pairIdx + 1;
          yBracketPos = (yMax + yRange * (0.05 + bracketLevel * 0.1)).toFixed(2);
        }
      }

      const showAll = settings.comparisonMode === 'all' || settings.comparisonMode === 'custom';

      if (showAll) {
        // Show all comparisons
        code += `comparison_df <- rbind(comparison_df, data.frame(
  group1 = '${cat1}',
  group2 = '${cat2}',
  p.signif = sig_${pairName},
  y.position = ${yBracketPos}
))
`;
      } else {
        // Show only significant comparisons
        code += `if (sig_${pairName} != 'ns') {
  comparison_df <- rbind(comparison_df, data.frame(
    group1 = '${cat1}',
    group2 = '${cat2}',
    p.signif = sig_${pairName},
    y.position = ${yBracketPos}
  ))
}
`;
      }
    });
  }

  // Add the ggpubr stat_pvalue_manual call
  code += `
# Add comparison brackets using ggpubr
if (nrow(comparison_df) > 0) {
  p <- p + stat_pvalue_manual(
    comparison_df,
    label = 'p.signif',
    y.position = 'y.position',
    size = ${settings.statSymbolSize || 7},
    size.line = ${settings.statLineSize || 1.0},
    tip.length = ${settings.statTipLength || 0.04}
  )
}

`;

  // Add appropriate note based on comparison mode
  if (settings.comparisonMode === 'all') {
    code += `# Note: All comparisons are shown on the plot (including non-significant 'ns').
`;
  } else if (settings.comparisonMode === 'custom') {
    const hasCustomPos = Object.keys(customPositions).length > 0;
    code += `# Note: Only selected custom comparisons are shown on the plot.`;
    if (hasCustomPos) {
      code += `
# Custom Y positions for brackets are included from your settings.`;
    }
    code += `

`;
  } else {
    code += `
# Note: Only significant comparisons (p < 0.05) are shown on the plot.
# To show all comparisons including 'ns', remove the if condition above.

`;
  }

  return code;
}

// Generate code for grouped box plots
function generateGroupedBoxCode(chartType, settings, groupColors) {
  const hasDots = chartType === 'box_grouped_dot';

  // Use col1, col2, col3 naming to match data frame creation
  const groupColName = `col${settings.groupColIndex}`;
  const xColName = `col${settings.xColIndex}`;
  const yColName = `col${settings.yColIndex}`;

  let code = '';

  // Add factor ordering if custom order is set
  // Get actual unique values from data to validate and use correct values
  let actualCategories = [];
  let actualGroups = [];

  if (window.lastProcessedData && window.lastProcessedData.length > 1) {
    const data = window.lastProcessedData;
    const headers = data[0];
    const catIdx = settings.xColIndex - 1;  // Convert from 1-based to 0-based
    const grpIdx = settings.groupColIndex - 1;

    if (catIdx >= 0 && catIdx < headers.length) {
      actualCategories = [...new Set(data.slice(1).map(row => row[catIdx]).filter(v => v))];
    }
    if (grpIdx >= 0 && grpIdx < headers.length) {
      actualGroups = [...new Set(data.slice(1).map(row => row[grpIdx]).filter(v => v))];
    }
  }

  // Handle category ordering based on dataOrder mode
  if (actualCategories.length > 0) {
    let finalCategories = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let categories = [];
      if (settings.customOrderCategory) {
        if (typeof settings.customOrderCategory === 'string') {
          categories = settings.customOrderCategory.split(',').map(c => c.trim()).filter(c => c);
        } else if (Array.isArray(settings.customOrderCategory)) {
          categories = settings.customOrderCategory;
        }
      }

      // Validate: only use custom order if values match actual categories
      const isValidCategoryOrder = categories.length > 0 &&
                                   categories.every(c => actualCategories.includes(c));

      if (isValidCategoryOrder) {
        finalCategories = categories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set custom category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalCategories = actualCategories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set category order from data\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort categories alphabetically
      finalCategories = [...actualCategories].sort((a, b) => String(a).localeCompare(String(b)));
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set alphabetical category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalCategories = actualCategories;
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set category order from data (original order)\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    }
  }

  // Handle group ordering based on dataOrder mode
  if (actualGroups.length > 0) {
    let finalGroups = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let groups = [];
      if (settings.customOrderGroup) {
        if (typeof settings.customOrderGroup === 'string') {
          groups = settings.customOrderGroup.split(',').map(g => g.trim()).filter(g => g);
        } else if (Array.isArray(settings.customOrderGroup)) {
          groups = settings.customOrderGroup;
        }
      }

      // Validate: only use custom order if values match actual groups
      const isValidGroupOrder = groups.length > 0 &&
                                groups.every(g => actualGroups.includes(g));

      if (isValidGroupOrder) {
        finalGroups = groups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set custom group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalGroups = actualGroups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set group order from data\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort groups alphabetically
      finalGroups = [...actualGroups].sort((a, b) => String(a).localeCompare(String(b)));
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set alphabetical group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalGroups = actualGroups;
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set group order from data (original order)\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    }
  }

  code += `# ============ Create Plot ============

# Create the base plot
p <- ggplot(dat, aes(x = ${xColName}, y = ${yColName}, fill = ${groupColName})) +

  # Add box plots
  geom_boxplot(
    position = position_dodge(width = ${settings.dodgeWidth}),
    width = ${settings.barWidth},
    alpha = ${settings.fillAlpha},
    linewidth = ${settings.lineWidth}
  ) +

`;

  if (hasDots) {
    code += `  # Add individual data points
  geom_point(
    position = position_jitterdodge(
      dodge.width = ${settings.dodgeWidth},
      jitter.width = 0.15,
      jitter.height = 0
    ),
    size = ${settings.dotSize},
    alpha = ${settings.dotAlpha},
    color = '${settings.dotColor}',
    shape = ${settings.dotShape}
  ) +

`;
  }

  code += `  # Set colors
  scale_fill_manual(values = ${groupColors}, name = '${settings.selectedGroupColumn || 'Group'}') +

  # Apply theme
  theme_${settings.themeName}(base_family = '${settings.fontFamily}') +

  # Customize theme elements
  theme(
    plot.background = element_rect(fill = 'white', color = NA),
    panel.background = element_rect(fill = 'white', color = NA),
    panel.grid.major.y = element_line(color = '#e5e7eb', linewidth = 0.6),
    panel.grid.minor = element_blank(),
    axis.line = element_line(linewidth = 0.6),
    axis.ticks = element_line(linewidth = 0.5),
    plot.title = element_text(size = ${settings.titleSize}, face = '${settings.titleWeight}', hjust = 0.5),
    axis.title.x = element_text(size = ${settings.xAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.title.y = element_text(size = ${settings.yAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.text.x = element_text(size = ${settings.xAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    axis.text.y = element_text(size = ${settings.yAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    legend.text = element_text(size = ${settings.legendTextSize}),
    legend.title = element_text(size = ${settings.legendTextSize})
  )

# Add labels
`;

  if (settings.showTitle) {
    code += `p <- p + ggtitle(${convertToRPlotmath(settings.title)})\n`;
  }
  if (settings.showXLabel) {
    code += `p <- p + xlab(${convertToRPlotmath(settings.xlab)})\n`;
  }
  if (settings.showYLabel) {
    code += `p <- p + ylab(${convertToRPlotmath(settings.ylab)})\n`;
  }

  // Add statistical comparisons using helper function
  code += generateStatisticalComparisonCode(settings, groupColName, xColName);

  code += `\n# Display the plot
print(p)
`;

  return code;
}

// Generate code for grouped violin plots
function generateGroupedViolinCode(chartType, settings, groupColors) {
  const hasDots = chartType === 'violin_grouped_dot';

  // Use col1, col2, col3 naming to match data frame creation
  const groupColName = `col${settings.groupColIndex}`;
  const xColName = `col${settings.xColIndex}`;
  const yColName = `col${settings.yColIndex}`;

  let code = '';

  // Add factor ordering if custom order is set
  // Get actual unique values from data to validate and use correct values
  let actualCategories = [];
  let actualGroups = [];

  if (window.lastProcessedData && window.lastProcessedData.length > 1) {
    const data = window.lastProcessedData;
    const headers = data[0];
    const catIdx = settings.xColIndex - 1;  // Convert from 1-based to 0-based
    const grpIdx = settings.groupColIndex - 1;

    if (catIdx >= 0 && catIdx < headers.length) {
      actualCategories = [...new Set(data.slice(1).map(row => row[catIdx]).filter(v => v))];
    }
    if (grpIdx >= 0 && grpIdx < headers.length) {
      actualGroups = [...new Set(data.slice(1).map(row => row[grpIdx]).filter(v => v))];
    }
  }

  // Handle category ordering based on dataOrder mode
  if (actualCategories.length > 0) {
    let finalCategories = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let categories = [];
      if (settings.customOrderCategory) {
        if (typeof settings.customOrderCategory === 'string') {
          categories = settings.customOrderCategory.split(',').map(c => c.trim()).filter(c => c);
        } else if (Array.isArray(settings.customOrderCategory)) {
          categories = settings.customOrderCategory;
        }
      }

      // Validate: only use custom order if values match actual categories
      const isValidCategoryOrder = categories.length > 0 &&
                                   categories.every(c => actualCategories.includes(c));

      if (isValidCategoryOrder) {
        finalCategories = categories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set custom category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalCategories = actualCategories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set category order from data\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort categories alphabetically
      finalCategories = [...actualCategories].sort((a, b) => String(a).localeCompare(String(b)));
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set alphabetical category order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalCategories = actualCategories;
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set category order from data (original order)\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    }
  }

  // Handle group ordering based on dataOrder mode
  if (actualGroups.length > 0) {
    let finalGroups = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let groups = [];
      if (settings.customOrderGroup) {
        if (typeof settings.customOrderGroup === 'string') {
          groups = settings.customOrderGroup.split(',').map(g => g.trim()).filter(g => g);
        } else if (Array.isArray(settings.customOrderGroup)) {
          groups = settings.customOrderGroup;
        }
      }

      // Validate: only use custom order if values match actual groups
      const isValidGroupOrder = groups.length > 0 &&
                                groups.every(g => actualGroups.includes(g));

      if (isValidGroupOrder) {
        finalGroups = groups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set custom group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalGroups = actualGroups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set group order from data\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort groups alphabetically
      finalGroups = [...actualGroups].sort((a, b) => String(a).localeCompare(String(b)));
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set alphabetical group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalGroups = actualGroups;
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set group order from data (original order)\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    }
  }

  code += `# ============ Create Plot ============

# Create the base plot
p <- ggplot(dat, aes(x = ${xColName}, y = ${yColName}, fill = ${groupColName})) +

  # Add violin plots
  geom_violin(
    position = position_dodge(width = ${settings.dodgeWidth}),
    width = ${settings.barWidth},
    alpha = ${settings.fillAlpha},
    linewidth = ${settings.lineWidth}
  ) +

`;

  if (hasDots) {
    code += `  # Add individual data points
  geom_point(
    position = position_jitterdodge(
      dodge.width = ${settings.dodgeWidth},
      jitter.width = 0.15,
      jitter.height = 0
    ),
    size = ${settings.dotSize},
    alpha = ${settings.dotAlpha},
    color = '${settings.dotColor}',
    shape = ${settings.dotShape}
  ) +

`;
  }

  code += `  # Set colors
  scale_fill_manual(values = ${groupColors}, name = '${settings.selectedGroupColumn || 'Group'}') +

  # Apply theme
  theme_${settings.themeName}(base_family = '${settings.fontFamily}') +

  # Customize theme elements
  theme(
    plot.background = element_rect(fill = 'white', color = NA),
    panel.background = element_rect(fill = 'white', color = NA),
    panel.grid.major.y = element_line(color = '#e5e7eb', linewidth = 0.6),
    panel.grid.minor = element_blank(),
    axis.line = element_line(linewidth = 0.6),
    axis.ticks = element_line(linewidth = 0.5),
    plot.title = element_text(size = ${settings.titleSize}, face = '${settings.titleWeight}', hjust = 0.5),
    axis.title.x = element_text(size = ${settings.xAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.title.y = element_text(size = ${settings.yAxisTitleSize}, face = '${settings.axisTitleWeight}'),
    axis.text.x = element_text(size = ${settings.xAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    axis.text.y = element_text(size = ${settings.yAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    legend.text = element_text(size = ${settings.legendTextSize}),
    legend.title = element_text(size = ${settings.legendTextSize})
  )

# Add labels
`;

  if (settings.showTitle) {
    code += `p <- p + ggtitle(${convertToRPlotmath(settings.title)})\n`;
  }
  if (settings.showXLabel) {
    code += `p <- p + xlab(${convertToRPlotmath(settings.xlab)})\n`;
  }
  if (settings.showYLabel) {
    code += `p <- p + ylab(${convertToRPlotmath(settings.ylab)})\n`;
  }

  // Add statistical comparisons using helper function
  code += generateStatisticalComparisonCode(settings, groupColName, xColName);

  code += `\n# Display the plot
print(p)
`;

  return code;
}

// Generate code for line plots with error bars
function generateLineGroupedCode(settings, groupColors) {
  // Use col1, col2, col3 naming to match data frame creation
  const groupColName = `col${settings.groupColIndex}`;
  const xColName = `col${settings.xColIndex}`;
  const yColName = `col${settings.yColIndex}`;

  // Get actual groups and categories from data for ordering
  let actualGroups = [];
  let actualCategories = [];
  if (window.lastProcessedData && window.lastProcessedData.length > 0) {
    const data = window.lastProcessedData;
    const grpIdx = settings.groupColIndex - 1;
    const catIdx = settings.xColIndex - 1;
    if (grpIdx >= 0) {
      // Get unique groups in data order
      actualGroups = [...new Set(data.slice(1).map(row => row[grpIdx]).filter(v => v))];
    }
    if (catIdx >= 0) {
      // Get unique categories (x-axis values) in data order
      actualCategories = [...new Set(data.slice(1).map(row => row[catIdx]).filter(v => v))];
    }
  }

  let code = '';

  // Handle category (x-axis) ordering based on dataOrder mode
  if (actualCategories.length > 0) {
    let finalCategories = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let categories = [];
      if (settings.customOrderCategory) {
        if (typeof settings.customOrderCategory === 'string') {
          categories = settings.customOrderCategory.split(',').map(c => c.trim()).filter(c => c);
        } else if (Array.isArray(settings.customOrderCategory)) {
          categories = settings.customOrderCategory;
        }
      }

      // Validate: only use custom order if values match actual categories
      const isValidCategoryOrder = categories.length > 0 &&
                                   categories.every(c => actualCategories.includes(c));

      if (isValidCategoryOrder) {
        finalCategories = categories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set custom x-axis (timepoint) order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalCategories = actualCategories;
        const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
        code += `# Set x-axis (timepoint) order from data\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort categories alphabetically
      finalCategories = [...actualCategories].sort((a, b) => String(a).localeCompare(String(b)));
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set alphabetical x-axis (timepoint) order\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalCategories = actualCategories;
      const categoryOrder = finalCategories.map(c => `'${c}'`).join(', ');
      code += `# Set x-axis (timepoint) order from data (original order)\ndat$${xColName} <- factor(dat$${xColName}, levels = c(${categoryOrder}))\n\n`;
    }
  }

  // Handle group ordering based on dataOrder mode
  if (actualGroups.length > 0) {
    let finalGroups = [];
    const dataOrder = settings.dataOrder || 'original';

    if (dataOrder === 'custom') {
      // Check if custom order is specified
      let groups = [];
      if (settings.customOrderGroup) {
        if (typeof settings.customOrderGroup === 'string') {
          groups = settings.customOrderGroup.split(',').map(g => g.trim()).filter(g => g);
        } else if (Array.isArray(settings.customOrderGroup)) {
          groups = settings.customOrderGroup;
        }
      }

      // Validate: only use custom order if values match actual groups
      const isValidGroupOrder = groups.length > 0 &&
                                groups.every(g => actualGroups.includes(g));

      if (isValidGroupOrder) {
        finalGroups = groups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set custom group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      } else {
        // Fall back to original order
        finalGroups = actualGroups;
        const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
        code += `# Set group order from data\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
      }
    } else if (dataOrder === 'alphabetical' || dataOrder === 'default') {
      // Sort groups alphabetically
      finalGroups = [...actualGroups].sort((a, b) => String(a).localeCompare(String(b)));
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set alphabetical group order\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    } else {
      // Original order: use order as values appear in data
      finalGroups = actualGroups;
      const groupOrder = finalGroups.map(g => `'${g}'`).join(', ');
      code += `# Set group order from data (original order)\ndat$${groupColName} <- factor(dat$${groupColName}, levels = c(${groupOrder}))\n\n`;
    }
  }

  code += `# ============ Calculate Summary Statistics ============

# Calculate mean and SD/SE for each group and time point
summary_data <- dat %>%
  group_by(${groupColName}, ${xColName}) %>%
  summarise(
    mean_value = mean(${yColName}, na.rm = TRUE),
    sd_value = sd(${yColName}, na.rm = TRUE),
    se_value = sd(${yColName}, na.rm = TRUE) / sqrt(n()),
    n = n(),
    .groups = 'drop'
  )

# Rename columns
colnames(summary_data) <- c('Group', 'TimePoint', 'Mean', 'SD', 'SE', 'N')
summary_data$Error <- summary_data$${settings.errorBarType === 'SE' ? 'SE' : 'SD'}

# ============ Create Plot ============

p <- ggplot(summary_data, aes(x = TimePoint, y = Mean, color = Group, group = Group)) +

  # Add lines
  geom_line(linewidth = ${settings.lineWidth}) +

  # Add points
  geom_point(size = 3) +

  # Add error bars
  geom_errorbar(
    aes(ymin = Mean - Error, ymax = Mean + Error),
    width = 0.2
  ) +

  # Set colors
  scale_color_manual(values = ${groupColors}) +

  # Apply theme
  theme_${settings.themeName}(base_family = '${settings.fontFamily}') +

  # Customize theme to match add-in appearance
  theme(
    # Background
    plot.background = element_rect(fill = 'white', color = NA),
    panel.background = element_rect(fill = 'white', color = NA),

    # Grid lines
    panel.grid.major.y = element_line(color = '#e5e7eb', linewidth = 0.6),
    panel.grid.minor = element_blank(),

    # Axis lines
    axis.line = element_line(linewidth = 0.6),
    axis.ticks = element_line(linewidth = 0.5),

    # Title and text
    plot.title = element_text(size = ${settings.titleSize}, face = '${settings.titleWeight}', hjust = 0.5, margin = margin(b = 6)),
    axis.title.x = element_text(size = ${settings.xAxisTitleSize}, face = '${settings.axisTitleWeight}', margin = margin(t = 6)),
    axis.title.y = element_text(size = ${settings.yAxisTitleSize}, face = '${settings.axisTitleWeight}', margin = margin(r = 6)),
    axis.text.x = element_text(size = ${settings.xAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    axis.text.y = element_text(size = ${settings.yAxisTextSize}, color = 'black', face = '${settings.axisTextWeight}'),
    legend.text = element_text(size = ${settings.legendTextSize})
  )

`;

  if (settings.showTitle) {
    code += `p <- p + ggtitle('${settings.title}')\n`;
  }
  if (settings.showXLabel) {
    code += `p <- p + xlab('${settings.xlab}')\n`;
  }
  if (settings.showYLabel) {
    code += `p <- p + ylab('${settings.ylab}')\n`;
  }

  // Add vbracket for statistical comparisons (3+ groups)
  if (settings.addStatistics && settings.numGroups >= 3) {
    const vbracketX = settings.vbracketX || 0.05;
    const vbracketY = settings.vbracketY || 0.99;
    const vbracketTextSize = settings.vbracketTextSize || 14;
    const vbracketMargin = settings.vbracketMargin || 0.06;
    const vbracketLineWidth = settings.vbracketLineWidth || 3;
    const vbracketSigSize = settings.vbracketSigSize || 20;

    // Determine which test to use based on settings
    const isNonParametric = settings.dataType === 'nonparametric' ||
                            settings.statisticalTest === 'nonparametric' ||
                            settings.statisticalTest === 'kruskal';
    const postHocTest = settings.postHocTest || 'tukey';
    const isVsControl = postHocTest === 'dunnett' || postHocTest === 'steel';
    const controlGroup = settings.controlGroup || '';

    code += `
# ============ Statistical Analysis with vbracket ============
# Install vbracket if needed: install.packages('vbracket', repos = 'https://h20gg702.r-universe.dev')
library(vbracket)
${isNonParametric && postHocTest === 'steel' ? `# Install and load kSamples for Steel test
if (!requireNamespace('kSamples', quietly = TRUE)) {
  install.packages('kSamples')
}
library(kSamples)` : ''}
${isNonParametric && postHocTest === 'dunn' ? `# Install and load dunn.test for Dunn test
if (!requireNamespace('dunn.test', quietly = TRUE)) {
  install.packages('dunn.test')
}
library(dunn.test)` : ''}

# Get group order (same as factor levels used in plot)
group_order <- levels(factor(summary_data$Group))
group_colors <- ${groupColors}

# Perform statistical tests at selected time point
selected_timepoint <- ${settings.vbracketTimepoint || 1}
test_data <- dat %>% filter(${xColName} == selected_timepoint)

# Get unique groups (in same order as plot)
n_groups <- length(group_order)

# Remove default legend (vbracket will replace it)
p <- p + theme(legend.position = 'none')

# Statistical test settings
data_type <- '${isNonParametric ? 'nonparametric' : 'parametric'}'
posthoc_type <- '${postHocTest}'
control_group <- '${controlGroup}'
if (control_group == '' || !(control_group %in% group_order)) {
  control_group <- group_order[1]  # Default to first group
}

# Test normality for each group (Shapiro-Wilk)
cat('Normality tests (Shapiro-Wilk) at timepoint', selected_timepoint, ':\\n')
normality_results <- list()
for (grp in group_order) {
  grp_data <- test_data[test_data$${groupColName} == grp, '${yColName}']
  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
    sw_result <- shapiro.test(grp_data)
    normality_results[[grp]] <- sw_result$p.value
    cat('  ', grp, ': p =', format(sw_result$p.value, digits = 4),
        if(sw_result$p.value < 0.05) '(non-normal)' else '(normal)', '\\n')
  } else {
    normality_results[[grp]] <- NA
    cat('  ', grp, ': n =', length(grp_data), '(skipped - need 3-5000 samples)\\n')
  }
}

# Check normality and show appropriate warning
any_non_normal <- any(sapply(normality_results, function(p) !is.na(p) && p < 0.05))
all_normal <- all(sapply(normality_results, function(p) is.na(p) || p >= 0.05))

if (data_type == 'parametric' && any_non_normal) {
  cat('\\n⚠️ WARNING: Some groups appear non-normal. Parametric test (ANOVA/', posthoc_type, ') was selected manually.\\n')
  cat('   Consider using non-parametric test (Kruskal-Wallis/Dunn or Steel) for more robust results.\\n\\n')
} else if (data_type == 'nonparametric' && all_normal) {
  cat('\\nℹ️ NOTE: All groups appear normally distributed. Non-parametric test was selected manually.\\n')
  cat('   Parametric test (ANOVA/Tukey or Dunnett) may have more statistical power.\\n\\n')
} else {
  cat('\\n')
}

# Perform statistical tests
if (n_groups >= 3) {
  comparisons_list <- list()
  labels_list <- c()

  if (data_type == 'nonparametric') {
    # Kruskal-Wallis test (non-parametric)
    kw_result <- kruskal.test(${yColName} ~ ${groupColName}, data = test_data)
    overall_p <- kw_result$p.value
    cat('Kruskal-Wallis test: p =', overall_p, '\\n')

    if (overall_p < 0.05) {
      if (posthoc_type == 'steel') {
        # Steel test - non-parametric vs control (using kSamples)
        cat('Running Steel test (vs control:', control_group, ')\\n')
        control_values <- test_data[test_data$${groupColName} == control_group, '${yColName}']
        treatment_groups <- group_order[group_order != control_group]

        for (trt in treatment_groups) {
          trt_values <- test_data[test_data$${groupColName} == trt, '${yColName}']
          steel_result <- Steel.test(list(control_values, trt_values))
          p_val <- steel_result$st[2]  # p-value is in st[2]

          cat(control_group, 'vs', trt, ': p =', p_val, '\\n')

          if (p_val < 0.05) {
            comparisons_list[[length(comparisons_list) + 1]] <- c(control_group, trt)
            if (p_val < 0.001) {
              labels_list <- c(labels_list, '***')
            } else if (p_val < 0.01) {
              labels_list <- c(labels_list, '**')
            } else {
              labels_list <- c(labels_list, '*')
            }
          }
        }
      } else {
        # Dunn test - non-parametric pairwise
        cat('Running Dunn test (pairwise comparisons)\\n')
        dunn_result <- dunn.test(test_data$${yColName}, test_data$${groupColName}, method = 'bonferroni')

        for (i in seq_along(dunn_result$comparisons)) {
          comp <- dunn_result$comparisons[i]
          p_val <- dunn_result$P.adjusted[i]

          grps <- strsplit(comp, ' - ')[[1]]
          cat(grps[1], 'vs', grps[2], ': p =', p_val, '\\n')

          if (p_val < 0.05) {
            comparisons_list[[length(comparisons_list) + 1]] <- grps
            if (p_val < 0.001) {
              labels_list <- c(labels_list, '***')
            } else if (p_val < 0.01) {
              labels_list <- c(labels_list, '**')
            } else {
              labels_list <- c(labels_list, '*')
            }
          }
        }
      }
    }
  } else {
    # Parametric tests (ANOVA)
    aov_result <- aov(${yColName} ~ ${groupColName}, data = test_data)
    overall_p <- summary(aov_result)[[1]][['Pr(>F)']][1]
    cat('ANOVA: p =', overall_p, '\\n')

    if (overall_p < 0.05) {
      if (posthoc_type == 'dunnett') {
        # Dunnett test - parametric vs control
        cat('Running Dunnett test (vs control:', control_group, ')\\n')
        library(multcomp)
        test_data$${groupColName} <- relevel(factor(test_data$${groupColName}), ref = control_group)
        dunnett_result <- glht(aov(${yColName} ~ ${groupColName}, data = test_data), linfct = mcp(${groupColName} = 'Dunnett'))
        dunnett_summary <- summary(dunnett_result)

        for (i in seq_along(dunnett_summary$test$coefficients)) {
          comp_name <- names(dunnett_summary$test$coefficients)[i]
          p_val <- dunnett_summary$test$pvalues[i]

          # Parse comparison name (format: "Treatment - Control")
          grps <- strsplit(comp_name, ' - ')[[1]]
          cat(grps[1], 'vs', grps[2], ': p =', p_val, '\\n')

          if (p_val < 0.05) {
            comparisons_list[[length(comparisons_list) + 1]] <- c(control_group, trimws(grps[1]))
            if (p_val < 0.001) {
              labels_list <- c(labels_list, '***')
            } else if (p_val < 0.01) {
              labels_list <- c(labels_list, '**')
            } else {
              labels_list <- c(labels_list, '*')
            }
          }
        }
      } else if (posthoc_type == 'bonferroni' || posthoc_type == 'holm') {
        # Bonferroni or Holm correction - parametric pairwise
        cat('Running pairwise t-test with', posthoc_type, 'correction\\n')
        pairwise_result <- pairwise.t.test(test_data$${yColName}, test_data$${groupColName},
                                            p.adjust.method = posthoc_type)

        # Extract p-values from the matrix
        p_matrix <- pairwise_result$p.value
        for (i in 1:nrow(p_matrix)) {
          for (j in 1:ncol(p_matrix)) {
            if (!is.na(p_matrix[i, j])) {
              grp1 <- colnames(p_matrix)[j]
              grp2 <- rownames(p_matrix)[i]
              p_val <- p_matrix[i, j]

              cat(grp1, 'vs', grp2, ': p =', p_val, '\\n')

              if (p_val < 0.05) {
                comparisons_list[[length(comparisons_list) + 1]] <- c(grp1, grp2)
                if (p_val < 0.001) {
                  labels_list <- c(labels_list, '***')
                } else if (p_val < 0.01) {
                  labels_list <- c(labels_list, '**')
                } else {
                  labels_list <- c(labels_list, '*')
                }
              }
            }
          }
        }
      } else {
        # Tukey HSD - parametric pairwise (default)
        cat('Running Tukey HSD test (pairwise comparisons)\\n')
        tukey_result <- TukeyHSD(aov_result)
        tukey_df <- as.data.frame(tukey_result[['${groupColName}']])
        tukey_df[['comparison']] <- rownames(tukey_df)

        for (i in 1:nrow(tukey_df)) {
          comp <- tukey_df[['comparison']][i]
          p_val <- tukey_df[['p adj']][i]

          grps <- strsplit(comp, '-')[[1]]
          cat(grps[1], 'vs', grps[2], ': p =', p_val, '\\n')

          if (p_val < 0.05) {
            comparisons_list[[length(comparisons_list) + 1]] <- grps
            if (p_val < 0.001) {
              labels_list <- c(labels_list, '***')
            } else if (p_val < 0.01) {
              labels_list <- c(labels_list, '**')
            } else {
              labels_list <- c(labels_list, '*')
            }
          }
        }
      }
    }
  }

  # Add vbracket legend
  if (length(comparisons_list) > 0) {
    vb_comparisons <- add_bracket_comparisons(
      groups1 = sapply(comparisons_list, function(x) x[1]),
      groups2 = sapply(comparisons_list, function(x) x[2]),
      labels = labels_list
    )

    p <- p + legend_bracket(
        labels = group_order,
        colors = group_colors,
        comparisons = vb_comparisons,
        legend_x = ${vbracketX},
        legend_y = ${vbracketY},
        text_size = ${vbracketTextSize},
        sig_size = ${vbracketSigSize},
        bracket_margin = ${vbracketMargin},
        output_width = ${settings.expWidth || 6},
        output_height = ${settings.expHeight || 4}
      )
  } else {
    # No significant comparisons - add legend without brackets
    p <- p + legend_bracket(
        labels = group_order,
        colors = group_colors,
        legend_x = ${vbracketX},
        legend_y = ${vbracketY},
        text_size = ${vbracketTextSize},
        output_width = ${settings.expWidth || 6},
        output_height = ${settings.expHeight || 4}
      )
  }
}

`;
  }

  code += `\nprint(p)\n`;

  return code;
}

// Download R code handler
async function downloadRCodeHandler() {
  try {
    const rCode = extractRelevantRCode();  // Try extraction again

    if (rCode.includes("Please click 'Preview' first")) {
      setStatus("⚠️ Please preview chart first before exporting R code");
      return;
    }

    const success = downloadRFile(rCode);
    if (success) {
      setStatus("✅ R code downloaded as 'generated_plot.R' to your Downloads folder");
    } else {
      setStatus("❌ Download failed - please try again");
    }
  } catch (e) {
    console.error('Download R code error:', e);
    setStatus("Error: " + (e?.message || e));
  }
}

// Write R code to clipboard handler
async function writeRCodeToCellHandler() {
  try {
    const rCode = extractRelevantRCode();  // Try extraction again

    if (rCode.includes("Please click 'Preview' first")) {
      setStatus("⚠️ Please preview chart first before exporting R code");
      return;
    }

    // Copy to clipboard directly (bypasses Excel's quote-wrapping behavior)
    await navigator.clipboard.writeText(rCode);
    setStatus("✅ R code copied to clipboard! Paste into R/RStudio.");
  } catch (e) {
    console.error('Copy R code error:', e);
    // Fallback to cell writing if clipboard fails
    try {
      await writeRCodeToCell(rCode);
    } catch (e2) {
      setStatus("Error: " + (e?.message || e));
    }
  }
}

// ========= Preset Settings Functions =========

// Collect all current settings from UI
function collectCurrentSettings() {
  const el = (id) => document.getElementById(id);

  const settings = {
    // Font settings
    fontFamily: el("fontFam")?.value || "Arial",
    titleSize: el("titleSize")?.value || "24",
    xAxisTitleSize: el("xAxisTitleSize")?.value || "20",
    yAxisTitleSize: el("yAxisTitleSize")?.value || "20",
    xAxisTextSize: el("xAxisTextSize")?.value || "18",
    yAxisTextSize: el("yAxisTextSize")?.value || "18",
    legendTextSize: el("legendTextSize")?.value || "16",
    titleWeight: el("titleWeight")?.value || "bold",
    axisTitleWeight: el("axisTitleWeight")?.value || "plain",
    axisTextWeight: el("axisTextWeight")?.value || "plain",

    // Theme
    themeName: el("ggtheme")?.value || "minimal",

    // Colors
    fillColor: el("fillColor")?.value || "#4C78A8",
    strokeColor: el("strokeColor")?.value || "#1f2937",
    fillAlpha: el("fillAlpha")?.value || "1",

    // Bar dimensions
    barWidth: el("barWidth")?.value || "0.8",
    dodgeWidth: el("dodgeWidth")?.value || "0.9",
    lineWidth: el("lineWidth")?.value || "0.5",

    // Dot settings (for plots with data points)
    dotSize: el("dotSize")?.value || "2",
    dotAlpha: el("dotAlpha")?.value || "0.5",
    dotColor: el("dotColor")?.value || "#000000",
    dotShape: el("dotShape")?.value || "16",

    // Rotation and layout
    rotation: el("rotation")?.value || "0",
    tableStyleLabels: el("tableStyleLabels")?.checked || false,
    xScale: el("xScale")?.value || "linear",
    yScale: el("yScale")?.value || "linear",

    // Axis rotation
    xAxisRotation: el("xAxisRotation")?.value || "0",
    yAxisRotation: el("yAxisRotation")?.value || "0",
    xAxisHjust: el("xAxisHjust")?.value || "0.5",
    xAxisVjust: el("xAxisVjust")?.value || "0.5",
    yAxisHjust: el("yAxisHjust")?.value || "0.5",
    yAxisVjust: el("yAxisVjust")?.value || "0.5",

    // Show/hide labels
    showTitle: el("showTitle")?.checked ? "true" : "false",
    showXLabel: el("showXLabel")?.checked ? "true" : "false",
    showYLabel: el("showYLabel")?.checked ? "true" : "false",

    // Label text content
    titleText: el("titleText")?.value || "",
    xLabel: el("xLabel")?.value || "",
    yLabel: el("yLabel")?.value || "",

    // Data order
    dataOrder: el("dataOrder")?.value || "original",
    customOrderGroup: (typeof getCustomOrderGroupString === 'function') ? getCustomOrderGroupString() : "",
    customOrderCategory: (typeof getCustomOrderCategoryString === 'function') ? getCustomOrderCategoryString() : "",

    // Export settings
    expWidth: el("expWidth")?.value || "6",
    expHeight: el("expHeight")?.value || "4",
    expUnits: el("expUnits")?.value || "in",
    expDpi: el("expDpi")?.value || "300",

    // Statistics
    addStatistics: el("addStatistics")?.checked ? "true" : "false",
    errorBarType: el("errorBarType")?.value || "sd",
    statisticalTestMode: el("statisticalTestMode")?.value || "auto",
    // When manual mode, use dataType to determine test; when auto, use "auto"
    statisticalTest: (el("statisticalTestMode")?.value || "auto") === "auto"
      ? "auto"
      : (el("dataTypeSelect")?.value === "nonparametric" ? "nonparametric" : "parametric"),
    varianceTest: el("varianceTest")?.value || "levene",
    statSymbolType: el("statSymbolType")?.value || "stars",
    customSymbol05: el("customSymbol05")?.value || "*",
    customSymbol01: el("customSymbol01")?.value || "**",
    customSymbol001: el("customSymbol001")?.value || "***",
    customSymbolNS: el("customSymbolNS")?.value || "ns",
    statSymbolSize: el("statSymbolSize")?.value || "6",
    significanceLevel: el("significanceLevel")?.value || "0.05",
    comparisonMode: document.querySelector('input[name="comparisonMode"]:checked')?.value || "significant",
    postHocTest: (el("dataTypeSelect")?.value === "nonparametric" ? el("postHocTestNonparam")?.value : el("postHocTest")?.value) || "tukey",
    controlGroup: el("dunnettControl")?.value || "",
    dataType: el("dataTypeSelect")?.value || "parametric",
    customComparisons: (typeof getSelectedCustomComparisons === 'function') ? JSON.stringify(getSelectedCustomComparisons()) : "[]",
    customPositions: (typeof getCustomBracketPositions === 'function') ? JSON.stringify(getCustomBracketPositions()) : "{}",

    // VBracket legend settings (for 3+ groups line plots)
    vbracketTimepoint: el("vbracketTimepoint")?.value || "",
    vbracketPosition: "custom",
    vbracketX: el("vbracketX")?.value || "0.05",
    vbracketY: el("vbracketY")?.value || "0.99",
    vbracketTextSize: el("vbracketTextSize")?.value || "14",
    vbracketSigSize: el("vbracketSigSize")?.value || "20",
    vbracketMargin: el("vbracketMargin")?.value || "0.06",
    vbracketLineWidth: el("vbracketLineWidth")?.value || "3",
    vbracketLegendLineLength: el("vbracketLegendLineLength")?.value || "0.05",
    vbracketLegendLineWidth: el("vbracketLegendLineWidth")?.value || "2",
    vbracketItemSpacing: el("vbracketItemSpacing")?.value || "0.1",
    vbracketBracketLayerSpacing: el("vbracketBracketLayerSpacing")?.value || "",

    // Number of bins (for histogram)
    bins: el("bins")?.value || "20",

    // Number of groups
    numGroups: el("numGroups")?.value || "2"
  };

  // Collect group colors dynamically
  const groupColors = {};
  for (let i = 1; i <= 10; i++) {
    const colorInput = el(`groupColor${i}`);
    if (colorInput && colorInput.value) {
      groupColors[`groupColor${i}`] = colorInput.value;
    }
  }

  // Add group colors to settings
  Object.assign(settings, groupColors);

  // Debug logging for custom symbols
  console.log('🔥 stat_symbol_type:', settings.statSymbolType);
  console.log('🔥 customSymbol05:', settings.customSymbol05);
  console.log('🔥 customSymbol01:', settings.customSymbol01);
  console.log('🔥 customSymbol001:', settings.customSymbol001);
  console.log('🔥 customSymbolNS:', settings.customSymbolNS);

  return settings;
}

// Save current settings to Excel sheet named "Settings_[name]"
async function saveSettingsToSheet() {
  try {
    // Get preset name from input field
    const presetNameInput = document.getElementById("presetName");
    const presetName = presetNameInput?.value || "";

    if (!presetName || presetName.trim() === "") {
      setStatus("❌ Please enter a preset name");
      return;
    }

    const sheetName = `Settings_${presetName.trim()}`;
    const settings = collectCurrentSettings();

    await Excel.run(async (context) => {
      const workbook = context.workbook;

      // Check if sheet already exists
      const sheets = workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const existingSheet = sheets.items.find(s => s.name === sheetName);
      const isOverwrite = !!existingSheet;

      if (existingSheet) {
        // Automatically overwrite existing preset
        existingSheet.delete();
        await context.sync();
        console.log(`Overwriting existing preset: ${sheetName}`);
      }

      // Create new sheet
      const newSheet = workbook.worksheets.add(sheetName);

      // Write headers
      newSheet.getRange("A1").values = [["Setting"]];
      newSheet.getRange("B1").values = [["Value"]];

      // Convert settings object to array
      const settingsArray = Object.entries(settings).map(([key, value]) => [key, value]);

      // Write settings data
      const dataRange = newSheet.getRange(`A2:B${settingsArray.length + 1}`);
      dataRange.values = settingsArray;

      // Format the sheet
      const headerRange = newSheet.getRange("A1:B1");
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#4472C4";
      headerRange.format.font.color = "white";

      // Auto-fit columns
      newSheet.getRange("A:B").format.autofitColumns();

      await context.sync();

      if (isOverwrite) {
        setStatus(`✅ Preset "${presetName}" updated (overwritten)`);
      } else {
        setStatus(`✅ Preset "${presetName}" saved to sheet "${sheetName}"`);
      }

      // Refresh preset dropdown
      await refreshPresetList();
    });

  } catch (e) {
    console.error('Save settings error:', e);
    setStatus("❌ Error saving preset: " + (e?.message || e));
  }
}

// Refresh the list of available preset sheets
async function refreshPresetList() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      // Find all sheets starting with "Settings_"
      const presetSheets = sheets.items
        .filter(s => s.name.startsWith("Settings_"))
        .map(s => s.name);

      const dropdown = document.getElementById("presetSheet");
      if (!dropdown) return;

      // Clear existing options
      dropdown.innerHTML = "";

      if (presetSheets.length === 0) {
        dropdown.innerHTML = '<option value="">-- No presets found --</option>';
      } else {
        presetSheets.forEach(sheetName => {
          const option = document.createElement("option");
          option.value = sheetName;
          option.textContent = sheetName.replace("Settings_", "");
          dropdown.appendChild(option);
        });
      }
    });
  } catch (e) {
    console.error('Refresh preset list error:', e);
  }
}

// Load settings from selected preset sheet
async function loadSettingsFromSheet() {
  try {
    const dropdown = document.getElementById("presetSheet");
    const sheetName = dropdown?.value;

    if (!sheetName || sheetName === "") {
      setStatus("⚠️ Please select a preset to load");
      return;
    }

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();

      const values = usedRange.values;

      // Skip header row and convert to object
      const settings = {};
      for (let i = 1; i < values.length; i++) {
        const [key, value] = values[i];
        if (key) settings[key] = value;
      }

      // Apply settings to UI
      applySettingsToUI(settings);

      const presetName = sheetName.replace("Settings_", "");
      setStatus(`✅ Preset "${presetName}" loaded successfully`);
    });

  } catch (e) {
    console.error('Load settings error:', e);
    setStatus("❌ Error loading preset: " + (e?.message || e));
  }
}

// Apply settings object to UI controls
function applySettingsToUI(settings) {
  const el = (id) => document.getElementById(id);

  // Helper to safely set value
  const setValue = (id, value) => {
    const element = el(id);
    if (element && value !== undefined && value !== null) {
      element.value = value;
    }
  };

  // Helper to safely set checkbox
  const setChecked = (id, value) => {
    const element = el(id);
    if (element && value !== undefined && value !== null) {
      element.checked = (value === "true" || value === true);
    }
  };

  // Font settings
  setValue("fontFam", settings.fontFamily);
  setValue("titleSize", settings.titleSize);
  setValue("xAxisTitleSize", settings.xAxisTitleSize);
  setValue("yAxisTitleSize", settings.yAxisTitleSize);
  setValue("xAxisTextSize", settings.xAxisTextSize);
  setValue("yAxisTextSize", settings.yAxisTextSize);
  setValue("legendTextSize", settings.legendTextSize);
  setValue("titleWeight", settings.titleWeight);
  setValue("axisTitleWeight", settings.axisTitleWeight || settings.axisWeight);
  setValue("axisTextWeight", settings.axisTextWeight || settings.axisWeight);

  // Theme
  setValue("ggtheme", settings.themeName);

  // Colors
  setValue("fillColor", settings.fillColor);
  setValue("strokeColor", settings.strokeColor);
  setValue("fillAlpha", settings.fillAlpha);

  // Update color pickers
  if (settings.fillColor && el("fillPicker")) {
    try {
      el("fillPicker").value = settings.fillColor;
    } catch (e) {
      // Ignore invalid color format
    }
  }
  if (settings.strokeColor && el("strokePicker")) {
    try {
      el("strokePicker").value = settings.strokeColor;
    } catch (e) {
      // Ignore invalid color format
    }
  }

  // Bar dimensions
  setValue("barWidth", settings.barWidth);
  setValue("dodgeWidth", settings.dodgeWidth);
  setValue("lineWidth", settings.lineWidth);

  // Rotation and layout
  setValue("rotation", settings.rotation);
  setValue("xScale", settings.xScale);
  setValue("yScale", settings.yScale);

  // Axis rotation
  setValue("xAxisRotation", settings.xAxisRotation);
  setValue("yAxisRotation", settings.yAxisRotation);
  setValue("xAxisHjust", settings.xAxisHjust);
  setValue("xAxisVjust", settings.xAxisVjust);
  setValue("yAxisHjust", settings.yAxisHjust);
  setValue("yAxisVjust", settings.yAxisVjust);

  // Show/hide labels
  setChecked("showTitle", settings.showTitle);
  setChecked("showXLabel", settings.showXLabel);
  setChecked("showYLabel", settings.showYLabel);

  // Data order
  setValue("dataOrder", settings.dataOrder);

  // Export settings
  setValue("expWidth", settings.expWidth);
  setValue("expHeight", settings.expHeight);
  setValue("expUnits", settings.expUnits);
  setValue("expDpi", settings.expDpi);

  // Statistics
  setChecked("addStatistics", settings.addStatistics);
  setValue("errorBarType", settings.errorBarType);
  setValue("statisticalTest", settings.statisticalTest);
  setValue("varianceTest", settings.varianceTest);
  setValue("symbolType", settings.symbolType);
  setValue("statSymbolSize", settings.statSymbolSize);
  setValue("significanceLevel", settings.significanceLevel);
  setValue("postHocTest", settings.postHocTest);

  // VBracket legend settings (position dropdown removed - manual X/Y only)
  setValue("vbracketX", settings.vbracketX);
  setValue("vbracketY", settings.vbracketY);
  setValue("vbracketTextSize", settings.vbracketTextSize);
  setValue("vbracketMargin", settings.vbracketMargin);
  setValue("vbracketLineWidth", settings.vbracketLineWidth);
  setValue("vbracketLegendLineLength", settings.vbracketLegendLineLength);
  setValue("vbracketLegendLineWidth", settings.vbracketLegendLineWidth);
  setValue("vbracketItemSpacing", settings.vbracketItemSpacing);

  // Comparison mode (radio buttons)
  if (settings.comparisonMode) {
    const radio = document.getElementById(settings.comparisonMode);
    if (radio) radio.checked = true;
  }

  // Number of bins
  setValue("bins", settings.bins);

  // Number of groups
  setValue("numGroups", settings.numGroups);

  // Restore group colors dynamically
  for (let i = 1; i <= 10; i++) {
    const colorKey = `groupColor${i}`;
    if (settings[colorKey]) {
      setValue(colorKey, settings[colorKey]);

      // Also update the color picker
      const pickerEl = el(`groupPicker${i}`);
      if (pickerEl) {
        try {
          pickerEl.value = settings[colorKey];
        } catch (e) {
          // Ignore invalid color format
        }
      }
    }
  }

  // Trigger numGroups change event to update UI
  const numGroupsEl = el("numGroups");
  if (numGroupsEl) {
    numGroupsEl.dispatchEvent(new Event('change'));
  }

  console.log("✅ Settings applied to UI");
}

// ========= イベントハンドラーの更新 =========

// 更新されたイベントハンドラー
function updateEventHandlersWithDebug() {
  // 既存のハンドラーを置換
  const previewBtn = document.getElementById("preview");
  if (previewBtn) {
    // 既存のハンドラーを削除してから新しいハンドラーを追加
    const newPreviewBtn = previewBtn.cloneNode(true);
    previewBtn.parentNode.replaceChild(newPreviewBtn, previewBtn);
    newPreviewBtn.addEventListener("click", previewPlotWithDebug);
  }

  const insertBtn = document.getElementById("insert");
  if (insertBtn) {
    const newInsertBtn = insertBtn.cloneNode(true);
    insertBtn.parentNode.replaceChild(newInsertBtn, insertBtn);
    newInsertBtn.addEventListener("click", insertIntoExcelFixed);
  }

  const downloadRCodeBtn = document.getElementById("downloadRCode");
  if (downloadRCodeBtn) {
    const newDownloadBtn = downloadRCodeBtn.cloneNode(true);
    downloadRCodeBtn.parentNode.replaceChild(newDownloadBtn, downloadRCodeBtn);
    newDownloadBtn.addEventListener("click", downloadRCodeHandler);
  }

  const writeRCodeBtn = document.getElementById("writeRCodeToCell");
  if (writeRCodeBtn) {
    const newWriteBtn = writeRCodeBtn.cloneNode(true);
    writeRCodeBtn.parentNode.replaceChild(newWriteBtn, writeRCodeBtn);
    newWriteBtn.addEventListener("click", writeRCodeToCellHandler);
  }

  // Preset buttons
  const savePresetBtn = document.getElementById("savePreset");
  if (savePresetBtn) {
    const newSavePresetBtn = savePresetBtn.cloneNode(true);
    savePresetBtn.parentNode.replaceChild(newSavePresetBtn, savePresetBtn);
    newSavePresetBtn.addEventListener("click", saveSettingsToSheet);
  }

  const loadPresetBtn = document.getElementById("loadPreset");
  if (loadPresetBtn) {
    const newLoadPresetBtn = loadPresetBtn.cloneNode(true);
    loadPresetBtn.parentNode.replaceChild(newLoadPresetBtn, loadPresetBtn);
    newLoadPresetBtn.addEventListener("click", loadSettingsFromSheet);
  }
}

// ---- 初期化 ----
Office.onReady(() => {
  // Display version in title
  const versionSpan = document.getElementById("appVersion");
  if (versionSpan) {
    versionSpan.textContent = `v${ADDIN_VERSION}`;
  }

  // === Advanced 決め打ちトグル（確実表示） ===
  const advBtn   = document.getElementById("toggleAdv");
  const advPanel = document.getElementById("advPanel");

  // id duplication/reference check
  console.log("advPanel nodes:", document.querySelectorAll("#advPanel").length);
  if (document.querySelectorAll("#advPanel").length !== 1) {
    setStatus("Advanced UI has 2+ or 0 instances. Please check id.");
  }

  let advOpen = false; // ← hidden には依存しない

  function renderAdvState() {
    if (!advPanel || !advBtn) return;
    // 強制的に表示/非表示を切替（CSSに勝つ）
    if (advOpen) {
      advPanel.removeAttribute("hidden");
      advPanel.style.display    = "block";
      advPanel.style.visibility = "visible";
      advPanel.style.opacity    = "1";
      advPanel.style.maxHeight  = "none";
      advPanel.style.minHeight  = "160px";
      advBtn.textContent = "Advanced ▾";
      advBtn.setAttribute("aria-expanded", "true");

      // 目で見えるデバッグ台
      const markId = "adv-debug-mark";
      if (!document.getElementById(markId)) {
        const m = document.createElement("div");
        m.id = markId;
        m.textContent = "【ADVANCED MODE";
        m.style.cssText = "margin:6px 0; padding:4px 8px; background:#ffe8a3; border:1px solid #e4c769; border-radius:6px; font-weight:700;";
        advPanel.insertBefore(m, advPanel.firstChild);
      }

      advPanel.scrollIntoView({ behavior: "instant", block: "start" });
    } else {
      advPanel.setAttribute("hidden", "");
      advPanel.style.display    = "none";
      advPanel.style.visibility = "hidden";
      advPanel.style.opacity    = "0";
      advPanel.style.maxHeight  = "0";
      advBtn.textContent = "Advanced ▸";
      advBtn.setAttribute("aria-expanded", "false");
    }

  }

  advBtn?.addEventListener("click", () => { advOpen = !advOpen; renderAdvState(); });
  renderAdvState(); // 初期は閉じる
  // === /Advanced 決め打ちトグル ===

  // Chart type change handler - handle column selector visibility

  // Define which chart types support statistical analysis
  const STATS_SUPPORTED_CHART_TYPES = [
    "box",
    "box_dot",
    "violin",
    "violin_dot",
    "bar_error_dot",
    "bar_grouped_error_dot",
    "box_grouped",
    "box_grouped_dot",
    "violin_grouped",
    "violin_grouped_dot",
    "line_grouped_error_raw"
  ];

  // Function to show/hide statistics controls based on chart type support
  function updateStatisticsAvailability() {
    const chartType = document.getElementById("chartType")?.value || "";
    const statsControlsContainer = document.getElementById("statsControlsContainer");
    const statsUnsupportedMessage = document.getElementById("statsUnsupportedMessage");
    const addStatisticsCheckbox = document.getElementById("addStatistics");

    if (!statsControlsContainer || !statsUnsupportedMessage) return;

    const isSupported = STATS_SUPPORTED_CHART_TYPES.includes(chartType);

    if (isSupported) {
      // Show controls, hide message
      statsControlsContainer.style.display = "flex";
      statsUnsupportedMessage.style.display = "none";
    } else {
      // Hide controls, show message, and UNCHECK the checkbox
      statsControlsContainer.style.display = "none";
      statsUnsupportedMessage.style.display = "block";

      // Uncheck the statistics checkbox to prevent auto-export attempts
      if (addStatisticsCheckbox) {
        addStatisticsCheckbox.checked = false;
      }
    }
  }

  // Function to show/hide detailed statistics controls based on checkbox state
  function updateStatsDetailedControlsVisibility() {
    const addStatistics = document.getElementById("addStatistics")?.checked || false;
    const statsDetailedControls = document.getElementById("statsDetailedControls");

    if (!statsDetailedControls) return;

    // Show detailed controls only when checkbox is checked
    statsDetailedControls.style.display = addStatistics ? "block" : "none";
  }

  // Function to show/hide VBracket controls based on chart type and statistics checkbox
  function updateVbracketVisibility() {
    const vbracketControls = document.getElementById("vbracketControls");
    const traditionalStatControls = document.getElementById("traditionalStatControls");
    if (!vbracketControls) return;

    const chartType = document.getElementById("chartType")?.value || "";
    const addStatistics = document.getElementById("addStatistics")?.checked || false;

    // Show VBracket controls only when:
    // 1. Chart type is "line_grouped_error_raw"
    // 2. Statistics checkbox is checked
    const shouldShowVbracket = (chartType === "line_grouped_error_raw") && addStatistics;

    vbracketControls.style.display = shouldShowVbracket ? "block" : "none";

    // Hide traditional controls when vbracket is shown, show them otherwise
    if (traditionalStatControls) {
      traditionalStatControls.style.display = shouldShowVbracket ? "none" : "block";
    }

    // Also populate timepoints when showing vbracket
    if (shouldShowVbracket && typeof populateVbracketTimepoints === 'function') {
      populateVbracketTimepoints();
    }
  }

  // Function to show/hide X/Y position inputs based on vbracket position dropdown
  function updateVbracketPositionInputs() {
    // X/Y position inputs are always visible now (no dropdown)
    // This function is kept for compatibility but does nothing
  }

  // Update Colors & Style tab controls visibility based on chart type
  function updateColorsStyleControlsVisibility() {
    const chartType = (document.getElementById("chartType")?.value || "").toLowerCase();

    // Get all control rows
    const barInteriorRow = document.getElementById("barInteriorRow");
    const barBorderRow = document.getElementById("barBorderRow");
    const groupColorRow = document.getElementById("groupColorRow");
    const lineWidthRow = document.getElementById("lineWidthRow");
    const barWidthRow = document.getElementById("barWidthRow");
    const dodgeWidthRow = document.getElementById("dodgeWidthRow");
    const dotSettingsRow = document.getElementById("dotSettingsRow");

    // Define chart type groups
    const singleGroupCharts = [
      'histogram', 'box', 'box_dot', 'violin', 'violin_dot',
      'dot', 'bar', 'bar_error', 'bar_error_dot', 'line'
    ];

    const groupedCharts = [
      'bar_grouped', 'bar_grouped_error', 'bar_grouped_error_dot',
      'box_grouped', 'box_grouped_dot',
      'violin_grouped', 'violin_grouped_dot',
      'line_grouped', 'line_grouped_error', 'line_grouped_error_raw'
    ];

    const barCharts = [
      'bar', 'bar_error', 'bar_error_dot',
      'bar_grouped', 'bar_grouped_error', 'bar_grouped_error_dot'
    ];
    const boxCharts = ['box', 'box_dot', 'box_grouped', 'box_grouped_dot'];
    const violinCharts = ['violin', 'violin_dot', 'violin_grouped', 'violin_grouped_dot'];

    const dotCharts = [
      'box_dot', 'violin_dot', 'dot', 'bar_error_dot',
      'bar_grouped_error_dot', 'box_grouped_dot', 'violin_grouped_dot'
    ];

    const groupedWithDodge = [
      'bar_grouped', 'bar_grouped_error', 'bar_grouped_error_dot',
      'box_grouped', 'box_grouped_dot',
      'violin_grouped', 'violin_grouped_dot'
    ];

    // IC50 chart type - hide all standard controls (has its own IC50 Style section)
    const isIC50 = chartType === 'ic50_dose_response';

    // Bar Interior: Show for single-group charts (NOT grouped, NOT line, NOT IC50)
    if (barInteriorRow) {
      const showBarInterior = !isIC50 && singleGroupCharts.includes(chartType) && chartType !== 'line';
      barInteriorRow.style.display = showBarInterior ? "flex" : "none";
    }

    // Bar Border: Show for single-group charts and grouped charts with borders (NOT IC50, NOT line_grouped)
    const groupedChartsWithBorder = [
      'bar_grouped', 'bar_grouped_error', 'bar_grouped_error_dot',
      'box_grouped', 'box_grouped_dot',
      'violin_grouped', 'violin_grouped_dot'
    ];
    if (barBorderRow) {
      const showBarBorder = !isIC50 && (singleGroupCharts.includes(chartType) || groupedChartsWithBorder.includes(chartType));
      barBorderRow.style.display = showBarBorder ? "flex" : "none";

      // Update label text based on chart type
      const label = barBorderRow.querySelector('label');
      if (label) {
        if (chartType === 'line') {
          label.textContent = 'Line Color';
        } else if (chartType.includes('box') || chartType.includes('violin')) {
          label.textContent = 'Border Color';
        } else {
          label.textContent = 'Bar Border';
        }
      }
    }

    // Group Colors: Show for grouped charts (already handled by handleChartTypeChange)
    // We don't need to change it here as it's managed elsewhere

    // Line Width: Show for all chart types except IC50 (IC50 has its own curve width)
    if (lineWidthRow) {
      lineWidthRow.style.display = isIC50 ? "none" : "flex";
    }

    // Bar/Box/Violin Width: Show for bar, box, and violin charts (NOT IC50)
    if (barWidthRow) {
      const showWidth = !isIC50 && (barCharts.includes(chartType) || boxCharts.includes(chartType) || violinCharts.includes(chartType));
      barWidthRow.style.display = showWidth ? "flex" : "none";

      // Update label based on chart type
      const label = barWidthRow.querySelector('label');
      if (label) {
        if (boxCharts.includes(chartType)) {
          label.textContent = 'Box Width';
        } else if (violinCharts.includes(chartType)) {
          label.textContent = 'Violin Width';
        } else {
          label.textContent = 'Bar Width';
        }
      }
    }

    // Dodge Width: Show only for grouped charts with dodge (NOT line_grouped*, NOT IC50)
    if (dodgeWidthRow) {
      const showDodgeWidth = !isIC50 && groupedWithDodge.includes(chartType);
      dodgeWidthRow.style.display = showDodgeWidth ? "flex" : "none";
    }

    // Dot Settings: Show only for charts with dots (NOT IC50 - has its own point settings)
    if (dotSettingsRow) {
      const showDotSettings = !isIC50 && dotCharts.includes(chartType);
      dotSettingsRow.style.display = showDotSettings ? "flex" : "none";
    }
  }

  function handleChartTypeChange() {
    // Clear old statistical results when chart type changes
    clearOldStatisticalResults().catch(e => console.log("Note: Could not clear stats on chart type change"));

    const v = (document.getElementById("chartType")?.value || "").toLowerCase();
    const lx = document.querySelector('label[for="xColumn"]');
    const ly = document.querySelector('label[for="yColumn"]');
    const groupColorRow = document.getElementById("groupColorRow");
    const groupColumnRow = document.getElementById("groupColumnRow");
    const errorColumnRow = document.getElementById("errorColumnRow");

    if (!lx || !ly) return;

    // Define grouped chart types
    const isGrouped = GROUPED_CHART_TYPES.includes(v);

    // Charts that need error column selector (pre-calculated error data)
    const needsErrorColumn = v === "line_grouped_error" ||
                            v === "bar_error" ||
                            v === "bar_grouped_error";

    if (v === "ic50_dose_response") {
      lx.textContent = "Concentration";
      ly.textContent = "Response (%)";
      // Set x-axis to log10 scale by default for dose-response curves
      const xScaleSelect = document.getElementById("xScale");
      if (xScaleSelect) {
        xScaleSelect.value = "log10";
      }
    } else if (v === "box") {
      lx.textContent = "X Axis (optional)";
      ly.textContent = "Y Axis";
    } else if (v === "bar") {
      lx.textContent = "X Axis";
      ly.textContent = "Y Axis";
    } else if (isGrouped) {
      lx.textContent = "X Axis";
      ly.textContent = "Y Axis";
      // Show group color controls for all grouped chart variants
      if (groupColorRow) groupColorRow.style.display = "flex";
      // Show group column selector for grouped charts
      if (groupColumnRow) groupColumnRow.style.display = "flex";
      // Set default bar width to 0.9 for grouped charts
      const barWidthInput = document.getElementById("barWidth");
      if (barWidthInput) {
        // Always set to 0.9 for grouped charts (override the 0.4 HTML default)
        barWidthInput.value = "0.9";
      }
    } else {
      // histogram and others
      lx.textContent = "X Axis";
      ly.textContent = "Y Axis";
    }

    // Hide group controls for non-grouped chart types
    if (!isGrouped) {
      if (groupColorRow) groupColorRow.style.display = "none";
      if (groupColumnRow) groupColumnRow.style.display = "none";
      // Reset bar width to default for non-grouped charts (only if it was 0.9)
      const barWidthInput = document.getElementById("barWidth");
      if (barWidthInput && barWidthInput.value === "0.9") {
        barWidthInput.value = "0.4";
      }
    }

    // Show/hide error column selector
    if (needsErrorColumn) {
      if (errorColumnRow) errorColumnRow.style.display = "flex";
    } else {
      if (errorColumnRow) errorColumnRow.style.display = "none";
    }

    // Update order containers based on chart type
    const groupOrderContainer = document.getElementById("groupOrderContainer");
    const xAxisOrderContainer = document.getElementById("xAxisOrderContainer");
    const groupOrderLabel = document.getElementById("groupOrderLabel");

    if (isGrouped) {
      // Grouped charts: show both Group order and X-axis order
      if (groupOrderLabel) groupOrderLabel.textContent = "Group order:";
      if (groupOrderContainer) groupOrderContainer.style.display = "block";
      if (xAxisOrderContainer) xAxisOrderContainer.style.display = "block";
    } else {
      // Single-group charts: show only one list labeled "X-axis order"
      if (groupOrderLabel) groupOrderLabel.textContent = "X-axis order:";
      if (groupOrderContainer) groupOrderContainer.style.display = "block";
      if (xAxisOrderContainer) xAxisOrderContainer.style.display = "none";
    }

    // Show/hide bins row (only for histogram)
    const binsRow = document.getElementById("binsRow");
    if (binsRow) {
      if (v === "histogram") {
        binsRow.style.display = "flex";
      } else {
        binsRow.style.display = "none";
      }
    }

    // Repopulate comparison checkboxes when chart type changes
    // This ensures line plots show time-point based comparisons
    if (typeof populateComparisonCheckboxes === 'function') {
      populateComparisonCheckboxes();
    }

    // Update statistics availability based on chart type
    updateStatisticsAvailability();

    // Update detailed statistics controls visibility based on checkbox state
    updateStatsDetailedControlsVisibility();

    // Update VBracket controls visibility
    updateVbracketVisibility();

    // Update Colors & Style tab controls visibility
    updateColorsStyleControlsVisibility();

    // Update IC50 Analysis controls
    updateIC50AnalysisVisibility();
  }

  // IC50 Analysis visibility control
  function updateIC50AnalysisVisibility() {
    const chartType = document.getElementById("chartType")?.value || "";
    const enableIC50Checkbox = document.getElementById("enableIC50Analysis");
    const ic50OptionsContainer = document.getElementById("ic50OptionsContainer");
    const ic50StyleSection = document.getElementById("ic50StyleSection");

    if (chartType === "ic50_dose_response") {
      // Auto-check and show options when IC50 chart type is selected
      if (enableIC50Checkbox) {
        enableIC50Checkbox.checked = true;
      }
      if (ic50OptionsContainer) {
        ic50OptionsContainer.style.display = "block";
      }
      // Show IC50 styling in Colors & Style tab
      if (ic50StyleSection) {
        ic50StyleSection.style.display = "block";
      }
    } else {
      // Uncheck and hide options for other chart types
      if (enableIC50Checkbox) {
        enableIC50Checkbox.checked = false;
      }
      if (ic50OptionsContainer) {
        ic50OptionsContainer.style.display = "none";
      }
      // Hide IC50 styling in Colors & Style tab
      if (ic50StyleSection) {
        ic50StyleSection.style.display = "none";
      }
    }
  }

  // IC50 checkbox toggle handler
  function updateIC50OptionsVisibility() {
    const enableIC50Checkbox = document.getElementById("enableIC50Analysis");
    const ic50OptionsContainer = document.getElementById("ic50OptionsContainer");

    if (enableIC50Checkbox && ic50OptionsContainer) {
      ic50OptionsContainer.style.display = enableIC50Checkbox.checked ? "block" : "none";
    }
  }

  // Set up IC50 checkbox listener
  document.getElementById("enableIC50Analysis")?.addEventListener("change", updateIC50OptionsVisibility);

  // Set up chart type change listener
  document.getElementById("chartType")?.addEventListener("change", handleChartTypeChange);

  // Trigger on initial page load to handle default selected chart type
  handleChartTypeChange();

  // Set up statistics checkbox listener to update detailed controls and VBracket visibility
  document.getElementById("addStatistics")?.addEventListener("change", function() {
    updateStatsDetailedControlsVisibility();
    updateVbracketVisibility();
  });

  // vbracket position dropdown removed - now uses manual X/Y positioning only

  // Set up group column change listener to auto-update number of groups and custom order
  document.getElementById("groupColumn")?.addEventListener("change", function() {
    const actualGroups = getActualGroupNames();
    if (actualGroups && actualGroups.length >= 2) {
      const numGroupsInput = document.getElementById("numGroups");
      if (numGroupsInput && parseInt(numGroupsInput.value) !== actualGroups.length) {
        numGroupsInput.value = actualGroups.length;
        // Trigger change event to show/hide color inputs
        numGroupsInput.dispatchEvent(new Event('change'));
        console.log("🔥 Auto-updated numGroups to:", actualGroups.length, "when Group Column changed");
      }
    }
    // Update custom order lists when column changes
    if (typeof detectAndStoreGroups === 'function') {
      detectAndStoreGroups();
    }
  });

  // Set up X column change listener to update custom order
  document.getElementById("xColumn")?.addEventListener("change", function() {
    // Update custom order lists when column changes
    if (typeof detectAndStoreGroups === 'function') {
      detectAndStoreGroups();
    }
  });

  // カラーピッカーとテキストの同期
  wireColorSync();

  // テキスト書式設定のヒントを追加
  addFormattingHints();

  // ハンドラの重複登録を防止
  if (window.__SATO_HANDLERS_BOUND__) return;
  window.__SATO_HANDLERS_BOUND__ = true;

  // 基本操作 - 修正版ハンドラーを使用
  document.getElementById("load")?.addEventListener("click", loadHeadersFromSelection);

  // Figure save/load with metadata
  document.getElementById("saveFigure")?.addEventListener("click", saveFigureWithMetadata);
  document.getElementById("loadFromFigure")?.addEventListener("click", loadFromFigure);
  document.getElementById("downloadMetadataJson")?.addEventListener("click", downloadMetadataJson);

  // 修正版のイベントハンドラーを適用
  updateEventHandlersWithDebug();

  // Initialize preset list on load
  refreshPresetList().catch(e => {
    console.log("Preset list initialization:", e);
  });

  document.getElementById("exportStatResults")?.addEventListener("click", async () => {
    try {
      await exportStatisticalResults();
    } catch (e) {
      console.error("Statistical export failed:", e);
      setStatus("Statistical results export error: " + (e?.message || e));
    }
  });

  // Data order control
  document.getElementById("dataOrder")?.addEventListener("change", function() {
    const customOrderContainer = document.getElementById("customOrderContainer");
    if (this.value === "custom") {
      customOrderContainer.style.display = "block";
      // Populate both lists if data is loaded
      window.detectAndStoreGroups();
    } else {
      customOrderContainer.style.display = "none";
    }
  });

  // Global variables to store detected groups and categories
  window.detectedGroups = [];
  window.detectedCategories = [];

  // Function to populate custom order list with draggable items
  function populateCustomOrderList(listId, items, emptyMessage) {
    const listContainer = document.getElementById(listId);
    if (!items || items.length === 0) {
      listContainer.innerHTML = `<div style="font-size:11px; color:#999; text-align:center;">${emptyMessage}</div>`;
      return;
    }

    // Clear and populate with draggable items
    listContainer.innerHTML = '';
    items.forEach((item, index) => {
      const div = document.createElement('div');
      div.className = 'draggable-item';
      div.draggable = true;
      div.textContent = item;
      div.dataset.value = item;
      div.dataset.listId = listId;
      div.style.cssText = 'padding:6px 10px; margin:4px 0; background:#fff; border:1px solid #ccc; border-radius:4px; cursor:move; user-select:none; font-size:12px;';

      // Drag event handlers
      div.addEventListener('dragstart', handleDragStart);
      div.addEventListener('dragover', handleDragOver);
      div.addEventListener('drop', handleDrop);
      div.addEventListener('dragend', handleDragEnd);

      listContainer.appendChild(div);
    });
  }

  // Drag-and-drop event handlers
  let draggedElement = null;

  function handleDragStart(e) {
    draggedElement = this;
    this.style.opacity = '0.5';
    e.dataTransfer.effectAllowed = 'move';
  }

  function handleDragOver(e) {
    if (e.preventDefault) {
      e.preventDefault();
    }
    e.dataTransfer.dropEffect = 'move';
    return false;
  }

  function handleDrop(e) {
    if (e.stopPropagation) {
      e.stopPropagation();
    }

    if (draggedElement !== this && draggedElement.dataset.listId === this.dataset.listId) {
      const listContainer = this.parentNode;
      const allItems = [...listContainer.children];
      const draggedIndex = allItems.indexOf(draggedElement);
      const targetIndex = allItems.indexOf(this);

      if (draggedIndex < targetIndex) {
        listContainer.insertBefore(draggedElement, this.nextSibling);
      } else {
        listContainer.insertBefore(draggedElement, this);
      }

      // Update the order arrays based on which list was modified
      updateGroupOrder(this.dataset.listId);
    }

    return false;
  }

  function handleDragEnd(e) {
    this.style.opacity = '1';
    draggedElement = null;
  }

  function updateGroupOrder(listId) {
    const listContainer = document.getElementById(listId);
    const items = listContainer.querySelectorAll('.draggable-item');
    const orderedValues = Array.from(items).map(item => item.dataset.value);

    if (listId === 'customOrderGroupList') {
      window.detectedGroups = orderedValues;
    } else if (listId === 'customOrderCategoryList') {
      window.detectedCategories = orderedValues;
    }
  }

  // Functions to get custom orders as comma-separated strings - make them globally accessible
  window.getCustomOrderGroupString = function() {
    return window.detectedGroups && window.detectedGroups.length > 0 ? window.detectedGroups.join(', ') : '';
  };

  window.getCustomOrderCategoryString = function() {
    return window.detectedCategories && window.detectedCategories.length > 0 ? window.detectedCategories.join(', ') : '';
  };

  // Function to detect groups and categories from loaded data (make it globally accessible)
  window.detectAndStoreGroups = async function() {
    try {
      const { values } = await getSelectedValues();
      if (!values || values.length < 2) return;

      const headers = values[0];

      // Detect groups from Column 1
      const groupCol = document.getElementById("groupColumn")?.value;
      if (groupCol) {
        const groupColIndex = headers.indexOf(groupCol);
        if (groupColIndex !== -1) {
          const groups = new Set();
          for (let i = 1; i < values.length; i++) {
            const value = values[i][groupColIndex];
            if (value !== null && value !== undefined && value !== "") {
              groups.add(String(value));
            }
          }
          window.detectedGroups = Array.from(groups);
          console.log("Detected groups:", window.detectedGroups);
        }
      }

      // Detect categories from X-axis column
      const xCol = document.getElementById("xColumn")?.value;
      if (xCol) {
        const xColIndex = headers.indexOf(xCol);
        if (xColIndex !== -1) {
          const categories = new Set();
          for (let i = 1; i < values.length; i++) {
            const value = values[i][xColIndex];
            if (value !== null && value !== undefined && value !== "") {
              categories.add(String(value));
            }
          }
          window.detectedCategories = Array.from(categories);
          console.log("Detected categories:", window.detectedCategories);
        }
      }

      // Check if this is a grouped chart type
      const chartType = document.getElementById("chartType")?.value || "";
      const isGrouped = GROUPED_CHART_TYPES.includes(chartType);

      // If custom order is currently selected, update lists
      const dataOrder = document.getElementById("dataOrder")?.value;
      if (dataOrder === "custom") {
        if (isGrouped) {
          // For grouped charts: populate both lists
          populateCustomOrderList('customOrderGroupList', window.detectedGroups, 'Load data to see groups');
          populateCustomOrderList('customOrderCategoryList', window.detectedCategories, 'Load data to see values');
        } else {
          // For single-group charts: use customOrderGroupList for X-axis values
          // Put X-axis values into detectedGroups (since that list works)
          window.detectedGroups = [...window.detectedCategories];
          populateCustomOrderList('customOrderGroupList', window.detectedGroups, 'Load data to see values');
        }
      }
    } catch (error) {
      console.error("Error detecting groups:", error);
    }
  };

  // Dynamic group color picker synchronization
  function setupGroupColorSync() {
    for (let i = 1; i <= 6; i++) {
      const colorInput = document.getElementById(`groupColor${i}`);
      const colorPicker = document.getElementById(`groupPicker${i}`);
      
      if (colorPicker) {
        colorPicker.addEventListener("change", function() {
          if (colorInput) colorInput.value = this.value;
        });
      }
      
      if (colorInput) {
        colorInput.addEventListener("change", function() {
          const colorValue = this.value;
          if (colorValue.startsWith("#") && colorPicker) {
            colorPicker.value = colorValue;
          }
        });
      }
    }
  }

  // Dot color picker sync
  function setupDotColorSync() {
    const dotColorInput = document.getElementById("dotColor");
    const dotColorPicker = document.getElementById("dotColorPicker");
    
    if (dotColorPicker) {
      dotColorPicker.addEventListener("change", function() {
        if (dotColorInput) dotColorInput.value = this.value;
      });
    }
    
    if (dotColorInput) {
      dotColorInput.addEventListener("change", function() {
        const colorValue = this.value;
        if (colorValue.startsWith("#") && dotColorPicker) {
          dotColorPicker.value = colorValue;
        }
      });
    }
  }
  
  // Number of groups control
  document.getElementById("numGroups")?.addEventListener("change", function() {
    const numGroups = parseInt(this.value);
    for (let i = 1; i <= 6; i++) {
      const groupItem = document.querySelector(`.group-color-item[data-group="${i}"]`);
      if (groupItem) {
        groupItem.style.display = i <= numGroups ? "flex" : "none";
      }
    }
  });
  
  setupGroupColorSync();
  setupDotColorSync();
  

  // Symbol size control
  document.getElementById("statSymbolSize")?.addEventListener("input", function() {
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      clearTimeout(window.symbolSizeUpdateTimeout);
      window.symbolSizeUpdateTimeout = setTimeout(() => {
        previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
      }, 300);
    }
  });

  // Note: Symbol size for ggpubr brackets is now handled by the existing statSymbolSize control

  document.getElementById("statLineSize")?.addEventListener("input", function() {
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      clearTimeout(window.statLineUpdateTimeout);
      window.statLineUpdateTimeout = setTimeout(() => {
        previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
      }, 300);
    }
  });

  document.getElementById("statTipLength")?.addEventListener("input", function() {
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      clearTimeout(window.statTipUpdateTimeout);
      window.statTipUpdateTimeout = setTimeout(() => {
        previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
      }, 300);
    }
  });

  document.getElementById("statVjust")?.addEventListener("input", function() {
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      clearTimeout(window.statVjustUpdateTimeout);
      window.statVjustUpdateTimeout = setTimeout(() => {
        previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
      }, 300);
    }
  });

  // Reset statistical parameters to default values
  document.getElementById("resetStatParams")?.addEventListener("click", function() {
    document.getElementById("statSymbolSize").value = "7";
    document.getElementById("statLineSize").value = "1.0";
    document.getElementById("statTipLength").value = "0.04";
    document.getElementById("statVjust").value = "-0.3";

    // Auto-update chart if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  });

  // Comparison mode controls
  document.querySelectorAll('input[name="comparisonMode"]').forEach(radio => {
    radio.addEventListener("change", function() {
      const customArea = document.getElementById("customComparisonArea");
      if (this.value === "custom") {
        customArea.style.display = "block";
        populateComparisonCheckboxes();
      } else {
        customArea.style.display = "none";
      }

      // Update chart if statistics are enabled
      const addStatistics = document.getElementById("addStatistics")?.checked;
      if (addStatistics && window.lastRender) {
        clearTimeout(window.comparisonModeUpdateTimeout);
        window.comparisonModeUpdateTimeout = setTimeout(() => {
          previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
        }, 300);
      }
    });
  });

  document.getElementById("statSymbolType")?.addEventListener("change", function() {
    // Show/hide custom symbol inputs
    const customSymbolInputs = document.getElementById("customSymbolInputs");
    const symbolType = this.value;

    if (customSymbolInputs) {
      customSymbolInputs.style.display = symbolType === "custom" ? "block" : "none";
    }

    // Change default symbol size to 4 when p-value or custom symbols are selected
    const statSymbolSize = document.getElementById("statSymbolSize");
    if (statSymbolSize && (symbolType === "custom" || symbolType === "pvalue")) {
      statSymbolSize.value = "4";
    }

    // Change vertical position adjustment to -0.8 for p-value and custom symbols
    const statVjust = document.getElementById("statVjust");
    if (statVjust && (symbolType === "custom" || symbolType === "pvalue")) {
      statVjust.value = "-0.8";
    }

    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  });

  // Show/hide manual test options based on statistical test mode selection
  document.getElementById("statisticalTestMode")?.addEventListener("change", function() {
    const mode = this.value;
    const manualTestOptions = document.getElementById("manualTestOptions");
    const autoModeExplanation = document.getElementById("autoModeExplanation");

    if (mode === "auto") {
      // Auto mode - hide manual options, show explanation
      manualTestOptions.style.display = "none";
      if (autoModeExplanation) autoModeExplanation.style.display = "block";
    } else {
      // Manual mode - show manual options, hide explanation
      manualTestOptions.style.display = "block";
      if (autoModeExplanation) autoModeExplanation.style.display = "none";
    }

    // Auto-update if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  });

  // Auto-update when manual test type changes
  document.getElementById("statisticalTest")?.addEventListener("change", function() {
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  });

  // Switch between parametric and non-parametric post-hoc options
  document.getElementById("dataTypeSelect")?.addEventListener("change", function() {
    const dataType = this.value;
    const parametricPostHoc = document.getElementById("parametricPostHoc");
    const nonparametricPostHoc = document.getElementById("nonparametricPostHoc");
    const dataTypeHint = document.getElementById("dataTypeHint");

    const varianceTestSection = document.getElementById("varianceTestSection");

    if (dataType === "parametric") {
      parametricPostHoc.style.display = "block";
      nonparametricPostHoc.style.display = "none";
      if (varianceTestSection) varianceTestSection.style.display = "block";
      dataTypeHint.textContent = "2 groups: t-test, 3+ groups: ANOVA";
    } else {
      parametricPostHoc.style.display = "none";
      nonparametricPostHoc.style.display = "block";
      if (varianceTestSection) varianceTestSection.style.display = "none";
      dataTypeHint.textContent = "2 groups: Wilcoxon, 3+ groups: Kruskal-Wallis";
    }

    // Update control group visibility based on current post-hoc selection
    updateControlGroupVisibility();

    // Auto-update if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  });

  // Helper function to get the currently active post-hoc test value
  function getActivePostHocTest() {
    const dataType = document.getElementById("dataTypeSelect")?.value || "parametric";
    if (dataType === "parametric") {
      return document.getElementById("postHocTest")?.value || "tukey";
    } else {
      return document.getElementById("postHocTestNonparam")?.value || "dunn";
    }
  }

  // Helper function to update control group visibility
  function updateControlGroupVisibility() {
    const postHocTest = getActivePostHocTest();
    const dunnettControlGroup = document.getElementById("dunnettControlGroup");

    if (postHocTest === "dunnett" || postHocTest === "steel") {
      dunnettControlGroup.style.display = "block";
      populateDunnettControl();
    } else {
      dunnettControlGroup.style.display = "none";
    }

    // Repopulate comparison checkboxes
    populateComparisonCheckboxes();
  }

  // Show/hide control group selection when Dunnett is selected (parametric)
  document.getElementById("postHocTest")?.addEventListener("change", function() {
    updateControlGroupVisibility();

    // Auto-update if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  });

  // Show/hide control group selection when Steel is selected (non-parametric)
  document.getElementById("postHocTestNonparam")?.addEventListener("change", function() {
    updateControlGroupVisibility();

    // Auto-update if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  });

  // Populate Dunnett control group dropdown with groups from current data
  function populateDunnettControl() {
    const dunnettControl = document.getElementById("dunnettControl");
    if (!dunnettControl) return;

    // Get actual group names from current data
    const groups = getActualGroupNames();
    if (groups.length < 2) {
      dunnettControl.innerHTML = "<option value=''>少なくとも2つのグループが必要です</option>";
      return;
    }

    // Clear and populate dropdown
    dunnettControl.innerHTML = "";
    groups.forEach((group, index) => {
      const option = document.createElement("option");
      option.value = group;
      option.textContent = group;
      if (index === 0) option.selected = true; // Select first group by default
      dunnettControl.appendChild(option);
    });

    // Repopulate comparison checkboxes when control group changes
    dunnettControl.removeEventListener("change", handleDunnettControlChange);
    dunnettControl.addEventListener("change", handleDunnettControlChange);
  }

  // Handler for dunnett control group changes
  function handleDunnettControlChange() {
    // Repopulate comparison checkboxes with new control group
    populateComparisonCheckboxes();

    // Auto-update if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (addStatistics && window.lastRender) {
      previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
    }
  }

  // Load and apply number of groups setting
  function updateGroupColorVisibility() {
    const numGroups = parseInt(document.getElementById("numGroups")?.value) || 2;
    for (let i = 1; i <= 6; i++) {
      const groupItem = document.querySelector(`.group-color-item[data-group="${i}"]`);
      if (groupItem) {
        groupItem.style.display = i <= numGroups ? "flex" : "none";
      }
    }
  }
  
  // Initial visibility setup
  updateGroupColorVisibility();
  
  // Initial bar width setup based on chart type
  function updateBarWidthDefault() {
    const chartType = document.getElementById("chartType")?.value;
    const barWidthInput = document.getElementById("barWidth");
    if (GROUPED_CHART_TYPES.includes(chartType) && barWidthInput) {
      barWidthInput.value = "0.9";
    }
  }
  updateBarWidthDefault();

  // 起動時の軽い初期化（失敗は無視）
  getSelectedValues().then(() => loadHeadersFromSelection()).catch(() => {});
});import { Canvg } from "canvg";

// Excel × WebR（R in WebAssembly）— カスタマイズ対応版
let webR, webrReady = false;

function el(id){ return document.getElementById(id); }
function setStatus(msg) {
  const el = document.getElementById("status");
  if (el) { el.textContent = msg; } else { console.log(msg); }
}

// Function to get actual group names from current data
function getActualGroupNames() {
  try {
    // Check if this is a line plot or other grouped chart with explicit group column
    const chartType = document.getElementById("chartType")?.value || "";
    const isLinePlot = chartType === "line_grouped_error_raw" || chartType === "line_grouped_error";
    const isGroupedChart = chartType.includes("grouped");

    // For line plots and explicit grouped charts, use the Group column selector
    if ((isLinePlot || isGroupedChart) && window.lastProcessedData && window.lastProcessedData.length > 0) {
      const groupColumn = document.getElementById("groupColumn")?.value;
      if (groupColumn) {
        const headers = window.lastProcessedData[0];
        const groupIndex = headers.findIndex(h => h === groupColumn);
        if (groupIndex >= 0) {
          const groupData = window.lastProcessedData.slice(1).map(row => row[groupIndex]);
          const uniqueGroups = [...new Set(groupData)].filter(g => g && g.trim() !== '');
          console.log("🔥 Found groups from Group column selector:", uniqueGroups);
          if (uniqueGroups.length >= 2) {
            return uniqueGroups;
          }
        }
      }
    }

    // Fallback: Try to get from last processed data first column
    if (window.lastProcessedData && window.lastProcessedData.length > 0) {
      const firstColumnData = window.lastProcessedData.map(row => row[0]);
      // Remove header and get unique values
      const uniqueGroups = [...new Set(firstColumnData.slice(1))].filter(g => g && g.trim() !== '');
      console.log("🔥 Found actual groups from lastProcessedData (first column):", uniqueGroups);

      // Only return if we have at least 2 groups
      if (uniqueGroups.length >= 2) {
        return uniqueGroups;
      }
    }

    // Fallback: try to get from current UI column selections
    const xColumn = document.getElementById("xColumn")?.value;
    if (xColumn && window.lastProcessedData && window.lastProcessedData.length > 0) {
      // Find the column index for the selected x column
      const headers = window.lastProcessedData[0];
      const xIndex = headers.findIndex(h => h === xColumn);
      if (xIndex >= 0) {
        const columnData = window.lastProcessedData.slice(1).map(row => row[xIndex]);
        const uniqueGroups = [...new Set(columnData)].filter(g => g && g.trim() !== '');
        console.log("🔥 Found groups from X column selection:", uniqueGroups);
        if (uniqueGroups.length >= 2) {
          return uniqueGroups;
        }
      }
    }

    console.log("🔥 No sufficient data available, using empty groups list");
    return []; // Return empty instead of fallback to force user to load data
  } catch (error) {
    console.log("🔥 Error getting actual group names:", error);
    return []; // Return empty on error
  }
}

// Helper function to get number of groups
async function getNumberOfGroups() {
  try {
    const groups = await getActualGroupNames();
    return groups.length;
  } catch (error) {
    console.log("Error getting number of groups:", error);
    return 0;
  }
}

// Function to populate comparison checkboxes dynamically
function populateVbracketTimepoints() {
  const timepointSelect = document.getElementById("vbracketTimepoint");
  if (!timepointSelect) return;

  // Get current selection before clearing
  const currentValue = timepointSelect.value;

  // Clear existing options
  timepointSelect.innerHTML = '<option value="">-- Select time point --</option>';

  // Only populate for line_grouped_error_raw chart type
  const chartType = document.getElementById("chartType")?.value || "";
  if (chartType !== "line_grouped_error_raw") {
    timepointSelect.innerHTML = '<option value="">-- Not applicable --</option>';
    return;
  }

  // Get unique X values (time points) from the data
  if (window.lastProcessedData && window.lastProcessedData.length > 1) {
    const headers = window.lastProcessedData[0];
    const selectedXColumn = document.getElementById("xColumn")?.value || "";
    const xColIndex = headers.findIndex(h => h === selectedXColumn);

    if (xColIndex >= 0) {
      // Get unique X values from the data
      const xValues = new Set();
      for (let i = 1; i < window.lastProcessedData.length; i++) {
        const xVal = window.lastProcessedData[i][xColIndex];
        if (xVal !== "" && xVal !== null && xVal !== undefined) {
          xValues.add(xVal);
        }
      }

      // Sort X values (numerically if possible, otherwise alphabetically)
      const sortedXValues = Array.from(xValues).sort((a, b) => {
        const numA = Number(a);
        const numB = Number(b);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        return String(a).localeCompare(String(b));
      });

      // Add options for each time point
      sortedXValues.forEach(xVal => {
        const option = new Option(xVal, xVal);
        timepointSelect.add(option);
      });

      // Restore previous selection if it still exists, otherwise select first
      if (currentValue && sortedXValues.includes(currentValue)) {
        timepointSelect.value = currentValue;
      } else if (sortedXValues.length > 0) {
        timepointSelect.value = sortedXValues[0];
      }

      console.log(`📊 Populated vbracket timepoints: ${sortedXValues.join(", ")}`);
    }
  }
}

function populateComparisonCheckboxes() {
  const checkboxContainer = document.getElementById("comparisonCheckboxes");
  if (!checkboxContainer) return;

  // Save current selections and positions before clearing
  const currentSelections = {};
  const currentPositions = {};
  const existingCheckboxes = checkboxContainer.querySelectorAll('input[type="checkbox"]');
  existingCheckboxes.forEach((checkbox, index) => {
    currentSelections[checkbox.value] = checkbox.checked;
    const posInput = document.getElementById(`pos_${index}`);
    if (posInput) {
      currentPositions[checkbox.value] = posInput.value;
    }
  });

  // Clear existing checkboxes
  checkboxContainer.innerHTML = "";

  // Check if this is a line plot with time series or grouped chart with categories
  const chartType = document.getElementById("chartType")?.value || "";
  const isLinePlot = chartType === "line_grouped_error_raw" || chartType === "line_grouped_error";
  const isGroupedWithCategories = chartType === "bar_grouped_error_dot" ||
                                   chartType === "box_grouped" ||
                                   chartType === "box_grouped_dot" ||
                                   chartType === "violin_grouped" ||
                                   chartType === "violin_grouped_dot";

  // Get actual group names from current data
  const actualGroups = getActualGroupNames();
  console.log("🔍 DEBUG: Column Selectors - Group:", document.getElementById("groupColumn")?.value, "X:", document.getElementById("xColumn")?.value, "Y:", document.getElementById("yColumn")?.value);
  console.log("🔍 DEBUG: Actual Groups detected:", actualGroups);
  if (actualGroups.length < 2) {
    checkboxContainer.innerHTML = "<p style='font-size:10px; color:#666;'>少なくとも2つのグループが必要です</p>";
    return;
  }

  // Generate pairwise combinations
  const combinations = [];

  if (isLinePlot || isGroupedWithCategories) {
    // For line plots and grouped charts: generate comparisons for each category/time point
    // Get unique X values (categories/time points) from the data
    if (window.lastProcessedData && window.lastProcessedData.length > 1) {
      const headers = window.lastProcessedData[0];
      const selectedXColumn = document.getElementById("xColumn")?.value || "";
      const xColIndex = headers.findIndex(h => h === selectedXColumn);

      if (xColIndex >= 0) {
        // Get unique X values from the data
        const xValues = new Set();
        console.log("🔍 DEBUG: Selected X Column:", selectedXColumn, "Index:", xColIndex);
        console.log("🔍 DEBUG: Headers:", headers);
        for (let i = 1; i < window.lastProcessedData.length; i++) {
          const xVal = window.lastProcessedData[i][xColIndex];
          if (xVal !== "" && xVal !== null && xVal !== undefined) {
            xValues.add(xVal);
          }
        }
        console.log("🔍 DEBUG: Extracted X Values:", Array.from(xValues));

        const sortedXValues = Array.from(xValues).sort((a, b) => {
          const numA = Number(a);
          const numB = Number(b);
          if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
          return String(a).localeCompare(String(b));
        });

        // For 2 groups: create "Group1 vs Group2 @ X" comparisons
        if (actualGroups.length === 2) {
          sortedXValues.forEach(xVal => {
            combinations.push([actualGroups[0], actualGroups[1], xVal]);
          });
        } else {
          // For 3+ groups: create all pairwise comparisons for each X value (category)
          sortedXValues.forEach(xVal => {
            for (let i = 0; i < actualGroups.length; i++) {
              for (let j = i + 1; j < actualGroups.length; j++) {
                combinations.push([actualGroups[i], actualGroups[j], xVal]);
              }
            }
          });
        }
      }
    }
  } else {
    // Standard grouped chart comparisons
    const dataType = document.getElementById("dataTypeSelect")?.value || "parametric";
    const postHocTest = dataType === "nonparametric"
      ? (document.getElementById("postHocTestNonparam")?.value || "dunn")
      : (document.getElementById("postHocTest")?.value || "tukey");
    const isVsControl = postHocTest === "dunnett" || postHocTest === "steel";
    const controlGroup = document.getElementById("dunnettControl")?.value || "";

    if (isVsControl && controlGroup) {
      // For Dunnett/Steel: only show comparisons with the control group
      for (let i = 0; i < actualGroups.length; i++) {
        if (actualGroups[i] !== controlGroup) {
          combinations.push([controlGroup, actualGroups[i]]);
        }
      }
    } else {
      // For other tests: show all pairwise combinations
      for (let i = 0; i < actualGroups.length; i++) {
        for (let j = i + 1; j < actualGroups.length; j++) {
          combinations.push([actualGroups[i], actualGroups[j]]);
        }
      }
    }
  }

  // Create checkboxes and position controls for each comparison
  combinations.forEach((comp, index) => {
    const label = document.createElement("label");
    label.style.cssText = "display:flex; align-items:center; gap:5px; font-size:10px; margin-bottom:5px;";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.id = `comp_${index}`;

    // For line plots with time points: comp = [group1, group2, xValue]
    // For standard grouped charts: comp = [group1, group2]
    if (comp.length === 3) {
      // Line plot format: "Group1-Group2@X"
      checkbox.value = `${comp[0]}-${comp[1]}@${comp[2]}`;
    } else {
      // Standard format: "Group1-Group2"
      checkbox.value = `${comp[0]}-${comp[1]}`;
    }
    console.log("🔥 Creating checkbox with value:", checkbox.value);

    // Restore previous selection or default to checked for new items
    checkbox.checked = currentSelections.hasOwnProperty(checkbox.value) ?
                      currentSelections[checkbox.value] : true;

    // Add event listener for real-time updates
    checkbox.addEventListener("change", function() {
      // Update position input visibility based on checkbox state
      const posInput = document.getElementById(`pos_${index}`);
      if (posInput) {
        posInput.style.visibility = this.checked ? 'visible' : 'hidden';
        posInput.style.opacity = this.checked ? '1' : '0.3';
      }

      const addStatistics = document.getElementById("addStatistics")?.checked;
      if (addStatistics && window.lastRender) {
        clearTimeout(window.customCompUpdateTimeout);
        window.customCompUpdateTimeout = setTimeout(() => {
          previewPlotWithDebug().catch(e => console.error("Auto-update failed:", e));
        }, 300);
      }
    });

    const text = document.createElement("span");
    // Display format based on comparison type
    if (comp.length === 3) {
      // Grouped chart with categories: "Group1 vs Group2 @ Category"
      text.textContent = `${comp[0]} vs ${comp[1]} @ ${comp[2]}`;
    } else {
      // Standard: "Group1 vs Group2"
      text.textContent = `${comp[0]} vs ${comp[1]}`;
    }
    text.style.cssText = "flex: 1; user-select: none; min-width: 120px;";

    // Add position control with absolute Y values
    const positionInput = document.createElement("input");
    positionInput.type = "number";
    positionInput.id = `pos_${index}`;
    positionInput.min = "0";
    positionInput.max = "10000";
    positionInput.step = "0.01";

    // Calculate default absolute Y position
    // Improved fallback: calculate from actual data range if available
    let defaultAbsoluteY = 100 + index * 15; // Fallback: Start at 100, increment by 15

    // Try to calculate better fallback from data
    if (window.lastProcessedData && window.lastProcessedData.length > 1) {
      const headers = window.lastProcessedData[0];
      const yColumnName = document.getElementById("yColumn")?.value;
      const yIdx = headers.findIndex(h => h === yColumnName);

      if (yIdx >= 0) {
        const allYValues = [];
        for (let i = 1; i < window.lastProcessedData.length; i++) {
          const yVal = parseFloat(window.lastProcessedData[i][yIdx]);
          if (!isNaN(yVal)) allYValues.push(yVal);
        }

        if (allYValues.length > 0) {
          const maxY = Math.max(...allYValues);
          const range = Math.max(...allYValues) - Math.min(...allYValues);
          // Start 20% above max, stack with 10% spacing
          defaultAbsoluteY = Math.round((maxY + (range * 0.20) + (index * range * 0.10)) * 100) / 100;
        }
      }
    }

    // For line plots and grouped charts, calculate actual position based on data
    if ((isLinePlot || isGroupedWithCategories) && comp.length === 3 && window.lastProcessedData && window.lastProcessedData.length > 1) {
      const headers = window.lastProcessedData[0];
      const groupColumnName = document.getElementById("groupColumn")?.value;
      const xColumnName = document.getElementById("xColumn")?.value;
      const yColumnName = document.getElementById("yColumn")?.value;

      const groupIdx = headers.findIndex(h => h === groupColumnName);
      const xIdx = headers.findIndex(h => h === xColumnName);
      const yIdx = headers.findIndex(h => h === yColumnName);

      if (groupIdx >= 0 && xIdx >= 0 && yIdx >= 0) {
        const xValue = comp[2]; // X value for this comparison

        if (isGroupedWithCategories) {
          // For grouped charts with categories: find maximum value across ALL groups at this X value
          let maxBarTop = 0;

          // Get all unique groups
          const allGroups = new Set();
          for (let i = 1; i < window.lastProcessedData.length; i++) {
            const group = window.lastProcessedData[i][groupIdx];
            if (group) allGroups.add(String(group));
          }

          // For each group, calculate mean + error at this X value
          allGroups.forEach(group => {
            const dataPoints = [];
            for (let i = 1; i < window.lastProcessedData.length; i++) {
              const row = window.lastProcessedData[i];
              if (String(row[groupIdx]) === group && String(row[xIdx]) === String(xValue)) {
                const yVal = parseFloat(row[yIdx]);
                if (!isNaN(yVal)) {
                  dataPoints.push(yVal);
                }
              }
            }

            if (dataPoints.length > 0) {
              const mean = dataPoints.reduce((a, b) => a + b, 0) / dataPoints.length;
              let sd = 0;
              if (dataPoints.length > 1) {
                const variance = dataPoints.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / dataPoints.length;
                sd = Math.sqrt(variance);
              }
              const barTop = mean + sd;
              if (barTop > maxBarTop) maxBarTop = barTop;
            }
          });

          if (maxBarTop > 0) {
            // Position ABOVE the tallest bar + error with spacing (20% above + stacking)
            const dataRange = maxBarTop - Math.min(...Array.from(allGroups).map(group => {
              const dataPoints = [];
              for (let i = 1; i < window.lastProcessedData.length; i++) {
                const row = window.lastProcessedData[i];
                if (String(row[groupIdx]) === group && String(row[xIdx]) === String(xValue)) {
                  const yVal = parseFloat(row[yIdx]);
                  if (!isNaN(yVal)) dataPoints.push(yVal);
                }
              }
              return dataPoints.length > 0 ? Math.min(...dataPoints) : Infinity;
            }));
            defaultAbsoluteY = Math.round((maxBarTop + (dataRange * 0.20) + (index * dataRange * 0.10)) * 100) / 100;
          }
        } else {
          // For line plots: position below the line (original logic for group2)
          const group2 = comp[1]; // Treatment group (second group)
          const dataPoints = [];
          for (let i = 1; i < window.lastProcessedData.length; i++) {
            const row = window.lastProcessedData[i];
            if (String(row[groupIdx]) === String(group2) && String(row[xIdx]) === String(xValue)) {
              const yVal = parseFloat(row[yIdx]);
              if (!isNaN(yVal)) {
                dataPoints.push(yVal);
              }
            }
          }

          if (dataPoints.length > 0) {
            const mean = dataPoints.reduce((a, b) => a + b, 0) / dataPoints.length;
            let sd = 0;
            if (dataPoints.length > 1) {
              const variance = dataPoints.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / dataPoints.length;
              sd = Math.sqrt(variance);
            }
            defaultAbsoluteY = Math.round(mean - sd - (mean * 0.03));
          }
        }
      }
    }

    // Restore previous position or use calculated default
    positionInput.value = currentPositions.hasOwnProperty(checkbox.value) ?
                         currentPositions[checkbox.value] :
                         defaultAbsoluteY;

    positionInput.style.cssText = "width: 50px; font-size: 9px; padding: 1px;";
    positionInput.title = "Bracket Y position (absolute chart value)";

    // Set initial visibility based on checkbox state
    positionInput.style.visibility = checkbox.checked ? 'visible' : 'hidden';
    positionInput.style.opacity = checkbox.checked ? '1' : '0.3';

    // Add event listener for position changes
    positionInput.addEventListener("input", function() {
      const addStatistics = document.getElementById("addStatistics")?.checked;
      if (addStatistics && window.lastRender) {
        clearTimeout(window.positionUpdateTimeout);
        window.positionUpdateTimeout = setTimeout(() => {
          previewPlotWithDebug().catch(e => console.error("Position update failed:", e));
        }, 300);
      }
    });

    label.appendChild(checkbox);
    label.appendChild(text);
    label.appendChild(positionInput);
    checkboxContainer.appendChild(label);
  });
}

// Function to get selected custom comparisons
function getSelectedCustomComparisons() {
  const checkboxes = document.querySelectorAll('#comparisonCheckboxes input[type="checkbox"]:checked');
  return Array.from(checkboxes).map(cb => cb.value);
}

// Function to get custom bracket positions
function getCustomBracketPositions() {
  const positions = {};
  const checkboxes = document.querySelectorAll('#comparisonCheckboxes input[type="checkbox"]');

  checkboxes.forEach((checkbox, index) => {
    const positionInput = document.getElementById(`pos_${index}`);
    if (positionInput) {
      positions[checkbox.value] = parseFloat(positionInput.value) || (1.15 + index * 0.15);
    }
  });

  console.log("🔥 Custom bracket positions:", positions);
  // Debug: show the keys we're sending
  Object.keys(positions).forEach(key => {
    console.log("🔥 Sending position key:", key, "=", positions[key]);
  });
  return positions;
}

// Persistent debug function (disabled for production)
function addDebugInfo(msg) {
  // Debug output disabled - uncomment below to enable
  // const debugPanel = document.getElementById("debugPanel");
  // const debugContent = document.getElementById("debugContent");
  // if (debugPanel && debugContent) {
  //   debugPanel.style.display = "block";
  //   debugContent.innerHTML += "<br>" + msg;
  // }
  // console.log("DEBUG:", msg);
}

// Global function to get statistical symbol size safely
function getStatSymbolSize() {
  try {
    const elem = document.getElementById("statSymbolSize");
    const val = elem ? Number(elem.value) : 10;
    const result = Math.max(6, Math.min(20, Math.round(val || 10)));
    console.log("getStatSymbolSize called - element:", elem, "raw value:", elem?.value, "result:", result);
    return result;
  } catch (e) {
    console.error("getStatSymbolSize error:", e);
    return 10;
  }
}

// Get statistical results as formatted text (shared by export and display functions)
async function getStatisticalResultsText() {
  // Check mode first: if "auto", use auto; otherwise use selected test
  const testMode = document.getElementById("statisticalTestMode")?.value || "auto";
  const statisticalTest = testMode === "auto" ? "auto" : (document.getElementById("statisticalTest")?.value || "auto");
  const varianceTest = document.getElementById("varianceTest")?.value || "levene";
  const dataType = document.getElementById("dataTypeSelect")?.value || "parametric";
  const postHocTest = dataType === "nonparametric"
    ? (document.getElementById("postHocTestNonparam")?.value || "dunn")
    : (document.getElementById("postHocTest")?.value || "tukey");
  const dunnettControl = document.getElementById("dunnettControl")?.value || "";

  // Get selected columns from UI
  const selectedXColumn = document.getElementById("xColumn")?.value || "";
  const selectedYColumn = document.getElementById("yColumn")?.value || "";

  // Find column indices in the data
  let xColIndex = 1, yColIndex = 2;
  if (window.lastProcessedData && window.lastProcessedData.length > 0) {
    const headers = window.lastProcessedData[0];
    const xIdx = headers.findIndex(h => h === selectedXColumn);
    const yIdx = headers.findIndex(h => h === selectedYColumn);
    if (xIdx >= 0) xColIndex = xIdx + 1;  // R uses 1-based indexing
    if (yIdx >= 0) yColIndex = yIdx + 1;
  }

  const detailedTestResult = await webR.evalR(`
    # Set the post-hoc test type directly
    selected_posthoc_test <- "${postHocTest}"

    # Define normality testing function
    sato_check_normality <- function(data, group_col=NULL) {
      if (is.null(group_col)) {
        # Single group normality test
        if (nrow(data) < 3) {
          return(list(
            is_normal = TRUE,  # Assume normal for small samples
            p_value = 1.0,
            test = "Sample too small for normality test"
          ))
        }

        shapiro_result <- shapiro.test(as.numeric(data[[ncol(data)]]))
        return(list(
          is_normal = shapiro_result$p.value > 0.05,
          p_value = shapiro_result$p.value,
          test = "Shapiro-Wilk"
        ))
      } else {
        # Multiple groups normality test
        groups <- unique(data[[group_col]])
        group_results <- list()
        overall_is_normal <- TRUE

        for (group in groups) {
          group_data <- data[data[[group_col]] == group, ]
          group_values <- as.numeric(group_data[[ncol(group_data)]])
          group_values <- group_values[!is.na(group_values)]  # Remove NAs

          if (length(group_values) >= 3 && length(group_values) <= 5000) {
            shapiro_result <- shapiro.test(group_values)
            group_results[[as.character(group)]] <- list(
              group = group,
              n = length(group_values),
              p_value = shapiro_result$p.value,
              is_normal = shapiro_result$p.value > 0.05
            )
            if (shapiro_result$p.value <= 0.05) {
              overall_is_normal <- FALSE
            }
          } else {
            group_results[[as.character(group)]] <- list(
              group = group,
              n = length(group_values),
              p_value = 1.0,
              is_normal = TRUE,
              note = "Sample too small or too large"
            )
          }
        }

        return(list(
          is_normal = overall_is_normal,
          group_results = group_results,
          test = "Shapiro-Wilk (per group)"
        ))
      }
    }

    # NOTE: sato_perform_statistical_test function is already defined earlier in the R code
    # (lines 2325-2418), so we don't need to redefine it here

    # Perform the statistical analysis
    if (exists("dat") && is.data.frame(dat) && ncol(dat) >= max(${xColIndex}, ${yColIndex})) {
      group_col <- dat[[${xColIndex}]]  # Use selected X column (categories)
      value_col <- as.numeric(as.character(dat[[${yColIndex}]]))  # Use selected Y column (values)
      groups <- unique(group_col)

      if (length(groups) >= 2) {
        # Use the existing statistical function
        stat_result <- sato_perform_statistical_test(data.frame(group_col = group_col, value_col = value_col), 1, 2, "${statisticalTest}", "${varianceTest}", selected_posthoc_test)

        p_val <- stat_result$p_value
        test_name <- stat_result$test_used

        # Get significance symbol (check for NA first)
        if (is.na(p_val)) {
          sig <- "ERROR"
        } else if (p_val < 0.001) sig <- "***"
        else if (p_val < 0.01) sig <- "**"
        else if (p_val < 0.05) sig <- "*"
        else sig <- "ns"

        # Update test_name to indicate whether post-hoc was performed
        # If omnibus test is not significant (p >= 0.05), post-hoc is not performed
        if (!is.na(p_val) && p_val >= 0.05) {
          # Replace the post-hoc test name with "not performed"
          test_name <- gsub("\\\\(post-hoc: [^)]+\\\\)", "(post-hoc: not performed)", test_name)
        } else if (grepl("Kruskal-Wallis", test_name)) {
          # For Kruskal-Wallis when significant, use appropriate non-parametric post-hoc
          # Dunnett → Steel (vs control), others → Dunn (pairwise)
          if ("${postHocTest}" == "dunnett" || "${postHocTest}" == "steel") {
            test_name <- gsub("\\\\(post-hoc: [^)]+\\\\)", "(post-hoc: steel)", test_name)
          } else {
            test_name <- gsub("\\\\(post-hoc: [^)]+\\\\)", "(post-hoc: dunn)", test_name)
          }
        }

        # Add group statistics (FIRST - Summary of data)
        group_stats <- "\\nGroup Statistics:"
        for (group in groups) {
          group_data <- value_col[group_col == group]
          group_stats <- paste0(group_stats, "\\n  ", group, ": n=", length(group_data), ", mean=", sprintf("%.2f", mean(group_data, na.rm=TRUE)), ", sd=", sprintf("%.2f", sd(group_data, na.rm=TRUE)))
        }

        # Add normality test results (SECOND - Check assumptions)
        normality_text <- ""
        if (!is.null(stat_result$normality_result)) {
          normality_result <- stat_result$normality_result
          normality_text <- paste0("\\n\\nNormality Testing (", normality_result$test, "):")
          normality_text <- paste0(normality_text, "\\n  Overall Assessment: Data appears ", if (normality_result$is_normal) "normal" else "non-normal")

          if (!is.null(normality_result$group_results)) {
            normality_text <- paste0(normality_text, "\\n  Per-group results:")
            for (group_name in names(normality_result$group_results)) {
              gr <- normality_result$group_results[[group_name]]
              if (!is.null(gr$note)) {
                normality_text <- paste0(normality_text, "\\n    ", gr$group, " (n=", gr$n, "): ", gr$note)
              } else if (!is.null(gr$is_normal)) {
                normality_text <- paste0(normality_text, "\\n    ", gr$group, " (n=", gr$n, "): p=", sprintf("%.4f", gr$p_value), " (", if (gr$is_normal) "normal" else "non-normal", ")")
              }
            }
          }
        }

        # Add variance test results (THIRD - For 2-group comparisons with t-test)
        variance_text <- ""
        if (length(groups) == 2 && !is.null(stat_result$variance_test) && !is.na(stat_result$variance_p_value)) {
          variance_status <- if (stat_result$equal_variances) "equal variances" else "unequal variances"
          variance_text <- paste0("\\nVariance Test (", stat_result$variance_test, "): p=", sprintf("%.4f", stat_result$variance_p_value), " (", variance_status, ")")
        }

        # Create main test result (FOURTH - Overall statistical test)
        main_result <- paste0("\\n\\n", test_name, ": p=", sprintf("%.4f", p_val), " (", sig, ")")

        # Add post-hoc results with detailed pairwise comparisons
        posthoc_results <- ""
        if (grepl("Kruskal-Wallis", test_name) && !is.na(p_val) && p_val < 0.05 && length(groups) > 2) {
          # For Kruskal-Wallis, use appropriate non-parametric post-hoc
          # Dunnett or Steel → Steel, others → Dunn
          cat("DEBUG (path 1): selected_posthoc_test =", "${postHocTest}", "\\n")
          actual_posthoc <- if ("${postHocTest}" == "dunnett" || "${postHocTest}" == "steel") "steel" else "dunn"
          cat("DEBUG (path 1): actual_posthoc =", actual_posthoc, "\\n")

          tryCatch({
            if (actual_posthoc == "steel") {
              # Steel test - non-parametric equivalent of Dunnett
              # Try kSamples::Steel.test
              steel_available <- tryCatch({
                if (!requireNamespace("kSamples", quietly = TRUE)) {
                  cat("Attempting to install kSamples package for Steel test\\n")
                  webr::install("kSamples")
                }
                library(kSamples)
                TRUE
              }, error = function(e) {
                cat("kSamples not available in webR:", e$message, "\\n")
                FALSE
              })

              if (!steel_available) {
                # Fallback to Dunn test with message
                cat("Steel test not available, falling back to Dunn test\\n")
                actual_posthoc <- "dunn"
                if ("${postHocTest}" == "dunnett") {
                  posthoc_results <- "\\n\\n(Note: Data is non-normal. Dunnett test requires normal data. Using Dunn test instead.)"
                } else {
                  posthoc_results <- "\\n\\n(Note: Steel test not available in webR, using Dunn test instead)"
                }
              } else {
                # Create data frame and set control group
                steel_data <- data.frame(
                  group = factor(group_col),
                  value = as.numeric(value_col)
                )

                cat("🔥 EXPORT Steel DEBUG: steel_data created with", nrow(steel_data), "rows\\n")
                cat("🔥 EXPORT Steel DEBUG: groups:", paste(levels(steel_data$group), collapse=", "), "\\n")

                control_group <- "${dunnettControl}"
                cat("🔥 EXPORT Steel DEBUG: raw control_group = '", control_group, "'\\n")
                if (control_group == "" || is.na(control_group)) {
                  control_group <- levels(steel_data$group)[1]
                }
                cat("🔥 EXPORT Steel DEBUG: final control_group = '", control_group, "'\\n")

                # Get all groups and identify treatment groups
                all_groups <- levels(steel_data$group)
                treatment_groups <- all_groups[all_groups != control_group]
                cat("🔥 EXPORT Steel DEBUG: treatment_groups:", paste(treatment_groups, collapse=", "), "\\n")

                posthoc_results <- paste0("\\n\\nSteel's Post-hoc Comparisons (vs ", control_group, "):")

                # Run Steel.test for each treatment vs control
                control_values <- steel_data$value[steel_data$group == control_group]
                cat("🔥 EXPORT Steel DEBUG: control_values (n=", length(control_values), "):", paste(control_values, collapse=", "), "\\n")

                for (trt in treatment_groups) {
                  trt_values <- steel_data$value[steel_data$group == trt]
                  cat("🔥 EXPORT Steel DEBUG: trt '", trt, "' values (n=", length(trt_values), "):", paste(trt_values, collapse=", "), "\\n")

                  # Steel.test compares two samples - p-value is in st[2]
                  steel_result <- Steel.test(list(control_values, trt_values))
                  cat("🔥 EXPORT Steel DEBUG: steel_result$st:", paste(steel_result$st, collapse=", "), "\\n")
                  p_val <- steel_result$st[2]
                  cat("🔥 EXPORT Steel DEBUG: p_val =", p_val, "\\n")

                  comparison <- paste0(trt, " - ", control_group)

                  if (p_val < 0.001) sig <- "***"
                  else if (p_val < 0.01) sig <- "**"
                  else if (p_val < 0.05) sig <- "*"
                  else sig <- "ns"

                  posthoc_results <- paste0(posthoc_results, "\\n  ", comparison, ": p=", sprintf("%.4f", p_val), " (", sig, ")")
                }
              }
            }  # end of actual_posthoc == "steel"

            # Run Dunn test if Steel was not available or if Dunn was originally selected
            if (actual_posthoc == "dunn") {
              # Dunn test - non-parametric pairwise comparisons
              if (!requireNamespace("dunn.test", quietly = TRUE)) {
                cat("Installing dunn.test package...\\n")
                webr::install("dunn.test")
              }
              library(dunn.test)

              dunn_result <- dunn.test(as.numeric(value_col), factor(group_col), method = "bonferroni")

              comparison_names <- dunn_result$comparisons
              p_values <- dunn_result$P.adjusted

              # Append to existing note (if Steel fallback) or create new
              posthoc_results <- paste0(posthoc_results, "\\n\\nDunn's Post-hoc Comparisons (Bonferroni-adjusted p-values):")

              for (i in 1:length(comparison_names)) {
                p_adj <- p_values[i]
                comparison <- comparison_names[i]

                if (!is.na(p_adj)) {
                  if (p_adj < 0.001) sig <- "***"
                  else if (p_adj < 0.01) sig <- "**"
                  else if (p_adj < 0.05) sig <- "*"
                  else sig <- "ns"

                  posthoc_results <- paste0(posthoc_results, "\\n  ", comparison, ": p=", sprintf("%.4f", p_adj), " (", sig, ")")
                }
              }
            }
          }, error = function(e) {
            posthoc_results <<- paste0("\\n\\nPost-hoc test Error: ", e$message)
          })
        } else if (grepl("ANOVA", test_name) && !is.na(p_val) && p_val < 0.05 && length(groups) > 2) {
          # Perform ANOVA and post-hoc test manually (same as export function)
          tryCatch({
            # Create proper data frame for ANOVA
            anova_data <- data.frame(
              group = factor(group_col),
              value = as.numeric(value_col)
            )

            # Perform ANOVA
            anova_result <- aov(value ~ group, data = anova_data)

            # Perform the selected post-hoc test
            if (selected_posthoc_test == "tukey") {
              posthoc_result <- TukeyHSD(anova_result)
              posthoc_summary <- posthoc_result$group
              posthoc_results <- "\\n\\nTukey HSD Post-hoc Comparisons:"

              # Format detailed results
              for (i in 1:nrow(posthoc_summary)) {
                comparison <- rownames(posthoc_summary)[i]
                p_adj <- posthoc_summary[i, "p adj"]
                diff <- posthoc_summary[i, "diff"]

                # Determine significance
                if (!is.na(p_adj)) {
                  if (p_adj < 0.001) sig <- "***"
                  else if (p_adj < 0.01) sig <- "**"
                  else if (p_adj < 0.05) sig <- "*"
                  else sig <- "ns"
                } else {
                  sig <- "ns"
                }

                posthoc_results <- paste0(posthoc_results, "\\n  ", comparison, ": diff=", round(diff, 2), ", p=", sprintf("%.4f", ifelse(is.na(p_adj), 1.0, p_adj)), " (", sig, ")")
              }
            } else if (selected_posthoc_test == "dunnett") {
              # Dunnett test - parametric vs control comparisons
              if (!requireNamespace("multcomp", quietly = TRUE)) {
                cat("Installing multcomp package for Dunnett test\\n")
                webr::install("multcomp")
              }
              library(multcomp)

              control_group <- "${dunnettControl}"
              if (control_group == "" || is.na(control_group)) {
                control_group <- levels(anova_data$group)[1]
              }
              anova_data$group <- relevel(factor(anova_data$group), ref = control_group)
              anova_result <- aov(value ~ group, data = anova_data)

              posthoc_result <- summary(glht(anova_result, linfct = mcp(group = "Dunnett")))
              estimates <- posthoc_result$test$coefficients
              p_values <- posthoc_result$test$pvalues
              comparison_names <- names(estimates)

              posthoc_results <- paste0("\\n\\nDunnett's Post-hoc Comparisons (vs ", control_group, "):")

              for (i in 1:length(comparison_names)) {
                p_adj <- p_values[i]
                comparison <- comparison_names[i]
                diff <- estimates[i]

                if (!is.na(p_adj)) {
                  if (p_adj < 0.001) sig <- "***"
                  else if (p_adj < 0.01) sig <- "**"
                  else if (p_adj < 0.05) sig <- "*"
                  else sig <- "ns"
                } else {
                  sig <- "ns"
                }

                posthoc_results <- paste0(posthoc_results, "\\n  ", comparison, ": diff=", round(diff, 3), ", p=", sprintf("%.4f", ifelse(is.na(p_adj), 1.0, p_adj)), " (", sig, ")")
              }
            } else if (selected_posthoc_test == "bonferroni" || selected_posthoc_test == "holm") {
              # Pairwise t-tests with correction
              method <- selected_posthoc_test
              pairwise_result <- pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = method)
              p_matrix <- pairwise_result$p.value

              posthoc_results <- paste0("\\n\\nPairwise t-test (", method, " correction):")

              groups <- rownames(p_matrix)
              for (i in 1:nrow(p_matrix)) {
                for (j in 1:ncol(p_matrix)) {
                  p_adj <- p_matrix[i, j]
                  if (!is.na(p_adj)) {
                    comparison <- paste0(groups[i], " - ", colnames(p_matrix)[j])

                    if (p_adj < 0.001) sig <- "***"
                    else if (p_adj < 0.01) sig <- "**"
                    else if (p_adj < 0.05) sig <- "*"
                    else sig <- "ns"

                    posthoc_results <- paste0(posthoc_results, "\\n  ", comparison, ": p=", sprintf("%.4f", p_adj), " (", sig, ")")
                  }
                }
              }
            } else {
              posthoc_results <- paste0("\\n\\nPost-hoc test (", selected_posthoc_test, ") performed - see brackets on chart")
            }
          }, error = function(e) {
            posthoc_results <<- paste0("\\n\\nPost-hoc Error: ", e$message)
          })
        } else if (is.na(p_val)) {
          posthoc_results <- "\\n\\nNo post-hoc test performed (statistical test failed - insufficient data)"
        } else if (p_val >= 0.05) {
          posthoc_results <- "\\n\\nNo post-hoc test performed (overall test not significant)"
        } else if (length(groups) == 2) {
          posthoc_results <- "\\n\\nDirect comparison shown on chart"
        } else {
          # For other tests: bonferroni and holm are p-value adjustments, not post-hoc tests
          if (selected_posthoc_test == "bonferroni") {
            posthoc_results <- "\\n\\nPairwise t-test with Bonferroni correction performed - see brackets on chart"
          } else if (selected_posthoc_test == "holm") {
            posthoc_results <- "\\n\\nPairwise t-test with Holm correction performed - see brackets on chart"
          } else {
            posthoc_results <- paste0("\\n\\nPost-hoc test (", selected_posthoc_test, ") performed - see brackets on chart")
          }
        }

        full_result <- paste0(group_stats, normality_text, variance_text, main_result, posthoc_results)
        full_result
      } else {
        "ERROR: Need at least 2 groups"
      }
    } else {
      "ERROR: Need at least 2 columns"
    }
  `);

  const testResult = await detailedTestResult.toJs();
  console.log("Statistical test result:", testResult);

  if (testResult && testResult.values && testResult.values[0]) {
    const resultText = testResult.values[0];
    return resultText;
  } else {
    return "Statistical analysis failed";
  }
}

// Get statistical results for line plots (time-point specific analysis)
async function getGroupedBarStatisticalResultsText() {
  try {
    await initWebR();

    // Check if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    console.log("DEBUG: addStatistics =", addStatistics);
    if (!addStatistics) {
      return "Statistics disabled";
    }

    // Try to retrieve the statistical results from the R global variable
    const statResultsR = await webR.evalR(`
      cat("DEBUG: Checking for grouped_bar_stat_results...\\n")
      if (exists("grouped_bar_stat_results")) {
        cat("DEBUG: Variable exists, length =", nchar(grouped_bar_stat_results), "\\n")
        if (nchar(grouped_bar_stat_results) > 0) {
          grouped_bar_stat_results
        } else {
          "EMPTY_RESULTS"
        }
      } else {
        cat("DEBUG: Variable does not exist\\n")
        "NO_RESULTS"
      }
    `);

    const statResultsJS = await statResultsR.toJs();
    const statText = statResultsJS && statResultsJS.values && statResultsJS.values[0]
                     ? statResultsJS.values[0]
                     : "NO_RESULTS";

    if (statText && statText !== "NO_RESULTS" && statText !== "EMPTY_RESULTS" && statText.trim() !== "") {
      return statText;
    } else {
      return `No statistical results available (status: ${statText})`;
    }
  } catch (error) {
    console.error("Get grouped bar statistical results error:", error);
    return `ERROR: ${error.message}`;
  }
}

async function getLineStatisticalResultsText() {
  try {
    await initWebR();

    // Check if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    console.log("DEBUG: addStatistics =", addStatistics);
    if (!addStatistics) {
      return "Statistics disabled";
    }

    // Try to retrieve the statistical results from the R global variable
    const statResultsR = await webR.evalR(`
      cat("DEBUG: Checking for line_plot_stat_results...\\n")
      if (exists("line_plot_stat_results")) {
        cat("DEBUG: Variable exists, length =", nchar(line_plot_stat_results), "\\n")
        if (nchar(line_plot_stat_results) > 0) {
          line_plot_stat_results
        } else {
          "EMPTY_RESULTS"
        }
      } else {
        cat("DEBUG: Variable does not exist\\n")
        "NO_RESULTS"
      }
    `);

    const statResultsJS = await statResultsR.toJs();
    const statText = statResultsJS && statResultsJS.values && statResultsJS.values[0]
                     ? statResultsJS.values[0]
                     : "NO_RESULTS";

    console.log("DEBUG: Retrieved stat text =", statText ? statText.substring(0, 100) : "null");
    console.log("DEBUG: Full stat text length =", statText ? statText.length : 0);

    if (statText && statText !== "NO_RESULTS" && statText !== "EMPTY_RESULTS" && statText.trim() !== "") {
      // Return the detailed statistical results from R
      console.log("DEBUG: Returning stat text with length", statText.length);
      return statText;
    } else {
      // Debug: Show what we got
      console.log("DEBUG: Failed to get results. Status:", statText);
      return `No statistical results available (status: ${statText})`;
    }

  } catch (error) {
    console.error("Get line statistical results error:", error);
    return `ERROR: ${error.message}`;
  }
}

// Display statistical results in UI only (without Excel export)
async function displayStatisticalResultsInUI() {
  try {
    await initWebR();

    // Check if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (!addStatistics) {
      // Statistics not enabled - don't update status (keep preview completion message)
      return;
    }

    // Check chart type - determine if this chart type supports overall statistical testing
    const chartType = document.getElementById("chartType")?.value || "";

    // Chart types that support overall statistical testing
    const statisticalChartTypes = [
      "bar_grouped_error_dot","bar_error_dot",
      "box_grouped_dot", "box_dot", "violin_dot",
      "violin_grouped", "violin_grouped_dot"
    ];

    // Line plots handle statistics differently (time-point specific, not overall)
    const isLinePlot = chartType === "line_grouped_error_raw" || chartType === "line_grouped_error";

    // Handle line plots (time-point specific comparisons)
    if (chartType === "line_grouped_error_raw") {
      const lineStatResults = await getLineStatisticalResultsText();
      // Debug output suppressed - results available in console and export
      // if (lineStatResults && !lineStatResults.startsWith("ERROR:") && !lineStatResults.includes("No statistical")) {
      //   setStatus(`Statistical results: ${lineStatResults}`);
      // } else {
      //   setStatus(`DEBUG: ${lineStatResults}`);
      // }
      return;
    }

    // Handle grouped charts with categories (category-specific comparisons)
    if (chartType === "bar_grouped_error_dot" || chartType === "box_grouped" || chartType === "box_grouped_dot" || chartType === "violin_grouped" || chartType === "violin_grouped_dot") {
      const groupedBarStatResults = await getGroupedBarStatisticalResultsText();
      // Debug output suppressed - results available in console and export
      // if (groupedBarStatResults && !groupedBarStatResults.startsWith("ERROR:") && !groupedBarStatResults.includes("No statistical")) {
      //   setStatus(`Statistical results: ${groupedBarStatResults}`);
      // } else {
      //   setStatus(`DEBUG: ${groupedBarStatResults}`);
      // }
      return;
    }

    // Check if this chart type supports overall statistical testing
    if (!statisticalChartTypes.includes(chartType)) {
      // Chart type doesn't support overall statistical testing
      // Show warning message to inform the user
      if (chartType === "line_grouped_error") {
        // Pre-calculated data - can't perform statistical tests
        setStatus("⚠️ Statistical testing is not available for this chart type. This chart accepts pre-calculated mean + error values. Use 'Grouped line with error (raw data)' to perform statistical tests.");
      } else {
        // Other chart types don't support statistics at all
        setStatus("⚠️ Statistical testing is not available for this chart type. Please use a grouped chart type (e.g., Grouped bar, Grouped box, Grouped violin) for statistical comparisons.");
      }
      return;
    }

    // Get statistical results (reuse same code as export function)
    const formattedResult = await getStatisticalResultsText();

    // Debug output suppressed - results available in console and export
    // if (formattedResult && !formattedResult.startsWith("ERROR:")) {
    //   // Check if user has overall symbol enabled with only 2 groups (educational warning)
    //   const showMainSymbol = document.getElementById("showMainStatSymbol")?.checked;
    //   const numGroups = await getNumberOfGroups();
    //
    //   let statusText = `Statistical results: ${formattedResult}`;
    //
    //   if (showMainSymbol && numGroups === 2) {
    //     statusText += " | ⚠️ Note: 'Show overall test symbol' is only for 3+ groups. For 2-group comparisons, significance is shown on the bracket.";
    //   }
    //
    //   setStatus(statusText);
    // }
  } catch (error) {
    console.error("Display statistical results error:", error);
    // Don't show error in UI - silently fail for auto-display
  }
}

async function exportStatisticalResults() {
  try {
    await initWebR();

    // Check if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    console.log("DEBUG exportStatisticalResults: addStatistics =", addStatistics);
    console.log("DEBUG exportStatisticalResults: checkbox element =", document.getElementById("addStatistics"));
    if (!addStatistics) {
      setStatus("Statistical analysis not enabled. Please check 'Show statistical significance' and run Preview first.");
      return;
    }

    // Check if this is a line plot or grouped bar - use same method as other chart types
    const chartType = document.getElementById("chartType")?.value || "";

    // Check if chart type supports statistics (silently return if not)
    const STATS_SUPPORTED_CHART_TYPES = [
      "bar_error_dot", "bar_grouped_error_dot",
      "box", "box_dot", "box_grouped", "box_grouped_dot",
      "violin_dot", "violin_grouped", "violin_grouped_dot",
      "line_grouped_error_raw"
    ];

    if (!STATS_SUPPORTED_CHART_TYPES.includes(chartType)) {
      // Silently return - UI already prevents checking statistics for unsupported types
      return;
    }

    if (chartType === "line_grouped_error_raw") {
      // For line plots, get the detailed statistical results text
      const lineStatResults = await getLineStatisticalResultsText();

      if (!lineStatResults || lineStatResults.includes("No statistical") || lineStatResults.includes("ERROR") || lineStatResults.includes("DEBUG")) {
        setStatus(`Statistical results not found: ${lineStatResults ? lineStatResults.substring(0, 100) : "null"}`);
        return;
      }

      // Use the same Excel export format as other chart types
      await Excel.run(async (context) => {
        try {
          // Get the selected range (user's selection)
          const selectedRange = context.workbook.getSelectedRange();
          selectedRange.load("address");
          await context.sync();

          // Convert \n string literals to actual newline characters for Excel
          const formattedResults = lineStatResults.replace(/\\n/g, '\n');

          // Write the results to the selected cell
          selectedRange.values = [[formattedResults]];

          await context.sync();
          setStatus(`Statistical results exported: ${formattedResults}`);
        } catch (error) {
          setStatus("Please select a cell before clicking 'Export statistics to Excel'.");
        }
      });
      return;
    } else if (chartType === "bar_grouped_error_dot" || chartType === "box_grouped" || chartType === "box_grouped_dot" || chartType === "violin_grouped" || chartType === "violin_grouped_dot") {
      // For grouped charts with categories, get the detailed statistical results text
      const groupedBarStatResults = await getGroupedBarStatisticalResultsText();

      if (!groupedBarStatResults || groupedBarStatResults.includes("No statistical") || groupedBarStatResults.includes("ERROR") || groupedBarStatResults.includes("DEBUG")) {
        setStatus(`Statistical results not found: ${groupedBarStatResults ? groupedBarStatResults.substring(0, 100) : "null"}`);
        return;
      }

      // Use the same Excel export format as other chart types
      await Excel.run(async (context) => {
        try {
          // Get the selected range (user's selection)
          const selectedRange = context.workbook.getSelectedRange();
          selectedRange.load("address");
          await context.sync();

          // Convert \n string literals to actual newline characters for Excel
          const formattedResults = groupedBarStatResults.replace(/\\n/g, '\n');

          // Write the results to the selected cell
          selectedRange.values = [[formattedResults]];

          await context.sync();
          setStatus(`Statistical results exported: ${formattedResults}`);
        } catch (error) {
          setStatus("Please select a cell before clicking 'Export statistics to Excel'.");
        }
      });
      return;
    }
    
    // Enhanced statistical analysis with normality testing and detailed results
    // Check mode first: if "auto", use auto; otherwise use selected test
    const testMode = document.getElementById("statisticalTestMode")?.value || "auto";
    const statisticalTest = testMode === "auto" ? "auto" : (document.getElementById("statisticalTest")?.value || "auto");
    const varianceTest = document.getElementById("varianceTest")?.value || "levene";
    const dataType = document.getElementById("dataTypeSelect")?.value || "parametric";
    const postHocTest = dataType === "nonparametric"
      ? (document.getElementById("postHocTestNonparam")?.value || "dunn")
      : (document.getElementById("postHocTest")?.value || "tukey");
    const dunnettControl = document.getElementById("dunnettControl")?.value || "";

    // Get selected columns from UI
    const selectedXColumn = document.getElementById("xColumn")?.value || "";
    const selectedYColumn = document.getElementById("yColumn")?.value || "";

    // Find column indices in the data
    let xColIndex = 1, yColIndex = 2;
    if (window.lastProcessedData && window.lastProcessedData.length > 0) {
      const headers = window.lastProcessedData[0];
      const xIdx = headers.findIndex(h => h === selectedXColumn);
      const yIdx = headers.findIndex(h => h === selectedYColumn);
      if (xIdx >= 0) xColIndex = xIdx + 1;  // R uses 1-based indexing
      if (yIdx >= 0) yColIndex = yIdx + 1;
    }

    const detailedTestResult = await webR.evalR(`
      # Set the post-hoc test type directly
      selected_posthoc_test <- "${postHocTest}"

      # Define normality testing function
      sato_check_normality <- function(data, group_col=NULL) {
        if (is.null(group_col)) {
          # Single group normality test
          if (nrow(data) < 3 || nrow(data) > 5000) {
            return(list(
              is_normal = TRUE,  # Assume normal for small samples
              p_value = 1.0,
              test = "Sample size outside range for Shapiro-Wilk test (requires 3-5000)"
            ))
          }

          shapiro_result <- shapiro.test(as.numeric(data[[ncol(data)]]))
          return(list(
            is_normal = shapiro_result$p.value > 0.05,
            p_value = shapiro_result$p.value,
            test = "Shapiro-Wilk"
          ))
        } else {
          # Multiple groups normality test
          groups <- unique(data[[group_col]])
          group_results <- list()
          overall_is_normal <- TRUE

          for (group in groups) {
            group_data <- data[data[[group_col]] == group, ]
            group_values <- as.numeric(group_data[[ncol(group_data)]])
            group_values <- group_values[!is.na(group_values)]  # Remove NAs

            if (length(group_values) >= 3 && length(group_values) <= 5000) {
              shapiro_result <- shapiro.test(group_values)
              group_results[[as.character(group)]] <- list(
                group = group,
                n = length(group_values),
                p_value = shapiro_result$p.value,
                is_normal = shapiro_result$p.value > 0.05
              )
              if (shapiro_result$p.value <= 0.05) {
                overall_is_normal <- FALSE
              }
            } else {
              group_results[[as.character(group)]] <- list(
                group = group,
                n = length(group_values),
                p_value = 1.0,
                is_normal = TRUE,
                note = "Sample too small or too large"
              )
            }
          }

          return(list(
            is_normal = overall_is_normal,
            group_results = group_results,
            test = "Shapiro-Wilk (per group)"
          ))
        }
      }

      # Define the statistical test function (with post-hoc support for export)
      sato_perform_statistical_test <- function(data, group_col, value_col, test_type="auto", variance_test="levene", posthoc_test="tukey") {
        groups <- unique(data[[group_col]])
        n_groups <- length(groups)

        # Perform normality testing
        normality_result <- sato_check_normality(data, group_col)

        # Auto-select test type based on normality if needed
        if (test_type == "auto") {
          if (n_groups == 2) {
            test_type <- if (normality_result$is_normal) "t.test" else "wilcox.test"
          } else {
            test_type <- if (normality_result$is_normal) "anova" else "kruskal.test"
          }
        }

        # For 2 groups
        if (n_groups == 2) {
          group1_data <- as.numeric(data[data[[group_col]] == groups[1], value_col])
          group2_data <- as.numeric(data[data[[group_col]] == groups[2], value_col])
          # Remove NAs
          group1_data <- group1_data[!is.na(group1_data)]
          group2_data <- group2_data[!is.na(group2_data)]

          if (test_type == "t.test") {
            result <- t.test(group1_data, group2_data)
            return(list(
              p_value = result$p.value,
              test_used = "t-test",
              normality_result = normality_result
            ))
          } else {
            result <- wilcox.test(group1_data, group2_data)
            return(list(
              p_value = result$p.value,
              test_used = "Wilcoxon test",
              normality_result = normality_result
            ))
          }
        }

        # For 3+ groups
        # Debug: Check data before cleaning
        cat("DEBUG: Before cleaning - rows:", nrow(data), "groups:", length(unique(data[[group_col]])), "\\n")
        cat("DEBUG: Value column type:", class(data[[value_col]]), "\\n")
        cat("DEBUG: First 10 values:", head(data[[value_col]], 10), "\\n")
        cat("DEBUG: NAs in value column:", sum(is.na(data[[value_col]])), "\\n")

        # Clean data by removing NAs from value column
        # First ensure it's numeric
        data[[value_col]] <- as.numeric(as.character(data[[value_col]]))
        clean_data <- data[!is.na(data[[value_col]]), ]

        cat("DEBUG: After cleaning - rows:", nrow(clean_data), "groups:", length(unique(clean_data[[group_col]])), "\\n")

        # Check if we still have enough groups after cleaning
        remaining_groups <- length(unique(clean_data[[group_col]]))
        if (remaining_groups < 2) {
          return(list(
            p_value = NA,
            test_used = paste0("ERROR: Insufficient groups after removing NAs (had ", nrow(data), " rows, now ", nrow(clean_data), " rows, ", remaining_groups, " groups)"),
            normality_result = normality_result
          ))
        }

        if (test_type == "anova") {
          result <- aov(as.formula(paste(names(clean_data)[value_col], "~", names(clean_data)[group_col])), data=clean_data)
          summary_result <- summary(result)
          return(list(
            p_value = summary_result[[1]][["Pr(>F)"]][1],
            test_used = paste("ANOVA (post-hoc:", selected_posthoc_test, ")"),
            normality_result = normality_result
          ))
        } else {
          result <- kruskal.test(as.formula(paste(names(clean_data)[value_col], "~", names(clean_data)[group_col])), data=clean_data)
          return(list(
            p_value = result$p.value,
            test_used = paste("Kruskal-Wallis test (post-hoc:", selected_posthoc_test, ")"),
            normality_result = normality_result
          ))
        }
      }

      # Check if data exists
      if (!exists("dat") || !is.data.frame(dat) || nrow(dat) < 2) {
        cat("ERROR: No valid data found\\n")
        "ERROR: No data"
      } else {
        cat("Data found - rows:", nrow(dat), "cols:", ncol(dat), "\\n")
        
        # Use selected columns from UI
        if (ncol(dat) >= max(${xColIndex}, ${yColIndex})) {
          group_col <- dat[[${xColIndex}]]  # X-axis column (categories to compare)
          value_col <- as.numeric(as.character(dat[[${yColIndex}]]))  # Y-axis column (values) - convert via character first
          
          # Get unique groups
          groups <- unique(group_col)
          cat("Groups found:", paste(groups, collapse=", "), "\\n")
          
          if (length(groups) == 2) {
            # Extract data for each group
            group1_values <- value_col[group_col == groups[1]]
            group2_values <- value_col[group_col == groups[2]]
            
            cat("Group 1 (", groups[1], ") - n=", length(group1_values), ", mean=", round(mean(group1_values), 2), ", sd=", round(sd(group1_values), 2), "\\n")
            cat("Group 2 (", groups[2], ") - n=", length(group2_values), ", mean=", round(mean(group2_values), 2), ", sd=", round(sd(group2_values), 2), "\\n")
            
            # Normality testing (if sample size >= 3)
            normality_text <- ""
            test_used <- ""
            
            if (length(group1_values) >= 3 && length(group1_values) <= 5000 &&
                length(group2_values) >= 3 && length(group2_values) <= 5000) {
              shapiro1 <- shapiro.test(group1_values)
              shapiro2 <- shapiro.test(group2_values)
              
              is_group1_normal <- shapiro1$p.value > 0.05
              is_group2_normal <- shapiro2$p.value > 0.05
              both_normal <- is_group1_normal && is_group2_normal
              
              cat("Normality test results:\\n")
              cat("  ", groups[1], ": Shapiro-Wilk p=", sprintf("%.4f", shapiro1$p.value), " (", ifelse(is_group1_normal, "assumed normal", "non-normal"), ")\\n")
              cat("  ", groups[2], ": Shapiro-Wilk p=", sprintf("%.4f", shapiro2$p.value), " (", ifelse(is_group2_normal, "assumed normal", "non-normal"), ")\\n")
              
              normality_text <- paste0("Normality test by Shapiro-Wilk test: ", groups[1], " p=", sprintf("%.4f", shapiro1$p.value), 
                                     " (", ifelse(is_group1_normal, "assumed normal", "non-normal"), "), ",
                                     groups[2], " p=", sprintf("%.4f", shapiro2$p.value), 
                                     " (", ifelse(is_group2_normal, "assumed normal", "non-normal"), ")")
              
              # Auto-select test based on normality if requested
              if ("` + (statisticalTest || 'auto') + `" == "auto") {
                if (both_normal) {
                  test_used <- "t-test"
                  cat("Auto-selected: t-test (both groups normal)\\n")
                } else {
                  test_used <- "Wilcoxon test"
                  cat("Auto-selected: Wilcoxon test (non-normal data detected)\\n")
                }
              } else {
                test_used <- "${statisticalTest}"
                cat("Manual selection:", test_used, "\\n")
              }
            } else {
              cat("Sample size too small for normality testing (using t-test)\\n")
              test_used <- if ("${statisticalTest}" == "auto") "t-test" else "${statisticalTest}"
              normality_text <- "Normality test by Shapiro-Wilk test: Sample size too small (normality assumed)"
            }
            
            # Perform the selected statistical test
            variance_text <- ""
            
            if (test_used == "wilcox.test" || test_used == "Wilcoxon test") {
              stat_result <- wilcox.test(group1_values, group2_values)
              test_name <- "Wilcoxon test"
              # No variance test needed for non-parametric test
            } else {
              # Check if variances are equal to choose appropriate t-test
              if ("` + (varianceTest || 'levene') + `" == "levene") {
                # Levene test for equality of variances
                combined_data <- data.frame(
                  values = c(group1_values, group2_values),
                  group = factor(c(rep(groups[1], length(group1_values)), rep(groups[2], length(group2_values))))
                )
                
                # Calculate Levene test manually
                group_means <- tapply(combined_data$values, combined_data$group, mean)
                abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
                levene_result <- anova(lm(abs_deviations ~ combined_data$group))
                levene_p <- levene_result$Pr[1]
                
                equal_variances <- levene_p > 0.05
                
                cat("Variance test: Levene test p=", sprintf("%.4f", levene_p), " (", ifelse(equal_variances, "equal variances", "unequal variances"), ")\\n")
                
                # Create variance test text for output
                variance_text <- paste0("Variance test by Levene test: p=", sprintf("%.4f", levene_p), 
                                      " (", ifelse(equal_variances, "equal variances", "unequal variances"), ")")
              } else {
                # F-test for equality of variances
                var_test <- var.test(group1_values, group2_values)
                equal_variances <- var_test$p.value > 0.05
                
                cat("Variance test: F-test p=", sprintf("%.4f", var_test$p.value), " (", ifelse(equal_variances, "equal variances", "unequal variances"), ")\\n")
                
                # Create variance test text for output
                variance_text <- paste0("Variance test by F-test: p=", sprintf("%.4f", var_test$p.value), 
                                      " (", ifelse(equal_variances, "equal variances", "unequal variances"), ")")
              }
              
              if (equal_variances) {
                stat_result <- t.test(group1_values, group2_values, var.equal = TRUE)
                test_name <- "Student's t-test"
              } else {
                stat_result <- t.test(group1_values, group2_values, var.equal = FALSE)
                test_name <- "Welch's t-test"
              }
            }
            
            p_val <- stat_result$p.value

            # Get significance symbol (check for NA first)
            if (is.na(p_val)) {
              sig <- "ERROR"
              main_result <- paste0(test_name, ": ERROR (insufficient data after removing NAs)")
            } else {
              if (p_val < 0.001) sig <- "***"
              else if (p_val < 0.01) sig <- "**"
              else if (p_val < 0.05) sig <- "*"
              else sig <- "ns"

              # Create comprehensive result
              main_result <- paste0(test_name, ": p=", sprintf("%.4f", p_val), " (", sig, ") [", groups[1], " vs ", groups[2], "]")
            }
            
            # Build comprehensive result with all test information
            result_parts <- c(main_result)
            if (nchar(normality_text) > 0) {
              result_parts <- c(result_parts, normality_text)
            }
            if (nchar(variance_text) > 0) {
              result_parts <- c(result_parts, variance_text)
            }
            
            full_result <- paste(result_parts, collapse="\\n")

            cat("\\nFINAL COMPREHENSIVE RESULT:\\n")
            cat(full_result, "\\n")
            full_result
          } else if (length(groups) > 2) {
            # Multiple groups - perform ANOVA or Kruskal-Wallis
            # Use the existing statistical function
            stat_result <- sato_perform_statistical_test(data.frame(group_col = group_col, value_col = value_col), 1, 2, "${statisticalTest}", "${varianceTest}", selected_posthoc_test)

            p_val <- stat_result$p_value
            test_name <- stat_result$test_used

            # Get significance symbol (check for NA first)
            if (is.na(p_val)) {
              sig <- "ERROR"
            } else if (p_val < 0.001) sig <- "***"
            else if (p_val < 0.01) sig <- "**"
            else if (p_val < 0.05) sig <- "*"
            else sig <- "ns"

            # Update test_name to indicate whether post-hoc was performed
            # If omnibus test is not significant (p >= 0.05), post-hoc is not performed
            if (!is.na(p_val) && p_val >= 0.05) {
              # Replace the post-hoc test name with "not performed"
              test_name <- gsub("\\\\(post-hoc: [^)]+\\\\)", "(post-hoc: not performed)", test_name)
            } else if (grepl("Kruskal-Wallis", test_name)) {
              # For Kruskal-Wallis when significant, use appropriate non-parametric post-hoc
              # Dunnett → Steel (vs control), others → Dunn (pairwise)
              if (selected_posthoc_test == "dunnett" || selected_posthoc_test == "steel") {
                test_name <- gsub("\\\\(post-hoc: [^)]+\\\\)", "(post-hoc: steel)", test_name)
              } else {
                test_name <- gsub("\\\\(post-hoc: [^)]+\\\\)", "(post-hoc: dunn)", test_name)
              }
            }

            # Create comprehensive result for multiple groups
            main_result <- paste0(test_name, ": p=", sprintf("%.4f", p_val), " (", sig, ") [Overall test for ", length(groups), " groups: ", paste(groups, collapse=", "), "]")

            # Add normality test results
            normality_text <- ""
            if (!is.null(stat_result$normality_result)) {
              normality_result <- stat_result$normality_result
              normality_text <- paste0("\\n\\nNormality Testing (", normality_result$test, "):")
              normality_text <- paste0(normality_text, "\\nOverall Assessment: ", ifelse(normality_result$is_normal, "Data appears normal", "Data appears non-normal"))

              if (!is.null(normality_result$group_results)) {
                normality_text <- paste0(normality_text, "\\nPer-group results:")
                for (group_name in names(normality_result$group_results)) {
                  group_result <- normality_result$group_results[[group_name]]
                  if (!is.null(group_result$note)) {
                    normality_text <- paste0(normality_text, "\\n  ", group_result$group, " (n=", group_result$n, "): ", group_result$note)
                  } else if (!is.null(group_result$is_normal)) {
                    normality_text <- paste0(normality_text, "\\n  ", group_result$group, " (n=", group_result$n, "): p=", sprintf("%.4f", group_result$p_value), " (", ifelse(group_result$is_normal, "normal", "non-normal"), ")")
                  }
                }
              }
            }

            # Add group summary statistics
            group_stats <- ""
            for (g in groups) {
              group_data <- value_col[group_col == g]
              group_stats <- paste0(group_stats, "\\n", g, ": n=", length(group_data), ", mean=", round(mean(group_data), 2), ", sd=", round(sd(group_data), 2))
            }

            # Add post-hoc test if ANOVA was significant
            posthoc_results <- ""
            if (grepl("ANOVA", test_name) && p_val < 0.05) {
              cat("Performing", selected_posthoc_test, "post-hoc test\\n")
              cat("DEBUG: selected_posthoc_test value is:", selected_posthoc_test, "\\n")

              # Perform ANOVA and post-hoc test
              tryCatch({
                # Create proper data frame for ANOVA
                anova_data <- data.frame(
                  group = factor(group_col),
                  value = as.numeric(value_col)
                )

                # Perform ANOVA
                anova_result <- aov(value ~ group, data = anova_data)

                # Perform the selected post-hoc test
                if (selected_posthoc_test == "tukey") {
                  posthoc_result <- TukeyHSD(anova_result)
                  posthoc_summary <- posthoc_result$group
                  posthoc_results <- "\\n\\nTukey HSD Post-hoc Comparisons:"

                  # Format detailed results (inside Tukey block)
                  for (i in 1:nrow(posthoc_summary)) {
                    comparison <- rownames(posthoc_summary)[i]
                    p_adj <- posthoc_summary[i, "p adj"]
                    diff <- posthoc_summary[i, "diff"]

                    # Determine significance
                    if (!is.na(p_adj)) {
                      if (p_adj < 0.001) sig <- "***"
                      else if (p_adj < 0.01) sig <- "**"
                      else if (p_adj < 0.05) sig <- "*"
                      else sig <- "ns"
                    } else {
                      sig <- "ns"
                    }

                    posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": diff=", round(diff, 2), ", p=", sprintf("%.4f", ifelse(is.na(p_adj), 1.0, p_adj)), " (", sig, ")")
                  }

                  # Make result globally available for plotting
                  assign("global_tukey_result", posthoc_result, envir = .GlobalEnv)

                } else if (selected_posthoc_test == "bonferroni") {
                  # Bonferroni correction using pairwise t-tests
                  posthoc_result <- pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "bonferroni")
                  posthoc_results <- "\\n\\nBonferroni Post-hoc Comparisons:"

                  # Convert to format similar to TukeyHSD for plotting
                  groups <- levels(anova_data$group)
                  n_groups <- length(groups)

                  # Create pairwise data frame
                  comparisons <- combn(groups, 2, simplify = FALSE)
                  p_adj_values <- c()
                  diff_values <- c()
                  comparison_names <- c()

                  for (comp in comparisons) {
                    g1 <- comp[1]
                    g2 <- comp[2]

                    # Get p-value from pairwise.t.test result
                    if (g2 %in% rownames(posthoc_result$p.value) && g1 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g2, g1]
                    } else if (g1 %in% rownames(posthoc_result$p.value) && g2 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g1, g2]
                    } else {
                      p_val <- NA
                    }

                    # Calculate mean difference
                    mean1 <- mean(anova_data$value[anova_data$group == g1], na.rm = TRUE)
                    mean2 <- mean(anova_data$value[anova_data$group == g2], na.rm = TRUE)
                    diff <- mean2 - mean1

                    comparison_names <- c(comparison_names, paste(g2, g1, sep = "-"))
                    p_adj_values <- c(p_adj_values, p_val)
                    diff_values <- c(diff_values, diff)
                  }

                  # Create summary matrix similar to TukeyHSD format
                  posthoc_summary <- data.frame(
                    diff = diff_values,
                    lwr = rep(NA, length(diff_values)),
                    upr = rep(NA, length(diff_values)),
                    "p adj" = p_adj_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )

                  # Format detailed results (inside Bonferroni block)
                  for (i in 1:nrow(posthoc_summary)) {
                    comparison <- rownames(posthoc_summary)[i]
                    p_adj <- posthoc_summary[i, "p adj"]
                    diff <- posthoc_summary[i, "diff"]

                    # Determine significance
                    if (!is.na(p_adj)) {
                      if (p_adj < 0.001) sig <- "***"
                      else if (p_adj < 0.01) sig <- "**"
                      else if (p_adj < 0.05) sig <- "*"
                      else sig <- "ns"
                    } else {
                      sig <- "ns"
                    }

                    posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": diff=", round(diff, 2), ", p=", sprintf("%.4f", ifelse(is.na(p_adj), 1.0, p_adj)), " (", sig, ")")
                  }

                } else if (selected_posthoc_test == "holm") {
                  # Holm correction using pairwise t-tests
                  posthoc_result <- pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "holm")
                  posthoc_results <- "\\n\\nHolm Post-hoc Comparisons:"

                  # Convert to format similar to TukeyHSD for plotting (same logic as Bonferroni)
                  groups <- levels(anova_data$group)
                  comparisons <- combn(groups, 2, simplify = FALSE)
                  p_adj_values <- c()
                  diff_values <- c()
                  comparison_names <- c()

                  for (comp in comparisons) {
                    g1 <- comp[1]
                    g2 <- comp[2]

                    if (g2 %in% rownames(posthoc_result$p.value) && g1 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g2, g1]
                    } else if (g1 %in% rownames(posthoc_result$p.value) && g2 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g1, g2]
                    } else {
                      p_val <- NA
                    }

                    mean1 <- mean(anova_data$value[anova_data$group == g1], na.rm = TRUE)
                    mean2 <- mean(anova_data$value[anova_data$group == g2], na.rm = TRUE)
                    diff <- mean2 - mean1

                    comparison_names <- c(comparison_names, paste(g2, g1, sep = "-"))
                    p_adj_values <- c(p_adj_values, p_val)
                    diff_values <- c(diff_values, diff)
                  }

                  posthoc_summary <- data.frame(
                    diff = diff_values,
                    lwr = rep(NA, length(diff_values)),
                    upr = rep(NA, length(diff_values)),
                    "p adj" = p_adj_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )

                  # Format detailed results (inside Holm block)
                  for (i in 1:nrow(posthoc_summary)) {
                    comparison <- rownames(posthoc_summary)[i]
                    p_adj <- posthoc_summary[i, "p adj"]
                    diff <- posthoc_summary[i, "diff"]

                    # Determine significance
                    if (!is.na(p_adj)) {
                      if (p_adj < 0.001) sig <- "***"
                      else if (p_adj < 0.01) sig <- "**"
                      else if (p_adj < 0.05) sig <- "*"
                      else sig <- "ns"
                    } else {
                      sig <- "ns"
                    }

                    posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": diff=", round(diff, 2), ", p=", sprintf("%.4f", ifelse(is.na(p_adj), 1.0, p_adj)), " (", sig, ")")
                  }

                } else if (selected_posthoc_test == "dunnett") {
                  # Dunnett test requires a control group - use first group as control
                  if (!requireNamespace("multcomp", quietly = TRUE)) {
                    cat("Installing multcomp package for Dunnett test\\n")
                    webr::install("multcomp")
                    library(multcomp)
                  } else {
                    library(multcomp)
                  }

                  # Set control group (from UI selection or default to first group)
                  control_group <- "${dunnettControl}"
                  if (control_group == "" || is.na(control_group)) {
                    control_group <- levels(anova_data$group)[1]
                  }
                  anova_data$group <- relevel(anova_data$group, ref = control_group)
                  anova_result <- aov(value ~ group, data = anova_data)

                  posthoc_result <- summary(glht(anova_result, linfct = mcp(group = "Dunnett")))
                  posthoc_results <- paste0("\\n\\nDunnett Post-hoc Comparisons (vs ", control_group, "):")

                  # Extract results
                  estimates <- posthoc_result$test$coefficients
                  p_values <- posthoc_result$test$pvalues
                  comparison_names <- names(estimates)

                  posthoc_summary <- data.frame(
                    diff = estimates,
                    lwr = rep(NA, length(estimates)),
                    upr = rep(NA, length(estimates)),
                    "p adj" = p_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )

                  # Format detailed results (inside Dunnett block)
                  for (i in 1:nrow(posthoc_summary)) {
                    comparison <- rownames(posthoc_summary)[i]
                    p_adj <- posthoc_summary[i, "p adj"]
                    diff <- posthoc_summary[i, "diff"]

                    # Determine significance
                    if (!is.na(p_adj)) {
                      if (p_adj < 0.001) sig <- "***"
                      else if (p_adj < 0.01) sig <- "**"
                      else if (p_adj < 0.05) sig <- "*"
                      else sig <- "ns"
                    } else {
                      sig <- "ns"
                    }

                    posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": diff=", round(diff, 2), ", p=", sprintf("%.4f", ifelse(is.na(p_adj), 1.0, p_adj)), " (", sig, ")")
                  }

                } else if (selected_posthoc_test == "steel") {
                  # Steel test - non-parametric equivalent of Dunnett (vs control)
                  # Try kSamples::Steel.test
                  steel_available <- tryCatch({
                    if (!requireNamespace("kSamples", quietly = TRUE)) {
                      cat("Installing kSamples package for Steel test\\n")
                      webr::install("kSamples")
                    }
                    library(kSamples)
                    TRUE
                  }, error = function(e) {
                    cat("kSamples not available:", e$message, "\\n")
                    FALSE
                  })

                  if (steel_available) {
                    # Set control group (from UI selection or default to first group)
                    control_group <- "${dunnettControl}"
                    if (control_group == "" || is.na(control_group)) {
                      control_group <- levels(anova_data$group)[1]
                    }

                    all_groups <- levels(anova_data$group)
                    treatment_groups <- all_groups[all_groups != control_group]

                    posthoc_results <- paste0("\\n\\nSteel Post-hoc Comparisons (vs ", control_group, "):")

                    control_values <- anova_data$value[anova_data$group == control_group]

                    for (trt in treatment_groups) {
                      trt_values <- anova_data$value[anova_data$group == trt]
                      steel_result <- Steel.test(list(control_values, trt_values))
                      p_val <- steel_result$st[2]

                      comparison <- paste0(trt, " - ", control_group)

                      if (p_val < 0.001) sig <- "***"
                      else if (p_val < 0.01) sig <- "**"
                      else if (p_val < 0.05) sig <- "*"
                      else sig <- "ns"

                      posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": p=", sprintf("%.4f", p_val), " (", sig, ")")
                    }
                  } else {
                    posthoc_results <- "\\n\\n(Note: Steel test not available in webR)"
                  }

                } else if (selected_posthoc_test == "dunn") {
                  # Dunn test for non-parametric post-hoc
                  if (!requireNamespace("dunn.test", quietly = TRUE)) {
                    cat("Installing dunn.test package...\\n")
                    webr::install("dunn.test")
                  }
                  library(dunn.test)

                  posthoc_result <- dunn.test(anova_data$value, anova_data$group, method = "bonferroni")
                  posthoc_results <- "\\n\\nDunn Post-hoc Comparisons (non-parametric):"

                  # Extract results from dunn.test
                  comparison_names <- posthoc_result$comparisons
                  p_values <- posthoc_result$P.adjusted
                  z_values <- posthoc_result$Z

                  posthoc_summary <- data.frame(
                    diff = z_values,  # Use Z-statistic as effect size
                    lwr = rep(NA, length(z_values)),
                    upr = rep(NA, length(z_values)),
                    "p adj" = p_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )

                  # Format detailed results (inside Dunn block)
                  for (i in 1:nrow(posthoc_summary)) {
                    comparison <- rownames(posthoc_summary)[i]
                    p_adj <- posthoc_summary[i, "p adj"]
                    diff <- posthoc_summary[i, "diff"]

                    # Determine significance
                    if (!is.na(p_adj)) {
                      if (p_adj < 0.001) sig <- "***"
                      else if (p_adj < 0.01) sig <- "**"
                      else if (p_adj < 0.05) sig <- "*"
                      else sig <- "ns"
                    } else {
                      sig <- "ns"
                    }

                    posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": Z=", round(diff, 2), ", p=", sprintf("%.4f", ifelse(is.na(p_adj), 1.0, p_adj)), " (", sig, ")")
                  }
                } else {
                  posthoc_results <- paste0("\\n\\nUnknown post-hoc test: ", selected_posthoc_test)
                }

                # Make post-hoc result globally available for plotting (all test types)
                assign("global_posthoc_result", posthoc_summary, envir = .GlobalEnv)

              }, error = function(e) {
                posthoc_results <<- paste0("\\n\\n", selected_posthoc_test, " Post-hoc Error: ", e$message)
              })
            } else if (grepl("Kruskal-Wallis", test_name) && p_val < 0.05) {
              # Kruskal-Wallis is significant - run non-parametric post-hoc test
              cat("Kruskal-Wallis significant, performing non-parametric post-hoc test\\n")
              cat("DEBUG: selected_posthoc_test =", selected_posthoc_test, "\\n")

              tryCatch({
                # Create proper data frame
                anova_data <- data.frame(
                  group = factor(group_col),
                  value = as.numeric(value_col)
                )

                # Determine which non-parametric post-hoc test to use
                # Dunnett → Steel, others → Dunn
                actual_posthoc <- if (selected_posthoc_test == "dunnett" || selected_posthoc_test == "steel") "steel" else "dunn"
                cat("DEBUG: actual_posthoc =", actual_posthoc, "\\n")
                cat("Using non-parametric post-hoc:", actual_posthoc, "\\n")

                if (actual_posthoc == "steel") {
                  # Steel test - non-parametric equivalent of Dunnett
                  # Try kSamples::Steel.test
                  steel_available <- tryCatch({
                    if (!requireNamespace("kSamples", quietly = TRUE)) {
                      cat("Attempting to install kSamples package for Steel test\\n")
                      webr::install("kSamples")
                    }
                    library(kSamples)
                    TRUE
                  }, error = function(e) {
                    cat("kSamples not available:", e$message, "\\n")
                    FALSE
                  })

                  if (steel_available) {
                    # Set control group
                    control_group <- "${dunnettControl}"
                    if (control_group == "" || is.na(control_group)) {
                      control_group <- levels(anova_data$group)[1]
                    }

                    # Get all groups and identify treatment groups
                    all_groups <- levels(anova_data$group)
                    treatment_groups <- all_groups[all_groups != control_group]

                    posthoc_results <- paste0("\\n\\nSteel's Post-hoc Comparisons (vs ", control_group, "):")

                    # Run Steel.test for each treatment vs control
                    control_values <- anova_data$value[anova_data$group == control_group]

                    for (trt in treatment_groups) {
                      trt_values <- anova_data$value[anova_data$group == trt]

                      # Steel.test compares two samples - p-value is in st[2]
                      steel_result <- Steel.test(list(control_values, trt_values))
                      p_val <- steel_result$st[2]

                      comparison <- paste0(trt, " - ", control_group)

                      if (p_val < 0.001) sig <- "***"
                      else if (p_val < 0.01) sig <- "**"
                      else if (p_val < 0.05) sig <- "*"
                      else sig <- "ns"

                      posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": p=", sprintf("%.4f", p_val), " (", sig, ")")
                    }
                  } else {
                    # Fallback to Dunn
                    cat("Steel test not available, falling back to Dunn test\\n")
                    actual_posthoc <- "dunn"
                    if (selected_posthoc_test == "dunnett") {
                      posthoc_results <- "\\n\\n(Note: Data is non-normal. Dunnett test requires normal data. Using Dunn test instead.)"
                    } else {
                      posthoc_results <- "\\n\\n(Note: Steel test not available in webR, using Dunn test instead)"
                    }
                  }
                }

                if (actual_posthoc == "dunn") {
                  # Dunn test - non-parametric pairwise comparisons
                  if (!requireNamespace("dunn.test", quietly = TRUE)) {
                    cat("Installing dunn.test package...\\n")
                    webr::install("dunn.test")
                  }
                  library(dunn.test)

                  posthoc_result <- dunn.test(anova_data$value, anova_data$group, method = "bonferroni")
                  # Append to existing note (if Steel fallback) or create new
                  posthoc_results <- paste0(posthoc_results, "\\n\\nDunn's Post-hoc Comparisons (Bonferroni-adjusted p-values):")

                  comparison_names <- posthoc_result$comparisons
                  p_values <- posthoc_result$P.adjusted

                  for (i in seq_along(comparison_names)) {
                    p_val <- p_values[i]
                    comparison <- comparison_names[i]

                    if (p_val < 0.001) sig <- "***"
                    else if (p_val < 0.01) sig <- "**"
                    else if (p_val < 0.05) sig <- "*"
                    else sig <- "ns"

                    posthoc_results <- paste0(posthoc_results, "\\n", comparison, ": p=", sprintf("%.4f", p_val), " (", sig, ")")
                  }
                }

              }, error = function(e) {
                posthoc_results <<- paste0("\\n\\nPost-hoc Error: ", e$message)
              })
            } else if (is.na(p_val)) {
              posthoc_results <- "\\n\\nNo post-hoc test performed (statistical test failed - insufficient data)"
            } else if (p_val >= 0.05) {
              posthoc_results <- "\\n\\nNo post-hoc test performed (overall test not significant)"
            } else {
              posthoc_results <- "\\n\\nPost-hoc test not performed"
            }

            full_result <- paste0(main_result, normality_text, group_stats, posthoc_results)
            full_result
          } else {
            "ERROR: Need at least 2 groups"
          }
        } else {
          "ERROR: Need at least 2 columns"
        }
      }
    `);
    
    const testResult = await detailedTestResult.toJs();
    console.log("Detailed test result:", testResult);
    
    let formattedResult;
    if (testResult && testResult.values && testResult.values[0]) {
      const resultText = testResult.values[0];
      if (resultText.startsWith("ERROR:")) {
        setStatus(resultText);
        return;
      }
      formattedResult = resultText;
    } else {
      formattedResult = "Statistical analysis failed";
    }
    
    // Write to selected Excel cell
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      
      range.values = [[formattedResult]];
      
      // Add formatting
      range.format.font.bold = true;
      range.format.font.size = 10;
      range.format.fill.color = formattedResult.includes("ns") ? "#FFEEEE" : "#EEFFEE";
      
      await context.sync();
      
      setStatus(`Statistical results exported: ${formattedResult}`);
    });

  } catch (error) {
    console.error("Statistical export error:", error);
    setStatus("Statistical results export error: " + (error?.message || error));
  }
}

// Export statistical results to a new Excel sheet
async function exportStatisticalResultsToNewSheet() {
  try {
    await initWebR();

    // Check if statistics are enabled
    const addStatistics = document.getElementById("addStatistics")?.checked;
    if (!addStatistics) {
      setStatus("Statistical analysis not enabled. Please check 'Show statistical significance' and run Preview first.");
      return;
    }

    // Get chart type
    const chartType = document.getElementById("chartType")?.value || "";

    // Check if chart type supports statistics (silently return if not)
    const STATS_SUPPORTED_CHART_TYPES = [
      "bar_error_dot", "bar_grouped_error_dot",
      "box", "box_dot", "box_grouped", "box_grouped_dot",
      "violin_dot", "violin_grouped", "violin_grouped_dot",
      "line_grouped_error_raw"
    ];

    if (!STATS_SUPPORTED_CHART_TYPES.includes(chartType)) {
      // Silently return - UI already prevents checking statistics for unsupported types
      return;
    }

    let statResults = null;

    // Get statistical results based on chart type
    if (chartType === "line_grouped_error_raw") {
      statResults = await getLineStatisticalResultsText();
    } else if (chartType === "bar_grouped" || chartType === "bar_grouped_error" || chartType === "bar_grouped_error_dot" || chartType === "box_grouped" || chartType === "box_grouped_dot" || chartType === "violin_grouped" || chartType === "violin_grouped_dot") {
      statResults = await getGroupedBarStatisticalResultsText();
    } else {
      // For other chart types, use the general function
      statResults = await getStatisticalResultsText();
    }

    // Check if we have valid results
    if (!statResults || statResults.includes("No statistical") || statResults.includes("ERROR") || statResults.includes("Statistics disabled")) {
      // Provide user-friendly error messages
      let errorMsg = "Statistical results not available";

      if (statResults && statResults.includes("missing value where TRUE/FALSE needed")) {
        // Count actual data points per group
        const chartType = document.getElementById('chartType').value;
        const data = window.lastProcessedData;
        const settings = window.lastPlotSettings;

        if (data && data.length > 1 && settings) {
          const isGrouped = chartType.includes('grouped');

          if (isGrouped) {
            const groupColIndex = (settings.groupColIndex || 1) - 1;
            const xColIndex = (settings.xColIndex || 2) - 1;

            // Count data points for each group-category combination
            const counts = {};
            for (let i = 1; i < data.length; i++) {
              const group = data[i][groupColIndex];
              const category = data[i][xColIndex];
              const key = `${group} (${category})`;
              counts[key] = (counts[key] || 0) + 1;
            }

            const countList = Object.entries(counts).map(([key, n]) => `${key}: n=${n}`).join(', ');
            errorMsg += `: Data points - ${countList}. Statistical tests require at least n=2 per group.`;
          } else {
            const xColIndex = (settings.xColIndex || 1) - 1;

            // Count data points for each category
            const counts = {};
            for (let i = 1; i < data.length; i++) {
              const category = data[i][xColIndex];
              counts[category] = (counts[category] || 0) + 1;
            }

            const countList = Object.entries(counts).map(([cat, n]) => `${cat}: n=${n}`).join(', ');
            errorMsg += `: Data points - ${countList}. Statistical tests require at least n=2 per group.`;
          }
        } else {
          errorMsg += ": Insufficient data. Statistical tests require at least 2 data points per group.";
        }
      } else if (statResults && statResults.includes("ERROR")) {
        // Don't show raw R error messages - they're too technical
        errorMsg += ": An error occurred during statistical analysis. Please check your data.";
      } else if (statResults) {
        errorMsg += `: ${statResults.substring(0, 100)}`;
      }

      setStatus(errorMsg);
      return;
    }

    // Create new sheet and insert results
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;

      // Create unique sheet name
      const chartTitle = document.getElementById("titleText")?.value?.trim() || chartType;
      const baseSheetName = `Stats_${chartTitle}`.substring(0, 25);
      let sheetName = baseSheetName;

      // Check if sheet exists and create unique name
      sheets.load("items/name");
      await context.sync();

      let counter = 1;
      while (sheets.items.some(s => s.name === sheetName)) {
        sheetName = `${baseSheetName}_${counter}`;
        counter++;
      }

      // Create new sheet
      const newSheet = sheets.add(sheetName);
      newSheet.activate();

      // Parse and format the results
      const formattedResults = statResults.replace(/\\n/g, '\n');
      const allLines = formattedResults.split('\n');

      // Filter out empty lines
      const lines = allLines.filter(line => line.trim().length > 0);

      // Insert results line by line for better formatting
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        const cellRange = newSheet.getRangeByIndexes(i, 0, 1, 1);
        cellRange.values = [[line]];

        // Format headers FIRST (with priority over significance colors)
        if (line.startsWith('===') || line.startsWith('Category') ||
            line.startsWith('Test:') || line.startsWith('Normality') ||
            line.startsWith('Post-hoc')) {
          cellRange.format.font.bold = true;
          cellRange.format.fill.color = "#E7E6E6";
        }
        // Only apply significance colors to NON-header lines
        else if (line.includes('***') || line.includes('**') || line.includes('*')) {
          cellRange.format.fill.color = "#C6EFCE"; // Light green
        }
        // Highlight non-significant
        else if (line.includes('(ns)') || line.includes('ns ')) {
          cellRange.format.fill.color = "#FFC7CE"; // Light red
        }
      }

      // Auto-fit column width
      const usedRange = newSheet.getUsedRange();
      usedRange.format.autofitColumns();

      // Add metadata at the top
      newSheet.getRangeByIndexes(0, 0, 4, 1).insert(Excel.InsertShiftDirection.down);

      const metaRange = newSheet.getRangeByIndexes(0, 0, 4, 1);
      metaRange.values = [
        [`STATISTICAL ANALYSIS RESULTS`],
        [`Chart: ${chartTitle}`],
        [`Date: ${new Date().toLocaleString()}`],
        [`Chart Type: ${chartType}`]
      ];
      metaRange.format.font.bold = true;
      metaRange.format.fill.color = "#4472C4";
      metaRange.format.font.color = "white";

      await context.sync();

      setStatus(`✅ Statistical results exported to new sheet: "${sheetName}"`);
      console.log(`Statistical results exported to sheet: ${sheetName}`);
    });

  } catch (error) {
    console.error("Export to new sheet error:", error);
    setStatus("❌ Error exporting statistics to new sheet: " + (error?.message || error));
  }
}

// ========= IC50 Curve Data Export Function =========

async function exportIC50CurveDataToExcel() {
  try {
    console.log("exportIC50CurveDataToExcel called");
    setStatus("Exporting IC50 results...");
    await initWebR();

    // Get IC50 results from R
    const ic50ResultR = await webR.evalR(`
      if (exists("ic50_result") && !is.na(ic50_result)) {
        as.numeric(ic50_result)
      } else {
        NA_real_
      }
    `);
    const ic50Value = await ic50ResultR.toNumber().catch(() => NaN);
    console.log("IC50 value:", ic50Value);

    // Get curve data from R as separate vectors
    const concR = await webR.evalR(`
      if (exists("ic50_curve_data") && !is.null(ic50_curve_data)) {
        as.numeric(ic50_curve_data$conc)
      } else {
        numeric(0)
      }
    `);
    const conc = await concR.toArray().catch(() => []);
    console.log("Concentration data length:", conc.length);

    const responseR = await webR.evalR(`
      if (exists("ic50_curve_data") && !is.null(ic50_curve_data)) {
        as.numeric(ic50_curve_data$response)
      } else {
        numeric(0)
      }
    `);
    const response = await responseR.toArray().catch(() => []);
    console.log("Response data length:", response.length);

    // Get model parameters from R - get names and values separately
    const paramNamesR = await webR.evalR(`
      if (exists("ic50_model_params") && length(ic50_model_params) > 0) {
        names(ic50_model_params)
      } else {
        character(0)
      }
    `);
    const paramNames = await paramNamesR.toArray().catch(() => []);

    const paramValuesR = await webR.evalR(`
      if (exists("ic50_model_params") && length(ic50_model_params) > 0) {
        as.numeric(unlist(ic50_model_params))
      } else {
        numeric(0)
      }
    `);
    const paramValues = await paramValuesR.toArray().catch(() => []);

    // Build model params object
    const modelParams = {};
    for (let i = 0; i < paramNames.length; i++) {
      modelParams[paramNames[i]] = paramValues[i];
    }
    console.log("Model params:", modelParams);

    // Check if we have valid data
    if (!conc || conc.length === 0) {
      console.log("IC50 curve data not available for export - no concentration data");
      return;
    }

    // Create new sheet and insert results
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;

      // Create unique sheet name
      const timestamp = new Date().toISOString().slice(11, 19).replace(/:/g, "");
      const sheetName = `IC50_Results_${timestamp}`;

      // Add new sheet
      const newSheet = sheets.add(sheetName);
      newSheet.activate();

      // Prepare data for Excel
      const outputData = [];

      // Header section - IC50 Result
      outputData.push(["IC50 Analysis Results"]);
      outputData.push([""]);
      outputData.push(["IC50 Value:", ic50Value && !isNaN(ic50Value) ? ic50Value : "Could not fit"]);
      outputData.push([""]);

      // Model parameters if available
      if (modelParams && Object.keys(modelParams).length > 0) {
        outputData.push(["Model Parameters:"]);
        for (const [key, value] of Object.entries(modelParams)) {
          outputData.push([`  ${key}:`, value]);
        }
        outputData.push([""]);
      }

      // Curve data header
      outputData.push(["Fitted Curve Data:"]);
      outputData.push(["Concentration", "Fitted Response"]);

      // Add all curve data points
      for (let i = 0; i < conc.length; i++) {
        outputData.push([conc[i], response[i]]);
      }

      // Write to Excel
      const range = newSheet.getRange(`A1:B${outputData.length}`);
      range.values = outputData.map(row => {
        // Ensure each row has exactly 2 columns
        if (row.length === 1) return [row[0], ""];
        return row.slice(0, 2);
      });

      // Format header
      const headerRange = newSheet.getRange("A1");
      headerRange.format.font.bold = true;
      headerRange.format.font.size = 14;

      const ic50LabelRange = newSheet.getRange("A3");
      ic50LabelRange.format.font.bold = true;

      // Auto-fit columns
      newSheet.getUsedRange().format.autofitColumns();

      await context.sync();

      setStatus(`✅ IC50 results exported to sheet: "${sheetName}"`);
      console.log(`IC50 curve data exported to sheet: ${sheetName}`);
    });

  } catch (error) {
    console.error("IC50 export error:", error);
    setStatus("❌ IC50 export error: " + (error?.message || error));
  }
}

// ========= Font issue complete fix functions =========

// 1. Font name validation and fallback (improved version)
function validateFont(fontName) {
  if (!fontName) return "Arial";
  
  const name = fontName.trim();
  
  // Only allow fonts selectable in UI
  const allowedFonts = [
    "Arial", "Times New Roman", "Georgia", "Verdana", 
    "Courier New", "Noto Sans JP", "Noto Serif JP"
  ];
  
  const normalized = allowedFonts.includes(name) ? name : "Arial";
  
  // Check font availability in browser (simplified version, not async)
  if (document.fonts && document.fonts.check) {
    const testSizes = ["12px", "14px"];
    const quotedName = normalized.includes(" ") ? `"${normalized}"` : normalized;
    
    for (const size of testSizes) {
      try {
        if (document.fonts.check(`${size} ${quotedName}`)) {
          console.log(`Font availability check: ${normalized} - AVAILABLE`);
          return normalized;
        }
      } catch (e) {
        // チェックエラーの場合は続行
        continue;
      }
    }
    
    // 利用不可の場合でも、WebRでマッピングできる可能性があるため警告のみ
    console.warn(`Font "${normalized}" availability check failed, but proceeding with WebR mapping`);
  }
  
  return normalized;
}

// 2. R用フォント名（UIで既に適切な名前が選択されているため、そのまま使用）
function getRFontName(uiFontName) {
  console.log(`DEBUG: getRFontName called with: "${uiFontName}"`);
  
  // UIで選択されたフォント名をそのまま使用（既に適切な形式）
  const validatedFont = validateFont(uiFontName);
  console.log(`DEBUG: Final R font name: "${validatedFont}"`);
  return validatedFont;
}

// 3. フォントスタック構築（使用するフォントのみ）
function buildCompleteFontStack(primaryFont) {
  const validatedFont = validateFont(primaryFont);
  
  const fontStacks = {
    "Times New Roman": "'Times New Roman', 'Times', serif",
    "Georgia": "'Georgia', 'Times New Roman', serif",
    "Arial": "'Arial', 'Helvetica', sans-serif",
    "Verdana": "'Verdana', 'Arial', sans-serif",
    "Courier New": "'Courier New', 'Courier', monospace",
    "Noto Sans JP": "'Noto Sans JP', 'Hiragino Kaku Gothic ProN', 'Yu Gothic', sans-serif",
    "Noto Serif JP": "'Noto Serif JP', 'Hiragino Kaku Gothic ProN', 'Yu Gothic', serif"
  };
  
  return fontStacks[validatedFont] || `'${validatedFont}', Arial, sans-serif`;
}
function ensureLocalFontFace(primary) {
  const family = validateFont(primary);
  const id = `localface-${family.replace(/\s+/g,'-')}`;
  if (document.getElementById(id)) return;

  const localSrcMap = {
    "Arial": "local('Arial'), local('ArialMT')",
    "Georgia": "local('Georgia')",
    "Verdana": "local('Verdana')",
    "Courier New": "local('Courier New'), local('CourierNewPSMT')",
    "Noto Sans JP": "local('Noto Sans JP')",
    "Noto Serif JP": "local('Noto Serif JP')",
    "Times New Roman": "local('Times New Roman'), local('TimesNewRomanPSMT'), local('Times-Roman')"
  };
  const src = localSrcMap[family];
  if (!src) return;

  const style = document.createElement('style');
  style.id = id;
  style.textContent = `
    @font-face {
      font-family: '${family}';
      src: ${src};
      font-weight: 400 700;
      font-style: normal italic;
    }`;
  document.head.appendChild(style);
}

function getChartType() {
  return document.getElementById("chartType")?.value || "histogram";
}

// 4. SVGフォント強制適用（デバッグ強化版）
function forceSvgFontSafe(svgText, primaryFont) {
  const validatedFont = validateFont(primaryFont);
  const fontName = validatedFont;

  let svg = String(svgText);
  console.log(`DEBUG: Safe font forcing for "${fontName}"`);
  
  // デバッグ: 元のSVG内のフォント情報を確認
  const fontMatches = svg.match(/font-family[^>]*/gi);
  console.log(`DEBUG: Original font references found:`, fontMatches);
  
  // より積極的にすべてのfont-family参照を置換
  
  // 1) すべての font-family="..." を置換
  svg = svg.replace(/font-family\s*=\s*"[^"]*"/gi, `font-family="${fontName}"`);
  
  // 2) すべての font-family='...' を置換  
  svg = svg.replace(/font-family\s*=\s*'[^']*'/gi, `font-family="${fontName}"`);
  
  // 3) すべての font-family: ... を置換
  svg = svg.replace(/font-family\s*:\s*[^;}\n]*/gi, `font-family:${fontName}`);
  
  // 4) 汎用フォント名を明示的に置換
  svg = svg.replace(/\bsans-serif\b/gi, fontName);
  svg = svg.replace(/\bserif\b/gi, fontName);
  svg = svg.replace(/\bmonospace\b/gi, fontName);
  
  // デバッグ: 置換後のフォント情報を確認
  const newFontMatches = svg.match(/font-family[^>]*/gi);
  console.log(`DEBUG: Font references after replacement:`, newFontMatches);
  console.log(`DEBUG: Safe font forcing complete for "${fontName}"`);
  
  return svg;
}




// 5. フォント読み込み確実化（簡素化版）
async function ensureFontLoaded(fontName) {
  if (!document.fonts || !document.fonts.load) return false;
  
  const validatedFont = validateFont(fontName);
  const quotedName = validatedFont.includes(" ") ? `"${validatedFont}"` : validatedFont;
  
  try {
    console.log(`Loading font: ${validatedFont}`);
    
    // 基本的な読み込み試行
    const loadPromises = [
      document.fonts.load(`12px ${quotedName}`).catch(() => null),
      document.fonts.load(`normal 12px ${quotedName}`).catch(() => null),
      document.fonts.load(`bold 12px ${quotedName}`).catch(() => null)
    ];
    
    await Promise.allSettled(loadPromises);
    
    // フォントシステムの準備完了を待つ
    await document.fonts.ready;
    
    // 少し待つ（フォントレンダリングシステムの安定化）
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // 読み込み確認
    const testConfigs = [
      `12px ${quotedName}`,
      `normal 12px ${quotedName}`
    ];
    
    let loaded = false;
    for (const config of testConfigs) {
      try {
        if (document.fonts.check(config)) {
          loaded = true;
          console.log(`Font "${validatedFont}" verified with config: ${config}`);
          break;
        }
      } catch (e) {
        continue;
      }
    }
    
    console.log(`Font "${validatedFont}" final status: ${loaded ? 'LOADED' : 'FAILED'}`);
    return loaded;
    
  } catch (e) {
    console.error(`Critical error loading font "${validatedFont}":`, e);
    return false;
  }
}

// 6. WebR環境でのフォント確認デバッグ関数（強化版）
async function debugWebRFonts() {
  await initWebR();
  
  const fontCheck = `
    # システムフォントの確認
    cat("=== WebR Font Debug ===\\n")
    
    # systemfonts パッケージのインストール試行
    if (!requireNamespace("systemfonts", quietly = TRUE)) {
      cat("Installing systemfonts...\\n")
      tryCatch({
        webr::install("systemfonts")
        library(systemfonts)
      }, error = function(e) {
        cat("systemfonts installation failed:", e$message, "\\n")
      })
    }
    
    # ggplot2でのフォント確認
    library(ggplot2)
    
    # Check ggpubr availability for statistical analysis
    if (${addStatistics ? 'TRUE' : 'FALSE'} == TRUE) {
      cat("*** CHECKING GGPUBR PACKAGE FOR STATISTICAL ANALYSIS ***\\n")
      if (requireNamespace("ggpubr", quietly = TRUE)) {
        cat("ggpubr package is AVAILABLE\\n")
        tryCatch({
          library(ggpubr)
          cat("ggpubr package loaded successfully\\n")
        }, error = function(e) {
          cat("ERROR loading ggpubr:", e$message, "\\n")
        })
      } else {
        cat("WARNING: ggpubr package is NOT AVAILABLE - statistical analysis will be skipped\\n")
        cat("To enable statistical analysis, install ggpubr package\\n")
      }
    }
    
    # テスト用の簡単なプロット
    test_data <- data.frame(x = 1:3, y = 1:3)
    
    fonts_to_test <- c("Arial", "Times New Roman", "Helvetica", "sans", "serif")
    
    for (font_name in fonts_to_test) {
      cat("Testing font:", font_name, "\\n")
      tryCatch({
        p <- ggplot(test_data, aes(x, y)) + 
          geom_point() + 
          ggtitle(paste("Test:", font_name)) +
          theme_minimal(base_family = font_name) +
          theme(
            text = element_text(family = font_name),
            plot.title = element_text(family = font_name, size = 24),
            axis.title = element_text(family = font_name, size = 12),
            axis.text = element_text(family = font_name, size = 10)
          )
        
        # SVG出力テスト
        tf <- tempfile(fileext = ".svg")
        svglite::svglite(tf, width = 4, height = 3)
        print(p)
        grDevices::dev.off()
        svg_content <- paste(readLines(tf, warn = FALSE), collapse = "\\n")
        
        # SVG内でのフォント確認
        if (grepl(font_name, svg_content, fixed = TRUE)) {
          cat("Font", font_name, "- SUCCESS (found in SVG)\\n")
        } else {
          cat("Font", font_name, "- WARNING (not found in SVG)\\n")
        }
        
      }, error = function(e) {
        cat("Font", font_name, "- ERROR:", e$message, "\\n")
      })
    }
    
    cat("Font debug complete\\n")
  `;
  
  try {
    await webR.evalRVoid(fontCheck);
    console.log("WebR font debug completed - check console for R output");
  } catch (e) {
    console.error("Font debug error:", e);
  }
}

// ========= 既存コード（WebR初期化など）=========

// ---- WebR を実行時に読み込む（Webpack にバンドルさせない）----
async function initWebR() {
  if (webrReady) return;
  setStatus("Loading data...");

  const WebR = window.WebR;
  if (!WebR) { setStatus("❌ Load failed. Please refresh and try again."); return; }

  webR = new WebR();
  await webR.init();
  await webR.evalRVoid('options(device = webr::canvas)');

  // Define variables used in R function templates
  const showMainStatSymbol = false; // Removed from UI - always disabled
  const showPairwiseComparisons = true; // Default enabled
  const yScale = "linear"; // Default scale for R function definitions

  // ggplot2 & dplyr（未導入なら取得）
  await webR.evalRVoid(`
    if (!requireNamespace("ggplot2", quietly = TRUE)) webr::install("ggplot2");
    if (!requireNamespace("dplyr", quietly = TRUE)) webr::install("dplyr");
    library(ggplot2)
    library(dplyr)

    suppressPackageStartupMessages(library(ggplot2))
    suppressPackageStartupMessages(library(dplyr))
    if (!requireNamespace("svglite", quietly = TRUE)) webr::install("svglite")

    # family をそのまま base_family へ渡す（Times/Helvetica/Arial/Noto…OK）
    sato_theme <- function(theme="classic", base_size=20, family="Arial",
  title_face="plain", axis_title_face="plain", axis_text_face="plain",
  title_size=24, x_axis_title_size=20, y_axis_title_size=20, x_axis_text_size=18, y_axis_text_size=18, legend_text_size=16) {

  th <- switch(tolower(theme),
    "minimal"  = theme_minimal(base_size=base_size, base_family=family),
    "light"    = theme_light  (base_size=base_size, base_family=family),
    "bw"       = theme_bw     (base_size=base_size, base_family=family),
    "void"     = theme_void   (base_size=base_size, base_family=family),
    "grey"     = theme_grey   (base_size=base_size, base_family=family),
    "gray"     = theme_grey   (base_size=base_size, base_family=family),
    "linedraw" = theme_linedraw(base_size=base_size, base_family=family),
    "dark"     = theme_dark   (base_size=base_size, base_family=family),
    "test"     = theme_test   (base_size=base_size, base_family=family),
                 theme_classic(base_size=base_size, base_family=family))

  th + theme(
    text               = element_text(family = family),

    plot.background    = element_rect(fill="white", color=NA),
    panel.background   = element_rect(fill="white", color=NA),
    panel.grid.major.y = element_line(color="#e5e7eb", linewidth=0.6),
    panel.grid.minor   = element_blank(),
    axis.line          = element_line(linewidth=0.6),
    axis.ticks         = element_line(linewidth=0.5),

    plot.title   = element_text(family=family, face=title_face, hjust=0.5, size=title_size, margin=margin(b=6)),
    axis.title.x = element_text(family=family, face=axis_title_face, size=x_axis_title_size, margin=margin(t=6)),
    axis.title.y = element_text(family=family, face=axis_title_face, size=y_axis_title_size, margin=margin(r=6)),
    axis.text.x  = element_text(family=family, face=axis_text_face, size=x_axis_text_size, color="black"),
    axis.text.y  = element_text(family=family, face=axis_text_face, size=y_axis_text_size, color="black"),
    legend.text  = element_text(family=family, size=legend_text_size),
    legend.title = element_text(family=family, size=legend_text_size)
  )
}

    sato_limits <- function(xlim=NULL, ylim=NULL) {
      L <- list()
      if (!is.null(xlim)) L <- c(L, list(scale_x_continuous(limits=xlim, expand=expansion(mult=c(0.02,0.02)))))
      if (!is.null(ylim)) L <- c(L, list(scale_y_continuous(limits=ylim, expand=expansion(mult=c(0.02,0.02)))))
      L
    }

    # Statistical analysis functions
    sato_check_normality <- function(data, group_col=NULL) {
      if (is.null(group_col)) {
        # Single group normality test
        if (length(data) < 3 || length(data) > 5000) {
          return(list(
            is_normal = TRUE,
            p_value = 1.0,
            test = "Sample size outside range for Shapiro-Wilk test"
          ))
        }
        shapiro_result <- shapiro.test(as.numeric(data))
        return(list(
          is_normal = shapiro_result$p.value > 0.05,
          p_value = shapiro_result$p.value,
          test = "Shapiro-Wilk"
        ))
      } else {
        # Multiple groups normality test
        groups <- unique(data[[group_col]])
        is_normal <- TRUE
        min_p <- 1

        for (group in groups) {
          group_data <- data[data[[group_col]] == group, ]
          group_values <- as.numeric(group_data[[ncol(group_data)]])
          group_values <- group_values[!is.na(group_values)]  # Remove NAs

          if (length(group_values) >= 3 && length(group_values) <= 5000) {
            shapiro_result <- shapiro.test(group_values)
            if (shapiro_result$p.value < 0.05) {
              is_normal <- FALSE
            }
            min_p <- min(min_p, shapiro_result$p.value)
          }
        }
        
        return(list(
          is_normal = is_normal,
          p_value = min_p,
          test = "Shapiro-Wilk (all groups)"
        ))
      }
    }
    
    sato_perform_statistical_test <- function(data, group_col, value_col, test_type="auto", variance_test="levene", posthoc_test="tukey") {
      normality_result <- NULL
      auto_selected <- FALSE
      
      if (test_type == "auto") {
        # Check normality first
        normality_result <- sato_check_normality(data, group_col)
        auto_selected <- TRUE
        
        # Count groups
        groups <- unique(data[[group_col]])
        n_groups <- length(groups)
        
        if (n_groups == 2) {
          test_type <- if (normality_result$is_normal) "t.test" else "wilcox.test"
        } else {
          test_type <- if (normality_result$is_normal) "anova" else "kruskal.test"
        }
      }
      
      groups <- unique(data[[group_col]])
      n_groups <- length(groups)
      
      if (n_groups == 2) {
        group1_data <- data[data[[group_col]] == groups[1], value_col]
        group2_data <- data[data[[group_col]] == groups[2], value_col]
        
        if (test_type == "t.test") {
          # Check variance equality to choose appropriate t-test
          variance_test_type <- variance_test
          variance_p_value <- NA
          variance_test_name <- ""

          if (variance_test_type == "levene") {
            # Levene test for equality of variances
            combined_data <- data.frame(
              values = c(group1_data, group2_data),
              group = factor(c(rep(groups[1], length(group1_data)), rep(groups[2], length(group2_data))))
            )

            # Calculate Levene test manually
            group_means <- tapply(combined_data$values, combined_data$group, mean)
            abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
            levene_result <- anova(lm(abs_deviations ~ combined_data$group))
            levene_p <- levene_result$Pr[1]
            equal_variances <- levene_p > 0.05
            variance_p_value <- levene_p
            variance_test_name <- "Levene"
          } else {
            # F-test for equality of variances
            var_test <- var.test(group1_data, group2_data)
            equal_variances <- var_test$p.value > 0.05
            variance_p_value <- var_test$p.value
            variance_test_name <- "F-test"
          }

          result <- t.test(group1_data, group2_data, var.equal = equal_variances)
          test_name <- if (equal_variances) "Student's t-test" else "Welch's t-test"

          return(list(
            p_value = result$p.value,
            test_used = test_name,
            comparison = paste(groups[1], "vs", groups[2]),
            normality_result = normality_result,
            auto_selected = auto_selected,
            variance_test = variance_test_name,
            variance_p_value = variance_p_value,
            equal_variances = equal_variances
          ))
        } else {
          result <- wilcox.test(group1_data, group2_data)
          return(list(
            p_value = result$p.value,
            test_used = "Wilcoxon test",
            comparison = paste(groups[1], "vs", groups[2]),
            normality_result = normality_result,
            auto_selected = auto_selected
          ))
        }
      } else {
        # Multiple groups
        # Clean data by removing NAs from value column
        clean_data <- data[!is.na(as.numeric(data[[value_col]])), ]
        # Also ensure value column is numeric
        clean_data[[value_col]] <- as.numeric(clean_data[[value_col]])

        # Check if we still have enough groups after cleaning
        remaining_groups <- length(unique(clean_data[[group_col]]))
        if (remaining_groups < 2) {
          return(list(
            p_value = NA,
            test_used = "ERROR: Insufficient groups after removing NAs",
            normality_result = normality_result,
            auto_selected = auto_selected
          ))
        }

        if (test_type == "anova") {
          result <- aov(as.formula(paste(names(clean_data)[value_col], "~", names(clean_data)[group_col])), data=clean_data)
          summary_result <- summary(result)
          return(list(
            p_value = summary_result[[1]][["Pr(>F)"]][1],
            test_used = "ANOVA",
            comparison = "Overall",
            normality_result = normality_result,
            auto_selected = auto_selected
          ))
        } else {
          result <- kruskal.test(as.formula(paste(names(clean_data)[value_col], "~", names(clean_data)[group_col])), data=clean_data)
          return(list(
            p_value = result$p.value,
            test_used = "Kruskal-Wallis test",
            comparison = "Overall",
            normality_result = normality_result,
            auto_selected = auto_selected
          ))
        }
      }
    }
    
    sato_get_significance_symbol <- function(p_value) {
      cat("DEBUGGING sato_get_significance_symbol:\\n")
      cat("  Input p_value:", p_value, "\\n")
      cat("  p_value type:", class(p_value), "\\n")
      cat("  p_value < 0.001:", p_value < 0.001, "\\n")
      cat("  p_value < 0.01:", p_value < 0.01, "\\n")
      cat("  p_value < 0.05:", p_value < 0.05, "\\n")
      
      if (p_value < 0.001) {
        cat("  Returning: ***\\n")
        return("***")
      }
      if (p_value < 0.01) {
        cat("  Returning: **\\n")
        return("**")
      }
      if (p_value < 0.05) {
        cat("  Returning: *\\n")
        return("*")
      }
      cat("  Returning: ns\\n")
      return("ns")
    }
    
    sato_add_statistics_to_plot <- function(p, data, group_col, value_col, test_type="auto", symbol_size=15, show_main_symbol=TRUE, show_pairwise=TRUE, ggpubr_symbol_size=10, ggpubr_line_size=1.2, ggpubr_tip_length=0.04, ggpubr_vjust=-0.3, comparison_mode="significant", custom_comparisons="[]", custom_positions="{}", posthoc_test="tukey", dunnett_control="", y_scale="linear", stat_symbol_type="stars", custom_symbol_05="*", custom_symbol_01="**", custom_symbol_001="***", custom_symbol_ns="ns") {
      cat("\\n\\n🔥🔥🔥 STAT FUNCTION CALLED 🔥🔥🔥\\n")
      cat("🔥 show_pairwise parameter value:", show_pairwise, "\\n")
      cat("🔥 posthoc_test:", posthoc_test, "\\n")
      cat("🔥 comparison_mode:", comparison_mode, "\\n")
      cat("🔥 stat_symbol_type:", stat_symbol_type, "\\n")

      # Helper function to generate significance symbol based on p-value and symbol type
      get_sig_symbol <- function(p_val) {
        cat("🔥 get_sig_symbol called - p_val:", p_val, "stat_symbol_type:", stat_symbol_type, "\\n")
        if (stat_symbol_type == "pvalue") {
          # Return p-value formatted
          result <- sprintf("p=%.3f", p_val)
          cat("🔥 Returning p-value:", result, "\\n")
          return(result)
        } else if (stat_symbol_type == "custom") {
          # Use custom symbols
          cat("🔥 Using custom symbols - custom_symbol_05:", custom_symbol_05, "custom_symbol_01:", custom_symbol_01, "custom_symbol_001:", custom_symbol_001, "\\n")
          if (p_val < 0.001) {
            cat("🔥 p < 0.001, returning:", custom_symbol_001, "\\n")
            return(custom_symbol_001)
          } else if (p_val < 0.01) {
            cat("🔥 p < 0.01, returning:", custom_symbol_01, "\\n")
            return(custom_symbol_01)
          } else if (p_val < 0.05) {
            cat("🔥 p < 0.05, returning:", custom_symbol_05, "\\n")
            return(custom_symbol_05)
          } else {
            cat("🔥 p >= 0.05, returning:", custom_symbol_ns, "\\n")
            return(custom_symbol_ns)
          }
        } else {
          # Default to stars
          if (p_val < 0.001) return("***")
          else if (p_val < 0.01) return("**")
          else if (p_val < 0.05) return("*")
          else return("ns")
        }
      }

      # DIRECT GGPUBR IMPLEMENTATION - NO COMPLEX NESTING
      cat("🔥 GGPUBR CHECK: show_pairwise =", show_pairwise, "🔥\\n")

      if (show_pairwise) {
        cat("🔥🔥🔥 ENTERED show_pairwise BLOCK 🔥🔥🔥\\n")

        tryCatch({
          cat("🔥🔥🔥 DIRECT GGPUBR IMPLEMENTATION STARTING 🔥🔥🔥\\n")

          # Install and load ggpubr
          if (!require("ggpubr", quietly = TRUE)) {
            webr::install("ggpubr")
          }
          library(ggpubr)

          # Create dynamic ggpubr data frame based on actual group names
          y_max <- max(data[[value_col]], na.rm = TRUE)

          # Get actual group names from the data
          actual_groups <- unique(data[[group_col]])
          cat("🔥 Actual groups found:", paste(actual_groups, collapse = ", "), "🔥\\n")

          # Create ALL pairwise combinations dynamically and get REAL statistical results
          if (length(actual_groups) >= 2) {
            cat("🔥🔥🔥 NUMBER OF GROUPS:", length(actual_groups), "🔥🔥🔥\\n")
            cat("🔥🔥🔥 GROUP NAMES:", paste(actual_groups, collapse=", "), "🔥🔥🔥\\n")

            # Perform REAL statistical analysis to get actual p-values
            tryCatch({
              cat("🔥 Performing REAL statistical analysis for ggpubr 🔥\\n")

              # Convert data to proper format for statistical test
              stat_data <- data.frame(
                group = data[[group_col]],
                value = data[[value_col]]
              )

              # Step 1: Determine which test to use
              # For 2 groups: use direct 2-sample test (t-test or Wilcoxon)
              # For 3+ groups: use omnibus test (ANOVA or Kruskal-Wallis) + post-hoc

              if (length(actual_groups) == 2) {
                # TWO GROUP COMPARISON - Use direct test
                cat("🔥 TWO GROUPS DETECTED - Using direct 2-sample test 🔥\\n")

                # Determine which test to use for 2 groups
                two_group_test <- test_type
                if (test_type == "auto") {
                  normality_check <- sato_check_normality(stat_data, "group")
                  two_group_test <- if (normality_check$is_normal) "t.test" else "wilcox.test"
                  cat("🔥 AUTO TEST SELECTION (2 groups): normality =", normality_check$is_normal, "→ using", two_group_test, "🔥\\n")
                } else if (test_type == "parametric") {
                  two_group_test <- "t.test"
                  cat("🔥 MANUAL MODE (2 groups): parametric → using t-test 🔥\\n")
                } else if (test_type == "nonparametric") {
                  two_group_test <- "wilcox.test"
                  cat("🔥 MANUAL MODE (2 groups): non-parametric → using Wilcoxon 🔥\\n")
                }

                # Perform the 2-sample test
                group1_data <- stat_data$value[stat_data$group == actual_groups[1]]
                group2_data <- stat_data$value[stat_data$group == actual_groups[2]]

                if (two_group_test == "t.test") {
                  test_result <- t.test(group1_data, group2_data)
                  omnibus_test_name <- "Student's t-test"
                } else if (two_group_test == "wilcox.test") {
                  test_result <- wilcox.test(group1_data, group2_data)
                  omnibus_test_name <- "Wilcoxon rank-sum test"
                } else {
                  # Default to t-test for unknown test types
                  test_result <- t.test(group1_data, group2_data)
                  omnibus_test_name <- "Student's t-test"
                }

                omnibus_p_value <- test_result$p.value
                cat("🔥", omnibus_test_name, "p-value:", omnibus_p_value, "🔥\\n")

                # For 2 groups, create the comparison list directly
                combinations <- list(c(actual_groups[1], actual_groups[2]))
                real_p_values <- c(omnibus_p_value)

                # Convert to significance symbol
                real_p_symbols <- c(get_sig_symbol(omnibus_p_value))

                cat("🔥", actual_groups[1], "vs", actual_groups[2], ": p =", omnibus_p_value, "→", real_p_symbols[1], "🔥\\n")

              } else {
                # THREE OR MORE GROUPS - Use omnibus test + post-hoc
                cat("🔥 3+ GROUPS DETECTED - Using omnibus test 🔥\\n")

                omnibus_test_type <- test_type
                if (test_type == "auto") {
                  normality_check <- sato_check_normality(stat_data, "group")
                  omnibus_test_type <- if (normality_check$is_normal) "anova" else "kruskal.test"
                  cat("🔥 AUTO TEST SELECTION: normality =", normality_check$is_normal, "→ using", omnibus_test_type, "🔥\\n")
                } else if (test_type == "parametric") {
                  omnibus_test_type <- "anova"
                  cat("🔥 MANUAL MODE: parametric → using ANOVA 🔥\\n")
                } else if (test_type == "nonparametric") {
                  omnibus_test_type <- "kruskal.test"
                  cat("🔥 MANUAL MODE: non-parametric → using Kruskal-Wallis 🔥\\n")
                }

                # Perform omnibus test
                if (omnibus_test_type == "anova" || omnibus_test_type == "t.test") {
                anova_result <- aov(value ~ group, data = stat_data)
                omnibus_summary <- summary(anova_result)
                omnibus_p_value <- omnibus_summary[[1]][["Pr(>F)"]][1]
                omnibus_test_name <- "ANOVA"
                cat("🔥 ANOVA p-value:", omnibus_p_value, "🔥\\n")
              } else {
                kruskal_result <- kruskal.test(value ~ group, data = stat_data)
                omnibus_p_value <- kruskal_result$p.value
                omnibus_test_name <- "Kruskal-Wallis"
                cat("🔥 Kruskal-Wallis p-value:", omnibus_p_value, "🔥\\n")
              }

              # Step 3: Check if omnibus test is significant
              if (omnibus_p_value >= 0.05) {
                cat("🔥🔥🔥", omnibus_test_name, "NOT SIGNIFICANT (p =", omnibus_p_value, ") - SKIPPING POST-HOC TESTS 🔥🔥🔥\\n")
                # Return empty results - no pairwise comparisons needed
                real_p_values <- rep(1.0, length(combn(actual_groups, 2, simplify = FALSE)))
                real_p_symbols <- rep("n.s.", length(combn(actual_groups, 2, simplify = FALSE)))
                combinations <- combn(actual_groups, 2, simplify = FALSE)
              } else {
                cat("🔥🔥🔥", omnibus_test_name, "SIGNIFICANT (p =", omnibus_p_value, ") - PROCEEDING WITH POST-HOC TESTS 🔥🔥🔥\\n")

                # Step 4: Select appropriate post-hoc test based on omnibus test type
                # For Kruskal-Wallis, use non-parametric equivalents:
                #   - Dunnett (vs control, parametric) → Steel (vs control, non-parametric)
                #   - Tukey/Bonferroni/Holm (pairwise, parametric) → Dunn (pairwise, non-parametric)
                # For ANOVA, use user-selected post-hoc test
                actual_posthoc_test <- posthoc_test
                if (omnibus_test_type == "kruskal.test") {
                  if (posthoc_test == "dunnett") {
                    actual_posthoc_test <- "steel"
                    cat("🔥 Kruskal-Wallis detected - automatically using Steel test (non-parametric equivalent of Dunnett) 🔥\\n")
                  } else if (posthoc_test != "dunn" && posthoc_test != "steel") {
                    actual_posthoc_test <- "dunn"
                    cat("🔥 Kruskal-Wallis detected - automatically using Dunn test for post-hoc 🔥\\n")
                  }
                }

                # Step 5: Perform post-hoc test
                # Create anova_result for parametric tests if not already created
                if (!exists("anova_result")) {
                  anova_result <- aov(value ~ group, data = stat_data)
                }

                if (actual_posthoc_test == "tukey") {
                  posthoc_result <- TukeyHSD(anova_result)
                  posthoc_summary <- posthoc_result$group
                } else if (actual_posthoc_test == "bonferroni") {
                  # Bonferroni correction using pairwise t-tests
                  posthoc_result <- pairwise.t.test(stat_data$value, stat_data$group, p.adjust.method = "bonferroni")

                  # Convert to format similar to TukeyHSD
                  groups <- levels(factor(stat_data$group))
                  comparisons <- combn(groups, 2, simplify = FALSE)
                  p_adj_values <- c()
                  diff_values <- c()
                  comparison_names <- c()

                  for (comp in comparisons) {
                    g1 <- comp[1]
                    g2 <- comp[2]

                    if (g2 %in% rownames(posthoc_result$p.value) && g1 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g2, g1]
                    } else if (g1 %in% rownames(posthoc_result$p.value) && g2 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g1, g2]
                    } else {
                      p_val <- NA
                    }

                    mean1 <- mean(stat_data$value[stat_data$group == g1], na.rm = TRUE)
                    mean2 <- mean(stat_data$value[stat_data$group == g2], na.rm = TRUE)
                    diff <- mean2 - mean1

                    comparison_names <- c(comparison_names, paste(g2, g1, sep = "-"))
                    p_adj_values <- c(p_adj_values, p_val)
                    diff_values <- c(diff_values, diff)
                  }

                  posthoc_summary <- data.frame(
                    diff = diff_values,
                    lwr = rep(NA, length(diff_values)),
                    upr = rep(NA, length(diff_values)),
                    "p adj" = p_adj_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )
                } else if (actual_posthoc_test == "holm") {
                  # Holm correction using pairwise t-tests
                  posthoc_result <- pairwise.t.test(stat_data$value, stat_data$group, p.adjust.method = "holm")

                  groups <- levels(factor(stat_data$group))
                  comparisons <- combn(groups, 2, simplify = FALSE)
                  p_adj_values <- c()
                  diff_values <- c()
                  comparison_names <- c()

                  for (comp in comparisons) {
                    g1 <- comp[1]
                    g2 <- comp[2]

                    if (g2 %in% rownames(posthoc_result$p.value) && g1 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g2, g1]
                    } else if (g1 %in% rownames(posthoc_result$p.value) && g2 %in% colnames(posthoc_result$p.value)) {
                      p_val <- posthoc_result$p.value[g1, g2]
                    } else {
                      p_val <- NA
                    }

                    mean1 <- mean(stat_data$value[stat_data$group == g1], na.rm = TRUE)
                    mean2 <- mean(stat_data$value[stat_data$group == g2], na.rm = TRUE)
                    diff <- mean2 - mean1

                    comparison_names <- c(comparison_names, paste(g2, g1, sep = "-"))
                    p_adj_values <- c(p_adj_values, p_val)
                    diff_values <- c(diff_values, diff)
                  }

                  posthoc_summary <- data.frame(
                    diff = diff_values,
                    lwr = rep(NA, length(diff_values)),
                    upr = rep(NA, length(diff_values)),
                    "p adj" = p_adj_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )
                } else if (actual_posthoc_test == "dunnett") {
                  # Dunnett test requires a control group
                  cat("🔥 DUNNETT TEST STARTING 🔥\\n")
                  if (!requireNamespace("multcomp", quietly = TRUE)) {
                    cat("Installing multcomp package for Dunnett test\\n")
                    webr::install("multcomp")
                    library(multcomp)
                  } else {
                    library(multcomp)
                    cat("multcomp package already loaded\\n")
                  }

                  # Set control group (from UI selection or default to first group)
                  control_group <- dunnett_control
                  cat("🔥 dunnett_control parameter value: '", dunnett_control, "' (length:", nchar(dunnett_control), ") 🔥\\n")
                  if (is.null(control_group) || is.na(control_group) || control_group == "" || nchar(control_group) == 0) {
                    control_group <- levels(factor(stat_data$group))[1]
                    cat("🔥 No valid control specified, using first group:", control_group, "🔥\\n")
                  } else {
                    cat("🔥 Using specified control group:", control_group, "🔥\\n")
                  }
                  stat_data$group <- relevel(factor(stat_data$group), ref = control_group)
                  anova_result <- aov(value ~ group, data = stat_data)

                  posthoc_result <- summary(glht(anova_result, linfct = mcp(group = "Dunnett")))
                  estimates <- posthoc_result$test$coefficients
                  p_values <- posthoc_result$test$pvalues
                  comparison_names <- names(estimates)
                  cat("🔥 Dunnett comparison names:", paste(comparison_names, collapse=", "), "🔥\\n")

                  posthoc_summary <- data.frame(
                    diff = estimates,
                    lwr = rep(NA, length(estimates)),
                    upr = rep(NA, length(estimates)),
                    "p adj" = p_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )
                  cat("🔥 Dunnett posthoc_summary created with", nrow(posthoc_summary), "rows 🔥\\n")
                } else if (actual_posthoc_test == "steel") {
                  # Steel test - non-parametric equivalent of Dunnett (vs control)
                  # Using kSamples::Steel.test
                  cat("🔥 STEEL TEST STARTING (kSamples) 🔥\\n")

                  steel_available <- tryCatch({
                    if (!requireNamespace("kSamples", quietly = TRUE)) {
                      cat("Installing kSamples package for Steel test\\n")
                      webr::install("kSamples")
                    }
                    library(kSamples)
                    TRUE
                  }, error = function(e) {
                    cat("kSamples not available:", e$message, "\\n")
                    FALSE
                  })

                  if (steel_available) {
                    # Set control group (from UI selection or default to first group)
                    control_group <- dunnett_control
                    cat("🔥 steel control parameter value: '", dunnett_control, "' (length:", nchar(dunnett_control), ") 🔥\\n")
                    if (is.null(control_group) || is.na(control_group) || control_group == "" || nchar(control_group) == 0) {
                      control_group <- levels(factor(stat_data$group))[1]
                      cat("🔥 No valid control specified, using first group:", control_group, "🔥\\n")
                    } else {
                      cat("🔥 Using specified control group:", control_group, "🔥\\n")
                    }

                    all_groups <- levels(factor(stat_data$group))
                    treatment_groups <- all_groups[all_groups != control_group]

                    comparison_names <- c()
                    p_values <- c()

                    control_values <- stat_data$value[stat_data$group == control_group]

                    for (trt in treatment_groups) {
                      trt_values <- stat_data$value[stat_data$group == trt]
                      steel_result <- Steel.test(list(control_values, trt_values))
                      comparison_names <- c(comparison_names, paste0(trt, " - ", control_group))
                      p_values <- c(p_values, steel_result$st[2])  # p-value is in st[2]
                    }

                    cat("🔥 Steel comparison names:", paste(comparison_names, collapse=", "), "🔥\\n")

                    posthoc_summary <- data.frame(
                      diff = rep(NA, length(p_values)),
                      lwr = rep(NA, length(p_values)),
                      upr = rep(NA, length(p_values)),
                      "p adj" = p_values,
                      row.names = comparison_names,
                      check.names = FALSE
                    )
                    cat("🔥 Steel posthoc_summary created with", nrow(posthoc_summary), "rows 🔥\\n")
                  } else {
                    # Fallback to Dunn
                    cat("🔥 Steel not available, falling back to Dunn 🔥\\n")
                    actual_posthoc_test <- "dunn"
                  }
                }

                if (actual_posthoc_test == "dunn") {
                  # Dunn test for non-parametric post-hoc
                  if (!requireNamespace("dunn.test", quietly = TRUE)) {
                    cat("Installing dunn.test package...\\n")
                    webr::install("dunn.test")
                  }
                  library(dunn.test)

                  posthoc_result <- dunn.test(stat_data$value, stat_data$group, method = "bonferroni")
                  comparison_names <- posthoc_result$comparisons
                  p_values <- posthoc_result$P.adjusted
                  z_values <- posthoc_result$Z

                  posthoc_summary <- data.frame(
                    diff = z_values,
                    lwr = rep(NA, length(z_values)),
                    upr = rep(NA, length(z_values)),
                    "p adj" = p_values,
                    row.names = comparison_names,
                    check.names = FALSE
                  )
                } else if (actual_posthoc_test != "dunnett" && actual_posthoc_test != "steel") {
                  # Default fallback to Tukey (only if not Dunnett or Steel)
                  posthoc_result <- TukeyHSD(anova_result)
                  posthoc_summary <- posthoc_result$group
                }

                cat("🔥 REAL", actual_posthoc_test, "post-hoc completed 🔥\\n")
                print(posthoc_summary)

                # Generate combinations based on post-hoc test type
                if (actual_posthoc_test == "dunnett" || actual_posthoc_test == "steel") {
                  # For Dunnett/Steel: only control vs each treatment group
                  control_group <- dunnett_control
                  if (control_group == "" || is.na(control_group)) {
                    control_group <- levels(factor(stat_data$group))[1]
                  }

                  # Create combinations: control vs each other group
                  combinations <- list()
                  for (grp in actual_groups) {
                    if (grp != control_group) {
                      combinations[[length(combinations) + 1]] <- c(control_group, grp)
                    }
                  }
                  cat("🔥", actual_posthoc_test, "comparisons: control (", control_group, ") vs treatments 🔥\\n")
                } else {
                  # For other tests: all pairwise combinations
                  combinations <- combn(actual_groups, 2, simplify = FALSE)
                }

                cat("🔥🔥🔥 NUMBER OF COMBINATIONS CREATED:", length(combinations), "🔥🔥🔥\\n")
                for (i in 1:length(combinations)) {
                  cat("🔥 Combination", i, ":", paste(combinations[[i]], collapse=" vs "), "🔥\\n")
                }

                # Extract REAL p-values and significance symbols
                real_p_values <- c()
                real_p_symbols <- c()

                for (i in 1:length(combinations)) {
                  comp_name1 <- paste(combinations[[i]][2], combinations[[i]][1], sep = "-")
                  comp_name2 <- paste(combinations[[i]][1], combinations[[i]][2], sep = "-")

                  # Try to find the comparison in post-hoc results
                  # Try multiple formats: "B-A", "A-B", "B - A", "A - B"
                  comp_name1_nospace <- paste(combinations[[i]][2], combinations[[i]][1], sep = "-")
                  comp_name2_nospace <- paste(combinations[[i]][1], combinations[[i]][2], sep = "-")
                  comp_name1_space <- paste(combinations[[i]][2], combinations[[i]][1], sep = " - ")
                  comp_name2_space <- paste(combinations[[i]][1], combinations[[i]][2], sep = " - ")

                  if (comp_name1_nospace %in% rownames(posthoc_summary)) {
                    p_val <- posthoc_summary[comp_name1_nospace, "p adj"]
                  } else if (comp_name2_nospace %in% rownames(posthoc_summary)) {
                    p_val <- posthoc_summary[comp_name2_nospace, "p adj"]
                  } else if (comp_name1_space %in% rownames(posthoc_summary)) {
                    p_val <- posthoc_summary[comp_name1_space, "p adj"]
                  } else if (comp_name2_space %in% rownames(posthoc_summary)) {
                    p_val <- posthoc_summary[comp_name2_space, "p adj"]
                  } else {
                    p_val <- 1.0  # Default non-significant
                    cat("⚠️ Could not find comparison:", combinations[[i]][1], "vs", combinations[[i]][2], "\\n")
                    cat("   Available rownames:", paste(rownames(posthoc_summary), collapse=", "), "\\n")
                  }

                  real_p_values <- c(real_p_values, p_val)

                  # Convert to significance symbol
                  real_p_symbols <- c(real_p_symbols, get_sig_symbol(p_val))

                  cat("🔥", combinations[[i]][1], "vs", combinations[[i]][2], ": p =", p_val, "→", tail(real_p_symbols, 1), "🔥\\n")
                }
              }  # Close the omnibus significance check (else block)
              }  # Close the 3+ groups else block

              # Calculate unit_step as a fraction of y_range (scale-aware)
              calculate_unit_step <- function(symbols, symbol_size, y_range, num_comparisons, y_scale="linear", y_min=NULL, y_max=NULL) {
                # Base fraction adapts to number of comparisons (more comparisons = more space)
                if (num_comparisons <= 3) {
                  base_fraction <- 0.25  # Current spacing for 2-3 comparisons
                } else if (num_comparisons <= 5) {
                  base_fraction <- 0.35  # More for 4-5 comparisons
                } else {
                  base_fraction <- 0.40  # Even more for 6+ comparisons
                }

                # Symbol count factor (more asterisks need more space)
                symbol_counts <- sapply(symbols, function(s) {
                  if (s == "***") return(2.0)
                  else if (s == "**") return(1.5)
                  else if (s == "*") return(1.2)
                  else return(1.0)  # for "n.s."
                })
                max_symbol_count <- max(symbol_counts)
                symbol_count_factor <- max_symbol_count

                # Symbol size factor (larger symbols need slightly more space)
                symbol_size_factor <- 1.0 + (symbol_size - 10) / 20  # 10=1.0, 15=1.25, 20=1.5

                # Calculate fraction
                fraction <- base_fraction * symbol_count_factor * symbol_size_factor

                # For all scales, use fraction of data range
                # ggplot2 will apply log transformation uniformly to all y-values
                unit_step <- fraction * y_range

                cat("🔥 Unit step calculation: num_comparisons=", num_comparisons,
                    ", base_fraction=", base_fraction,
                    ", symbol_factor=", symbol_count_factor,
                    ", size_factor=", symbol_size_factor,
                    ", fraction=", fraction,
                    ", y_range=", y_range,
                    ", unit_step=", unit_step, "🔥\\n")

                return(max(unit_step, y_range * 0.08))  # Minimum 8% of range for linear
              }

              # Calculate the dynamic step for this data (unit_step)
              y_min <- min(data[[value_col]], na.rm = TRUE)
              y_range <- y_max - y_min
              num_comparisons <- length(combinations)
              unit_step <- calculate_unit_step(real_p_symbols, ggpubr_symbol_size, y_range, num_comparisons, y_scale, y_min, y_max)
              cat("🔥 CALCULATED unit_step =", unit_step, "for y_range =", y_range, ", num_comparisons =", num_comparisons, ", y_scale =", y_scale, "🔥\\n")

              # Parse custom positions ONLY if user is in "custom" comparison mode
              user_positions <- list()
              has_custom_pos <- FALSE

              # Only use custom positions when explicitly in custom mode
              if (comparison_mode == "custom" && custom_positions != "{}" && custom_positions != "") {
                library(jsonlite)
                user_positions <- fromJSON(custom_positions)
                has_custom_pos <- length(user_positions) > 0
                cat("🔥 Using CUSTOM positions:", length(user_positions), "entries 🔥\\n")
              } else {
                cat("🔥 Using DATA-DRIVEN bracket positioning algorithm 🔥\\n")
              }

              # NEW ALGORITHM: Data-driven bracket positioning
              # For each comparison, calculate baseHeight = max of the two groups being compared

              # DEBUG: Print combinations RIGHT BEFORE copying
              cat("🔥🔥🔥 ABOUT TO COPY combinations 🔥🔥🔥\\n")
              cat("🔥 length(combinations) =", length(combinations), "🔥\\n")
              cat("🔥 typeof(combinations) =", typeof(combinations), "🔥\\n")
              cat("🔥 class(combinations) =", class(combinations), "🔥\\n")
              if (length(combinations) > 0) {
                cat("🔥 combinations[[1]] =", paste(combinations[[1]], collapse=", "), "🔥\\n")
                cat("🔥 typeof(combinations[[1]]) =", typeof(combinations[[1]]), "🔥\\n")
                cat("🔥 length(combinations[[1]]) =", length(combinations[[1]]), "🔥\\n")
              }

              # CRITICAL FIX: Save combinations to a new variable to avoid any scope issues
              combinations_for_positioning <- combinations

              cat("🔥🔥🔥 AFTER COPY 🔥🔥🔥\\n")
              if (length(combinations_for_positioning) > 0) {
                cat("🔥 combinations_for_positioning[[1]] =", paste(combinations_for_positioning[[1]], collapse=", "), "🔥\\n")
              }

              # PRE-EXTRACT all group names using sapply (more robust than manual loop)
              num_combos <- length(combinations_for_positioning)
              group1_names <- sapply(combinations_for_positioning, function(x) as.character(x[1]))
              group2_names <- sapply(combinations_for_positioning, function(x) as.character(x[2]))

              cat("🔥🔥🔥 PRE-EXTRACTED GROUPS using sapply 🔥🔥🔥\\n")
              cat("🔥 group1_names:", paste(group1_names, collapse=", "), "🔥\\n")
              cat("🔥 group2_names:", paste(group2_names, collapse=", "), "🔥\\n")

              y_positions <- numeric(num_combos)
              base_heights <- numeric(num_combos)  # Store for debug

              # Store first comparison's extracted values for debug
              first_g1_vals <- NULL
              first_g2_vals <- NULL
              first_g1_name <- NULL
              first_g2_name <- NULL
              unique_groups_in_data <- unique(data[[group_col]])
              num_combinations_before_loop <- length(combinations_for_positioning)

              # Debug: extract first combo details
              first_combo_str <- if (num_combinations_before_loop > 0) {
                paste(combinations_for_positioning[[1]], collapse=" vs ")
              } else {
                "NO COMBOS"
              }
              first_combo_len <- if (num_combinations_before_loop > 0) {
                length(combinations_for_positioning[[1]])
              } else {
                0
              }
              first_combo_type <- if (num_combinations_before_loop > 0) {
                typeof(combinations_for_positioning[[1]])
              } else {
                "none"
              }

              cat("🔥🔥🔥 BEFORE LOOP: combinations has", num_combinations_before_loop, "items 🔥🔥🔥\\n")
              if (num_combinations_before_loop > 0) {
                cat("🔥 First combination:", paste(combinations_for_positioning[[1]], collapse=" vs "), "🔥\\n")
                cat("🔥 First combination[1]:", combinations_for_positioning[[1]][1], "🔥\\n")
                cat("🔥 First combination[2]:", combinations_for_positioning[[1]][2], "🔥\\n")
                cat("🔥 Structure of combinations_for_positioning[[1]]:\\n")
                print(str(combinations_for_positioning[[1]]))
              }

              for (idx in seq_along(combinations_for_positioning)) {
                # Use pre-extracted group names
                g1 <- as.character(group1_names[idx])
                g2 <- as.character(group2_names[idx])

                # DEBUG: Store first iteration values immediately
                if (idx == 1) {
                  first_g1_name <- g1
                  first_g2_name <- g2
                }

                # Create comparison keys
                combo_key1 <- paste(g1, g2, sep="-")
                combo_key2 <- paste(g2, g1, sep="-")

                # Verify they're not empty
                if (is.na(g1) || g1 == "" || is.na(g2) || g2 == "") {
                  next  # Skip this iteration
                }

                # Initialize baseHeight to 0 (will be overwritten)
                baseHeight <- 0

                # Check for custom position first
                custom_pos <- NULL
                for (user_key in names(user_positions)) {
                  if (user_key == combo_key1 || user_key == combo_key2) {
                    custom_pos <- user_positions[[user_key]]
                    break
                  }
                }

                if (!is.null(custom_pos) && is.numeric(custom_pos)) {
                  # User provided absolute Y value, use it directly
                  y_positions[idx] <- as.numeric(custom_pos)
                  baseHeight <- as.numeric(custom_pos) - unit_step  # Reverse-calculate for debug
                  base_heights[idx] <- baseHeight
                  cat("🔥 CUSTOM: Using custom position", custom_pos, "for", combo_key1, "🔥\\n")
                } else {
                  # ORIGINAL ALGORITHM: Calculate baseHeight = max(maxValue(g1), maxValue(g2))
                  # Need to account for error bars extending above the data

                  # CRITICAL FIX: Convert group column to character to match g1/g2 type
                  data_groups_char <- as.character(data[[group_col]])
                  g1_char <- as.character(g1)
                  g2_char <- as.character(g2)

                  cat("🔥🔥🔥 EXTRACTING VALUES FOR GROUPS:", g1_char, "and", g2_char, "🔥🔥🔥\\n")
                  cat("🔥 Unique groups in data:", paste(unique(data_groups_char), collapse=", "), "🔥\\n")

                  g1_values <- data[[value_col]][data_groups_char == g1_char]
                  g2_values <- data[[value_col]][data_groups_char == g2_char]

                  cat("🔥 g1_values (", g1, "):", paste(g1_values, collapse=", "), "🔥\\n")
                  cat("🔥 g2_values (", g2, "):", paste(g2_values, collapse=", "), "🔥\\n")
                  cat("🔥 length(g1_values) =", length(g1_values), ", length(g2_values) =", length(g2_values), "🔥\\n")

                  # Store first comparison data values for chart annotation
                  if (idx == 1) {
                    first_g1_vals <- g1_values
                    first_g2_vals <- g2_values
                    # Note: first_g1_name and first_g2_name already stored at loop start
                  }

                  # Store lengths for debug
                  if (idx == 1) {
                    first_g1_len <- length(g1_values)
                    first_g2_len <- length(g2_values)
                  }

                  # Check for empty values
                  if (length(g1_values) == 0 || length(g2_values) == 0) {
                    cat("🔥 WARNING: Empty values for", g1, "or", g2, "- using y_max as fallback 🔥\\n")
                    baseHeight <- y_max
                  } else {
                    cat("\\n🔥🔥🔥 NEW BASE HEIGHT CALCULATION CODE RUNNING! 🔥🔥🔥\\n")

                    # CRITICAL FIX: The chart displays mean ± SD (error bars), not raw data!
                    # So baseHeight must be the TOP of the error bar = mean + SD
                    g1_mean <- mean(g1_values, na.rm = TRUE)
                    g2_mean <- mean(g2_values, na.rm = TRUE)
                    g1_sd <- if(length(g1_values) > 1) sd(g1_values, na.rm = TRUE) else 0
                    g2_sd <- if(length(g2_values) > 1) sd(g2_values, na.rm = TRUE) else 0

                    # Top of error bar for each group
                    g1_error_top <- g1_mean + g1_sd
                    g2_error_top <- g2_mean + g2_sd

                    # Also get max data point for comparison
                    g1_max_data <- max(g1_values, na.rm = TRUE)
                    g2_max_data <- max(g2_values, na.rm = TRUE)

                    # Use the higher of: (error bar top) or (max data point)
                    g1_top <- max(g1_error_top, g1_max_data)
                    g2_top <- max(g2_error_top, g2_max_data)

                    # baseHeight = max of both groups' tops
                    baseHeight <- max(g1_top, g2_top)

                    cat("🔥 CALCULATED: baseHeight =", round(baseHeight, 2), "(from error_bar_tops) 🔥\\n")
                    cat("🔥", g1, ": values=", paste(g1_values, collapse=","), "🔥\\n")
                    cat("🔥", g1, ": mean=", round(g1_mean,1), ", sd=", round(g1_sd,1), ", max_data=", g1_max_data, ", error_top=", round(g1_error_top,1), ", final_top=", round(g1_top,1), "🔥\\n")
                    cat("🔥", g2, ": values=", paste(g2_values, collapse=","), "🔥\\n")
                    cat("🔥", g2, ": mean=", round(g2_mean,1), ", sd=", round(g2_sd,1), ", max_data=", g2_max_data, ", error_top=", round(g2_error_top,1), ", final_top=", round(g2_top,1), "🔥\\n")
                    cat("🔥 baseHeight (max of both groups) =", round(baseHeight,1), "🔥\\n")
                  }

                  # Store baseHeight for debug annotation
                  base_heights[idx] <- baseHeight
                  cat("🔥🔥🔥 STORED base_heights[", idx, "] =", base_heights[idx], "🔥🔥🔥\\n")

                  # Calculate bracket position (scale-aware)
                  if (y_scale %in% c("log", "log10", "log2")) {
                    # LOG SCALE: Work in log-transformed coordinate space
                    # When scale_y_log2() is applied, coordinates are in log2 space
                    # E.g., data value 12 appears at coordinate log2(12) ≈ 3.585

                    # Transform baseHeight to log coordinate space
                    if (y_scale == "log2") {
                      base_coord <- log2(baseHeight)
                    } else if (y_scale == "log10") {
                      base_coord <- log10(baseHeight)
                    } else {  # "log" = natural log
                      base_coord <- log(baseHeight)
                    }

                    # Add small increment in log space (0.15 for first, 0.15 for subsequent)
                    if (idx == 1) {
                      y_positions[idx] <- base_coord + 0.15
                    } else {
                      y_positions[idx] <- y_positions[idx-1] + 0.15
                    }

                    cat("🔥 LOG SCALE bracket: baseHeight=", baseHeight,
                        "→ log_coord=", base_coord,
                        "→ bracket_coord=", y_positions[idx], "🔥\\n")
                  } else {
                    # LINEAR SCALE: Use additive spacing in data space
                    if (idx == 1) {
                      y_positions[idx] <- baseHeight + (unit_step * 0.6)  # First bracket closer to data
                    } else {
                      y_positions[idx] <- max(baseHeight + unit_step, y_positions[idx-1] + unit_step)
                    }
                  }

                  cat("🔥 AUTO: ", combo_key1, " - baseHeight=", baseHeight,
                      " (", g1, "max=", ifelse(length(g1_values) > 0, max(g1_values, na.rm=TRUE), "NA"),
                      ", ", g2, "max=", ifelse(length(g2_values) > 0, max(g2_values, na.rm=TRUE), "NA"), ")",
                      ", y_pos=", y_positions[idx], "🔥\\n")
                }
              }

              # Create pairwise data with different symbol sizes for n.s. vs asterisks
              symbol_sizes <- sapply(real_p_symbols, function(s) {
                if (s == "n.s.") return(ggpubr_symbol_size * 0.7)  # Make n.s. 30% smaller
                else return(ggpubr_symbol_size)  # Keep asterisks at full size
              })

              pairwise_data <- data.frame(
                group1 = sapply(combinations, function(x) x[1]),
                group2 = sapply(combinations, function(x) x[2]),
                p.adj = real_p_values,
                p.signif = real_p_symbols,
                y.position = y_positions,
                symbol_size = symbol_sizes
              )

              # Store bracket data globally for educational R code export
              bracket_export_data <<- pairwise_data

              cat("🔥 Created pairwise_data with", nrow(pairwise_data), "rows 🔥\\n")
              cat("🔥 Y positions:", paste(y_positions, collapse=", "), "🔥\\n")
              cat("🔥 P-values:", paste(real_p_values, collapse=", "), "🔥\\n")
              cat("🔥 Symbols:", paste(real_p_symbols, collapse=", "), "🔥\\n")
              cat("🔥 unit_step value:", unit_step, "🔥\\n")
              cat("🔥 y_max value:", y_max, "🔥\\n")
              cat("🔥 y_range value:", y_range, "🔥\\n")

              # Check for invalid positions
              if (any(is.na(y_positions)) || any(is.infinite(y_positions))) {
                cat("🔥 ERROR: Invalid y_positions detected! 🔥\\n")
                cat("🔥 NA positions:", which(is.na(y_positions)), "🔥\\n")
                cat("🔥 Infinite positions:", which(is.infinite(y_positions)), "🔥\\n")
              }

              print(pairwise_data)

              # Filter based on comparison mode (use function parameter)
              cat("🔥 Comparison mode:", comparison_mode, "🔥\\n")

              if (comparison_mode == "significant") {
                # Keep only significant comparisons (p < 0.05)
                significant_indices <- which(pairwise_data$p.adj < 0.05)
                if (length(significant_indices) > 0) {
                  pairwise_data <- pairwise_data[significant_indices, ]
                  cat("🔥 Filtered to", nrow(pairwise_data), "significant comparisons 🔥\\n")

                  # RECALCULATE y positions to ensure equal spacing for filtered brackets
                  # (unless using custom positions)
                  if (!has_custom_pos) {
                    num_filtered <- nrow(pairwise_data)

                    # Recalculate unit_step based on ACTUAL number of brackets to be shown
                    unit_step_filtered <- calculate_unit_step(pairwise_data$p.signif, ggpubr_symbol_size, y_range, num_filtered, y_scale, y_min, y_max)
                    cat("🔥 Recalculated unit_step for", num_filtered, "filtered brackets: ", unit_step_filtered, "(was", unit_step, ") 🔥\\n")

                    new_y_positions <- numeric(num_filtered)

                    for (i in 1:num_filtered) {
                      # Get group names for this comparison
                      g1_filtered <- as.character(pairwise_data$group1[i])
                      g2_filtered <- as.character(pairwise_data$group2[i])

                      # Extract data values
                      data_groups_char <- as.character(data[[group_col]])
                      g1_values_f <- data[[value_col]][data_groups_char == g1_filtered]
                      g2_values_f <- data[[value_col]][data_groups_char == g2_filtered]

                      # Calculate baseHeight
                      if (length(g1_values_f) > 0 && length(g2_values_f) > 0) {
                        g1_mean_f <- mean(g1_values_f, na.rm = TRUE)
                        g2_mean_f <- mean(g2_values_f, na.rm = TRUE)
                        g1_sd_f <- if(length(g1_values_f) > 1) sd(g1_values_f, na.rm = TRUE) else 0
                        g2_sd_f <- if(length(g2_values_f) > 1) sd(g2_values_f, na.rm = TRUE) else 0
                        g1_top_f <- max(g1_mean_f + g1_sd_f, max(g1_values_f, na.rm = TRUE))
                        g2_top_f <- max(g2_mean_f + g2_sd_f, max(g2_values_f, na.rm = TRUE))
                        baseHeight_f <- max(g1_top_f, g2_top_f)
                      } else {
                        baseHeight_f <- y_max
                      }

                      # Calculate position with consistent spacing - scale-aware
                      if (y_scale %in% c("log", "log10", "log2")) {
                        # LOG SCALE: Work in log-transformed coordinate space
                        if (y_scale == "log2") {
                          base_coord_f <- log2(baseHeight_f)
                        } else if (y_scale == "log10") {
                          base_coord_f <- log10(baseHeight_f)
                        } else {
                          base_coord_f <- log(baseHeight_f)
                        }

                        if (i == 1) {
                          new_y_positions[i] <- base_coord_f + 0.15
                        } else {
                          new_y_positions[i] <- new_y_positions[i-1] + 0.15
                        }
                      } else {
                        # LINEAR SCALE: Use data space
                        if (i == 1) {
                          new_y_positions[i] <- baseHeight_f + (unit_step_filtered * 0.6)
                        } else {
                          new_y_positions[i] <- max(baseHeight_f + unit_step_filtered, new_y_positions[i-1] + unit_step_filtered)
                        }
                      }
                    }

                    # Update pairwise_data with recalculated positions
                    pairwise_data$y.position <- new_y_positions
                    cat("🔥 Recalculated y positions for equal spacing:", paste(round(new_y_positions, 1), collapse=", "), "🔥\\n")
                  }
                } else {
                  # No significant comparisons found
                  pairwise_data <- pairwise_data[0, ]  # Empty data frame
                  cat("🔥 No significant comparisons found 🔥\\n")
                }
              } else if (comparison_mode == "all") {
                cat("🔥 Showing all", nrow(pairwise_data), "comparisons 🔥\\n")
              } else if (comparison_mode == "custom") {
                # Custom selection - use actual selected comparisons
                cat("🔥 Custom comparisons JSON:", custom_comparisons, "🔥\\n")

                # Parse the custom comparisons JSON
                if (custom_comparisons != "[]" && custom_comparisons != "") {
                  # Convert JSON string to R list
                  library(jsonlite)
                  selected_pairs <- fromJSON(custom_comparisons)
                  cat("🔥 Parsed custom selections:", length(selected_pairs), "pairs 🔥\\n")
                  for (sp in selected_pairs) {
                    cat("🔥 Selected pair:", sp, "🔥\\n")
                  }

                  # Filter pairwise_data to only include selected comparisons
                  selected_indices <- c()
                  for (i in 1:nrow(pairwise_data)) {
                    # Check multiple possible formats
                    pair_str_vs <- paste(pairwise_data$group1[i], "vs", pairwise_data$group2[i])
                    reverse_pair_str_vs <- paste(pairwise_data$group2[i], "vs", pairwise_data$group1[i])
                    pair_str_dash <- paste(pairwise_data$group1[i], pairwise_data$group2[i], sep="-")
                    reverse_pair_str_dash <- paste(pairwise_data$group2[i], pairwise_data$group1[i], sep="-")

                    if (pair_str_vs %in% selected_pairs || reverse_pair_str_vs %in% selected_pairs ||
                        pair_str_dash %in% selected_pairs || reverse_pair_str_dash %in% selected_pairs) {
                      selected_indices <- c(selected_indices, i)
                      cat("🔥 Including comparison:", pair_str_vs, "(matched with selection) 🔥\\n")
                    }
                  }

                  if (length(selected_indices) > 0) {
                    pairwise_data <- pairwise_data[selected_indices, ]
                    cat("🔥 Custom selection: showing", nrow(pairwise_data), "selected comparisons - PRESERVING custom positions 🔥\\n")
                    # DON'T recalculate y positions - preserve custom user positions!
                  } else {
                    pairwise_data <- pairwise_data[0, ]  # Empty data frame
                    cat("🔥 Custom selection: no matching comparisons found 🔥\\n")
                  }
                } else {
                  pairwise_data <- pairwise_data[0, ]  # Empty data frame
                  cat("🔥 Custom selection: no comparisons selected 🔥\\n")
                }
              } else {
                cat("🔥 Unknown comparison mode, showing all", nrow(pairwise_data), "comparisons 🔥\\n")
              }

            }, error = function(e) {
              cat("🔥 Error in real statistical analysis:", e$message, "🔥\\n")
              # Fallback to combinations without statistical analysis
              combinations <- combn(actual_groups, 2, simplify = FALSE)
              pairwise_data <- data.frame(
                group1 = sapply(combinations, function(x) x[1]),
                group2 = sapply(combinations, function(x) x[2]),
                p.adj = rep(1.0, length(combinations)),
                p.signif = rep("n.s.", length(combinations)),
                y.position = y_max * (1.3 + (0:(length(combinations)-1)) * 0.15)
              )
            })
          } else {
            # Fallback if less than 2 groups
            pairwise_data <- data.frame(
              group1 = character(0),
              group2 = character(0),
              p.adj = numeric(0),
              p.signif = character(0),
              y.position = numeric(0)
            )
          }

          cat("🔥 Dynamic pairwise data created:", nrow(pairwise_data), "comparisons 🔥\\n")
          if (nrow(pairwise_data) > 0) {
            print(pairwise_data)

            # Debug: Show exact y.position values being sent to ggpubr
            cat("🔥 EXACT Y.POSITIONS being sent to ggpubr: 🔥\\n")
            for (k in 1:nrow(pairwise_data)) {
              cat("🔥", pairwise_data$group1[k], "vs", pairwise_data$group2[k], "→ Y =", pairwise_data$y.position[k], "🔥\\n")
            }

            # Debug: Check for duplicate y.positions
            unique_positions <- unique(pairwise_data$y.position)
            if (length(unique_positions) != nrow(pairwise_data)) {
              cat("🔥 WARNING: Duplicate y.positions found! ggpubr may auto-adjust these 🔥\\n")
              cat("🔥 Unique positions:", length(unique_positions), "vs rows:", nrow(pairwise_data), "🔥\\n")
              cat("🔥 ggpubr will likely move overlapping brackets to prevent collision 🔥\\n")
              print(pairwise_data$y.position)
            }
          }

          # Determine if user has set custom positions AND is in custom mode
          has_custom_positions <- FALSE
          cat("🔥 DEBUG: custom_positions string = '", custom_positions, "' 🔥\\n")
          cat("🔥 DEBUG: comparison_mode = '", comparison_mode, "' 🔥\\n")

          # Custom positions should ONLY be used when comparison_mode is "custom"
          if (comparison_mode == "custom" && custom_positions != "{}" && custom_positions != "") {
            has_custom_positions <- TRUE
            cat("🔥 User in CUSTOM mode with custom positions - will use exact positioning 🔥\\n")
          } else {
            cat("🔥 User in AUTOMATIC mode (", comparison_mode, ") - will use dynamic spacing 🔥\\n")
          }

          # NEW: Data-driven Y-axis limit scales with number of comparisons
          max_bracket_y <<- max(pairwise_data$y.position)  # Make global so other functions can access it
          num_bracket_pairs <- nrow(pairwise_data)

          # Scale expansion factor: more brackets = more space needed
          expansion_factor <- 0.15 + (0.03 * num_bracket_pairs)

          # Calculate y-axis limit - coordinate system aware
          if (y_scale %in% c("log", "log10", "log2")) {
            # LOG SCALE: Work in log coordinate space
            # Need MUCH MORE expansion for log scales to show asterisk symbols
            # Add fixed amount (1.5) plus proportional expansion
            if (y_scale == "log2") {
              y_range_log <- log2(y_max) - log2(y_min)
            } else if (y_scale == "log10") {
              y_range_log <- log10(y_max) - log10(y_min)
            } else {  # "log"
              y_range_log <- log(y_max) - log(y_min)
            }
            # For log scales, add 2.5 units in log space to show symbols with better spacing
            y_axis_upper_limit <- max_bracket_y + 2.5 + (expansion_factor * y_range_log)
            cat("🔥 LOG SCALE Y-axis limit: max_bracket=", max_bracket_y,
                " + 2.5 + ", expansion_factor, "*y_range_log(", y_range_log, ") = ",
                y_axis_upper_limit, "🔥\\n")
          } else {
            # LINEAR SCALE: Use data space
            y_axis_upper_limit <- max_bracket_y + expansion_factor * y_range
            cat("🔥 Y-axis limit: max_bracket=", max_bracket_y,
                " + ", expansion_factor, "*y_range(", y_range, ") = ",
                y_axis_upper_limit, "🔥\\n")
          }

          p <- p + expand_limits(y = y_axis_upper_limit)
          cat("🔥 Expansion factor scales with", num_bracket_pairs, "bracket pairs 🔥\\n")

          # Get parameter values from function parameters
          bracket_size <- ggpubr_symbol_size
          line_size <- ggpubr_line_size
          tip_length <- ggpubr_tip_length
          v_just <- ggpubr_vjust + 0.6  # Move symbols much closer to brackets (less negative = closer)

          cat("🔥 Adding ggpubr brackets with custom parameters 🔥\\n")
          cat("🔥 Parameters: size=", bracket_size, ", line=", line_size, ", tip=", tip_length, ", vjust=", v_just, "🔥\\n")

          # Always use step.increase = 0 since we provide manual y.position values
          step_increase_value <- 0
          cat("🔥 Using step.increase =", step_increase_value, "(manual y.position provided) 🔥\\n")

          # Split data by symbol type to apply different sizes
          asterisk_data <- pairwise_data[pairwise_data$p.signif != "n.s.", ]
          ns_data <- pairwise_data[pairwise_data$p.signif == "n.s.", ]

          # Add asterisk symbols with full size
          if (nrow(asterisk_data) > 0) {
            p <- p + stat_pvalue_manual(asterisk_data,
                                       label = "p.signif",
                                       size = bracket_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = v_just,
                                       hjust = 0.5,
                                       step.increase = step_increase_value,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
            cat("🔥 Added", nrow(asterisk_data), "asterisk symbols at full size 🔥\\n")
          }

          # Add n.s. symbols with smaller size and higher position
          if (nrow(ns_data) > 0) {
            smaller_size <- bracket_size * 0.7  # 30% smaller for n.s.
            ns_vjust <- v_just - 0.5  # Position n.s. higher on the bracket (more negative = higher)
            p <- p + stat_pvalue_manual(ns_data,
                                       label = "p.signif",
                                       size = smaller_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = ns_vjust,
                                       hjust = 0.5,
                                       step.increase = step_increase_value,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
            cat("🔥 Added", nrow(ns_data), "n.s. symbols at smaller size (", smaller_size, ") 🔥\\n")
          }

          # ggpubr statistical brackets added successfully
          cat("🔥 DIRECT GGPUBR COMPLETED 🔥\\n")

          # ADD OVERALL TEST SYMBOL if checkbox is enabled and test is significant
          # This should be added ABOVE the pairwise brackets
          if (show_main_symbol && length(actual_groups) > 2) {
            cat("🔥🔥🔥 CHECKING FOR OVERALL SYMBOL - show_main_symbol=", show_main_symbol, "🔥🔥🔥\\n")

            # Get omnibus test p-value (it was calculated earlier in this same tryCatch block at lines 2678-2755)
            # The omnibus_p_value variable should still be available
            omnibus_sig_symbol <- sato_get_significance_symbol(omnibus_p_value)
            cat("🔥 Omnibus p-value:", omnibus_p_value, ", symbol:", omnibus_sig_symbol, "🔥\\n")

            if (omnibus_sig_symbol != "ns") {
              cat("🔥🔥🔥 ADDING OVERALL SYMBOL:", omnibus_sig_symbol, "🔥🔥🔥\\n")

              # Calculate position - at the very top, above all brackets
              y_max_data <- max(data[[value_col]], na.rm = TRUE)

              # Position should be ABOVE the highest bracket - coordinate system aware
              if (y_scale %in% c("log", "log10", "log2")) {
                # LOG SCALE: Work in log coordinate space
                if (y_scale == "log2") {
                  y_range_log <- log2(y_max) - log2(y_min)
                } else if (y_scale == "log10") {
                  y_range_log <- log10(y_max) - log10(y_min)
                } else {
                  y_range_log <- log(y_max) - log(y_min)
                }
                overall_symbol_y <- max_bracket_y + (0.08 * y_range_log)
                cat("🔥 LOG SCALE overall symbol Y:", overall_symbol_y, "🔥\\n")
              } else {
                # LINEAR SCALE: Use data space
                overall_symbol_y <- max_bracket_y + (0.08 * y_range)
              }

              # Center x position
              n_groups <- length(actual_groups)
              if (n_groups == 2) {
                x_center <- 1.5
              } else if (n_groups == 3) {
                x_center <- 2
              } else if (n_groups == 4) {
                x_center <- 2.5
              } else {
                x_center <- (n_groups + 1) / 2
              }

              cat("🔥 Overall symbol position: x=", x_center, ", y=", overall_symbol_y, "🔥\\n")

              # Add the overall symbol annotation
              p <- p + annotate("text",
                               x = x_center,
                               y = overall_symbol_y,
                               label = omnibus_sig_symbol,
                               size = symbol_size,
                               hjust = 0.5,
                               vjust = 0,
                               color = "black",
                               fontface = "bold")

              # Expand y-axis to accommodate the overall symbol - coordinate system aware
              if (y_scale %in% c("log", "log10", "log2")) {
                # LOG SCALE: Work in log coordinate space
                # Add fixed amount (2.5) to ensure symbol is visible
                if (y_scale == "log2") {
                  y_range_log <- log2(y_max) - log2(y_min)
                } else if (y_scale == "log10") {
                  y_range_log <- log10(y_max) - log10(y_min)
                } else {
                  y_range_log <- log(y_max) - log(y_min)
                }
                y_axis_upper_limit_with_symbol <- overall_symbol_y + 2.5 + (0.1 * y_range_log)
              } else {
                # LINEAR SCALE: Use data space
                y_axis_upper_limit_with_symbol <- overall_symbol_y + (0.1 * y_range)
              }

              p <- p + expand_limits(y = y_axis_upper_limit_with_symbol)

              cat("🔥🔥🔥 OVERALL SYMBOL ADDED SUCCESSFULLY 🔥🔥🔥\\n")
              cat("🔥 Expanded y-axis to:", y_axis_upper_limit_with_symbol, "🔥\\n")
            } else {
              cat("🔥 Omnibus test not significant - no overall symbol 🔥\\n")
            }
          } else {
            if (!show_main_symbol) {
              cat("🔥 show_main_symbol is FALSE - skipping overall symbol 🔥\\n")
            }
            if (show_main_symbol && length(actual_groups) <= 2) {
              cat("🔥 Only", length(actual_groups), "groups - overall symbol not needed 🔥\\n")
              # Add warning to Excel status for user education
              if (exists("setStatus")) {
                warning_msg <- "Note: 'Show overall test symbol' is only for 3+ groups. For 2-group comparisons, the significance is already shown on the bracket. Please uncheck this option or use 3+ groups."
                tryCatch(setStatus(warning_msg), error = function(e) {})
              }
            }
          }

        }, error = function(e) {
          cat("🔥 GGPUBR ERROR:", e$message, "🔥\\n")
        })
      }

      tryCatch({
        cat("\\n🔥🔥🔥 CUSTOM STATISTICAL FUNCTION CALLED 🔥🔥🔥\\n")
        cat("🔥🔥🔥 THIS FUNCTION SHOULD ADD CUSTOM ANNOTATIONS 🔥🔥🔥\\n")
        cat("🔥🔥🔥 PARAMETERS: group_col=", group_col, ", value_col=", value_col, ", test_type=", test_type, ", symbol_size=", symbol_size, "🔥🔥🔥\\n")
        cat("🔥🔥🔥 UI OPTIONS: show_main_symbol=", show_main_symbol, ", show_pairwise=", show_pairwise, "🔥🔥🔥\\n")
        cat("🔥🔥🔥 CRITICAL: show_pairwise type =", typeof(show_pairwise), ", value =", show_pairwise, "🔥🔥🔥\\n")
        cat("🔥🔥🔥 DATA CHECK: nrow=", nrow(data), ", ncol=", ncol(data), ", column names=", paste(names(data), collapse=", "), "🔥🔥🔥\\n")
        
        # Show status in Excel interface
        if (exists("setStatus")) {
          tryCatch(setStatus("🔥 STATISTICAL FUNCTION CALLED"), error = function(e) {})
        }
        cat("Group column:", group_col, "\\n")
        cat("Value column:", value_col, "\\n")
        cat("Test type:", test_type, "\\n")
        cat("Data dimensions:", nrow(data), "x", ncol(data), "\\n")
        cat("Data column names:", paste(names(data), collapse=", "), "\\n")
        
        # Check if we have the required columns
        if (!group_col %in% names(data) || !value_col %in% names(data)) {
          cat("Required columns not found in data\\n")
          return(p)
        }
        
        # ENABLE GGPUBR FOR PROPER STATISTICAL BRACKETS
        use_ggpubr <- TRUE && show_pairwise  # Only use ggpubr if pairwise comparisons are requested
        cat("🔥🔥🔥 GGPUBR MODE:", ifelse(use_ggpubr, "ENABLED", "DISABLED"), "🔥🔥🔥\\n")
        
        # Get unique groups and convert to character
        groups <- as.character(unique(data[[group_col]]))
        cat("Groups found:", paste(groups, collapse=", "), "\\n")
        
        if (length(groups) < 2) {
          cat("Less than 2 groups found, no comparison possible\\n")
          return(p)
        }
        
        # Perform statistical test
        cat("🔥 ABOUT TO PERFORM STATISTICAL TEST 🔥\\n")
        stat_result <- sato_perform_statistical_test(data, group_col, value_col, test_type, "${varianceTest}", "${postHocTest}")
        cat("🔥 STATISTICAL TEST COMPLETED 🔥\\n")

        cat("🔥🔥🔥 STATISTICAL TEST RESULT 🔥🔥🔥\\n")
        cat("Test used:", stat_result$test_used, "\\n")
        cat("P-value:", stat_result$p_value, "\\n")
        cat("P-value (scientific):", format(stat_result$p_value, scientific=TRUE), "\\n")
        cat("P-value < 0.001?", stat_result$p_value < 0.001, "\\n")
        cat("P-value < 0.01?", stat_result$p_value < 0.01, "\\n")
        cat("P-value < 0.05?", stat_result$p_value < 0.05, "\\n")
        sig_symbol_calculated <- sato_get_significance_symbol(stat_result$p_value)
        cat("Calculated significance symbol:", sig_symbol_calculated, "\\n")
        cat("Will add annotation if not 'ns':", sig_symbol_calculated != "ns", "\\n")
        
        # Store results globally for Excel export
        cat("=== STORING STATISTICAL RESULTS ===\\n")
        cat("stat_result structure:\\n")
        print(str(stat_result))
        cat("Groups found:", paste(groups, collapse=", "), "\\n")
        cat("Test method:", stat_result$test_used, "\\n")
        cat("P-value:", stat_result$p_value, "\\n")
        cat("Significance symbol:", sato_get_significance_symbol(stat_result$p_value), "\\n")
        
        sato_last_stat_result <<- list(
          test_method = stat_result$test_used,
          p_value = stat_result$p_value,
          significance = sato_get_significance_symbol(stat_result$p_value),
          comparison = paste(groups, collapse=" vs "),
          n_groups = length(groups),
          group_column = group_col,
          value_column = value_col,
          normality_result = stat_result$normality_result,
          auto_selected = stat_result$auto_selected
        )
        
        cat("Stored sato_last_stat_result:\\n")
        print(str(sato_last_stat_result))
        cat("Statistical results stored globally for export\\n")
        
        # Add statistical annotation (y-axis expansion handled in theme function)
        cat("DEBUGGING STATISTICAL ANNOTATION SECTION\\n")
        cat("use_ggpubr value:", use_ggpubr, "\\n")
        cat("Number of groups:", length(groups), "\\n")
        cat("Statistical result p-value:", stat_result$p_value, "\\n")
        
        if (use_ggpubr) {
          cat("Using ggpubr for statistical annotations\\n")
          if (length(groups) == 2) {
            cat("Adding pairwise comparison for 2 groups\\n")
            # For two groups, use pairwise comparison
            # Use our calculated significance symbol instead of ggpubr's automatic calculation
            our_significance <- sato_get_significance_symbol(stat_result$p_value)
            
            cat("DEBUG: Chart annotation significance calculation:\\n")
            cat("  P-value for chart:", stat_result$p_value, "\\n")
            cat("  P-value < 0.001?", stat_result$p_value < 0.001, "\\n")
            cat("  P-value < 0.01?", stat_result$p_value < 0.01, "\\n")
            cat("  P-value < 0.05?", stat_result$p_value < 0.05, "\\n")
            cat("  Our calculated significance:", our_significance, "\\n")
            
            if (our_significance != "ns") {
              cat("  CREATING CUSTOM ANNOTATION\\n")
              # Use direct annotation instead of stat_compare_means to ensure our significance is displayed
              y_max <- max(data[[value_col]], na.rm = TRUE)
              y_min <- min(data[[value_col]], na.rm = TRUE)

              # Calculate y_pos (multiplicative for all scales - log scales display as powers)
              y_pos <- y_max * 1.15
              bracket_offset <- y_max * 0.02
              tick_offset <- y_max * 0.02

              x_pos <- 1.5  # Between the two groups (assuming positions 1 and 2)

              cat("  y_max:", y_max, "y_pos:", y_pos, "x_pos:", x_pos, "\\n")

              # Add bracket lines
              p <- p +
                annotate("segment", x = 1, xend = 2, y = y_pos - bracket_offset, yend = y_pos - bracket_offset, color = "black") +
                annotate("segment", x = 1, xend = 1, y = y_pos - bracket_offset, yend = y_pos - bracket_offset - tick_offset, color = "black") +
                annotate("segment", x = 2, xend = 2, y = y_pos - bracket_offset, yend = y_pos - bracket_offset - tick_offset, color = "black") +
                annotate("text", x = x_pos, y = y_pos, label = our_significance,
                        size = 8, hjust = 0.5, vjust = 0.5, color = "black")

              cat("  SUCCESSFULLY ADDED DIRECT SIGNIFICANCE ANNOTATION:", our_significance, "at position y =", y_pos, "\\n")
            } else {
              cat("  No significance annotation added (ns)\\n")
            }
          } else {
            cat("Adding overall test result for", length(groups), "groups\\n")

            # Add overall symbol if checkbox is enabled and test is significant
            if (show_main_symbol) {
              sig_symbol <- sato_get_significance_symbol(stat_result$p_value)
              cat("🔥🔥🔥 GGPUBR PATH - Overall symbol check: sig_symbol='", sig_symbol, "', show_main_symbol=", show_main_symbol, "🔥🔥🔥\\n")

              if (sig_symbol != "ns") {
                cat("🔥🔥🔥 ADDING OVERALL SYMBOL IN GGPUBR PATH 🔥🔥🔥\\n")
                # Get y-axis max for positioning
                y_max <- max(data[[value_col]], na.rm = TRUE)
                y_min <- min(data[[value_col]], na.rm = TRUE)

                # Calculate y_pos (multiplicative for all scales)
                y_pos <- y_max * 1.15

                # Calculate center position for multiple groups
                n_groups <- length(groups)
                if (n_groups == 2) {
                  x_pos <- 1.5
                } else if (n_groups == 3) {
                  x_pos <- 2
                } else if (n_groups == 4) {
                  x_pos <- 2.5
                } else {
                  x_pos <- (n_groups + 1) / 2
                }

                cat("  Adding overall symbol '", sig_symbol, "' at x=", x_pos, ", y=", y_pos, "\\n")

                p <- p + annotate("text",
                                 x = x_pos,
                                 y = y_pos,
                                 label = sig_symbol,
                                 size = symbol_size,
                                 hjust = 0.5,
                                 vjust = -0.5,
                                 color = "black",
                                 fontface = "bold")

                cat("🔥🔥🔥 OVERALL SYMBOL ADDED SUCCESSFULLY 🔥🔥🔥\\n")
              } else {
                cat("🔥🔥🔥 Overall test not significant - no symbol added 🔥🔥🔥\\n")
              }
            } else {
              cat("🔥🔥🔥 show_main_symbol is FALSE - skipping overall symbol 🔥🔥🔥\\n")
            }
          }
        } else {
          cat("USING BASIC TEXT ANNOTATION FOR STATISTICAL RESULTS\\n")
          # Basic fallback: add only significance symbol
          sig_symbol <- sato_get_significance_symbol(stat_result$p_value)

          cat("FALLBACK ANNOTATION DEBUG:\\n")
          cat("  P-value:", stat_result$p_value, "\\n")
          cat("  P-value < 0.001?", stat_result$p_value < 0.001, "\\n")
          cat("  P-value < 0.01?", stat_result$p_value < 0.01, "\\n")
          cat("  P-value < 0.05?", stat_result$p_value < 0.05, "\\n")
          cat("  Calculated significance symbol:", sig_symbol, "\\n")

          # Only show symbol if significant AND if main symbol is enabled
          cat("🔥🔥🔥 CHECKING IF SHOULD ADD ANNOTATION: sig_symbol='", sig_symbol, "', not ns?", sig_symbol != "ns", ", show_main_symbol=", show_main_symbol, "🔥🔥🔥\\n")
          if (sig_symbol != "ns" && show_main_symbol) {
            cat("🔥🔥🔥 ENTERING ANNOTATION BLOCK - WILL ADD MAIN SYMBOL 🔥🔥🔥\\n")
            # Get y-axis max for positioning
            y_max <- max(data[[value_col]], na.rm = TRUE)
            y_min <- min(data[[value_col]], na.rm = TRUE)

            # Calculate y_pos (ggplot2 will apply log transformation if needed)
            y_pos <- y_max * 1.15

            # Determine x position based on group structure
            # If groups are categorical names, position based on their order
            unique_groups <- unique(data[[group_col]])
            n_groups <- length(unique_groups)

            cat("  Groups for positioning:", paste(unique_groups, collapse=", "), "\\n")
            cat("  Number of groups:", n_groups, "\\n")

            # For categorical data, center position is the middle of the discrete axis
            if (n_groups == 2) {
              x_pos <- 1.5  # Between positions 1 and 2
            } else if (n_groups == 3) {
              x_pos <- 2    # Center of positions 1, 2, 3
            } else if (n_groups == 4) {
              x_pos <- 2.5  # Center of positions 1, 2, 3, 4
            } else {
              x_pos <- (n_groups + 1) / 2  # General formula
            }

            cat("  Adding basic annotation at x_pos:", x_pos, "y_pos:", y_pos, "\\n")

            p <- p + annotate("text",
                             x = x_pos,
                             y = y_pos,
                             label = sig_symbol,
                             size = 8,  # Increased size for visibility
                             hjust = 0.5)

            cat("  SUCCESSFULLY ADDED BASIC ANNOTATION:", sig_symbol, "at x =", x_pos, ", y =", y_pos, "\\n")

            # Expand y-axis to make annotations visible
            # Calculate the maximum y position we'll need for all annotations
            max_annotation_y <- y_pos
            if (n_groups > 2) {
              # Account for pairwise comparison lines which will be higher
              max_annotation_y <- y_max * (1.15 + 0.04 * (n_groups - 1))
            }
            y_axis_limit <- max_annotation_y * 1.1  # Add 10% more space above highest annotation

            # Check if we have ggpubr brackets - use data-driven limit
            if (exists("max_bracket_y") && !is.null(max_bracket_y) && max_bracket_y > y_axis_limit) {
              data_y_min <- min(data[[value_col]], na.rm = TRUE)
              data_y_max <- max(data[[value_col]], na.rm = TRUE)
              data_y_range <- data_y_max - data_y_min

              # Calculate y-axis limit - same for all scales
              # ggplot2 handles log transformation internally
              y_axis_limit <- max_bracket_y + 0.15 * data_y_range
              cat("  Adjusting y-axis for ggpubr brackets: max_bracket=", max_bracket_y, " + 0.15*range = ", y_axis_limit, "\\n")
            }
            cat("  Expanding y-axis to:", y_axis_limit, "to accommodate annotations\\n")
            p <- p + ylim(0, y_axis_limit)

            # Add pairwise comparison lines if we have more than 2 groups AND pairwise is enabled
            if (n_groups > 2 && show_pairwise) {
              cat("  Adding pairwise comparison lines for", n_groups, "groups\\n")

              # Perform pairwise statistical tests
              for (i in 1:(n_groups-1)) {
                for (j in (i+1):n_groups) {
                  group1 <- unique_groups[i]
                  group2 <- unique_groups[j]

                  # Get data for these two groups
                  group1_data <- data[data[[group_col]] == group1, value_col]
                  group2_data <- data[data[[group_col]] == group2, value_col]

                  # Perform t-test or Mann-Whitney U test
                  tryCatch({
                    # Simple t-test for pairwise comparison
                    pairwise_test <- t.test(group1_data, group2_data)
                    pairwise_p <- pairwise_test$p.value
                    pairwise_sig <- sato_get_significance_symbol(pairwise_p)

                    cat("    Pairwise test", group1, "vs", group2, "p =", pairwise_p, "sig =", pairwise_sig, "\\n")

                    # Only add line if significant
                    if (pairwise_sig != "ns") {
                      # Calculate positions for comparison line (scale-aware)
                      x1 <- i
                      x2 <- j

                      # Calculate line_y position (multiplicative for all scales)
                      line_y <- y_max * (1.05 + 0.04 * (j - i - 1))
                      tick_offset <- y_max * 0.01
                      text_offset <- y_max * 0.02

                      cat("    Adding comparison line from", x1, "to", x2, "at y =", line_y, "\\n")

                      # Add comparison line
                      p <- p +
                        annotate("segment", x = x1, xend = x2, y = line_y, yend = line_y, color = "black", linewidth = 0.5) +
                        annotate("segment", x = x1, xend = x1, y = line_y - tick_offset, yend = line_y + tick_offset, color = "black", linewidth = 0.5) +
                        annotate("segment", x = x2, xend = x2, y = line_y - tick_offset, yend = line_y + tick_offset, color = "black", linewidth = 0.5) +
                        annotate("text", x = (x1 + x2) / 2, y = line_y + text_offset, label = pairwise_sig, size = 6, hjust = 0.5)
                    }
                  }, error = function(e) {
                    cat("    Error in pairwise test for", group1, "vs", group2, ":", e$message, "\\n")
                  })
                }
              }
            }

          } else {
            cat("No significant difference found or main symbol disabled, no main symbol added\\n")
          }

          # Handle pairwise comparisons even if main symbol is not shown
          if (!show_main_symbol && show_pairwise && n_groups > 1) {
            cat("🔥🔥🔥 MAIN SYMBOL DISABLED BUT PAIRWISE ENABLED - ADDING PAIRWISE ONLY 🔥🔥🔥\\n")

            # Get y-axis max for positioning
            y_max <- max(data[[value_col]], na.rm = TRUE)

            # Calculate y-axis expansion for pairwise lines only
            max_annotation_y <- y_max * (1.1 + 0.04 * (n_groups - 1))
            y_axis_limit <- max_annotation_y * 1.1

            # Check if we have ggpubr brackets - use data-driven limit
            if (exists("max_bracket_y") && !is.null(max_bracket_y) && max_bracket_y > y_axis_limit) {
              data_y_min <- min(data[[value_col]], na.rm = TRUE)
              data_y_max <- max(data[[value_col]], na.rm = TRUE)
              data_y_range <- data_y_max - data_y_min

              # Calculate y-axis limit - same for all scales
              # ggplot2 handles log transformation internally
              y_axis_limit <- max_bracket_y + 0.15 * data_y_range
              cat("  Adjusting y-axis for ggpubr brackets: max_bracket=", max_bracket_y, " + 0.15*range = ", y_axis_limit, "\\n")
            }
            cat("  Expanding y-axis to:", y_axis_limit, "for pairwise comparisons only\\n")
            p <- p + ylim(0, y_axis_limit)

            # Add pairwise comparison lines
            if (n_groups > 1) {
              cat("  Adding pairwise comparison lines for", n_groups, "groups (no main symbol)\\n")

              # Perform pairwise statistical tests
              for (i in 1:(n_groups-1)) {
                for (j in (i+1):n_groups) {
                  group1 <- unique_groups[i]
                  group2 <- unique_groups[j]

                  # Get data for these two groups
                  group1_data <- data[data[[group_col]] == group1, value_col]
                  group2_data <- data[data[[group_col]] == group2, value_col]

                  # Perform t-test or Mann-Whitney U test
                  tryCatch({
                    # Simple t-test for pairwise comparison
                    pairwise_test <- t.test(group1_data, group2_data)
                    pairwise_p <- pairwise_test$p.value
                    pairwise_sig <- sato_get_significance_symbol(pairwise_p)

                    cat("    Pairwise test", group1, "vs", group2, "p =", pairwise_p, "sig =", pairwise_sig, "\\n")

                    # Only add line if significant
                    if (pairwise_sig != "ns") {
                      # Calculate positions for comparison line (scale-aware)
                      x1 <- i
                      x2 <- j

                      # Calculate line_y position (multiplicative for all scales)
                      line_y <- y_max * (1.05 + 0.04 * (j - i - 1))
                      tick_offset <- y_max * 0.01
                      text_offset <- y_max * 0.02

                      cat("    Adding comparison line from", x1, "to", x2, "at y =", line_y, "\\n")

                      # Add comparison line
                      p <- p +
                        annotate("segment", x = x1, xend = x2, y = line_y, yend = line_y, color = "black", linewidth = 0.5) +
                        annotate("segment", x = x1, xend = x1, y = line_y - tick_offset, yend = line_y + tick_offset, color = "black", linewidth = 0.5) +
                        annotate("segment", x = x2, xend = x2, y = line_y - tick_offset, yend = line_y + tick_offset, color = "black", linewidth = 0.5) +
                        annotate("text", x = (x1 + x2) / 2, y = line_y + text_offset, label = pairwise_sig, size = 6, hjust = 0.5)
                    }
                  }, error = function(e) {
                    cat("    Error in pairwise test for", group1, "vs", group2, ":", e$message, "\\n")
                  })
                }
              }
            }
          }
        }

        cat("Statistical annotations added successfully\\n")
        return(p)
      }, error = function(e) {
        cat("Error adding statistical annotations:", e$message, "\\n")
        cat("Continuing without statistical annotations\\n")
        return(p)
      })
    }

    # Test if the function was defined correctly
    cat("🔥🔥🔥 TESTING IF sato_add_statistics_to_plot FUNCTION EXISTS:", exists("sato_add_statistics_to_plot"), "🔥🔥🔥\\n")

    # Data transformation function for negative log scales
    sato_transform_data <- function(dat, x_col=NULL, y_col=NULL, x_scale="linear", y_scale="linear") {
      # Always convert Y column to numeric (value column)
      if (!is.null(y_col)) {
        dat[[y_col]] <- as.numeric(dat[[y_col]])
      }

      # Apply scale transformations
      if (!is.null(x_col) && x_scale == "-log10") {
        dat[[x_col]] <- -log10(as.numeric(dat[[x_col]]))
      } else if (!is.null(x_col) && x_scale == "-log2") {
        dat[[x_col]] <- -log2(as.numeric(dat[[x_col]]))
      }

      if (!is.null(y_col) && y_scale == "-log10") {
        dat[[y_col]] <- -log10(dat[[y_col]])
      } else if (!is.null(y_col) && y_scale == "-log2") {
        dat[[y_col]] <- -log2(dat[[y_col]])
      }

      return(dat)
    }

    # Master theme and scaling function - handles all advanced features
    sato_apply_theme <- function(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                                title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                                title_text, x_text, y_text,
                                show_title, show_x_label, show_y_label,
                                x_scale, y_scale,
                                theme_name="minimal",
                                x_axis_rotation=0, y_axis_rotation=0,
                                x_axis_hjust=0.5, x_axis_vjust=0.5,
                                y_axis_hjust=0.5, y_axis_vjust=0.5,
                                add_statistics=FALSE, statistical_test="auto") {

      # Apply base theme with comprehensive styling
      p <- p + sato_theme(theme=theme_name, base_size=11, family=target_font,
                         title_face=title_weight, axis_title_face=axis_title_weight, axis_text_face=axis_text_weight,
                         title_size=title_size, x_axis_title_size=x_axis_title_size, y_axis_title_size=y_axis_title_size,
                         x_axis_text_size=x_axis_text_size, y_axis_text_size=y_axis_text_size, legend_text_size=legend_text_size) +
        theme(
          # Basic text elements
          text = element_text(family = target_font, size = 11),

          # Plot-related text
          plot.title = element_text(
            family = target_font,
            size = title_size,
            hjust = 0.5,
            face = title_weight
          ),
          plot.subtitle = element_text(family = target_font),
          plot.caption = element_text(family = target_font),
          plot.tag = element_text(family = target_font),

          # Axis-related text
          axis.title = element_text(family = target_font, face = axis_title_weight),
          axis.title.x = element_text(
            family = target_font,
            size = x_axis_title_size,
            face = axis_title_weight
          ),
          axis.title.y = element_text(
            family = target_font,
            size = y_axis_title_size,
            face = axis_title_weight
          ),
          axis.text = element_text(family = target_font, color = "black", face = axis_text_weight),
          axis.text.x = element_text(
            family = target_font,
            size = x_axis_text_size,
            color = "black",
            face = axis_text_weight,
            angle = x_axis_rotation,
            hjust = x_axis_hjust,
            vjust = x_axis_vjust
          ),
          axis.text.y = element_text(
            family = target_font,
            size = y_axis_text_size,
            color = "black",
            face = axis_text_weight,
            angle = y_axis_rotation,
            hjust = y_axis_hjust,
            vjust = y_axis_vjust
          ),
          
          # Legend-related text
          legend.text = element_text(family = target_font, size = legend_text_size),
          legend.title = element_text(family = target_font, size = legend_text_size),
          
          # Facet/strip text
          strip.text = element_text(family = target_font),
          strip.text.x = element_text(family = target_font),
          strip.text.y = element_text(family = target_font),
          
          # Background settings
          plot.background = element_rect(fill = "white", color = NA),
          panel.background = element_rect(fill = "white", color = NA)
        )
      
      # Apply labels (conditional show/hide)
      if (show_title) {
        p <- p + ggtitle(title_text)
      } else {
        p <- p + ggtitle("")
      }
      if (show_x_label) {
        p <- p + xlab(x_text)
      } else {
        p <- p + xlab("")
      }
      if (show_y_label) {
        p <- p + ylab(y_text)
      } else {
        p <- p + ylab("")
      }
      
      # Apply axis scales
      if (x_scale == "log10") {
        p <- p + scale_x_log10()
      } else if (x_scale == "log2") {
        p <- p + scale_x_continuous(trans = "log2")
      } else if (x_scale == "log") {
        p <- p + scale_x_continuous(trans = "log")
      }

      # Y-axis scale with statistical analysis expansion if needed
      stat_expansion <- if (add_statistics) expansion(mult = c(0.05, 0.15)) else expansion(mult = c(0.02, 0.02))
      
      if (y_scale == "log10") {
        p <- p + scale_y_log10(expand = stat_expansion)
      } else if (y_scale == "log2") {
        p <- p + scale_y_continuous(trans = "log2", expand = stat_expansion)
      } else if (y_scale == "log") {
        p <- p + scale_y_continuous(trans = "log", expand = stat_expansion)
      } else if (add_statistics) {
        # Add expansion for linear scale when statistics are enabled
        p <- p + scale_y_continuous(expand = stat_expansion)
      }
      
      return(p)
    }

    # Updated chart functions using common theme
    sato_hist <- function(dat, x_col=1, bins=20, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9,
                         target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                         title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                         title_text="Histogram", x_text="Value", y_text="Frequency",
                         show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                         x_scale="linear", y_scale="linear",
                         theme_name="minimal",
                         x_axis_rotation=0, y_axis_rotation=0,
                         x_axis_hjust=0.5, x_axis_vjust=0.5,
                         y_axis_hjust=0.5, y_axis_vjust=0.5,
                         add_statistics=FALSE, statistical_test="auto") {
      
      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, NULL, x_scale, y_scale)
      
      p <- ggplot(dat, aes(x=dat[[x_col]])) +
           geom_histogram(bins=bins, fill=fill, color=color, linewidth=linewidth, alpha=alpha, boundary=0)
      
      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    sato_box <- function(dat, x_col=1, y_col=NULL, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9, width=0.7,
                        target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                        title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                        title_text="Box plot", x_text="", y_text="Value",
                        show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                        x_scale="linear", y_scale="linear",
                        theme_name="minimal",
                        x_axis_rotation=0, y_axis_rotation=0,
                        x_axis_hjust=0.5, x_axis_vjust=0.5,
                        y_axis_hjust=0.5, y_axis_vjust=0.5,
                        add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      if (is.null(y_col) || ncol(dat) == 1) {
        # Single column: box plot of x_col values
        p <- ggplot(dat, aes(x="", y=dat[[x_col]])) +
             geom_boxplot(fill=fill, color=color, linewidth=linewidth, alpha=alpha, width=width,
                         outlier.shape=16, outlier.alpha=0.5)
      } else {
        # Two columns: grouped box plot
        plot_data <- data.frame(
          group = factor(dat[[x_col]]),
          value = as.numeric(dat[[y_col]])
        )
        plot_data <- plot_data[complete.cases(plot_data), ]
        
        p <- ggplot(plot_data, aes(x=group, y=value)) +
             geom_boxplot(fill=fill, color=color, linewidth=linewidth, alpha=alpha, width=width,
                         outlier.shape=16, outlier.alpha=0.5)

        # Add statistical analysis if requested and we have multiple groups
        if (add_statistics && length(unique(plot_data$group)) > 1) {
          p <- sato_add_statistics_to_plot(p, plot_data, "group", "value", statistical_test, sato_symbol_size, posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale=y_scale, stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
        }
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    sato_bar <- function(dat, x_col=1, y_col=NULL, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9,
                        width=0.4,
                        target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                        title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                        title_text="Bar plot", x_text="Category", y_text="Count",
                        show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                        x_scale="linear", y_scale="linear",
                        theme_name="minimal",
                        x_axis_rotation=0, y_axis_rotation=0,
                        x_axis_hjust=0.5, x_axis_vjust=0.5,
                        y_axis_hjust=0.5, y_axis_vjust=0.5,
                        add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      if (is.null(y_col) || ncol(dat) == 1) {
        # Single column: frequency bar chart
        p <- ggplot(dat, aes(x=factor(dat[[x_col]]))) +
             geom_bar(fill=fill, color=color, linewidth=linewidth, alpha=alpha, width=width)
      } else {
        # Two columns: height bar chart
        p <- ggplot(dat, aes(x=factor(dat[[x_col]]), y=dat[[y_col]])) +
             geom_col(fill=fill, color=color, linewidth=linewidth, alpha=alpha, width=width)
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    sato_dot <- function(dat, x_col=1, y_col=NULL, fill="#4C78A8", color="#1F2937", size=2, alpha=0.9, shape=16,
                        target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                        title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                        title_text="Dot plot", x_text="X", y_text="Y",
                        show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                        x_scale="linear", y_scale="linear",
                        theme_name="minimal",
                        x_axis_rotation=0, y_axis_rotation=0,
                        x_axis_hjust=0.5, x_axis_vjust=0.5,
                        y_axis_hjust=0.5, y_axis_vjust=0.5,
                        add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      if (is.null(y_col) || ncol(dat) == 1) {
        # Single column: dot plot with row indices
        p <- ggplot(dat, aes(x=1:nrow(dat), y=dat[[x_col]])) +
             geom_point(fill=fill, color=color, size=size, alpha=alpha, shape=shape)
      } else {
        # Two columns: x-y scatter plot
        p <- ggplot(dat, aes(x=dat[[x_col]], y=dat[[y_col]])) +
             geom_point(fill=fill, color=color, size=size, alpha=alpha, shape=shape)
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    sato_line <- function(dat, x_col=1, y_col=NULL, color="#1F2937", linewidth=1, alpha=0.9,
                         target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                         title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                         title_text="Line plot", x_text="X", y_text="Y",
                         show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                         x_scale="linear", y_scale="linear",
                         theme_name="minimal",
                         x_axis_rotation=0, y_axis_rotation=0,
                         x_axis_hjust=0.5, x_axis_vjust=0.5,
                         y_axis_hjust=0.5, y_axis_vjust=0.5,
                         add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      if (is.null(y_col) || ncol(dat) == 1) {
        # Single column: line plot with row indices
        df <- data.frame(x=1:nrow(dat), y=dat[[x_col]])
        p <- ggplot(df, aes(x=x, y=y)) +
             geom_line(color=color, linewidth=linewidth, alpha=alpha)
      } else {
        # Two columns: x-y line plot
        df <- data.frame(x=dat[[x_col]], y=dat[[y_col]])
        df <- df[order(df$x), ]  # Sort by x for proper line connection
        p <- ggplot(df, aes(x=x, y=y)) +
             geom_line(color=color, linewidth=linewidth, alpha=alpha)
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped line chart function - for data with Group/X/Y columns
    sato_line_grouped <- function(dat, group_col=1, x_col=2, y_col=3,
                                  line_colors=c("#4C78A8", "#E15759", "#76B7B2", "#F28E2B"),
                                  linewidth=1.5, alpha=0.9,
                                  group_name="Group", x_name="X", value_name="Value",
                                  target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                  title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                  title_text="Grouped Line Plot", x_text="X", y_text="Value",
                                  show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                  x_scale="linear", y_scale="linear",
                                  theme_name="minimal",
                                  x_axis_rotation=0, y_axis_rotation=0,
                                  x_axis_hjust=0.5, x_axis_vjust=0.5,
                                  y_axis_hjust=0.5, y_axis_vjust=0.5,
                                  add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Check data structure
      if (ncol(dat) < 3) {
        stop("Grouped line chart requires at least 3 columns: Group, X, Value")
      }

      cat("DEBUG - Grouped line plot - Column names:", paste(names(dat), collapse=", "), "\\n")

      # Get actual column names (fallback to parameter names if V1, V2, V3)
      col_names <- names(dat)
      actual_group_name <- if(!is.null(col_names) && length(col_names) >= group_col && !grepl("^V[0-9]+$", col_names[group_col])) {
        col_names[group_col]
      } else {
        group_name
      }

      actual_x_name <- if(!is.null(col_names) && length(col_names) >= x_col && !grepl("^V[0-9]+$", col_names[x_col])) {
        col_names[x_col]
      } else {
        x_name
      }

      actual_value_name <- if(!is.null(col_names) && length(col_names) >= y_col && !grepl("^V[0-9]+$", col_names[y_col])) {
        col_names[y_col]
      } else {
        value_name
      }

      cat("DEBUG - Final names - Group:", actual_group_name, "X:", actual_x_name, "Value:", actual_value_name, "\\n")

      # Prepare data
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        x = as.numeric(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )

      # Remove NA values
      plot_data <- plot_data[complete.cases(plot_data), ]

      if (nrow(plot_data) == 0) {
        stop("No valid data points after removing NAs")
      }

      # Sort by group and x for proper line connection
      plot_data <- plot_data[order(plot_data$group, plot_data$x), ]

      # Get unique groups and assign colors
      unique_groups <- unique(as.character(dat[[group_col]]))
      plot_data$group <- factor(plot_data$group, levels = unique_groups)

      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)

      if (length(line_colors) < n_groups) {
        # Extend colors if not enough
        line_colors <- rep(line_colors, length.out = n_groups)
      }

      cat("Plot data prepared:", nrow(plot_data), "rows\\n")
      cat("Groups:", paste(levels(plot_data$group), collapse=", "), "\\n")
      cat("Assigned colors:", paste(line_colors[1:n_groups], collapse=", "), "\\n")

      # Create grouped line plot
      p <- ggplot(plot_data, aes(x = x, y = value, color = group, group = group)) +
           geom_line(linewidth = linewidth, alpha = alpha) +
           geom_point(size = 2, alpha = alpha) +  # Add points for clarity
           scale_color_manual(values = setNames(line_colors[1:n_groups], group_levels)) +
           labs(color = actual_group_name)  # Legend title

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped line chart with error bars/ribbons - for PRE-CALCULATED mean + SD/SE
    # Data format: Group, X, Mean, Error (4 columns)
    sato_line_grouped_error <- function(dat, group_col=1, x_col=2, y_col=3, error_col=4,
                                        line_colors=c("#4C78A8", "#E15759", "#76B7B2", "#F28E2B"),
                                        linewidth=1.5, alpha=0.9, ribbon_alpha=0.2,
                                        error_type="sd", use_ribbon=TRUE,
                                        group_name="Group", x_name="X", value_name="Value",
                                        target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                        title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                        title_text="Grouped Line Plot with Error", x_text="X", y_text="Value",
                                        show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                        x_scale="linear", y_scale="linear",
                                        theme_name="minimal",
                                        x_axis_rotation=0, y_axis_rotation=0,
                                        x_axis_hjust=0.5, x_axis_vjust=0.5,
                                        y_axis_hjust=0.5, y_axis_vjust=0.5,
                                        add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Check data structure
      if (ncol(dat) < 4) {
        stop("This chart requires 4 columns: Group, X, Mean, Error. Use 'Grouped line plot with error (raw data)' for 3-column data.")
      }

      # Get actual column names
      col_names <- names(dat)
      actual_group_name <- if(!is.null(col_names) && length(col_names) >= group_col && !grepl("^V[0-9]+$", col_names[group_col])) {
        col_names[group_col]
      } else {
        group_name
      }

      actual_x_name <- if(!is.null(col_names) && length(col_names) >= x_col && !grepl("^V[0-9]+$", col_names[x_col])) {
        col_names[x_col]
      } else {
        x_name
      }

      actual_value_name <- if(!is.null(col_names) && length(col_names) >= y_col && !grepl("^V[0-9]+$", col_names[y_col])) {
        col_names[y_col]
      } else {
        value_name
      }

      # Format: Group, X, Mean, SD/SE (4 columns)
      unique_groups <- unique(as.character(dat[[group_col]]))

      plot_data <- data.frame(
        group = factor(dat[[group_col]], levels = unique_groups),
        x = as.numeric(dat[[x_col]]),
        mean_val = as.numeric(dat[[y_col]]),
        error_val = as.numeric(dat[[error_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]

      if (nrow(plot_data) == 0) {
        stop("No valid data points for plotting")
      }

      # Sort by group and x
      plot_data <- plot_data[order(plot_data$group, plot_data$x), ]

      cat("Using pre-calculated statistics - rows:", nrow(plot_data), "\\n")

      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)

      if (length(line_colors) < n_groups) {
        line_colors <- rep(line_colors, length.out = n_groups)
      }

      cat("Groups:", paste(levels(plot_data$group), collapse=", "), "\\n")
      cat("Colors:", paste(line_colors[1:n_groups], collapse=", "), "\\n")

      # Create grouped line plot with error
      p <- ggplot(plot_data, aes(x = x, y = mean_val, color = group, fill = group, group = group))

      if (use_ribbon) {
        # Use ribbon for error visualization (shaded area)
        p <- p +
             geom_ribbon(aes(ymin = mean_val - error_val, ymax = mean_val + error_val),
                        alpha = ribbon_alpha, color = NA) +
             geom_line(linewidth = linewidth, alpha = alpha) +
             geom_point(size = 2, alpha = alpha)
      } else {
        # Use error bars
        p <- p +
             geom_line(linewidth = linewidth, alpha = alpha) +
             geom_point(size = 2, alpha = alpha) +
             geom_errorbar(aes(ymin = mean_val - error_val, ymax = mean_val + error_val),
                          width = 0.1, linewidth = linewidth * 0.7, alpha = alpha)
      }

      p <- p +
           scale_color_manual(values = setNames(line_colors[1:n_groups], group_levels)) +
           scale_fill_manual(values = setNames(line_colors[1:n_groups], group_levels)) +
           labs(color = actual_group_name, fill = actual_group_name)

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped line chart with error - for RAW DATA with automatic calculation
    # Data format: Group, X, Value (3 columns with replicates)
    sato_line_grouped_error_raw <- function(dat, group_col=1, x_col=2, y_col=3,
                                            line_colors=c("#4C78A8", "#E15759", "#76B7B2", "#F28E2B"),
                                            linewidth=1.5, alpha=0.9, ribbon_alpha=0.2,
                                            error_type="sd", use_ribbon=TRUE,
                                            group_name="Group", x_name="X", value_name="Value",
                                            target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                            title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                            title_text="Grouped Line Plot with Error", x_text="X", y_text="Value",
                                            show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                            x_scale="linear", y_scale="linear",
                                            theme_name="minimal",
                                            x_axis_rotation=0, y_axis_rotation=0,
                                            x_axis_hjust=0.5, x_axis_vjust=0.5,
                                            y_axis_hjust=0.5, y_axis_vjust=0.5,
                                            add_statistics=FALSE, statistical_test="auto", variance_test="levene", stat_symbol_size=7,
                                            comparison_mode="all", custom_comparisons="[]", custom_positions="{}",
                                            stat_symbol_type="stars", custom_symbol_05="*", custom_symbol_01="**", custom_symbol_001="***", custom_symbol_ns="ns",
                                            vbracket_timepoint="", vbracket_position="topleft", vbracket_x=0.08, vbracket_y=0.92,
                                            vbracket_text_size=14, vbracket_sig_size=14, vbracket_margin=0.03, vbracket_line_width=0.5,
                                            vbracket_legend_line_length=NULL, vbracket_legend_line_width=NULL, vbracket_item_spacing=NULL,
                                            vbracket_bracket_layer_spacing=NULL,
                                            output_width=6, output_height=4) {

      # Helper function to generate significance symbol based on p-value and symbol type
      get_sig_symbol <- function(p_val) {
        if (stat_symbol_type == "pvalue") {
          return(sprintf("p=%.3f", p_val))
        } else if (stat_symbol_type == "custom") {
          if (p_val < 0.001) return(custom_symbol_001)
          else if (p_val < 0.01) return(custom_symbol_01)
          else if (p_val < 0.05) return(custom_symbol_05)
          else return(custom_symbol_ns)
        } else {
          # Default to stars
          if (p_val < 0.001) return("***")
          else if (p_val < 0.01) return("**")
          else if (p_val < 0.05) return("*")
          else return("ns")
        }
      }

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Check data structure
      if (ncol(dat) < 3) {
        stop("This chart requires 3 columns: Group, X, Value")
      }

      # Get actual column names
      col_names <- names(dat)
      actual_group_name <- if(!is.null(col_names) && length(col_names) >= group_col && !grepl("^V[0-9]+$", col_names[group_col])) {
        col_names[group_col]
      } else {
        group_name
      }

      actual_x_name <- if(!is.null(col_names) && length(col_names) >= x_col && !grepl("^V[0-9]+$", col_names[x_col])) {
        col_names[x_col]
      } else {
        x_name
      }

      actual_value_name <- if(!is.null(col_names) && length(col_names) >= y_col && !grepl("^V[0-9]+$", col_names[y_col])) {
        col_names[y_col]
      } else {
        value_name
      }

      # Debug: Print column info
      cat("DEBUG - Total columns:", ncol(dat), "\\n")
      cat("DEBUG - Column names:", paste(names(dat), collapse=", "), "\\n")
      cat("DEBUG - Using group_col:", group_col, "x_col:", x_col, "y_col:", y_col, "\\n")

      # Get unique groups BEFORE creating factors to preserve original order
      unique_groups <- unique(as.character(dat[[group_col]]))
      cat("DEBUG - Unique groups found:", paste(unique_groups, collapse=", "), "\\n")

      # Create raw data frame
      raw_data <- data.frame(
        group = as.character(dat[[group_col]]),
        x = dat[[x_col]],
        value = dat[[y_col]],
        stringsAsFactors = FALSE
      )

      cat("DEBUG - Raw data before cleaning - rows:", nrow(raw_data), "\\n")
      cat("DEBUG - Sample of raw data:\\n")
      print(head(raw_data, 10))

      # Convert to proper types
      raw_data$group <- factor(raw_data$group, levels = unique_groups)
      raw_data$x <- as.numeric(raw_data$x)
      raw_data$value <- as.numeric(raw_data$value)

      # Remove NA values
      raw_data_before <- raw_data
      raw_data <- raw_data[complete.cases(raw_data), ]

      if (nrow(raw_data) == 0) {
        # Provide detailed error message
        msg <- paste0(
          "ERROR: No valid data after removing NAs.\\n",
          "Data before cleaning: ", nrow(raw_data_before), " rows\\n",
          "Columns used: group_col=", group_col, ", x_col=", x_col, ", y_col=", y_col, "\\n",
          "NA counts - group:", sum(is.na(raw_data_before$group)),
          ", x:", sum(is.na(raw_data_before$x)),
          ", value:", sum(is.na(raw_data_before$value)), "\\n",
          "Sample group values:", paste(head(raw_data_before$group, 3), collapse=", "), "\\n",
          "Sample x values:", paste(head(raw_data_before$x, 3), collapse=", "), "\\n",
          "Sample value values:", paste(head(raw_data_before$value, 3), collapse=", ")
        )
        stop(msg)
      }

      cat("DEBUG - Raw data after cleaning - rows:", nrow(raw_data), "\\n")
      cat("DEBUG - Unique groups:", paste(levels(raw_data$group), collapse=", "), "\\n")

      # Calculate summary statistics using dplyr
      if (requireNamespace("dplyr", quietly = TRUE)) {
        plot_data <- raw_data %>%
          dplyr::group_by(group, x) %>%
          dplyr::summarise(
            mean_val = mean(value, na.rm = TRUE),
            sd_val = sd(value, na.rm = TRUE),
            se_val = sd_val / sqrt(dplyr::n()),
            n = dplyr::n(),
            .groups = "drop"
          )

        # Calculate error based on error_type
        if (error_type == "se") {
          plot_data$error_val <- plot_data$se_val
        } else if (error_type == "ci95") {
          # 95% CI = 1.96 * SE
          plot_data$error_val <- 1.96 * plot_data$se_val
        } else {
          # Default to SD
          plot_data$error_val <- plot_data$sd_val
        }

        # Replace NA errors with 0 (for single data points)
        plot_data$error_val[is.na(plot_data$error_val)] <- 0

        # Ensure factor levels are preserved
        plot_data$group <- factor(plot_data$group, levels = unique_groups)

        # Sort by group and x
        plot_data <- plot_data[order(plot_data$group, plot_data$x), ]

        cat("Calculated statistics - rows:", nrow(plot_data), "error_type:", error_type, "\\n")

      } else {
        stop("dplyr package required for automatic statistics calculation")
      }

      if (nrow(plot_data) == 0) {
        stop("No valid data points for plotting")
      }

      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)

      if (length(line_colors) < n_groups) {
        line_colors <- rep(line_colors, length.out = n_groups)
      }

      cat("Groups:", paste(group_levels, collapse=", "), "\\n")
      cat("Colors:", paste(line_colors[1:n_groups], collapse=", "), "\\n")

      # Create grouped line plot with error
      p <- ggplot(plot_data, aes(x = x, y = mean_val, color = group, fill = group, group = group))

      if (use_ribbon) {
        # Use ribbon for error visualization (shaded area)
        p <- p +
             geom_ribbon(aes(ymin = mean_val - error_val, ymax = mean_val + error_val),
                        alpha = ribbon_alpha, color = NA) +
             geom_line(linewidth = linewidth, alpha = alpha) +
             geom_point(size = 2, alpha = alpha)
      } else {
        # Use error bars
        p <- p +
             geom_line(linewidth = linewidth, alpha = alpha) +
             geom_point(size = 2, alpha = alpha) +
             geom_errorbar(aes(ymin = mean_val - error_val, ymax = mean_val + error_val),
                          width = 0.1, linewidth = linewidth * 0.7, alpha = alpha)
      }

      p <- p +
           scale_color_manual(values = setNames(line_colors[1:n_groups], group_levels)) +
           scale_fill_manual(values = setNames(line_colors[1:n_groups], group_levels)) +
           labs(color = actual_group_name, fill = actual_group_name)

      # Initialize global variable for statistical results
      line_plot_stat_results <<- ""

      # Add statistical comparisons for line plots
      if (add_statistics && n_groups >= 2) {
        cat("Adding statistical comparisons at each time point...\\n")

        # Parse custom comparisons if in custom mode
        selected_x_values <- NULL
        custom_y_positions <- list()

        if (comparison_mode == "custom") {
          # Parse custom comparisons: format is "Control-Treatment@1", "Control-Treatment@2", etc.
          custom_comps <- tryCatch(jsonlite::fromJSON(custom_comparisons), error = function(e) c())
          cat("Custom comparisons:", custom_comparisons, "\\n")

          if (length(custom_comps) > 0) {
            # Extract X values from custom comparisons
            selected_x_values <- c()
            for (comp_str in custom_comps) {
              # Parse format: "Group1-Group2@XValue"
              parts <- strsplit(comp_str, "@", fixed = TRUE)[[1]]
              if (length(parts) == 2) {
                x_val <- as.numeric(parts[2])
                if (!is.na(x_val)) {
                  selected_x_values <- c(selected_x_values, x_val)
                }
              }
            }
            cat("Selected X values from custom comparisons:", paste(selected_x_values, collapse=", "), "\\n")

            # Parse custom positions
            custom_pos_list <- tryCatch(jsonlite::fromJSON(custom_positions), error = function(e) list())
            if (length(custom_pos_list) > 0) {
              for (comp_str in names(custom_pos_list)) {
                custom_y_positions[[comp_str]] <- as.numeric(custom_pos_list[[comp_str]])
              }
            }
          }
        }

        # Determine which X values to test
        if (is.null(selected_x_values) || length(selected_x_values) == 0) {
          # Test all time points (default or "all" mode)
          unique_x <- sort(unique(plot_data$x))
        } else {
          # Test only selected time points
          unique_x <- sort(unique(selected_x_values))
        }

        # For each time point, perform statistical tests
        stat_results <- data.frame(x = numeric(), y_pos = numeric(), label = character(), stringsAsFactors = FALSE)

        # Store statistical results text for UI display
        stat_text_results <- c()

        for (x_val in unique_x) {
          # Get data for all groups at this X value
          data_at_x <- raw_data[raw_data$x == x_val, ]

          if (nrow(data_at_x) > 0) {
            # Count how many groups have data at this time point
            groups_with_data <- unique(data_at_x$group[!is.na(data_at_x$value)])
            n_groups_at_x <- length(groups_with_data)

            cat(sprintf("\\nTime point X=%.1f: %d groups with data\\n", x_val, n_groups_at_x))

            if (n_groups_at_x == 2) {
              # TWO GROUPS: Use t-test or Wilcoxon
              group1_data <- data_at_x[data_at_x$group == groups_with_data[1], "value"]
              group2_data <- data_at_x[data_at_x$group == groups_with_data[2], "value"]

              # Only test if both groups have data
              if (length(group1_data) > 0 && length(group2_data) > 0) {
                # Perform normality test for both groups
                normality_text <- c()
                is_group1_normal <- TRUE  # default assumption
                is_group2_normal <- TRUE
                both_normal <- TRUE

                if (length(group1_data) >= 3 && length(group1_data) <= 5000) {
                  shapiro_result1 <- tryCatch(
                    shapiro.test(group1_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result1)) {
                    is_group1_normal <- shapiro_result1$p.value >= 0.05
                    norm_status <- if (is_group1_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[1], shapiro_result1$p.value, norm_status))
                    cat(sprintf("    %s: Shapiro-Wilk p=%.4f (%s)\\n", groups_with_data[1], shapiro_result1$p.value, norm_status))
                  }
                }

                if (length(group2_data) >= 3 && length(group2_data) <= 5000) {
                  shapiro_result2 <- tryCatch(
                    shapiro.test(group2_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result2)) {
                    is_group2_normal <- shapiro_result2$p.value >= 0.05
                    norm_status <- if (is_group2_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[2], shapiro_result2$p.value, norm_status))
                    cat(sprintf("    %s: Shapiro-Wilk p=%.4f (%s)\\n", groups_with_data[2], shapiro_result2$p.value, norm_status))
                  }
                }

                both_normal <- is_group1_normal && is_group2_normal

                # Select test based on mode
                test_to_use <- statistical_test
                if (statistical_test == "auto") {
                  if (both_normal) {
                    test_to_use <- "t-test"
                    cat("  Auto-selected: t-test (both groups normal)\\n")
                  } else {
                    test_to_use <- "wilcoxon"
                    cat("  Auto-selected: Wilcoxon test (non-normal data detected)\\n")
                  }
                } else if (statistical_test == "parametric") {
                  test_to_use <- "t-test"
                  cat("  Manual mode: parametric t-test\\n")
                } else if (statistical_test == "nonparametric") {
                  test_to_use <- "wilcoxon"
                  cat("  Manual mode: non-parametric Wilcoxon test\\n")
                }

                # Perform variance test (only for parametric tests)
                variance_text <- ""
                equal_variances <- TRUE  # default
                if (test_to_use != "wilcoxon") {
                  if (variance_test == "levene") {
                    # Levene test for equality of variances
                    combined_data <- data.frame(
                      values = c(group1_data, group2_data),
                      group = factor(c(rep(groups_with_data[1], length(group1_data)),
                                      rep(groups_with_data[2], length(group2_data))))
                    )

                    # Calculate Levene test manually
                    group_means <- tapply(combined_data$values, combined_data$group, mean)
                    abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
                    levene_result <- tryCatch(
                      anova(lm(abs_deviations ~ combined_data$group)),
                      error = function(e) NULL
                    )

                    if (!is.null(levene_result)) {
                      levene_p <- levene_result$\`Pr(>F)\`[1]
                      equal_variances <- levene_p > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, Levene)", levene_p, variance_status)
                      cat(sprintf("    Levene test: p=%.4f (%s)\\n", levene_p, variance_status))
                    }
                  } else {
                    # F-test for equality of variances
                    var_test <- tryCatch(
                      var.test(group1_data, group2_data),
                      error = function(e) NULL
                    )

                    if (!is.null(var_test)) {
                      equal_variances <- var_test$p.value > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, F-test)", var_test$p.value, variance_status)
                      cat(sprintf("    F-test: p=%.4f (%s)\\n", var_test$p.value, variance_status))
                    }
                  }
                }

                # Perform statistical test based on auto-selection
                if (test_to_use == "wilcoxon") {
                  test_result <- tryCatch(
                    wilcox.test(group1_data, group2_data),
                    error = function(e) NULL
                  )
                  test_name <- "Wilcoxon"
                } else {
                  # Use appropriate t-test based on variance equality
                  test_result <- tryCatch(
                    t.test(group1_data, group2_data, var.equal = equal_variances),
                    error = function(e) NULL
                  )
                  test_name <- if (equal_variances) "Student's t-test" else "Welch's t-test"
                }

                if (!is.null(test_result)) {
                  p_val <- test_result$p.value
                  cat(sprintf("  2-group test: %s vs %s, %s p=%.4f\\n",
                              groups_with_data[1], groups_with_data[2], test_name, p_val))

                  # Determine significance symbol using helper function
                  sig_label <- get_sig_symbol(p_val)

                  # Store result text for UI with normality and variance info

                  result_text <- sprintf("Time point %.1f: %s vs %s, %s p=%.4f (%s)",
                                        x_val, groups_with_data[1], groups_with_data[2],
                                        test_name, p_val, sig_label)
                  if (length(normality_text) > 0) {
                    result_text <- paste0(result_text, "\\nNormality (Shapiro-Wilk):\\n", paste(normality_text, collapse="\\n"))
                  }
                  if (nchar(variance_text) > 0) {
                    result_text <- paste0(result_text, "\\n", variance_text)
                  }
                  stat_text_results <- c(stat_text_results, result_text)

                  # Add to results for visual annotation (only if significant)
                  if (sig_label != "" && sig_label != "ns") {
                    stat_results <- rbind(stat_results, data.frame(
                      x = x_val,
                      y_pos = 0,  # Will be calculated later
                      label = sig_label
                    ))
                  }
                }
              }

            } else if (n_groups_at_x >= 3) {
              # THREE OR MORE GROUPS: Use ANOVA + post-hoc
              cat("  Performing ANOVA at this time point...\\n")

              # Prepare data for ANOVA
              anova_data <- data.frame(
                group = factor(data_at_x$group),
                value = as.numeric(data_at_x$value)
              )
              anova_data <- anova_data[complete.cases(anova_data), ]

              # Perform normality test for each group
              normality_text <- c()
              for (grp in unique(anova_data$group)) {
                grp_data <- anova_data$value[anova_data$group == grp]
                if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
                  shapiro_result <- tryCatch(
                    shapiro.test(grp_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result)) {
                    norm_status <- if (shapiro_result$p.value >= 0.05) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", grp, shapiro_result$p.value, norm_status))
                    cat(sprintf("    %s: Shapiro-Wilk p=%.4f\\n", grp, shapiro_result$p.value))
                  }
                }
              }

              # Perform ANOVA
              anova_result <- tryCatch(
                aov(value ~ group, data = anova_data),
                error = function(e) {
                  cat("  ANOVA error:", e$message, "\\n")
                  NULL
                }
              )

              if (!is.null(anova_result)) {
                anova_summary <- summary(anova_result)
                anova_p <- anova_summary[[1]][["Pr(>F)"]][1]
                cat(sprintf("  ANOVA p-value: %.4f\\n", anova_p))

                # Determine ANOVA significance using helper function
                anova_sig <- ""
                if (!is.na(anova_p)) {
                  anova_sig <- get_sig_symbol(anova_p)
                }

                # Build organized output similar to other chart types
                # Header
                anova_text <- sprintf("Time point %.1f:", x_val)

                # Normality tests
                if (length(normality_text) > 0) {
                  anova_text <- paste0(anova_text, "\\nNormality (Shapiro-Wilk):\\n", paste(normality_text, collapse="\\n"))
                }

                # ANOVA result
                anova_text <- paste0(anova_text, sprintf("\\nANOVA p=%.4f (%s)", anova_p, anova_sig))

                # If ANOVA is significant, perform post-hoc test
                if (!is.na(anova_p) && anova_p < 0.05) {
                  cat("  ANOVA significant - performing Tukey post-hoc test...\\n")

                  # Perform Tukey HSD post-hoc test
                  tukey_result <- tryCatch(
                    TukeyHSD(anova_result),
                    error = function(e) {
                      cat("  Tukey error:", e$message, "\\n")
                      NULL
                    }
                  )

                  if (!is.null(tukey_result)) {
                    tukey_summary <- tukey_result$group
                    cat(sprintf("  Tukey found %d pairwise comparisons\\n", nrow(tukey_summary)))

                    # Add post-hoc section header
                    anova_text <- paste0(anova_text, "\\nPost-hoc pairwise comparisons (Tukey HSD):")

                    # Store all pairwise comparisons
                    tukey_text <- c()
                    for (i in 1:nrow(tukey_summary)) {
                      comparison <- rownames(tukey_summary)[i]
                      p_adj <- tukey_summary[i, "p adj"]
                      diff <- tukey_summary[i, "diff"]

                      sig_label <- ""
                      if (!is.na(p_adj)) {
                        sig_label <- get_sig_symbol(p_adj)
                      } else {
                        sig_label <- "ns"
                      }

                      cat(sprintf("    %s: diff=%.2f, p=%.4f (%s)\\n",
                                  comparison, diff, p_adj, sig_label))

                      # Format for UI
                      tukey_text <- c(tukey_text, sprintf("  %s: diff=%.2f, p=%.4f (%s)",
                                                          comparison, diff, ifelse(is.na(p_adj), 1.0, p_adj), sig_label))
                    }

                    # Add pairwise comparisons
                    anova_text <- paste0(anova_text, "\\n", paste(tukey_text, collapse="\\n"))
                  }
                }

                # Add to results
                stat_text_results <- c(stat_text_results, anova_text)
              }
            }
          }
        }

        # For 2-group data only: Add visual symbols
        if (n_groups == 2 && nrow(stat_results) > 0) {
          # Calculate X offset for symbol positioning
          x_range <- diff(range(unique_x))
          x_offset <- x_range * 0.075  # Offset to the right

          for (i in 1:nrow(stat_results)) {
            x_val <- stat_results$x[i]

            # Get the mean Y values for both groups at this X
            summary_at_x <- plot_data[plot_data$x == x_val, ]

            # Sort by group to ensure consistent ordering
            summary_at_x <- summary_at_x[order(summary_at_x$group), ]

            # Second group is the treatment (non-control)
            treatment_mean <- summary_at_x$mean_val[2]
            treatment_error <- summary_at_x$error_val[2]
            treatment_y_bottom <- treatment_mean - treatment_error

            # Position symbol below the treatment group's error bar, slightly to the right
            # Get Y range to calculate appropriate offset
            y_range <- diff(range(c(plot_data$mean_val + plot_data$error_val,
                                    plot_data$mean_val - plot_data$error_val), na.rm = TRUE))
            y_offset <- y_range * 0.03  # 3% of Y range below the error bar

            symbol_y <- treatment_y_bottom - y_offset
            symbol_x <- x_val + x_offset

            # Add significance symbol (use user-specified symbol size)
            p <- p + annotate("text", x = symbol_x, y = symbol_y,
                             label = stat_results$label[i], size = stat_symbol_size, fontface = "plain")
          }

          cat(sprintf("Added %d statistical annotations\\n", nrow(stat_results)))
        } else if (n_groups >= 3) {
          cat("Multi-group line plot (3+ groups): Using vbracket for statistical legend\\n")

          # Install and load vbracket package from webR-compatible repository
          vbracket_loaded <- FALSE
          tryCatch({
            # Try loading package first
            if (!require("vbracket", quietly = TRUE)) {
              cat("Installing vbracket package from r-universe (WebAssembly)...\\n")
              # Use webR's install function with r-universe repository
              webr::install("vbracket",
                          repos = c("https://h20gg702.r-universe.dev", "https://repo.r-wasm.org"))
              cat("✓ vbracket package installed\\n")
            }
            suppressPackageStartupMessages(library(vbracket))
            vbracket_loaded <- TRUE
            cat("✓ vbracket package loaded successfully\\n")
          }, error = function(e1) {
            cat(sprintf("✗ Failed to load vbracket: %s\\n", e1$message))
          })

          # Only proceed if vbracket loaded successfully
          if (vbracket_loaded) {
            tryCatch({

              # Log comparison mode for debugging
              cat(sprintf("VBracket comparison_mode: %s\\n", comparison_mode))

              # Parse custom_comparisons from JSON if needed for custom mode
              custom_comparisons_vec <- c()
              if (comparison_mode == "custom" && nchar(custom_comparisons) > 2) {
                tryCatch({
                  if (!requireNamespace("jsonlite", quietly = TRUE)) {
                    webr::install("jsonlite")
                  }
                  library(jsonlite)
                  custom_comparisons_list <- fromJSON(custom_comparisons)
                  if (length(custom_comparisons_list) > 0) {
                    custom_comparisons_vec <- unlist(custom_comparisons_list)
                    cat(sprintf("Custom comparisons: %s\\n", paste(custom_comparisons_vec, collapse=", ")))
                  }
                }, error = function(e) {
                  cat("Warning: Failed to parse custom_comparisons JSON\\n")
                })
              }

              # Parse stat_text_results to extract significant comparisons
              # Format: "Group1-Group2: diff=X.XX, p=X.XXXX (**)"
              # Build vectors for data frame columns
              groups1 <- c()
              groups2 <- c()
              labels <- c()

              if (length(stat_text_results) > 0) {
                # Filter stat_text_results by selected timepoint
                filtered_results <- stat_text_results
                if (nchar(vbracket_timepoint) > 0) {
                  # Create pattern to match "Time point X:" or "Time point X.0:" where X is the selected timepoint
                  # Handle both integer and decimal formats (e.g., "4" should match "4.0")
                  timepoint_num <- as.numeric(vbracket_timepoint)
                  # Create regex pattern that matches "Time point 4:" or "Time point 4.0:"
                  timepoint_pattern <- sprintf("Time point %s(\\\\.0)?:", gsub("\\\\.", "\\\\\\\\.", vbracket_timepoint))
                  cat(sprintf("Filtering statistical results for timepoint: %s (pattern: %s)\\n", vbracket_timepoint, timepoint_pattern))

                  # Keep only results that start with the selected timepoint
                  filtered_results <- stat_text_results[grepl(timepoint_pattern, stat_text_results)]

                  if (length(filtered_results) == 0) {
                    cat(sprintf("⚠️ No statistical results found for timepoint %s\\n", vbracket_timepoint))
                    cat(sprintf("Available results:\\n%s\\n", paste(stat_text_results, collapse="\\n")))
                  } else {
                    cat(sprintf("Found %d result(s) for timepoint %s\\n", length(filtered_results), vbracket_timepoint))
                  }
                }

                for (result_text in filtered_results) {
                  cat(sprintf("DEBUG: Processing result_text: %s\\n", substr(result_text, 1, 200)))

                  # Split by newline to get individual comparisons
                  lines <- strsplit(result_text, "\\\\n")[[1]]
                  cat(sprintf("DEBUG: Found %d lines\\n", length(lines)))

                  for (line in lines) {
                    cat(sprintf("DEBUG: Checking line: %s\\n", line))

                    # Look for comparison lines with significance symbol in parentheses
                    # Matches any symbol type: (*), (**), (***), (A), (B), (C), (ns), (p=0.023)
                    if (grepl("-", line) && grepl("\\\\([^)]+\\\\)$", line)) {
                      cat(sprintf("DEBUG: Line has comparison with significance!\\n"))

                      # Extract groups: "Group1-Group2: diff=..."
                      match <- regmatches(line, regexpr("^[^:]+(?=:)", line, perl=TRUE))
                      if (length(match) > 0) {
                        groups_str <- gsub("^\\\\s+", "", match[1])  # Trim leading space
                        cat(sprintf("DEBUG: Extracted groups string: '%s'\\n", groups_str))

                        group_pair <- strsplit(groups_str, "-")[[1]]
                        cat(sprintf("DEBUG: Split into %d parts\\n", length(group_pair)))

                        # Extract significance symbol (anything in parentheses at end of line)
                        sig_match <- regmatches(line, regexpr("\\\\([^)]+\\\\)$", line))
                        sig_label <- if (length(sig_match) > 0) gsub("[()]", "", sig_match[1]) else "*"
                        cat(sprintf("DEBUG: Extracted sig label: '%s'\\n", sig_label))

                        if (length(group_pair) == 2) {
                          # Apply comparison_mode filtering
                          should_add <- FALSE

                          if (comparison_mode == "significant") {
                            # Only add if significant (not "ns")
                            should_add <- (sig_label != "ns" && sig_label != "")
                          } else if (comparison_mode == "all") {
                            # Add all comparisons
                            should_add <- TRUE
                          } else if (comparison_mode == "custom") {
                            # Only add if in custom_comparisons list
                            comparison_name <- paste0(group_pair[1], "-", group_pair[2])
                            should_add <- comparison_name %in% custom_comparisons_vec
                          }

                          if (should_add) {
                            groups1 <- c(groups1, group_pair[1])
                            groups2 <- c(groups2, group_pair[2])
                            labels <- c(labels, sig_label)
                            cat(sprintf("DEBUG: Added comparison: %s -> %s (%s)\\n", group_pair[1], group_pair[2], sig_label))
                          } else {
                            cat(sprintf("DEBUG: Skipped comparison (mode=%s): %s -> %s (%s)\\n", comparison_mode, group_pair[1], group_pair[2], sig_label))
                          }
                        }
                      }
                    }
                  }
                }
              }

              # Summary of filtering
              cat(sprintf("VBracket filtering complete: %d comparisons selected (mode: %s)\\n", length(groups1), comparison_mode))

              # Only add vbracket if we have significant comparisons
              if (length(groups1) > 0) {
                # Create comparisons data frame (same format as add_bracket_comparisons())
                comparisons_df <- data.frame(
                  group1 = groups1,
                  group2 = groups2,
                  label = labels,
                  stringsAsFactors = FALSE
                )

                cat(sprintf("DEBUG: Comparisons data frame:\\n"))
                print(comparisons_df)
                cat(sprintf("Adding vbracket legend with %d comparisons\\n", nrow(comparisons_df)))

                # Get group labels and colors
                group_labels <- levels(plot_data$group)
                legend_colors <- line_colors[1:length(group_labels)]

                # Add vbracket custom legend (default legend already suppressed above)

                # Use manual X/Y positioning
                cat(sprintf("Using vbracket position: x=%.2f, y=%.2f\\n", vbracket_x, vbracket_y))
                p <- p + legend_bracket(
                  labels = group_labels,
                  colors = legend_colors,
                  comparisons = comparisons_df,
                  legend_x = vbracket_x,
                  legend_y = vbracket_y,
                  text_size = vbracket_text_size,
                  sig_size = vbracket_sig_size,
                  bracket_margin = vbracket_margin,
                  line_length = vbracket_legend_line_length,
                  line_width = vbracket_legend_line_width,
                  item_spacing = vbracket_item_spacing,
                  bracket_layer_spacing = vbracket_bracket_layer_spacing,
                  output_width = output_width,
                  output_height = output_height,
                  text_family = target_font
                )

                cat("✅ Vbracket legend added successfully\\n")
              } else {
                cat("No significant comparisons found for vbracket\\n")
              }
            }, error = function(e) {
              cat("Error adding vbracket:", e$message, "\\n")
            })
          } else {
            cat("❌ vbracket could not be loaded\\n")
          }
        }

        # Store statistical results in global variable for JavaScript to retrieve
        cat("\\nDEBUG: Number of stat_text_results =", length(stat_text_results), "\\n")
        if (length(stat_text_results) > 0) {
          # Use \\n\\n to separate time points (same as other chart types)
          line_plot_stat_results <<- paste(stat_text_results, collapse="\\n\\n")
          cat("\\n=== STATISTICAL RESULTS SUMMARY ===\\n")
          cat(line_plot_stat_results, "\\n")
          cat("===================================\\n")
          cat("DEBUG: Stored", nchar(line_plot_stat_results), "characters in line_plot_stat_results\\n")
        } else {
          cat("DEBUG: No stat_text_results to store\\n")
          line_plot_stat_results <<- ""
        }
      } else {
        # Statistics not enabled
        cat("DEBUG: Statistics not enabled (add_statistics=", add_statistics, ", n_groups=", n_groups, ")\\n")
        line_plot_stat_results <<- ""
      }

      p <- sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)

      # For 3+ groups with statistics: Suppress default legend AFTER theme is applied
      # (theme application resets legend.position, so must suppress afterwards)
      if (add_statistics && n_groups >= 3) {
        p <- p + theme(legend.position = "none")
      }

      p
    }

    sato_bar_error <- function(dat, x_col=1, y_col=2, error_col=3, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9,
                              width=0.4, errorbar_width=0.2,
                              target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                              title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                              title_text="Bar plot with error bars", x_text="Category", y_text="Mean",
                              show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                              x_scale="linear", y_scale="linear",
                              theme_name="minimal",
                              x_axis_rotation=0, y_axis_rotation=0,
                              x_axis_hjust=0.5, x_axis_vjust=0.5,
                              y_axis_hjust=0.5, y_axis_vjust=0.5,
                              add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales  
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Prepare data frame with categories, means, and errors
      df <- data.frame(
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        mean_val = as.numeric(dat[[y_col]]),
        error = as.numeric(dat[[error_col]])
      )
      
      # Create bar plot with error bars
      p <- ggplot(df, aes(x=category, y=mean_val)) +
           geom_col(fill=fill, color=color, linewidth=linewidth, alpha=alpha, width=width) +
           geom_errorbar(aes(ymin=mean_val-error, ymax=mean_val+error), 
                        width=errorbar_width, color=color, linewidth=linewidth*0.8)

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    sato_bar_error_dot <- function(dat, x_col=1, y_col=2, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9,
                                  width=0.4, errorbar_width=0.2, dot_size=4, dot_alpha=1.0, dot_color="#333333", dot_shape=16, jitter_width=0.2,
                                  error_type="sd",
                                  target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                  title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                  title_text="Bar plot with error bars and data points", x_text="Category", y_text="Value",
                                  show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                  x_scale="linear", y_scale="linear",
                                  theme_name="minimal",
                                  x_axis_rotation=0, y_axis_rotation=0,
                                  x_axis_hjust=0.5, x_axis_vjust=0.5,
                                  y_axis_hjust=0.5, y_axis_vjust=0.5,
                                  add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Prepare data frame
      df <- data.frame(
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      
      # For bar_error_dot, always use column names unless user specifically entered text
      # Strip quotes and check if it's meaningful text
      x_clean <- gsub('^["\\\']|["\\\']$', '', x_text)  # Remove surrounding quotes
      y_clean <- gsub('^["\\\']|["\\\']$', '', y_text)  # Remove surrounding quotes
      
      # Use column names by default, unless user entered meaningful custom text
      if (is.null(x_clean) || x_clean == "" || x_clean == "NULL") {
        x_text <- names(dat)[x_col]
      } else {
        x_text <- x_clean  # Use the cleaned custom text
      }
      
      if (is.null(y_clean) || y_clean == "" || y_clean == "NULL") {
        y_text <- names(dat)[y_col]
      } else {
        y_text <- y_clean  # Use the cleaned custom text
      }
      
      # Calculate summary statistics by category
      if (requireNamespace("dplyr", quietly = TRUE)) {
        library(dplyr)
        summary_df <- df %>%
          group_by(category) %>%
          summarise(
            mean_val = mean(value, na.rm = TRUE),
            sd_val = ifelse(n() > 1, sd(value, na.rm = TRUE), 0),
            n_val = n(),
            se_val = ifelse(n() > 1, sd(value, na.rm = TRUE) / sqrt(n()), 0),
            .groups = "drop"
          )
      } else {
        # Manual calculation if dplyr not available
        categories <- levels(df$category)
        summary_list <- list()
        for (cat in categories) {
          cat_data <- df$value[df$category == cat]
          n_obs <- length(cat_data[!is.na(cat_data)])
          sd_calc <- if(n_obs > 1) sd(cat_data, na.rm = TRUE) else 0
          se_calc <- if(n_obs > 1) sd_calc / sqrt(n_obs) else 0
          summary_list[[length(summary_list) + 1]] <- data.frame(
            category = factor(cat, levels = categories),
            mean_val = mean(cat_data, na.rm = TRUE),
            sd_val = sd_calc,
            n_val = n_obs,
            se_val = se_calc
          )
        }
        summary_df <- do.call(rbind, summary_list)
      }
      
      # Select error value based on error_type
      if (error_type == "se") {
        summary_df$error_val <- summary_df$se_val
      } else {
        summary_df$error_val <- summary_df$sd_val
      }
      
      
      # Create combined plot: bars + error bars + individual points
      p <- ggplot() +
           # Bar layer (means)
           geom_col(data = summary_df, aes(x = category, y = mean_val), 
                   fill = fill, color = color, linewidth = linewidth, alpha = alpha, width = width) +
           # Error bar layer
           geom_errorbar(data = summary_df, aes(x = category, ymin = mean_val - error_val, ymax = mean_val + error_val), 
                        width = errorbar_width, color = color, linewidth = linewidth * 0.8) +
           # Individual data points (jittered horizontally only)
           geom_point(data = df, aes(x = category, y = value),
                     position = position_jitter(width = jitter_width, height = 0),
                     size = dot_size, alpha = dot_alpha, color = dot_color, shape = dot_shape)

      # Add statistical analysis if requested and we have multiple groups
      cat("\\n=== BAR ERROR DOT STATISTICAL CHECK ===\\n")
      cat("add_statistics:", add_statistics, "\\n")
      cat("Number of unique categories:", length(unique(df$category)), "\\n")
      cat("Categories found:", paste(unique(df$category), collapse=", "), "\\n")
      
      cat("🔥🔥🔥 BAR_ERROR_DOT STATISTICAL CHECK 🔥🔥🔥\\n")
      cat("add_statistics:", add_statistics, "\\n")
      cat("unique categories:", length(unique(df$category)), "\\n")
      cat("categories:", paste(unique(df$category), collapse=", "), "\\n")
      
      if (add_statistics && length(unique(df$category)) > 1) {
        cat("🔥🔥🔥 CALLING STATISTICAL ANALYSIS FUNCTION FROM BAR_ERROR_DOT 🔥🔥🔥\\n")
        p <- sato_add_statistics_to_plot(p, df, "category", "value", statistical_test, sato_symbol_size, show_main_symbol=${showMainStatSymbol ? 'TRUE' : 'FALSE'}, show_pairwise=${showPairwiseComparisons ? 'TRUE' : 'FALSE'}, posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale="${yScale}", stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      } else {
        if (!add_statistics) {
          cat("🔥🔥🔥 STATISTICAL ANALYSIS IS DISABLED IN BAR_ERROR_DOT 🔥🔥🔥\\n")
        } else {
          cat("🔥🔥🔥 NOT ENOUGH GROUPS FOR STATISTICAL ANALYSIS IN BAR_ERROR_DOT 🔥🔥🔥\\n")
        }
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    sato_box_dot <- function(dat, x_col=1, y_col=2, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9, width=0.7,
                            dot_size=4, dot_alpha=1.0, dot_color="#333333", dot_shape=16, jitter_width=0.2,
                            target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                            title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                            title_text="Box plot with data points", x_text="Category", y_text="Value",
                            show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                            x_scale="linear", y_scale="linear",
                            theme_name="minimal",
                            x_axis_rotation=0, y_axis_rotation=0,
                            x_axis_hjust=0.5, x_axis_vjust=0.5,
                            y_axis_hjust=0.5, y_axis_vjust=0.5,
                            add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Prepare data frame
      df <- data.frame(
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      
      # For box_dot, always use column names unless user specifically entered text
      # Strip quotes and check if it's meaningful text
      x_clean <- gsub('^["\\\']|["\\\']$', '', x_text)  # Remove surrounding quotes
      y_clean <- gsub('^["\\\']|["\\\']$', '', y_text)  # Remove surrounding quotes
      
      # Use column names by default, unless user entered meaningful custom text
      if (is.null(x_clean) || x_clean == "" || x_clean == "NULL") {
        x_text <- names(dat)[x_col]
      } else {
        x_text <- x_clean  # Use the cleaned custom text
      }
      
      if (is.null(y_clean) || y_clean == "" || y_clean == "NULL") {
        y_text <- names(dat)[y_col]
      } else {
        y_text <- y_clean  # Use the cleaned custom text
      }
      
      # Create combined plot: box plot + individual points
      p <- ggplot(df, aes(x = category, y = value)) +
           # Box plot layer
           geom_boxplot(fill = fill, color = color, linewidth = linewidth, alpha = alpha, width = width,
                       outlier.shape = NA) +  # Hide default outliers to avoid duplication
           # Individual data points (jittered horizontally only)
           geom_point(position = position_jitter(width = jitter_width, height = 0),
                     size = dot_size, alpha = dot_alpha, color = dot_color, shape = dot_shape)

      # Add statistical analysis if requested and we have multiple groups
      if (add_statistics && length(unique(df$category)) > 1) {
        p <- sato_add_statistics_to_plot(p, df, "category", "value", statistical_test, sato_symbol_size, posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale=y_scale, stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped bar chart function - for data with Group/Category columns
    sato_bar_grouped <- function(dat, group_col=1, x_col=2, y_col=3,
                                fill_colors=c("#4C78A8", "#E15759"), stroke_color="#1f2937",
                                alpha=0.9, linewidth=0.7, width=0.7, dodge_width=0.9, position="dodge",
                                group_name="Group", category_name="Treatment", value_name="Value",
                                target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                title_text="Grouped Bar Plot", x_text="Category", y_text="Value",
                                show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                x_scale="linear", y_scale="linear",
                                theme_name="minimal",
                                x_axis_rotation=0, y_axis_rotation=0,
                                x_axis_hjust=0.5, x_axis_vjust=0.5,
                                y_axis_hjust=0.5, y_axis_vjust=0.5,
                                add_statistics=FALSE, statistical_test="auto") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # データフレーム構造確認
      if (ncol(dat) < 3) {
        stop("Grouped bar chart requires at least 3 columns: Group, Category, Value")
      }
      
      # 列名を取得（データから実際の名前を取得、フォールバックあり）
      cat("DEBUG - Original column names:", paste(names(dat), collapse=", "), "\\n")
      cat("DEBUG - Data structure:", str(dat), "\\n")
      
      # データの実際の列名をチェック（V1, V2, V3 の場合はパラメータ名を使用）
      col_names <- names(dat)
      actual_group_name <- if(!is.null(col_names) && length(col_names) >= group_col && !grepl("^V[0-9]+$", col_names[group_col])) {
        col_names[group_col]
      } else {
        group_name
      }
      
      actual_category_name <- if(!is.null(col_names) && length(col_names) >= x_col && !grepl("^V[0-9]+$", col_names[x_col])) {
        col_names[x_col] 
      } else {
        category_name
      }
      
      actual_value_name <- if(!is.null(col_names) && length(col_names) >= y_col && !grepl("^V[0-9]+$", col_names[y_col])) {
        col_names[y_col]
      } else {
        value_name
      }
      
      cat("DEBUG - Final names used - Group:", actual_group_name, "Category:", actual_category_name, "Value:", actual_value_name, "\\n")
      
      # データ準備
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      
      # グループの順序を確認・修正
      unique_groups <- unique(as.character(dat[[group_col]]))
      cat("Original group order:", paste(unique_groups, collapse=", "), "\\n")
      
      # グループを文字順でソートして一貫性を保つ
      # Normal が Tumor より前に来るように、または最初に出現した順序を保つ
      if ("Normal" %in% unique_groups && "Tumor" %in% unique_groups) {
        # Normal/Tumor の場合は Normal を最初に
        plot_data$group <- factor(plot_data$group, levels = c("Normal", "Tumor"))
      } else {
        # その他の場合は最初に出現した順序を維持
        plot_data$group <- factor(plot_data$group, levels = unique_groups)
      }
      
      # NA値を除去
      plot_data <- plot_data[complete.cases(plot_data), ]
      
      if (nrow(plot_data) == 0) {
        stop("No valid data points after removing NAs")
      }
      
      cat("Plot data prepared:", nrow(plot_data), "rows\\n")
      cat("Groups:", paste(levels(plot_data$group), collapse=", "), "\\n")
      cat("Categories:", paste(levels(plot_data$category), collapse=", "), "\\n")
      
      # グループ数に応じて色を設定
      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)
      
      if (length(fill_colors) < n_groups) {
        # 足りない場合は色を拡張
        fill_colors <- rep(fill_colors, length.out = n_groups)
      }
      
      # 色の割り当てをログ出力
      cat("Group levels:", paste(group_levels, collapse=", "), "\\n")
      cat("Assigned colors:", paste(fill_colors[1:n_groups], collapse=", "), "\\n")
      
      # グループ化されたバープロット作成
      p <- ggplot(plot_data, aes(x = category, y = value, fill = group)) +
           geom_col(position = position_dodge(width = dodge_width), 
                   width = width,
                   alpha = alpha, 
                   linewidth = linewidth,
                   color = stroke_color) +  # Use controllable stroke color
           scale_fill_manual(values = setNames(fill_colors[1:n_groups], group_levels)) +
           labs(fill = actual_group_name)  # Legend title using actual column name
      
      # Add statistical analysis if requested
      if (add_statistics && ncol(dat) >= 3) {
        p <- sato_add_statistics_to_plot(p, plot_data, "group", "value", statistical_test, sato_symbol_size, posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale=y_scale, stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped bar chart with error bars
    sato_bar_grouped_error <- function(dat, group_col=1, x_col=2, y_col=3, error_col=4,
                                      fill_colors=c("#4C78A8", "#E15759"), stroke_color="#1f2937",
                                      alpha=0.9, linewidth=0.7, width=0.7, dodge_width=0.9,
                                      group_name="Group", category_name="Category", value_name="Mean",
                                      target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                      title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                      title_text="Grouped Bar Plot with Error", x_text="Category", y_text="Mean",
                                      show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                      x_scale="linear", y_scale="linear",
                                      theme_name="minimal",
                                      x_axis_rotation=0, y_axis_rotation=0,
                                      x_axis_hjust=0.5, x_axis_vjust=0.5,
                                      y_axis_hjust=0.5, y_axis_vjust=0.5) {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      if (ncol(dat) < 4) {
        stop("Grouped bar chart with error requires 4 columns: Group, Category, Mean, Error")
      }

      # Detect column name for group
      actual_group_name <- if(!is.null(names(dat)) && length(names(dat)) >= group_col && !grepl("^V[0-9]+$", names(dat)[group_col])) {
        names(dat)[group_col]
      } else {
        group_name
      }

      # Prepare data using PRE-CALCULATED mean and error values (no calculation)
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        mean_val = as.numeric(dat[[y_col]]),
        error_val = as.numeric(dat[[error_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]

      # Special handling for Normal/Tumor ordering
      if ("Normal" %in% unique(as.character(plot_data$group)) && "Tumor" %in% unique(as.character(plot_data$group))) {
        plot_data$group <- factor(plot_data$group, levels = c("Normal", "Tumor"))
      }

      # Color setup
      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)
      if (length(fill_colors) < n_groups) {
        fill_colors <- rep(fill_colors, length.out = n_groups)
      }

      # Create plot with pre-calculated error bars
      p <- ggplot(plot_data, aes(x = category, y = mean_val, fill = group)) +
           geom_col(position = position_dodge(width = dodge_width),
                   width = width,
                   alpha = alpha,
                   linewidth = linewidth,
                   color = stroke_color) +
           geom_errorbar(aes(ymin = mean_val - error_val, ymax = mean_val + error_val),
                        position = position_dodge(width = dodge_width),
                        width = 0.25, linewidth = linewidth * 0.8) +
           scale_fill_manual(values = setNames(fill_colors[1:n_groups], group_levels)) +
           labs(fill = actual_group_name)

      # NO statistical analysis for pre-calculated data (no raw data points available)

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      FALSE, "none")  # No statistics for pre-calculated data
    }

    # Grouped bar chart with error bars and individual data points
    sato_bar_grouped_error_dot <- function(dat, group_col=1, x_col=2, y_col=3,
                                          fill_colors=c("#4C78A8", "#E15759", "#76B7B2", "#F28E2B"), stroke_color="#1f2937",
                                          alpha=0.9, linewidth=0.7, width=0.7, dodge_width=0.9, error_type="sd",
                                          dot_size=4, dot_alpha=1.0, dot_color="#333333", dot_shape=16, jitter_width=0.15,
                                          group_name="Group", category_name="Treatment", value_name="Value",
                                          target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                          title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                          title_text="Grouped Bar + Dot Plot with Error", x_text="Category", y_text="Value",
                                          show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                          x_scale="linear", y_scale="linear",
                                          theme_name="minimal",
                                          x_axis_rotation=0, y_axis_rotation=0,
                                          x_axis_hjust=0.5, x_axis_vjust=0.5,
                                          y_axis_hjust=0.5, y_axis_vjust=0.5,
                                          add_statistics=FALSE, statistical_test="auto", variance_test="levene", symbol_size=8,
                                          ggpubr_line_size=1.0, ggpubr_tip_length=0.04, ggpubr_vjust=-0.3,
                                          comparison_mode="all", custom_comparisons="[]", custom_positions="{}",
                                          stat_symbol_type="stars", custom_symbol_05="*", custom_symbol_01="**", custom_symbol_001="***", custom_symbol_ns="ns") {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Helper function to generate significance symbol based on p-value and symbol type
      get_sig_symbol <- function(p_val) {
        if (stat_symbol_type == "pvalue") {
          return(sprintf("p=%.3f", p_val))
        } else if (stat_symbol_type == "custom") {
          if (p_val < 0.001) return(custom_symbol_001)
          else if (p_val < 0.01) return(custom_symbol_01)
          else if (p_val < 0.05) return(custom_symbol_05)
          else return(custom_symbol_ns)
        } else {
          # Default to stars
          if (p_val < 0.001) return("***")
          else if (p_val < 0.01) return("**")
          else if (p_val < 0.05) return("*")
          else return("ns")
        }
      }
      
      if (ncol(dat) < 3) {
        stop("Grouped bar chart requires at least 3 columns: Group, Category, Value")
      }
      
      # 列名検出
      actual_group_name <- if(!is.null(names(dat)) && length(names(dat)) >= group_col && !grepl("^V[0-9]+$", names(dat)[group_col])) {
        names(dat)[group_col]
      } else {
        group_name
      }
      
      # データ準備
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]
      
      if ("Normal" %in% unique(as.character(plot_data$group)) && "Tumor" %in% unique(as.character(plot_data$group))) {
        plot_data$group <- factor(plot_data$group, levels = c("Normal", "Tumor"))
      }
      
      # 要約統計を計算
      if (requireNamespace("dplyr", quietly = TRUE)) {
        summary_data <- plot_data %>%
          dplyr::group_by(group, category) %>%
          dplyr::summarise(
            mean_val = mean(value, na.rm = TRUE),
            sd_val = sd(value, na.rm = TRUE),
            se_val = sd_val / sqrt(dplyr::n()),
            .groups = "drop"
          )
        summary_data$error_val <- if(error_type == "se") summary_data$se_val else summary_data$sd_val
      } else {
        stop("dplyr package required for grouped bar chart with error")
      }
      
      # 色設定
      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)
      if (length(fill_colors) < n_groups) {
        fill_colors <- rep(fill_colors, length.out = n_groups)
      }
      
      # プロット作成
      p <- ggplot(summary_data, aes(x = category, y = mean_val, fill = group)) +
           geom_col(position = position_dodge(width = dodge_width), 
                   width = width,
                   alpha = alpha, 
                   linewidth = linewidth,
                   color = stroke_color) +
           geom_errorbar(aes(ymin = mean_val - error_val, ymax = mean_val + error_val),
                        position = position_dodge(width = dodge_width),
                        width = 0.25, linewidth = linewidth * 0.8) +
           geom_point(data = plot_data, aes(x = category, y = value, fill = group),
                     position = position_jitterdodge(dodge.width = dodge_width, jitter.width = jitter_width, jitter.height = 0),
                     size = dot_size, alpha = dot_alpha, shape = dot_shape, color = dot_color) +
           scale_fill_manual(values = setNames(fill_colors[1:n_groups], group_levels)) +
           labs(fill = actual_group_name)

      # Initialize global variable for statistical results
      grouped_bar_stat_results <<- ""

      # Add statistical analysis if requested - compare groups within each category
      if (add_statistics && n_groups >= 2) {
        tryCatch({
          # Load ggpubr package for stat_pvalue_manual
          if (!require("ggpubr", quietly = TRUE)) {
            webr::install("ggpubr")
          }
          library(ggpubr)

          # Get unique categories (filter out any categories with no data)
          all_categories <- levels(plot_data$category)
          unique_categories <- all_categories[all_categories %in% unique(plot_data$category)]

          # Handle custom comparisons
          selected_categories <- NULL
          custom_y_positions <- list()
          custom_comps <- c()  # Initialize to empty array

          if (comparison_mode == "custom") {
            # Parse custom comparisons: format is "Group1-Group2@CategoryA", "Group1-Group2@CategoryB", etc.
            custom_comps <- tryCatch(jsonlite::fromJSON(custom_comparisons), error = function(e) c())
            cat("Custom comparisons:", custom_comparisons, "\\n")
            cat("Custom positions JSON:", custom_positions, "\\n")

            if (length(custom_comps) > 0) {
              # Extract category values from custom comparisons
              selected_categories <- c()
              for (comp_str in custom_comps) {
                # Parse format: "Group1-Group2@Category"
                parts <- strsplit(comp_str, "@", fixed = TRUE)[[1]]
                if (length(parts) == 2) {
                  cat_val <- parts[2]
                  selected_categories <- c(selected_categories, cat_val)
                }
              }
              cat("Selected categories from custom comparisons:", paste(selected_categories, collapse=", "), "\\n")

              # Parse custom positions
              custom_pos_list <- tryCatch(jsonlite::fromJSON(custom_positions), error = function(e) list())
              if (length(custom_pos_list) > 0) {
                for (comp_str in names(custom_pos_list)) {
                  custom_y_positions[[comp_str]] <- as.numeric(custom_pos_list[[comp_str]])
                }
              }
            }
          }

          # Determine which categories to test
          if (is.null(selected_categories) || length(selected_categories) == 0) {
            # Test all categories (default or "all" mode)
            categories_to_test <- unique_categories
          } else {
            # Test only selected categories (deduplicate to avoid testing same category multiple times)
            categories_to_test <- unique(selected_categories)
          }

          # Debug: Check what we're testing
          cat("DEBUG: comparison_mode =", comparison_mode, "\\n")
          cat("DEBUG: categories_to_test =", paste(categories_to_test, collapse=", "), "\\n")
          cat("DEBUG: Number of categories to test:", length(categories_to_test), "\\n")

          stat_text_results <- c()

        # Data frame to store ggpubr bracket information
        bracket_data <- data.frame(
          group1 = character(),
          group2 = character(),
          p.signif = character(),
          x.position = numeric(),
          y.position = numeric(),
          stringsAsFactors = FALSE
        )

        # For each category, perform statistical tests between groups
        for (cat_idx in seq_along(categories_to_test)) {
          cat_val <- categories_to_test[cat_idx]

          # Get data for this category
          data_at_cat <- plot_data[plot_data$category == cat_val, ]

          if (nrow(data_at_cat) > 0) {
            # Get groups with data at this category
            groups_with_data <- unique(data_at_cat$group[!is.na(data_at_cat$value)])
            n_groups_at_cat <- length(groups_with_data)

            if (n_groups_at_cat == 2) {
              # TWO GROUPS: Use t-test or Wilcoxon
              group1_data <- data_at_cat[data_at_cat$group == groups_with_data[1], "value"]
              group2_data <- data_at_cat[data_at_cat$group == groups_with_data[2], "value"]

              # Only test if both groups have data
              if (length(group1_data) > 0 && length(group2_data) > 0) {
                # Perform normality test for both groups
                normality_text <- c()
                is_group1_normal <- TRUE
                is_group2_normal <- TRUE

                if (length(group1_data) >= 3 && length(group1_data) <= 5000) {
                  shapiro_result1 <- tryCatch(
                    shapiro.test(group1_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result1)) {
                    is_group1_normal <- shapiro_result1$p.value >= 0.05
                    norm_status <- if (is_group1_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[1], shapiro_result1$p.value, norm_status))
                  }
                }

                if (length(group2_data) >= 3 && length(group2_data) <= 5000) {
                  shapiro_result2 <- tryCatch(
                    shapiro.test(group2_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result2)) {
                    is_group2_normal <- shapiro_result2$p.value >= 0.05
                    norm_status <- if (is_group2_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[2], shapiro_result2$p.value, norm_status))
                  }
                }

                both_normal <- is_group1_normal && is_group2_normal

                # Select test based on mode
                test_to_use <- statistical_test
                if (statistical_test == "auto") {
                  if (both_normal) {
                    test_to_use <- "t-test"
                    cat("  Auto-selected: t-test (both groups normal)\\n")
                  } else {
                    test_to_use <- "wilcoxon"
                    cat("  Auto-selected: Wilcoxon test (non-normal data detected)\\n")
                  }
                } else if (statistical_test == "parametric") {
                  test_to_use <- "t-test"
                  cat("  Manual mode: parametric t-test\\n")
                } else if (statistical_test == "nonparametric") {
                  test_to_use <- "wilcoxon"
                  cat("  Manual mode: non-parametric Wilcoxon test\\n")
                }

                # Perform variance test (only for parametric tests)
                variance_text <- ""
                equal_variances <- TRUE
                if (test_to_use != "wilcoxon") {
                  if (variance_test == "levene") {
                    # Levene test for equality of variances
                    combined_data <- data.frame(
                      values = c(group1_data, group2_data),
                      group = factor(c(rep(groups_with_data[1], length(group1_data)),
                                      rep(groups_with_data[2], length(group2_data))))
                    )

                    # Calculate Levene test manually
                    group_means <- tapply(combined_data$values, combined_data$group, mean)
                    abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
                    levene_result <- tryCatch(
                      anova(lm(abs_deviations ~ combined_data$group)),
                      error = function(e) NULL
                    )

                    if (!is.null(levene_result)) {
                      levene_p <- levene_result$\`Pr(>F)\`[1]
                      equal_variances <- levene_p > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, Levene)", levene_p, variance_status)
                      cat(sprintf("    Levene test: p=%.4f (%s)\\n", levene_p, variance_status))
                    }
                  } else {
                    # F-test for equality of variances
                    var_test <- tryCatch(
                      var.test(group1_data, group2_data),
                      error = function(e) NULL
                    )

                    if (!is.null(var_test)) {
                      equal_variances <- var_test$p.value > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, F-test)", var_test$p.value, variance_status)
                      cat(sprintf("    F-test: p=%.4f (%s)\\n", var_test$p.value, variance_status))
                    }
                  }
                }

                # Perform statistical test based on auto-selection
                if (test_to_use == "wilcoxon") {
                  test_result <- tryCatch(
                    wilcox.test(group1_data, group2_data),
                    error = function(e) NULL
                  )
                  test_name <- "Wilcoxon"
                } else {
                  # Use appropriate t-test based on variance equality
                  test_result <- tryCatch(
                    t.test(group1_data, group2_data, var.equal = equal_variances),
                    error = function(e) NULL
                  )
                  test_name <- if (equal_variances) "Student's t-test" else "Welch's t-test"
                }

                if (!is.null(test_result)) {
                  p_val <- test_result$p.value
                  cat(sprintf("  2-group test: %s vs %s, %s p=%.4f\\n",
                              groups_with_data[1], groups_with_data[2], test_name, p_val))

                  # Determine significance symbol using helper function
                  sig_label <- get_sig_symbol(p_val)

                  # Store result text for UI with proper order: Category -> Group Stats -> Normality -> Variance -> Test result
                  result_text <- sprintf("Category %s:", cat_val)

                  # Add group statistics (n, mean, sd) for each group
                  result_text <- paste0(result_text, "\\nGroup Statistics:")
                  for (g in groups_with_data) {
                    g_data <- plot_data[plot_data$category == cat_val & plot_data$group == g, "value"]
                    result_text <- paste0(result_text, sprintf("\\n  %s: n=%d, mean=%.2f, sd=%.2f",
                                          g, length(g_data), mean(g_data, na.rm=TRUE), sd(g_data, na.rm=TRUE)))
                  }

                  # Add normality testing results
                  if (length(normality_text) > 0) {
                    result_text <- paste0(result_text, "\\n\\nNormality Testing (Shapiro-Wilk):\\n  ", paste(normality_text, collapse="\\n  "))
                  }

                  # Add variance test results
                  if (nchar(variance_text) > 0) {
                    result_text <- paste0(result_text, "\\n", variance_text)
                  }

                  # Add main test result
                  result_text <- paste0(result_text, "\\n\\n", sprintf("%s: p=%.4f (%s)",
                                        test_name, p_val, sig_label))
                  stat_text_results <- c(stat_text_results, result_text)

                  # Add ggpubr bracket for this comparison
                  # Calculate x position (find position in all unique_categories, not just selected ones)
                  cat_index <- which(unique_categories == cat_val)

                  # Calculate y position (above the highest bar/point in this category)
                  # Check if custom position is provided (check both possible orderings)
                  comp_key1 <- paste0(groups_with_data[1], "-", groups_with_data[2], "@", cat_val)
                  comp_key2 <- paste0(groups_with_data[2], "-", groups_with_data[1], "@", cat_val)

                  custom_y_pos <- NULL
                  if (!is.null(custom_y_positions[[comp_key1]])) {
                    custom_y_pos <- custom_y_positions[[comp_key1]]
                    cat("Using custom Y position for", comp_key1, ":", custom_y_pos, "\\n")
                  } else if (!is.null(custom_y_positions[[comp_key2]])) {
                    custom_y_pos <- custom_y_positions[[comp_key2]]
                    cat("Using custom Y position for", comp_key2, ":", custom_y_pos, "\\n")
                  }

                  if (!is.null(custom_y_pos)) {
                    # Use custom Y position
                    y_pos <- custom_y_pos
                  } else {
                    # Calculate default position
                    # Get the summary data for this category to find the bar heights
                    # Get all data for this category
                    data_at_cat_for_pos <- plot_data[plot_data$category == cat_val, ]
                    max_value <- max(data_at_cat_for_pos$value, na.rm = TRUE)
                    # Box plots have no error bars

                    # Position symbol above the tallest bar + error bar with some spacing
                    # Increased spacing to avoid overlapping with dots (0.055 -> 0.15)
                    y_pos <- max_value + (max_value * 0.20)  # 20% above max value
                  }

                  # Add to bracket data based on comparison mode
                  should_add_bracket <- FALSE
                  if (comparison_mode == "significant") {
                    # Only add if significant
                    should_add_bracket <- (sig_label != "ns")
                  } else if (comparison_mode == "all") {
                    # Add all comparisons
                    should_add_bracket <- TRUE
                  } else if (comparison_mode == "custom") {
                    # Only add if this comparison is selected
                    # Check if this comparison is in the custom selections
                    comp_key1 <- paste0(groups_with_data[1], "-", groups_with_data[2], "@", cat_val)
                    comp_key2 <- paste0(groups_with_data[2], "-", groups_with_data[1], "@", cat_val)
                    should_add_bracket <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                  }

                  if (should_add_bracket) {
                    bracket_data <- rbind(bracket_data, data.frame(
                      group1 = as.character(groups_with_data[1]),
                      group2 = as.character(groups_with_data[2]),
                      p.signif = sig_label,
                      x.position = cat_index,
                      y.position = y_pos,
                      stringsAsFactors = FALSE
                    ))
                  }
                }
              }
            } else if (n_groups_at_cat >= 3) {
              # THREE OR MORE GROUPS: Use ANOVA/Kruskal-Wallis + post-hoc tests
              cat("  Performing ANOVA/Kruskal-Wallis for", n_groups_at_cat, "groups at category", cat_val, "\\n")

              # Calculate x position (find position in all unique_categories, not just selected ones)
              cat_index <- which(unique_categories == cat_val)

              # Prepare data for ANOVA
              anova_data <- data.frame(
                group = factor(data_at_cat$group),
                value = as.numeric(data_at_cat$value)
              )
              anova_data <- anova_data[complete.cases(anova_data), ]

              if (nrow(anova_data) >= 3) {
                # Perform normality test for auto-selection
                normality_text <- c()
                all_normal <- TRUE

                for (grp in groups_with_data) {
                  grp_data <- data_at_cat[data_at_cat$group == grp, "value"]
                  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
                    shapiro_result <- tryCatch(
                      shapiro.test(grp_data),
                      error = function(e) NULL
                    )
                    if (!is.null(shapiro_result)) {
                      is_normal <- shapiro_result$p.value >= 0.05
                      norm_status <- if (is_normal) "normal" else "non-normal"
                      normality_text <- c(normality_text,
                        sprintf("  %s: p=%.4f (%s)", grp, shapiro_result$p.value, norm_status))
                      if (!is_normal) all_normal <- FALSE
                    }
                  }
                }

                # Auto-select test
                test_to_use <- statistical_test
                if (statistical_test == "auto") {
                  test_to_use <- if (all_normal) "anova" else "kruskal"
                  cat("  Auto-selected:", test_to_use, "(normality:", all_normal, ")\\n")
                }

                # Perform omnibus test
                omnibus_p <- NA
                test_name <- ""
                if (test_to_use == "kruskal") {
                  kruskal_result <- kruskal.test(value ~ group, data = anova_data)
                  omnibus_p <- kruskal_result$p.value
                  test_name <- "Kruskal-Wallis"
                } else {
                  anova_result <- aov(value ~ group, data = anova_data)
                  anova_summary <- summary(anova_result)
                  omnibus_p <- anova_summary[[1]][["Pr(>F)"]][1]
                  test_name <- "ANOVA"
                }

                cat("  ", test_name, "p-value:", omnibus_p, "\\n")

                # Build result text - ORDER: Normality, ANOVA, Post-hoc
                result_text <- sprintf("Category %s:", cat_val)
                if (length(normality_text) > 0) {
                  result_text <- paste0(result_text, "\\nNormality (Shapiro-Wilk):\\n", paste(normality_text, collapse="\\n"))
                }
                # Add ANOVA result with significance symbol
                anova_sig_label <- if (!is.na(omnibus_p)) {
                  if (omnibus_p < 0.001) "***" else if (omnibus_p < 0.01) "**" else if (omnibus_p < 0.05) "*" else "ns"
                } else "ns"
                result_text <- paste0(result_text, "\\n", test_name, " p=", sprintf("%.4f", omnibus_p), " (", anova_sig_label, ")")

                # If significant, perform post-hoc tests
                if (!is.na(omnibus_p) && omnibus_p < 0.05) {
                  cat("  Omnibus test significant, performing post-hoc tests\\n")

                  # Perform appropriate post-hoc test based on omnibus test type
                  posthoc_result <- NULL
                  posthoc_text <- c()

                  if (test_to_use == "kruskal") {
                    # For Kruskal-Wallis, use Dunn test
                    cat("  Performing Dunn post-hoc test...\\n")
                    posthoc_result <- tryCatch({
                      if (!requireNamespace("dunn.test", quietly = TRUE)) {
                        cat("Installing dunn.test package...\\n")
                        webr::install("dunn.test")
                      }
                      library(dunn.test)

                      dunn_result <- dunn.test(anova_data$value, anova_data$group, method = "bonferroni")

                      # Format results into a data frame
                      data.frame(
                        Comparison = dunn_result$comparisons,
                        P.adj = dunn_result$P.adjusted,
                        stringsAsFactors = FALSE
                      )
                    }, error = function(e) {
                      cat("  Dunn test error:", e$message, "\\n")
                      # Add error as annotation text for debugging
                      result_text <<- paste0(result_text, "\\n[DUNN ERROR: ", e$message, "]")
                      NULL
                    })

                    if (!is.null(posthoc_result)) {
                      cat(sprintf("  Dunn found %d pairwise comparisons\\n", nrow(posthoc_result)))
                      for (i in 1:nrow(posthoc_result)) {
                        comparison <- posthoc_result$Comparison[i]
                        p_adj <- posthoc_result$P.adj[i]

                        sig_label <- ""
                        if (!is.na(p_adj)) {
                          sig_label <- get_sig_symbol(p_adj)
                        } else {
                          sig_label <- "ns"
                        }

                        cat(sprintf("    %s: p=%.4f (%s)\\n", comparison, p_adj, sig_label))
                        posthoc_text <- c(posthoc_text, sprintf("  %s: p=%.4f (%s)",
                                                                comparison, ifelse(is.na(p_adj), 1.0, p_adj), sig_label))

                        # Add comparisons to bracket data based on comparison mode
                        should_add_posthoc <- FALSE
                        if (comparison_mode == "significant") {
                          should_add_posthoc <- (sig_label != "ns")
                        } else if (comparison_mode == "all") {
                          should_add_posthoc <- TRUE
                        } else if (comparison_mode == "custom") {
                          # Parse comparison to check if selected
                          comp_parts_temp <- strsplit(comparison, " - ")[[1]]
                          if (length(comp_parts_temp) == 2) {
                            comp_key1 <- paste0(comp_parts_temp[1], "-", comp_parts_temp[2], "@", cat_val)
                            comp_key2 <- paste0(comp_parts_temp[2], "-", comp_parts_temp[1], "@", cat_val)
                            should_add_posthoc <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                          }
                        }

                        if (should_add_posthoc) {
                          # Parse comparison string (format: "group1 - group2")
                          comp_parts <- strsplit(comparison, " - ")[[1]]
                          if (length(comp_parts) == 2) {
                            bracket_data <- rbind(bracket_data, data.frame(
                              group1 = comp_parts[1],
                              group2 = comp_parts[2],
                              p.signif = sig_label,
                              x.position = cat_index,
                              y.position = NA,  # Will be calculated later
                              stringsAsFactors = FALSE
                            ))
                          }
                        }
                      }
                    }
                  } else {
                    # For ANOVA, use selected post-hoc test
                    cat(sprintf("  Performing %s post-hoc test...\\n", selected_posthoc_test))
                    posthoc_result <- NULL
                    posthoc_summary <- NULL

                    if (selected_posthoc_test == "tukey") {
                      # Tukey HSD test
                      posthoc_result <- tryCatch(
                        TukeyHSD(anova_result),
                        error = function(e) {
                          cat("  Tukey error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        posthoc_summary <- posthoc_result$group
                      }
                    } else if (selected_posthoc_test == "bonferroni") {
                      # Bonferroni correction using pairwise t-tests
                      posthoc_result <- tryCatch(
                        pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "bonferroni"),
                        error = function(e) {
                          cat("  Bonferroni error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        # Convert to Tukey-like format
                        groups <- levels(anova_data$group)
                        comparisons <- combn(groups, 2, simplify = FALSE)
                        diff_values <- c()
                        p_adj_values <- c()
                        comparison_names <- c()

                        for (comp in comparisons) {
                          g1 <- comp[1]
                          g2 <- comp[2]
                          g1_idx <- which(rownames(posthoc_result$p.value) == g1)
                          g2_idx <- which(colnames(posthoc_result$p.value) == g2)

                          if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                            p_val <- posthoc_result$p.value[g1_idx, g2_idx]
                          } else {
                            g1_idx <- which(rownames(posthoc_result$p.value) == g2)
                            g2_idx <- which(colnames(posthoc_result$p.value) == g1)
                            p_val <- if (length(g1_idx) > 0 && length(g2_idx) > 0) posthoc_result$p.value[g1_idx, g2_idx] else NA
                          }

                          g1_values <- anova_data$value[anova_data$group == g1]
                          g2_values <- anova_data$value[anova_data$group == g2]
                          diff <- mean(g2_values, na.rm = TRUE) - mean(g1_values, na.rm = TRUE)

                          comparison_names <- c(comparison_names, paste0(g2, "-", g1))
                          diff_values <- c(diff_values, diff)
                          p_adj_values <- c(p_adj_values, p_val)
                        }

                        posthoc_summary <- data.frame(
                          diff = diff_values,
                          lwr = rep(NA, length(diff_values)),
                          upr = rep(NA, length(diff_values)),
                          "p adj" = p_adj_values,
                          row.names = comparison_names,
                          check.names = FALSE
                        )
                      }
                    } else if (selected_posthoc_test == "holm") {
                      # Holm correction using pairwise t-tests
                      posthoc_result <- tryCatch(
                        pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "holm"),
                        error = function(e) {
                          cat("  Holm error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        # Convert to Tukey-like format (same as Bonferroni)
                        groups <- levels(anova_data$group)
                        comparisons <- combn(groups, 2, simplify = FALSE)
                        diff_values <- c()
                        p_adj_values <- c()
                        comparison_names <- c()

                        for (comp in comparisons) {
                          g1 <- comp[1]
                          g2 <- comp[2]
                          g1_idx <- which(rownames(posthoc_result$p.value) == g1)
                          g2_idx <- which(colnames(posthoc_result$p.value) == g2)

                          if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                            p_val <- posthoc_result$p.value[g1_idx, g2_idx]
                          } else {
                            g1_idx <- which(rownames(posthoc_result$p.value) == g2)
                            g2_idx <- which(colnames(posthoc_result$p.value) == g1)
                            p_val <- if (length(g1_idx) > 0 && length(g2_idx) > 0) posthoc_result$p.value[g1_idx, g2_idx] else NA
                          }

                          g1_values <- anova_data$value[anova_data$group == g1]
                          g2_values <- anova_data$value[anova_data$group == g2]
                          diff <- mean(g2_values, na.rm = TRUE) - mean(g1_values, na.rm = TRUE)

                          comparison_names <- c(comparison_names, paste0(g2, "-", g1))
                          diff_values <- c(diff_values, diff)
                          p_adj_values <- c(p_adj_values, p_val)
                        }

                        posthoc_summary <- data.frame(
                          diff = diff_values,
                          lwr = rep(NA, length(diff_values)),
                          upr = rep(NA, length(diff_values)),
                          "p adj" = p_adj_values,
                          row.names = comparison_names,
                          check.names = FALSE
                        )
                      }
                    } else if (selected_posthoc_test == "dunnett") {
                      # Dunnett test (not typically used for grouped bar, default to Tukey)
                      cat("  Warning: Dunnett test not implemented for grouped bar, using Tukey\\n")
                      posthoc_result <- tryCatch(
                        TukeyHSD(anova_result),
                        error = function(e) {
                          cat("  Tukey error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        posthoc_summary <- posthoc_result$group
                      }
                    }

                    if (!is.null(posthoc_summary)) {
                      cat(sprintf("  Post-hoc found %d pairwise comparisons\\n", nrow(posthoc_summary)))

                      for (i in 1:nrow(posthoc_summary)) {
                        comparison <- rownames(posthoc_summary)[i]
                        p_adj <- posthoc_summary[i, "p adj"]
                        diff <- posthoc_summary[i, "diff"]

                        sig_label <- ""
                        if (!is.na(p_adj)) {
                          sig_label <- get_sig_symbol(p_adj)
                        } else {
                          sig_label <- "ns"
                        }

                        cat(sprintf("    %s: diff=%.2f, p=%.4f (%s)\\n", comparison, diff, p_adj, sig_label))
                        posthoc_text <- c(posthoc_text, sprintf("  %s: diff=%.2f, p=%.4f (%s)",
                                                                comparison, diff, ifelse(is.na(p_adj), 1.0, p_adj), sig_label))

                        # Add comparisons to bracket data based on comparison mode
                        should_add_posthoc <- FALSE
                        if (comparison_mode == "significant") {
                          should_add_posthoc <- (sig_label != "ns")
                        } else if (comparison_mode == "all") {
                          should_add_posthoc <- TRUE
                        } else if (comparison_mode == "custom") {
                          # Parse comparison to check if selected
                          comp_parts_temp <- strsplit(comparison, "-")[[1]]
                          if (length(comp_parts_temp) == 2) {
                            comp_key1 <- paste0(comp_parts_temp[1], "-", comp_parts_temp[2], "@", cat_val)
                            comp_key2 <- paste0(comp_parts_temp[2], "-", comp_parts_temp[1], "@", cat_val)
                            should_add_posthoc <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                          }
                        }

                        if (should_add_posthoc) {
                          # Parse comparison string (format: "group2-group1")
                          comp_parts <- strsplit(comparison, "-")[[1]]
                          if (length(comp_parts) == 2) {
                            bracket_data <- rbind(bracket_data, data.frame(
                              group1 = comp_parts[1],
                              group2 = comp_parts[2],
                              p.signif = sig_label,
                              x.position = cat_index,
                              y.position = NA,  # Will be calculated later
                              stringsAsFactors = FALSE
                            ))
                          }
                        }
                      }
                    }
                  }

                  # Combine omnibus and post-hoc results (sig_label already added above)
                  if (length(posthoc_text) > 0) {
                    # Add post-hoc test type header with correct test name
                    posthoc_header <- if (test_to_use == "kruskal") {
                      "Post-hoc pairwise comparisons (Dunn test):"
                    } else {
                      # For ANOVA, show the actual selected post-hoc test
                      # Note: Bonferroni and Holm are p-value adjustments, not post-hoc tests
                      if (selected_posthoc_test == "tukey") {
                        "Post-hoc pairwise comparisons (Tukey HSD):"
                      } else if (selected_posthoc_test == "bonferroni") {
                        "Pairwise t-test with Bonferroni correction:"
                      } else if (selected_posthoc_test == "holm") {
                        "Pairwise t-test with Holm correction:"
                      } else if (selected_posthoc_test == "dunnett") {
                        "Post-hoc pairwise comparisons (Dunnett):"
                      } else {
                        "Post-hoc pairwise comparisons (Tukey HSD):"
                      }
                    }
                    result_text <- paste0(result_text, "\\n", posthoc_header, "\\n", paste(posthoc_text, collapse="\\n"))
                  }

                  # Add to stat text results
                  stat_text_results <- c(stat_text_results, result_text)
                } else {
                  # Not significant - sig_label already added above
                  stat_text_results <- c(stat_text_results, result_text)
                }
              }
            }
          }
        }

        # Calculate Y positions for brackets using proper stacking logic (same as regular bar)
        if (nrow(bracket_data) > 0) {
          # Calculate Y range for unit_step
          y_range <- diff(range(plot_data$value, na.rm = TRUE))

          # Calculate unit_step (spacing between brackets) based on symbol size and Y range
          # This matches the logic from regular bar chart
          unit_step <- y_range * 0.08  # 8% of Y range as default spacing

          # Group brackets by category for proper stacking
          for (cat_idx in unique(bracket_data$x.position)) {
            na_positions <- which(bracket_data$x.position == cat_idx & is.na(bracket_data$y.position))

            if (length(na_positions) > 0) {
              cat_val <- unique_categories[cat_idx]

              # Stack brackets at this category
              for (j in seq_along(na_positions)) {
                bracket_idx <- na_positions[j]

                # Get the two groups being compared
                g1 <- bracket_data$group1[bracket_idx]
                g2 <- bracket_data$group2[bracket_idx]

                # Check for custom Y position first (both possible orderings)
                comp_key1 <- paste0(g1, "-", g2, "@", cat_val)
                comp_key2 <- paste0(g2, "-", g1, "@", cat_val)

                custom_y_found <- FALSE
                if (!is.null(custom_y_positions[[comp_key1]])) {
                  bracket_data$y.position[bracket_idx] <- as.numeric(custom_y_positions[[comp_key1]])
                  cat("Using custom Y position for", comp_key1, ":", custom_y_positions[[comp_key1]], "\\n")
                  custom_y_found <- TRUE
                } else if (!is.null(custom_y_positions[[comp_key2]])) {
                  bracket_data$y.position[bracket_idx] <- as.numeric(custom_y_positions[[comp_key2]])
                  cat("Using custom Y position for", comp_key2, ":", custom_y_positions[[comp_key2]], "\\n")
                  custom_y_found <- TRUE
                }

                # If custom position found, skip auto-calculation
                if (custom_y_found) {
                  next
                }

                # Calculate baseHeight for this specific comparison
                # baseHeight = max of (mean + error) for the two groups
                # Get all data for this category
                data_at_cat_all <- plot_data[plot_data$category == cat_val, ]
                g1_data <- data_at_cat_all[data_at_cat_all$group == g1, ]
                g2_data <- data_at_cat_all[data_at_cat_all$group == g2, ]

                if (nrow(g1_data) > 0 && nrow(g2_data) > 0) {
                  # Match regular bar's baseHeight calculation:
                  # baseHeight = max of (error_bar_top, max_data_point) for each group
                  # No error bars in box/violin plots
                  # No error bars in box/violin plots

                  # For box/violin plots, calculate max values from data
                  g1_max_val <- max(g1_data$value, na.rm = TRUE)
                  g2_max_val <- max(g2_data$value, na.rm = TRUE)
                  
                  g1_top <- g1_max_val  # Use max value directly
                  g2_top <- g2_max_val  # Use max value directly

                  baseHeight <- max(g1_top, g2_top, na.rm = TRUE)
                } else {
                  # Fallback: use max of all groups at this category
                  baseHeight <- max(data_at_cat_all$value, na.rm = TRUE)
                }

                # Calculate Y position with stacking - same for all scales
                # ggplot2 handles log transformation internally
                if (j == 1) {
                  # First bracket at this category (higher to avoid overlap with dots)
                  bracket_data$y.position[bracket_idx] <- baseHeight + (unit_step * 4.0)
                } else {
                  # Stack subsequent brackets (more spacing to avoid overlap)
                  prev_bracket_idx <- na_positions[j-1]
                  bracket_data$y.position[bracket_idx] <- max(
                    baseHeight + (unit_step * 5.0),
                    bracket_data$y.position[prev_bracket_idx] + (unit_step * 5.0)
                  )
                }
              }
            }
          }
        }

        # Add statistical brackets using stat_pvalue_manual (same as regular bar)
        if (nrow(bracket_data) > 0) {
          # Calculate X positions for each group within each category
          # IMPORTANT: Use only groups that actually have data, not all factor levels
          unique_groups <- levels(droplevels(plot_data$group))
          n_groups <- length(unique_groups)

          # Calculate xmin and xmax for each bracket
          for (i in 1:nrow(bracket_data)) {
            if (!is.na(bracket_data$group1[i]) && bracket_data$group1[i] != "" &&
                !is.na(bracket_data$group2[i]) && bracket_data$group2[i] != "") {

              cat_idx <- bracket_data$x.position[i]
              g1_idx <- which(unique_groups == bracket_data$group1[i])
              g2_idx <- which(unique_groups == bracket_data$group2[i])

              if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                # Calculate X positions to match the CENTER of each dodged bar
                # ggplot2's position_dodge formula for bar centers:
                # x = category_x + (index - 1 - (n-1)/2) * (dodge_width / n)
                # where index is 1-based

                # Calculate the center position of each bar
                g1_x <- cat_idx + (g1_idx - 1 - (n_groups - 1) / 2) * (dodge_width / n_groups)
                g2_x <- cat_idx + (g2_idx - 1 - (n_groups - 1) / 2) * (dodge_width / n_groups)

                # Store xmin/xmax for stat_pvalue_manual
                # IMPORTANT: Ensure xmin is leftmost and xmax is rightmost
                bracket_data$xmin[i] <- min(g1_x, g2_x)
                bracket_data$xmax[i] <- max(g1_x, g2_x)
              }
            }
          }

          # Store bracket data globally for educational R code export
          bracket_export_data <<- bracket_data

          # Get parameter values (same as regular bar)
          bracket_size <- symbol_size
          line_size <- ggpubr_line_size
          tip_length <- ggpubr_tip_length
          v_just <- ggpubr_vjust + 0.6

          # Split data by symbol type to apply different sizes (same as regular bar)
          asterisk_data <- bracket_data[bracket_data$p.signif != "ns", ]
          ns_data <- bracket_data[bracket_data$p.signif == "ns", ]

          # Add asterisk symbols with full size
          if (nrow(asterisk_data) > 0) {
            p <- p + stat_pvalue_manual(asterisk_data,
                                       label = "p.signif",
                                       xmin = "xmin",
                                       xmax = "xmax",
                                       y.position = "y.position",
                                       size = bracket_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = v_just,
                                       hjust = 0.5,
                                       step.increase = 0,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
          }

          # Add n.s. symbols with smaller size and higher position
          if (nrow(ns_data) > 0) {
            smaller_size <- bracket_size * 0.7
            ns_vjust <- v_just - 0.5
            p <- p + stat_pvalue_manual(ns_data,
                                       label = "p.signif",
                                       xmin = "xmin",
                                       xmax = "xmax",
                                       y.position = "y.position",
                                       size = smaller_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = ns_vjust,
                                       hjust = 0.5,
                                       step.increase = 0,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
          }
        }

        # Store statistical results in global variable for JavaScript to retrieve
        if (length(stat_text_results) > 0) {
          grouped_bar_stat_results <<- paste(stat_text_results, collapse="\\n\\n")
          cat("\\n=== STATISTICAL RESULTS SUMMARY ===\\n")
          cat(grouped_bar_stat_results, "\\n")
          cat("===================================\\n")
          cat("DEBUG: Stored", nchar(grouped_bar_stat_results), "characters in grouped_bar_stat_results\\n")
        } else {
          cat("DEBUG: No stat_text_results to store\\n")
          grouped_bar_stat_results <<- ""
        }
        }, error = function(e) {
          grouped_bar_stat_results <<- paste(grouped_bar_stat_results, "\\nERROR in statistical analysis:", e$message)
        })
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped box plot function - for data with Group/Category columns
    sato_box_grouped <- function(dat, group_col=1, x_col=2, y_col=3,
                                fill_colors=c("#4C78A8", "#E15759"), stroke_color="#1f2937",
                                alpha=0.9, linewidth=0.7, width=0.7, dodge_width=0.9,
                                group_name="Group", category_name="Treatment", value_name="Value",
                                target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                title_text="Grouped Box Plot", x_text="Category", y_text="Value",
                                show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                x_scale="linear", y_scale="linear",
                                theme_name="minimal",
                                x_axis_rotation=0, y_axis_rotation=0,
                                x_axis_hjust=0.5, x_axis_vjust=0.5,
                                y_axis_hjust=0.5, y_axis_vjust=0.5,
                                add_statistics=FALSE, statistical_test="auto", variance_test="levene", symbol_size=8,
                                ggpubr_line_size=1.0, ggpubr_tip_length=0.04, ggpubr_vjust=-0.3,
                                comparison_mode="all", custom_comparisons="[]", custom_positions="{}",
                                selected_posthoc_test="tukey", dunnett_control="", sato_symbol_size=7) {
      
      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)
      
      if (ncol(dat) < 3) {
        stop("Grouped box chart requires at least 3 columns: Group, Category, Value")
      }
      
      # 列名検出
      actual_group_name <- if(!is.null(names(dat)) && length(names(dat)) >= group_col && !grepl("^V[0-9]+$", names(dat)[group_col])) {
        names(dat)[group_col]
      } else {
        group_name
      }
      
      # データ準備とグループ順序設定
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]
      
      if ("Normal" %in% unique(as.character(plot_data$group)) && "Tumor" %in% unique(as.character(plot_data$group))) {
        plot_data$group <- factor(plot_data$group, levels = c("Normal", "Tumor"))
      }
      
      # グループ数に応じて色を設定
      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)
      
      if (length(fill_colors) < n_groups) {
        fill_colors <- rep(fill_colors, length.out = n_groups)
      }
      
      # グループ化されたボックスプロット作成
      p <- ggplot(plot_data, aes(x = category, y = value, fill = group)) +
           geom_boxplot(position = position_dodge(width = dodge_width), 
                       width = width,
                       alpha = alpha, 
                       linewidth = linewidth,
                       color = stroke_color,
                       outlier.shape = 16) +
           scale_fill_manual(values = setNames(fill_colors[1:n_groups], group_levels)) +
           labs(fill = actual_group_name)
      
      # Add statistical analysis if requested
      if (add_statistics && ncol(dat) >= 3) {
        p <- sato_add_statistics_to_plot(p, plot_data, "group", "value", statistical_test, sato_symbol_size, posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale=y_scale, stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped box plot with individual data points
    sato_box_grouped_dot <- function(dat, group_col=1, x_col=2, y_col=3,
                                    fill_colors=c("#4C78A8", "#E15759"), stroke_color="#1f2937",
                                    alpha=0.9, linewidth=0.7, width=0.7, dodge_width=0.9,
                                    dot_size=4, dot_alpha=1.0, dot_color="#333333", dot_shape=16, jitter_width=0.15,
                                    group_name="Group", category_name="Treatment", value_name="Value",
                                    target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                    title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                    title_text="Grouped Box Plot with Dots", x_text="Category", y_text="Value",
                                    show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                    x_scale="linear", y_scale="linear",
                                    theme_name="minimal",
                                    x_axis_rotation=0, y_axis_rotation=0,
                                    x_axis_hjust=0.5, x_axis_vjust=0.5,
                                    y_axis_hjust=0.5, y_axis_vjust=0.5,
                                    add_statistics=FALSE, statistical_test="auto", variance_test="levene", symbol_size=8,
                                    ggpubr_line_size=1.0, ggpubr_tip_length=0.04, ggpubr_vjust=-0.3,
                                    comparison_mode="all", custom_comparisons="[]", custom_positions="{}",
                                    selected_posthoc_test="tukey", dunnett_control="", sato_symbol_size=7,
                                    stat_symbol_type="stars", custom_symbol_05="*", custom_symbol_01="**", custom_symbol_001="***", custom_symbol_ns="ns") {

      # Helper function to generate significance symbol based on p-value and symbol type
      get_sig_symbol <- function(p_val) {
        if (stat_symbol_type == "pvalue") {
          return(sprintf("p=%.3f", p_val))
        } else if (stat_symbol_type == "custom") {
          if (p_val < 0.001) return(custom_symbol_001)
          else if (p_val < 0.01) return(custom_symbol_01)
          else if (p_val < 0.05) return(custom_symbol_05)
          else return(custom_symbol_ns)
        } else {
          # Default to stars
          if (p_val < 0.001) return("***")
          else if (p_val < 0.01) return("**")
          else if (p_val < 0.05) return("*")
          else return("ns")
        }
      }

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)
      
      if (ncol(dat) < 3) {
        stop("Grouped box chart requires at least 3 columns: Group, Category, Value")
      }
      
      # 列名検出
      actual_group_name <- if(!is.null(names(dat)) && length(names(dat)) >= group_col && !grepl("^V[0-9]+$", names(dat)[group_col])) {
        names(dat)[group_col]
      } else {
        group_name
      }
      
      # データ準備とグループ順序設定
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]
      
      if ("Normal" %in% unique(as.character(plot_data$group)) && "Tumor" %in% unique(as.character(plot_data$group))) {
        plot_data$group <- factor(plot_data$group, levels = c("Normal", "Tumor"))
      }
      
      # グループ数に応じて色を設定
      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)
      
      if (length(fill_colors) < n_groups) {
        fill_colors <- rep(fill_colors, length.out = n_groups)
      }
      
      # グループ化されたボックスプロット + dots作成
      p <- ggplot(plot_data, aes(x = category, y = value, fill = group)) +
           geom_boxplot(position = position_dodge(width = dodge_width),
                       width = width,
                       alpha = alpha,
                       linewidth = linewidth,
                       color = stroke_color,
                       outlier.shape = NA) +  # Hide default outliers to avoid duplication
           geom_point(aes(fill = group),
                     position = position_jitterdodge(dodge.width = dodge_width, jitter.width = jitter_width, jitter.height = 0),
                     size = dot_size, alpha = dot_alpha, shape = dot_shape, color = dot_color) +
           scale_fill_manual(values = setNames(fill_colors[1:n_groups], group_levels)) +
           labs(fill = actual_group_name)

      # Initialize global variable for statistical results
      grouped_bar_stat_results <<- ""

      # Add statistical analysis if requested - compare groups within each category
      if (add_statistics && n_groups >= 2) {
        tryCatch({
          # Load ggpubr package for stat_pvalue_manual
          if (!require("ggpubr", quietly = TRUE)) {
            webr::install("ggpubr")
          }
          library(ggpubr)

          # Get unique categories (filter out any categories with no data)
          all_categories <- levels(plot_data$category)
          unique_categories <- all_categories[all_categories %in% unique(plot_data$category)]

          # Handle custom comparisons
          selected_categories <- NULL
          custom_y_positions <- list()
          custom_comps <- c()  # Initialize to empty array

          if (comparison_mode == "custom") {
            # Parse custom comparisons: format is "Group1-Group2@CategoryA", "Group1-Group2@CategoryB", etc.
            custom_comps <- tryCatch(jsonlite::fromJSON(custom_comparisons), error = function(e) c())
            cat("Custom comparisons:", custom_comparisons, "\\n")
            cat("Custom positions JSON:", custom_positions, "\\n")

            if (length(custom_comps) > 0) {
              # Extract category values from custom comparisons
              selected_categories <- c()
              for (comp_str in custom_comps) {
                # Parse format: "Group1-Group2@Category"
                parts <- strsplit(comp_str, "@", fixed = TRUE)[[1]]
                if (length(parts) == 2) {
                  cat_val <- parts[2]
                  selected_categories <- c(selected_categories, cat_val)
                }
              }
              cat("Selected categories from custom comparisons:", paste(selected_categories, collapse=", "), "\\n")

              # Parse custom positions
              custom_pos_list <- tryCatch(jsonlite::fromJSON(custom_positions), error = function(e) list())
              if (length(custom_pos_list) > 0) {
                for (comp_str in names(custom_pos_list)) {
                  custom_y_positions[[comp_str]] <- as.numeric(custom_pos_list[[comp_str]])
                }
              }
            }
          }

          # Determine which categories to test
          if (is.null(selected_categories) || length(selected_categories) == 0) {
            # Test all categories (default or "all" mode)
            categories_to_test <- unique_categories
          } else {
            # Test only selected categories (deduplicate to avoid testing same category multiple times)
            categories_to_test <- unique(selected_categories)
          }

          # Debug: Check what we're testing
          cat("DEBUG: comparison_mode =", comparison_mode, "\\n")
          cat("DEBUG: categories_to_test =", paste(categories_to_test, collapse=", "), "\\n")
          cat("DEBUG: Number of categories to test:", length(categories_to_test), "\\n")

          stat_text_results <- c()

        # Data frame to store ggpubr bracket information
        bracket_data <- data.frame(
          group1 = character(),
          group2 = character(),
          p.signif = character(),
          x.position = numeric(),
          y.position = numeric(),
          stringsAsFactors = FALSE
        )

        # For each category, perform statistical tests between groups
        for (cat_idx in seq_along(categories_to_test)) {
          cat_val <- categories_to_test[cat_idx]

          # Get data for this category
          data_at_cat <- plot_data[plot_data$category == cat_val, ]

          if (nrow(data_at_cat) > 0) {
            # Get groups with data at this category
            groups_with_data <- unique(data_at_cat$group[!is.na(data_at_cat$value)])
            n_groups_at_cat <- length(groups_with_data)

            if (n_groups_at_cat == 2) {
              # TWO GROUPS: Use t-test or Wilcoxon
              group1_data <- data_at_cat[data_at_cat$group == groups_with_data[1], "value"]
              group2_data <- data_at_cat[data_at_cat$group == groups_with_data[2], "value"]

              # Only test if both groups have data
              if (length(group1_data) > 0 && length(group2_data) > 0) {
                # Perform normality test for both groups
                normality_text <- c()
                is_group1_normal <- TRUE
                is_group2_normal <- TRUE

                if (length(group1_data) >= 3 && length(group1_data) <= 5000) {
                  shapiro_result1 <- tryCatch(
                    shapiro.test(group1_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result1)) {
                    is_group1_normal <- shapiro_result1$p.value >= 0.05
                    norm_status <- if (is_group1_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[1], shapiro_result1$p.value, norm_status))
                  }
                }

                if (length(group2_data) >= 3 && length(group2_data) <= 5000) {
                  shapiro_result2 <- tryCatch(
                    shapiro.test(group2_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result2)) {
                    is_group2_normal <- shapiro_result2$p.value >= 0.05
                    norm_status <- if (is_group2_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[2], shapiro_result2$p.value, norm_status))
                  }
                }

                both_normal <- is_group1_normal && is_group2_normal

                # Select test based on mode
                test_to_use <- statistical_test
                if (statistical_test == "auto") {
                  if (both_normal) {
                    test_to_use <- "t-test"
                    cat("  Auto-selected: t-test (both groups normal)\\n")
                  } else {
                    test_to_use <- "wilcoxon"
                    cat("  Auto-selected: Wilcoxon test (non-normal data detected)\\n")
                  }
                } else if (statistical_test == "parametric") {
                  test_to_use <- "t-test"
                  cat("  Manual mode: parametric t-test\\n")
                } else if (statistical_test == "nonparametric") {
                  test_to_use <- "wilcoxon"
                  cat("  Manual mode: non-parametric Wilcoxon test\\n")
                }

                # Perform variance test (only for parametric tests)
                variance_text <- ""
                equal_variances <- TRUE
                if (test_to_use != "wilcoxon") {
                  if (variance_test == "levene") {
                    # Levene test for equality of variances
                    combined_data <- data.frame(
                      values = c(group1_data, group2_data),
                      group = factor(c(rep(groups_with_data[1], length(group1_data)),
                                      rep(groups_with_data[2], length(group2_data))))
                    )

                    # Calculate Levene test manually
                    group_means <- tapply(combined_data$values, combined_data$group, mean)
                    abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
                    levene_result <- tryCatch(
                      anova(lm(abs_deviations ~ combined_data$group)),
                      error = function(e) NULL
                    )

                    if (!is.null(levene_result)) {
                      levene_p <- levene_result$\`Pr(>F)\`[1]
                      equal_variances <- levene_p > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, Levene)", levene_p, variance_status)
                      cat(sprintf("    Levene test: p=%.4f (%s)\\n", levene_p, variance_status))
                    }
                  } else {
                    # F-test for equality of variances
                    var_test <- tryCatch(
                      var.test(group1_data, group2_data),
                      error = function(e) NULL
                    )

                    if (!is.null(var_test)) {
                      equal_variances <- var_test$p.value > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, F-test)", var_test$p.value, variance_status)
                      cat(sprintf("    F-test: p=%.4f (%s)\\n", var_test$p.value, variance_status))
                    }
                  }
                }

                # Perform statistical test based on auto-selection
                if (test_to_use == "wilcoxon") {
                  test_result <- tryCatch(
                    wilcox.test(group1_data, group2_data),
                    error = function(e) NULL
                  )
                  test_name <- "Wilcoxon"
                } else {
                  # Use appropriate t-test based on variance equality
                  test_result <- tryCatch(
                    t.test(group1_data, group2_data, var.equal = equal_variances),
                    error = function(e) NULL
                  )
                  test_name <- if (equal_variances) "Student's t-test" else "Welch's t-test"
                }

                if (!is.null(test_result)) {
                  p_val <- test_result$p.value
                  cat(sprintf("  2-group test: %s vs %s, %s p=%.4f\\n",
                              groups_with_data[1], groups_with_data[2], test_name, p_val))

                  # Determine significance symbol using helper function
                  sig_label <- get_sig_symbol(p_val)

                  # Store result text for UI with proper order: Category -> Group Stats -> Normality -> Variance -> Test result
                  result_text <- sprintf("Category %s:", cat_val)

                  # Add group statistics (n, mean, sd) for each group
                  result_text <- paste0(result_text, "\\nGroup Statistics:")
                  for (g in groups_with_data) {
                    g_data <- plot_data[plot_data$category == cat_val & plot_data$group == g, "value"]
                    result_text <- paste0(result_text, sprintf("\\n  %s: n=%d, mean=%.2f, sd=%.2f",
                                          g, length(g_data), mean(g_data, na.rm=TRUE), sd(g_data, na.rm=TRUE)))
                  }

                  # Add normality testing results
                  if (length(normality_text) > 0) {
                    result_text <- paste0(result_text, "\\n\\nNormality Testing (Shapiro-Wilk):\\n  ", paste(normality_text, collapse="\\n  "))
                  }

                  # Add variance test results
                  if (nchar(variance_text) > 0) {
                    result_text <- paste0(result_text, "\\n", variance_text)
                  }

                  # Add main test result
                  result_text <- paste0(result_text, "\\n\\n", sprintf("%s: p=%.4f (%s)",
                                        test_name, p_val, sig_label))
                  stat_text_results <- c(stat_text_results, result_text)

                  # Add ggpubr bracket for this comparison
                  # Calculate x position (find position in all unique_categories, not just selected ones)
                  cat_index <- which(unique_categories == cat_val)

                  # Calculate y position (above the highest bar/point in this category)
                  # Check if custom position is provided (check both possible orderings)
                  comp_key1 <- paste0(groups_with_data[1], "-", groups_with_data[2], "@", cat_val)
                  comp_key2 <- paste0(groups_with_data[2], "-", groups_with_data[1], "@", cat_val)

                  custom_y_pos <- NULL
                  if (!is.null(custom_y_positions[[comp_key1]])) {
                    custom_y_pos <- custom_y_positions[[comp_key1]]
                    cat("Using custom Y position for", comp_key1, ":", custom_y_pos, "\\n")
                  } else if (!is.null(custom_y_positions[[comp_key2]])) {
                    custom_y_pos <- custom_y_positions[[comp_key2]]
                    cat("Using custom Y position for", comp_key2, ":", custom_y_pos, "\\n")
                  }

                  if (!is.null(custom_y_pos)) {
                    # Use custom Y position
                    y_pos <- custom_y_pos
                  } else {
                    # Calculate default position
                    # Get the summary data for this category to find the bar heights
                    data_at_cat_all <- plot_data[plot_data$category == cat_val, ]
                    max_value <- max(data_at_cat_all$value, na.rm = TRUE)
                    # No error bars in box plots

                    # Position symbol above the tallest bar + error bar with some spacing
                    # Increased spacing to avoid overlapping with dots (0.055 -> 0.15)
                    y_pos <- max_value + (max_value * 0.20)  # 20% above highest value
                  }

                  # Add to bracket data based on comparison mode
                  should_add_bracket <- FALSE
                  if (comparison_mode == "significant") {
                    # Only add if significant
                    should_add_bracket <- (sig_label != "ns")
                  } else if (comparison_mode == "all") {
                    # Add all comparisons
                    should_add_bracket <- TRUE
                  } else if (comparison_mode == "custom") {
                    # Only add if this comparison is selected
                    # Check if this comparison is in the custom selections
                    comp_key1 <- paste0(groups_with_data[1], "-", groups_with_data[2], "@", cat_val)
                    comp_key2 <- paste0(groups_with_data[2], "-", groups_with_data[1], "@", cat_val)
                    should_add_bracket <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                  }

                  if (should_add_bracket) {
                    bracket_data <- rbind(bracket_data, data.frame(
                      group1 = as.character(groups_with_data[1]),
                      group2 = as.character(groups_with_data[2]),
                      p.signif = sig_label,
                      x.position = cat_index,
                      y.position = y_pos,
                      stringsAsFactors = FALSE
                    ))
                  }
                }
              }
            } else if (n_groups_at_cat >= 3) {
              # THREE OR MORE GROUPS: Use ANOVA/Kruskal-Wallis + post-hoc tests
              cat("  Performing ANOVA/Kruskal-Wallis for", n_groups_at_cat, "groups at category", cat_val, "\\n")

              # Calculate x position (find position in all unique_categories, not just selected ones)
              cat_index <- which(unique_categories == cat_val)

              # Prepare data for ANOVA
              anova_data <- data.frame(
                group = factor(data_at_cat$group),
                value = as.numeric(data_at_cat$value)
              )
              anova_data <- anova_data[complete.cases(anova_data), ]

              if (nrow(anova_data) >= 3) {
                # Perform normality test for auto-selection
                normality_text <- c()
                all_normal <- TRUE

                for (grp in groups_with_data) {
                  grp_data <- data_at_cat[data_at_cat$group == grp, "value"]
                  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
                    shapiro_result <- tryCatch(
                      shapiro.test(grp_data),
                      error = function(e) NULL
                    )
                    if (!is.null(shapiro_result)) {
                      is_normal <- shapiro_result$p.value >= 0.05
                      norm_status <- if (is_normal) "normal" else "non-normal"
                      normality_text <- c(normality_text,
                        sprintf("  %s: p=%.4f (%s)", grp, shapiro_result$p.value, norm_status))
                      if (!is_normal) all_normal <- FALSE
                    }
                  }
                }

                # Auto-select test
                test_to_use <- statistical_test
                if (statistical_test == "auto") {
                  test_to_use <- if (all_normal) "anova" else "kruskal"
                  cat("  Auto-selected:", test_to_use, "(normality:", all_normal, ")\\n")
                }

                # Perform omnibus test
                omnibus_p <- NA
                test_name <- ""
                if (test_to_use == "kruskal") {
                  kruskal_result <- kruskal.test(value ~ group, data = anova_data)
                  omnibus_p <- kruskal_result$p.value
                  test_name <- "Kruskal-Wallis"
                } else {
                  anova_result <- aov(value ~ group, data = anova_data)
                  anova_summary <- summary(anova_result)
                  omnibus_p <- anova_summary[[1]][["Pr(>F)"]][1]
                  test_name <- "ANOVA"
                }

                cat("  ", test_name, "p-value:", omnibus_p, "\\n")

                # Build result text - ORDER: Normality, ANOVA, Post-hoc
                result_text <- sprintf("Category %s:", cat_val)
                if (length(normality_text) > 0) {
                  result_text <- paste0(result_text, "\\nNormality (Shapiro-Wilk):\\n", paste(normality_text, collapse="\\n"))
                }
                # Add ANOVA result with significance symbol
                anova_sig_label <- if (!is.na(omnibus_p)) {
                  if (omnibus_p < 0.001) "***" else if (omnibus_p < 0.01) "**" else if (omnibus_p < 0.05) "*" else "ns"
                } else "ns"
                result_text <- paste0(result_text, "\\n", test_name, " p=", sprintf("%.4f", omnibus_p), " (", anova_sig_label, ")")

                # If significant, perform post-hoc tests
                if (!is.na(omnibus_p) && omnibus_p < 0.05) {
                  cat("  Omnibus test significant, performing post-hoc tests\\n")

                  # Perform appropriate post-hoc test based on omnibus test type
                  posthoc_result <- NULL
                  posthoc_text <- c()

                  if (test_to_use == "kruskal") {
                    # For Kruskal-Wallis, use Dunn test
                    cat("  Performing Dunn post-hoc test...\\n")
                    posthoc_result <- tryCatch({
                      if (!requireNamespace("dunn.test", quietly = TRUE)) {
                        cat("Installing dunn.test package...\\n")
                        webr::install("dunn.test")
                      }
                      library(dunn.test)

                      dunn_result <- dunn.test(anova_data$value, anova_data$group, method = "bonferroni")

                      # Format results into a data frame
                      data.frame(
                        Comparison = dunn_result$comparisons,
                        P.adj = dunn_result$P.adjusted,
                        stringsAsFactors = FALSE
                      )
                    }, error = function(e) {
                      cat("  Dunn test error:", e$message, "\\n")
                      # Add error as annotation text for debugging
                      result_text <<- paste0(result_text, "\\n[DUNN ERROR: ", e$message, "]")
                      NULL
                    })

                    if (!is.null(posthoc_result)) {
                      cat(sprintf("  Dunn found %d pairwise comparisons\\n", nrow(posthoc_result)))
                      for (i in 1:nrow(posthoc_result)) {
                        comparison <- posthoc_result$Comparison[i]
                        p_adj <- posthoc_result$P.adj[i]

                        sig_label <- ""
                        if (!is.na(p_adj)) {
                          sig_label <- get_sig_symbol(p_adj)
                        } else {
                          sig_label <- "ns"
                        }

                        cat(sprintf("    %s: p=%.4f (%s)\\n", comparison, p_adj, sig_label))
                        posthoc_text <- c(posthoc_text, sprintf("  %s: p=%.4f (%s)",
                                                                comparison, ifelse(is.na(p_adj), 1.0, p_adj), sig_label))

                        # Add comparisons to bracket data based on comparison mode
                        should_add_posthoc <- FALSE
                        if (comparison_mode == "significant") {
                          should_add_posthoc <- (sig_label != "ns")
                        } else if (comparison_mode == "all") {
                          should_add_posthoc <- TRUE
                        } else if (comparison_mode == "custom") {
                          # Parse comparison to check if selected
                          comp_parts_temp <- strsplit(comparison, " - ")[[1]]
                          if (length(comp_parts_temp) == 2) {
                            comp_key1 <- paste0(comp_parts_temp[1], "-", comp_parts_temp[2], "@", cat_val)
                            comp_key2 <- paste0(comp_parts_temp[2], "-", comp_parts_temp[1], "@", cat_val)
                            should_add_posthoc <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                          }
                        }

                        if (should_add_posthoc) {
                          # Parse comparison string (format: "group1 - group2")
                          comp_parts <- strsplit(comparison, " - ")[[1]]
                          if (length(comp_parts) == 2) {
                            bracket_data <- rbind(bracket_data, data.frame(
                              group1 = comp_parts[1],
                              group2 = comp_parts[2],
                              p.signif = sig_label,
                              x.position = cat_index,
                              y.position = NA,  # Will be calculated later
                              stringsAsFactors = FALSE
                            ))
                          }
                        }
                      }
                    }
                  } else {
                    # For ANOVA, use selected post-hoc test
                    cat(sprintf("  Performing %s post-hoc test...\\n", selected_posthoc_test))
                    posthoc_result <- NULL
                    posthoc_summary <- NULL

                    if (selected_posthoc_test == "tukey") {
                      # Tukey HSD test
                      posthoc_result <- tryCatch(
                        TukeyHSD(anova_result),
                        error = function(e) {
                          cat("  Tukey error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        posthoc_summary <- posthoc_result$group
                      }
                    } else if (selected_posthoc_test == "bonferroni") {
                      # Bonferroni correction using pairwise t-tests
                      posthoc_result <- tryCatch(
                        pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "bonferroni"),
                        error = function(e) {
                          cat("  Bonferroni error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        # Convert to Tukey-like format
                        groups <- levels(anova_data$group)
                        comparisons <- combn(groups, 2, simplify = FALSE)
                        diff_values <- c()
                        p_adj_values <- c()
                        comparison_names <- c()

                        for (comp in comparisons) {
                          g1 <- comp[1]
                          g2 <- comp[2]
                          g1_idx <- which(rownames(posthoc_result$p.value) == g1)
                          g2_idx <- which(colnames(posthoc_result$p.value) == g2)

                          if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                            p_val <- posthoc_result$p.value[g1_idx, g2_idx]
                          } else {
                            g1_idx <- which(rownames(posthoc_result$p.value) == g2)
                            g2_idx <- which(colnames(posthoc_result$p.value) == g1)
                            p_val <- if (length(g1_idx) > 0 && length(g2_idx) > 0) posthoc_result$p.value[g1_idx, g2_idx] else NA
                          }

                          g1_values <- anova_data$value[anova_data$group == g1]
                          g2_values <- anova_data$value[anova_data$group == g2]
                          diff <- mean(g2_values, na.rm = TRUE) - mean(g1_values, na.rm = TRUE)

                          comparison_names <- c(comparison_names, paste0(g2, "-", g1))
                          diff_values <- c(diff_values, diff)
                          p_adj_values <- c(p_adj_values, p_val)
                        }

                        posthoc_summary <- data.frame(
                          diff = diff_values,
                          lwr = rep(NA, length(diff_values)),
                          upr = rep(NA, length(diff_values)),
                          "p adj" = p_adj_values,
                          row.names = comparison_names,
                          check.names = FALSE
                        )
                      }
                    } else if (selected_posthoc_test == "holm") {
                      # Holm correction using pairwise t-tests
                      posthoc_result <- tryCatch(
                        pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "holm"),
                        error = function(e) {
                          cat("  Holm error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        # Convert to Tukey-like format (same as Bonferroni)
                        groups <- levels(anova_data$group)
                        comparisons <- combn(groups, 2, simplify = FALSE)
                        diff_values <- c()
                        p_adj_values <- c()
                        comparison_names <- c()

                        for (comp in comparisons) {
                          g1 <- comp[1]
                          g2 <- comp[2]
                          g1_idx <- which(rownames(posthoc_result$p.value) == g1)
                          g2_idx <- which(colnames(posthoc_result$p.value) == g2)

                          if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                            p_val <- posthoc_result$p.value[g1_idx, g2_idx]
                          } else {
                            g1_idx <- which(rownames(posthoc_result$p.value) == g2)
                            g2_idx <- which(colnames(posthoc_result$p.value) == g1)
                            p_val <- if (length(g1_idx) > 0 && length(g2_idx) > 0) posthoc_result$p.value[g1_idx, g2_idx] else NA
                          }

                          g1_values <- anova_data$value[anova_data$group == g1]
                          g2_values <- anova_data$value[anova_data$group == g2]
                          diff <- mean(g2_values, na.rm = TRUE) - mean(g1_values, na.rm = TRUE)

                          comparison_names <- c(comparison_names, paste0(g2, "-", g1))
                          diff_values <- c(diff_values, diff)
                          p_adj_values <- c(p_adj_values, p_val)
                        }

                        posthoc_summary <- data.frame(
                          diff = diff_values,
                          lwr = rep(NA, length(diff_values)),
                          upr = rep(NA, length(diff_values)),
                          "p adj" = p_adj_values,
                          row.names = comparison_names,
                          check.names = FALSE
                        )
                      }
                    } else if (selected_posthoc_test == "dunnett") {
                      # Dunnett test (not typically used for grouped bar, default to Tukey)
                      cat("  Warning: Dunnett test not implemented for grouped bar, using Tukey\\n")
                      posthoc_result <- tryCatch(
                        TukeyHSD(anova_result),
                        error = function(e) {
                          cat("  Tukey error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        posthoc_summary <- posthoc_result$group
                      }
                    }

                    if (!is.null(posthoc_summary)) {
                      cat(sprintf("  Post-hoc found %d pairwise comparisons\\n", nrow(posthoc_summary)))

                      for (i in 1:nrow(posthoc_summary)) {
                        comparison <- rownames(posthoc_summary)[i]
                        p_adj <- posthoc_summary[i, "p adj"]
                        diff <- posthoc_summary[i, "diff"]

                        sig_label <- ""
                        if (!is.na(p_adj)) {
                          sig_label <- get_sig_symbol(p_adj)
                        } else {
                          sig_label <- "ns"
                        }

                        cat(sprintf("    %s: diff=%.2f, p=%.4f (%s)\\n", comparison, diff, p_adj, sig_label))
                        posthoc_text <- c(posthoc_text, sprintf("  %s: diff=%.2f, p=%.4f (%s)",
                                                                comparison, diff, ifelse(is.na(p_adj), 1.0, p_adj), sig_label))

                        # Add comparisons to bracket data based on comparison mode
                        should_add_posthoc <- FALSE
                        if (comparison_mode == "significant") {
                          should_add_posthoc <- (sig_label != "ns")
                        } else if (comparison_mode == "all") {
                          should_add_posthoc <- TRUE
                        } else if (comparison_mode == "custom") {
                          # Parse comparison to check if selected
                          comp_parts_temp <- strsplit(comparison, "-")[[1]]
                          if (length(comp_parts_temp) == 2) {
                            comp_key1 <- paste0(comp_parts_temp[1], "-", comp_parts_temp[2], "@", cat_val)
                            comp_key2 <- paste0(comp_parts_temp[2], "-", comp_parts_temp[1], "@", cat_val)
                            should_add_posthoc <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                          }
                        }

                        if (should_add_posthoc) {
                          # Parse comparison string (format: "group2-group1")
                          comp_parts <- strsplit(comparison, "-")[[1]]
                          if (length(comp_parts) == 2) {
                            bracket_data <- rbind(bracket_data, data.frame(
                              group1 = comp_parts[1],
                              group2 = comp_parts[2],
                              p.signif = sig_label,
                              x.position = cat_index,
                              y.position = NA,  # Will be calculated later
                              stringsAsFactors = FALSE
                            ))
                          }
                        }
                      }
                    }
                  }

                  # Combine omnibus and post-hoc results (sig_label already added above)
                  if (length(posthoc_text) > 0) {
                    # Add post-hoc test type header with correct test name
                    posthoc_header <- if (test_to_use == "kruskal") {
                      "Post-hoc pairwise comparisons (Dunn test):"
                    } else {
                      # For ANOVA, show the actual selected post-hoc test
                      # Note: Bonferroni and Holm are p-value adjustments, not post-hoc tests
                      if (selected_posthoc_test == "tukey") {
                        "Post-hoc pairwise comparisons (Tukey HSD):"
                      } else if (selected_posthoc_test == "bonferroni") {
                        "Pairwise t-test with Bonferroni correction:"
                      } else if (selected_posthoc_test == "holm") {
                        "Pairwise t-test with Holm correction:"
                      } else if (selected_posthoc_test == "dunnett") {
                        "Post-hoc pairwise comparisons (Dunnett):"
                      } else {
                        "Post-hoc pairwise comparisons (Tukey HSD):"
                      }
                    }
                    result_text <- paste0(result_text, "\\n", posthoc_header, "\\n", paste(posthoc_text, collapse="\\n"))
                  }

                  # Add to stat text results
                  stat_text_results <- c(stat_text_results, result_text)
                } else {
                  # Not significant - sig_label already added above
                  stat_text_results <- c(stat_text_results, result_text)
                }
              }
            }
          }
        }

        # Calculate Y positions for brackets using proper stacking logic (same as regular bar)
        if (nrow(bracket_data) > 0) {
          # Calculate Y range for unit_step
          y_range <- diff(range(plot_data$value,
                                 na.rm = TRUE))

          # Calculate unit_step (spacing between brackets) based on symbol size and Y range
          # This matches the logic from regular bar chart
          unit_step <- y_range * 0.08  # 8% of Y range as default spacing

          # Group brackets by category for proper stacking
          for (cat_idx in unique(bracket_data$x.position)) {
            na_positions <- which(bracket_data$x.position == cat_idx & is.na(bracket_data$y.position))

            if (length(na_positions) > 0) {
              cat_val <- unique_categories[cat_idx]

              # Stack brackets at this category
              for (j in seq_along(na_positions)) {
                bracket_idx <- na_positions[j]

                # Get the two groups being compared
                g1 <- bracket_data$group1[bracket_idx]
                g2 <- bracket_data$group2[bracket_idx]

                # Check for custom Y position first (both possible orderings)
                comp_key1 <- paste0(g1, "-", g2, "@", cat_val)
                comp_key2 <- paste0(g2, "-", g1, "@", cat_val)

                custom_y_found <- FALSE
                if (!is.null(custom_y_positions[[comp_key1]])) {
                  bracket_data$y.position[bracket_idx] <- as.numeric(custom_y_positions[[comp_key1]])
                  cat("Using custom Y position for", comp_key1, ":", custom_y_positions[[comp_key1]], "\\n")
                  custom_y_found <- TRUE
                } else if (!is.null(custom_y_positions[[comp_key2]])) {
                  bracket_data$y.position[bracket_idx] <- as.numeric(custom_y_positions[[comp_key2]])
                  cat("Using custom Y position for", comp_key2, ":", custom_y_positions[[comp_key2]], "\\n")
                  custom_y_found <- TRUE
                }

                # If custom position found, skip auto-calculation
                if (custom_y_found) {
                  next
                }

                # Calculate baseHeight for this specific comparison
                # baseHeight = max of (mean + error) for the two groups
                data_at_cat_all <- plot_data[plot_data$category == cat_val, ]
                g1_data <- data_at_cat_all[data_at_cat_all$group == g1, ]
                g2_data <- data_at_cat_all[data_at_cat_all$group == g2, ]

                if (nrow(g1_data) > 0 && nrow(g2_data) > 0) {
                  # Match regular bar's baseHeight calculation:
                  # baseHeight = max of (error_bar_top, max_data_point) for each group
                  # No error bars in box/violin plots
                  # No error bars in box/violin plots

                  # For box/violin plots, calculate max values from data
                  g1_max_val <- max(g1_data$value, na.rm = TRUE)
                  g2_max_val <- max(g2_data$value, na.rm = TRUE)
                  
                  g1_top <- g1_max_val  # Use max value directly
                  g2_top <- g2_max_val  # Use max value directly

                  baseHeight <- max(g1_top, g2_top, na.rm = TRUE)
                } else {
                  # Fallback: use max of all groups at this category
                  baseHeight <- max(data_at_cat_all$value, na.rm = TRUE)
                }

                # Calculate Y position with stacking - same for all scales
                # ggplot2 handles log transformation internally
                if (j == 1) {
                  # First bracket at this category (higher to avoid overlap with dots)
                  bracket_data$y.position[bracket_idx] <- baseHeight + (unit_step * 4.0)
                } else {
                  # Stack subsequent brackets (more spacing to avoid overlap)
                  prev_bracket_idx <- na_positions[j-1]
                  bracket_data$y.position[bracket_idx] <- max(
                    baseHeight + (unit_step * 5.0),
                    bracket_data$y.position[prev_bracket_idx] + (unit_step * 5.0)
                  )
                }
              }
            }
          }
        }

        # Add statistical brackets using stat_pvalue_manual (same as regular bar)
        if (nrow(bracket_data) > 0) {
          # Calculate X positions for each group within each category
          # IMPORTANT: Use only groups that actually have data, not all factor levels
          unique_groups <- levels(droplevels(plot_data$group))
          n_groups <- length(unique_groups)

          # Calculate xmin and xmax for each bracket
          for (i in 1:nrow(bracket_data)) {
            if (!is.na(bracket_data$group1[i]) && bracket_data$group1[i] != "" &&
                !is.na(bracket_data$group2[i]) && bracket_data$group2[i] != "") {

              cat_idx <- bracket_data$x.position[i]
              g1_idx <- which(unique_groups == bracket_data$group1[i])
              g2_idx <- which(unique_groups == bracket_data$group2[i])

              if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                # Calculate X positions to match the CENTER of each dodged bar
                # ggplot2's position_dodge formula for bar centers:
                # x = category_x + (index - 1 - (n-1)/2) * (dodge_width / n)
                # where index is 1-based

                # Calculate the center position of each bar
                g1_x <- cat_idx + (g1_idx - 1 - (n_groups - 1) / 2) * (dodge_width / n_groups)
                g2_x <- cat_idx + (g2_idx - 1 - (n_groups - 1) / 2) * (dodge_width / n_groups)

                # Store xmin/xmax for stat_pvalue_manual
                # IMPORTANT: Ensure xmin is leftmost and xmax is rightmost
                bracket_data$xmin[i] <- min(g1_x, g2_x)
                bracket_data$xmax[i] <- max(g1_x, g2_x)
              }
            }
          }

          # Store bracket data globally for educational R code export
          bracket_export_data <<- bracket_data

          # Get parameter values (same as regular bar)
          bracket_size <- symbol_size
          line_size <- ggpubr_line_size
          tip_length <- ggpubr_tip_length
          v_just <- ggpubr_vjust + 0.6

          # Split data by symbol type to apply different sizes (same as regular bar)
          asterisk_data <- bracket_data[bracket_data$p.signif != "ns", ]
          ns_data <- bracket_data[bracket_data$p.signif == "ns", ]

          # Add asterisk symbols with full size
          if (nrow(asterisk_data) > 0) {
            p <- p + stat_pvalue_manual(asterisk_data,
                                       label = "p.signif",
                                       xmin = "xmin",
                                       xmax = "xmax",
                                       y.position = "y.position",
                                       size = bracket_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = v_just,
                                       hjust = 0.5,
                                       step.increase = 0,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
          }

          # Add n.s. symbols with smaller size and higher position
          if (nrow(ns_data) > 0) {
            smaller_size <- bracket_size * 0.7
            ns_vjust <- v_just - 0.5
            p <- p + stat_pvalue_manual(ns_data,
                                       label = "p.signif",
                                       xmin = "xmin",
                                       xmax = "xmax",
                                       y.position = "y.position",
                                       size = smaller_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = ns_vjust,
                                       hjust = 0.5,
                                       step.increase = 0,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
          }
        }

        # Store statistical results in global variable for JavaScript to retrieve
        if (length(stat_text_results) > 0) {
          grouped_bar_stat_results <<- paste(stat_text_results, collapse="\\n\\n")
          cat("\\n=== STATISTICAL RESULTS SUMMARY ===\\n")
          cat(grouped_bar_stat_results, "\\n")
          cat("===================================\\n")
          cat("DEBUG: Stored", nchar(grouped_bar_stat_results), "characters in grouped_bar_stat_results\\n")
        } else {
          cat("DEBUG: No stat_text_results to store\\n")
          grouped_bar_stat_results <<- ""
        }
        }, error = function(e) {
          grouped_bar_stat_results <<- paste(grouped_bar_stat_results, "\\nERROR in statistical analysis:", e$message)
        })
      }
      
      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Basic violin plot function
    sato_violin <- function(dat, x_col=1, y_col=2, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9, width=0.7,
                           target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                           title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                           title_text="Violin plot", x_text="Category", y_text="Value",
                           show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                           x_scale="linear", y_scale="linear",
                           theme_name="minimal",
                           x_axis_rotation=0, y_axis_rotation=0,
                           x_axis_hjust=0.5, x_axis_vjust=0.5,
                           y_axis_hjust=0.5, y_axis_vjust=0.5) {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Prepare data frame
      df <- data.frame(
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      
      p <- ggplot(df, aes(x = category, y = value)) +
           geom_violin(fill = fill, color = color, linewidth = linewidth, alpha = alpha, width = width)

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Violin plot with individual data points
    sato_violin_dot <- function(dat, x_col=1, y_col=2, fill="#4C78A8", color="#1F2937", linewidth=0.7, alpha=0.9, width=0.7,
                               dot_size=4, dot_alpha=1.0, dot_color="#333333", dot_shape=16, jitter_width=0.2,
                               target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                               title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                               title_text="Violin plot with data points", x_text="Category", y_text="Value",
                               show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                               x_scale="linear", y_scale="linear",
                               theme_name="minimal",
                               x_axis_rotation=0, y_axis_rotation=0,
                               x_axis_hjust=0.5, x_axis_vjust=0.5,
                               y_axis_hjust=0.5, y_axis_vjust=0.5) {

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)

      # Prepare data frame
      df <- data.frame(
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      
      p <- ggplot(df, aes(x = category, y = value)) +
           # Violin layer
           geom_violin(fill = fill, color = color, linewidth = linewidth, alpha = alpha, width = width) +
           # Individual data points (jittered horizontally only)
           geom_point(position = position_jitter(width = jitter_width, height = 0),
                     size = dot_size, alpha = dot_alpha, color = dot_color, shape = dot_shape)

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped violin plot function - for data with Group/Category columns
    sato_violin_grouped <- function(dat, group_col=1, x_col=2, y_col=3,
                                   fill_colors=c("#4C78A8", "#E15759"), stroke_color="#1f2937",
                                   alpha=0.9, linewidth=0.7, width=0.7, dodge_width=0.9,
                                   group_name="Group", category_name="Treatment", value_name="Value",
                                   target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                   title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                   title_text="Grouped Violin Plot", x_text="Category", y_text="Value",
                                   show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                   x_scale="linear", y_scale="linear",
                                   theme_name="minimal",
                                   x_axis_rotation=0, y_axis_rotation=0,
                                   x_axis_hjust=0.5, x_axis_vjust=0.5,
                                   y_axis_hjust=0.5, y_axis_vjust=0.5,
                                   add_statistics=FALSE, statistical_test="auto", variance_test="levene", symbol_size=8,
                                   ggpubr_line_size=1.0, ggpubr_tip_length=0.04, ggpubr_vjust=-0.3,
                                   comparison_mode="all", custom_comparisons="[]", custom_positions="{}",
                                   selected_posthoc_test="tukey", dunnett_control="", sato_symbol_size=7) {
      
      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)
      
      if (ncol(dat) < 3) {
        stop("Grouped violin chart requires at least 3 columns: Group, Category, Value")
      }
      
      # 列名検出
      actual_group_name <- if(!is.null(names(dat)) && length(names(dat)) >= group_col && !grepl("^V[0-9]+$", names(dat)[group_col])) {
        names(dat)[group_col]
      } else {
        group_name
      }
      
      # データ準備とグループ順序設定
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]
      
      if ("Normal" %in% unique(as.character(plot_data$group)) && "Tumor" %in% unique(as.character(plot_data$group))) {
        plot_data$group <- factor(plot_data$group, levels = c("Normal", "Tumor"))
      }
      
      # グループ数に応じて色を設定
      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)
      
      if (length(fill_colors) < n_groups) {
        fill_colors <- rep(fill_colors, length.out = n_groups)
      }
      
      # グループ化されたバイオリンプロット作成
      p <- ggplot(plot_data, aes(x = category, y = value, fill = group)) +
           geom_violin(position = position_dodge(width = dodge_width), 
                      width = width,
                      alpha = alpha, 
                      linewidth = linewidth,
                      color = stroke_color) +
           scale_fill_manual(values = setNames(fill_colors[1:n_groups], group_levels)) +
           labs(fill = actual_group_name)
      
      # Add statistical analysis if requested
      if (add_statistics && ncol(dat) >= 3) {
        p <- sato_add_statistics_to_plot(p, plot_data, "group", "value", statistical_test, sato_symbol_size, posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale=y_scale, stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # Grouped violin plot with individual data points
    sato_violin_grouped_dot <- function(dat, group_col=1, x_col=2, y_col=3,
                                       fill_colors=c("#4C78A8", "#E15759"), stroke_color="#1f2937",
                                       alpha=0.9, linewidth=0.7, width=0.7, dodge_width=0.9,
                                       dot_size=4, dot_alpha=1.0, dot_color="#333333", dot_shape=16, jitter_width=0.15,
                                       group_name="Group", category_name="Treatment", value_name="Value",
                                       target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                                       title_size=14, x_axis_title_size=12, y_axis_title_size=12, x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                                       title_text="Grouped Violin Plot with Dots", x_text="Category", y_text="Value",
                                       show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                                       x_scale="linear", y_scale="linear",
                                       theme_name="minimal",
                                       x_axis_rotation=0, y_axis_rotation=0,
                                       x_axis_hjust=0.5, x_axis_vjust=0.5,
                                       y_axis_hjust=0.5, y_axis_vjust=0.5,
                                       add_statistics=FALSE, statistical_test="auto", variance_test="levene", symbol_size=8,
                                       ggpubr_line_size=1.0, ggpubr_tip_length=0.04, ggpubr_vjust=-0.3,
                                       comparison_mode="all", custom_comparisons="[]", custom_positions="{}",
                                       selected_posthoc_test="tukey", dunnett_control="", sato_symbol_size=7,
                                       stat_symbol_type="stars", custom_symbol_05="*", custom_symbol_01="**", custom_symbol_001="***", custom_symbol_ns="ns") {

      # Helper function to generate significance symbol based on p-value and symbol type
      get_sig_symbol <- function(p_val) {
        if (stat_symbol_type == "pvalue") {
          return(sprintf("p=%.3f", p_val))
        } else if (stat_symbol_type == "custom") {
          if (p_val < 0.001) return(custom_symbol_001)
          else if (p_val < 0.01) return(custom_symbol_01)
          else if (p_val < 0.05) return(custom_symbol_05)
          else return(custom_symbol_ns)
        } else {
          # Default to stars
          if (p_val < 0.001) return("***")
          else if (p_val < 0.01) return("**")
          else if (p_val < 0.05) return("*")
          else return("ns")
        }
      }

      # Transform data for negative log scales
      dat <- sato_transform_data(dat, x_col, y_col, x_scale, y_scale)
      
      if (ncol(dat) < 3) {
        stop("Grouped violin chart requires at least 3 columns: Group, Category, Value")
      }
      
      # 列名検出
      actual_group_name <- if(!is.null(names(dat)) && length(names(dat)) >= group_col && !grepl("^V[0-9]+$", names(dat)[group_col])) {
        names(dat)[group_col]
      } else {
        group_name
      }
      
      # データ準備とグループ順序設定
      plot_data <- data.frame(
        group = if(is.factor(dat[[group_col]])) dat[[group_col]] else factor(dat[[group_col]]),
        category = if(is.factor(dat[[x_col]])) dat[[x_col]] else factor(dat[[x_col]]),
        value = as.numeric(dat[[y_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]
      
      if ("Normal" %in% unique(as.character(plot_data$group)) && "Tumor" %in% unique(as.character(plot_data$group))) {
        plot_data$group <- factor(plot_data$group, levels = c("Normal", "Tumor"))
      }
      
      # グループ数に応じて色を設定
      n_groups <- length(levels(plot_data$group))
      group_levels <- levels(plot_data$group)
      
      if (length(fill_colors) < n_groups) {
        fill_colors <- rep(fill_colors, length.out = n_groups)
      }
      
      # グループ化されたバイオリンプロット + dots作成
      p <- ggplot(plot_data, aes(x = category, y = value, fill = group)) +
           geom_violin(position = position_dodge(width = dodge_width),
                      width = width,
                      alpha = alpha,
                      linewidth = linewidth,
                      color = stroke_color) +
           geom_point(aes(fill = group),
                     position = position_jitterdodge(dodge.width = dodge_width, jitter.width = jitter_width, jitter.height = 0),
                     size = dot_size, alpha = dot_alpha, shape = dot_shape, color = dot_color) +
           scale_fill_manual(values = setNames(fill_colors[1:n_groups], group_levels)) +
           labs(fill = actual_group_name)

      # Initialize global variable for statistical results
      grouped_bar_stat_results <<- ""

      # Add statistical analysis if requested - compare groups within each category
      if (add_statistics && n_groups >= 2) {
        tryCatch({
          # Load ggpubr package for stat_pvalue_manual
          if (!require("ggpubr", quietly = TRUE)) {
            webr::install("ggpubr")
          }
          library(ggpubr)

          # Get unique categories (filter out any categories with no data)
          all_categories <- levels(plot_data$category)
          unique_categories <- all_categories[all_categories %in% unique(plot_data$category)]

          # Handle custom comparisons
          selected_categories <- NULL
          custom_y_positions <- list()
          custom_comps <- c()  # Initialize to empty array

          if (comparison_mode == "custom") {
            # Parse custom comparisons: format is "Group1-Group2@CategoryA", "Group1-Group2@CategoryB", etc.
            custom_comps <- tryCatch(jsonlite::fromJSON(custom_comparisons), error = function(e) c())
            cat("Custom comparisons:", custom_comparisons, "\\n")
            cat("Custom positions JSON:", custom_positions, "\\n")

            if (length(custom_comps) > 0) {
              # Extract category values from custom comparisons
              selected_categories <- c()
              for (comp_str in custom_comps) {
                # Parse format: "Group1-Group2@Category"
                parts <- strsplit(comp_str, "@", fixed = TRUE)[[1]]
                if (length(parts) == 2) {
                  cat_val <- parts[2]
                  selected_categories <- c(selected_categories, cat_val)
                }
              }
              cat("Selected categories from custom comparisons:", paste(selected_categories, collapse=", "), "\\n")

              # Parse custom positions
              custom_pos_list <- tryCatch(jsonlite::fromJSON(custom_positions), error = function(e) list())
              if (length(custom_pos_list) > 0) {
                for (comp_str in names(custom_pos_list)) {
                  custom_y_positions[[comp_str]] <- as.numeric(custom_pos_list[[comp_str]])
                }
              }
            }
          }

          # Determine which categories to test
          if (is.null(selected_categories) || length(selected_categories) == 0) {
            # Test all categories (default or "all" mode)
            categories_to_test <- unique_categories
          } else {
            # Test only selected categories (deduplicate to avoid testing same category multiple times)
            categories_to_test <- unique(selected_categories)
          }

          # Debug: Check what we're testing
          cat("DEBUG: comparison_mode =", comparison_mode, "\\n")
          cat("DEBUG: categories_to_test =", paste(categories_to_test, collapse=", "), "\\n")
          cat("DEBUG: Number of categories to test:", length(categories_to_test), "\\n")

          stat_text_results <- c()

        # Data frame to store ggpubr bracket information
        bracket_data <- data.frame(
          group1 = character(),
          group2 = character(),
          p.signif = character(),
          x.position = numeric(),
          y.position = numeric(),
          stringsAsFactors = FALSE
        )

        # For each category, perform statistical tests between groups
        for (cat_idx in seq_along(categories_to_test)) {
          cat_val <- categories_to_test[cat_idx]

          # Get data for this category
          data_at_cat <- plot_data[plot_data$category == cat_val, ]

          if (nrow(data_at_cat) > 0) {
            # Get groups with data at this category
            groups_with_data <- unique(data_at_cat$group[!is.na(data_at_cat$value)])
            n_groups_at_cat <- length(groups_with_data)

            if (n_groups_at_cat == 2) {
              # TWO GROUPS: Use t-test or Wilcoxon
              group1_data <- data_at_cat[data_at_cat$group == groups_with_data[1], "value"]
              group2_data <- data_at_cat[data_at_cat$group == groups_with_data[2], "value"]

              # Only test if both groups have data
              if (length(group1_data) > 0 && length(group2_data) > 0) {
                # Perform normality test for both groups
                normality_text <- c()
                is_group1_normal <- TRUE
                is_group2_normal <- TRUE

                if (length(group1_data) >= 3 && length(group1_data) <= 5000) {
                  shapiro_result1 <- tryCatch(
                    shapiro.test(group1_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result1)) {
                    is_group1_normal <- shapiro_result1$p.value >= 0.05
                    norm_status <- if (is_group1_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[1], shapiro_result1$p.value, norm_status))
                  }
                }

                if (length(group2_data) >= 3 && length(group2_data) <= 5000) {
                  shapiro_result2 <- tryCatch(
                    shapiro.test(group2_data),
                    error = function(e) NULL
                  )
                  if (!is.null(shapiro_result2)) {
                    is_group2_normal <- shapiro_result2$p.value >= 0.05
                    norm_status <- if (is_group2_normal) "normal" else "non-normal"
                    normality_text <- c(normality_text,
                      sprintf("  %s: p=%.4f (%s)", groups_with_data[2], shapiro_result2$p.value, norm_status))
                  }
                }

                both_normal <- is_group1_normal && is_group2_normal

                # Select test based on mode
                test_to_use <- statistical_test
                if (statistical_test == "auto") {
                  if (both_normal) {
                    test_to_use <- "t-test"
                    cat("  Auto-selected: t-test (both groups normal)\\n")
                  } else {
                    test_to_use <- "wilcoxon"
                    cat("  Auto-selected: Wilcoxon test (non-normal data detected)\\n")
                  }
                } else if (statistical_test == "parametric") {
                  test_to_use <- "t-test"
                  cat("  Manual mode: parametric t-test\\n")
                } else if (statistical_test == "nonparametric") {
                  test_to_use <- "wilcoxon"
                  cat("  Manual mode: non-parametric Wilcoxon test\\n")
                }

                # Perform variance test (only for parametric tests)
                variance_text <- ""
                equal_variances <- TRUE
                if (test_to_use != "wilcoxon") {
                  if (variance_test == "levene") {
                    # Levene test for equality of variances
                    combined_data <- data.frame(
                      values = c(group1_data, group2_data),
                      group = factor(c(rep(groups_with_data[1], length(group1_data)),
                                      rep(groups_with_data[2], length(group2_data))))
                    )

                    # Calculate Levene test manually
                    group_means <- tapply(combined_data$values, combined_data$group, mean)
                    abs_deviations <- abs(combined_data$values - group_means[combined_data$group])
                    levene_result <- tryCatch(
                      anova(lm(abs_deviations ~ combined_data$group)),
                      error = function(e) NULL
                    )

                    if (!is.null(levene_result)) {
                      levene_p <- levene_result$\`Pr(>F)\`[1]
                      equal_variances <- levene_p > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, Levene)", levene_p, variance_status)
                      cat(sprintf("    Levene test: p=%.4f (%s)\\n", levene_p, variance_status))
                    }
                  } else {
                    # F-test for equality of variances
                    var_test <- tryCatch(
                      var.test(group1_data, group2_data),
                      error = function(e) NULL
                    )

                    if (!is.null(var_test)) {
                      equal_variances <- var_test$p.value > 0.05
                      variance_status <- if (equal_variances) "equal variances" else "unequal variances"
                      variance_text <- sprintf("Variance test: p=%.4f (%s, F-test)", var_test$p.value, variance_status)
                      cat(sprintf("    F-test: p=%.4f (%s)\\n", var_test$p.value, variance_status))
                    }
                  }
                }

                # Perform statistical test based on auto-selection
                if (test_to_use == "wilcoxon") {
                  test_result <- tryCatch(
                    wilcox.test(group1_data, group2_data),
                    error = function(e) NULL
                  )
                  test_name <- "Wilcoxon"
                } else {
                  # Use appropriate t-test based on variance equality
                  test_result <- tryCatch(
                    t.test(group1_data, group2_data, var.equal = equal_variances),
                    error = function(e) NULL
                  )
                  test_name <- if (equal_variances) "Student's t-test" else "Welch's t-test"
                }

                if (!is.null(test_result)) {
                  p_val <- test_result$p.value
                  cat(sprintf("  2-group test: %s vs %s, %s p=%.4f\\n",
                              groups_with_data[1], groups_with_data[2], test_name, p_val))

                  # Determine significance symbol using helper function
                  sig_label <- get_sig_symbol(p_val)

                  # Store result text for UI with proper order: Category -> Group Stats -> Normality -> Variance -> Test result
                  result_text <- sprintf("Category %s:", cat_val)

                  # Add group statistics (n, mean, sd) for each group
                  result_text <- paste0(result_text, "\\nGroup Statistics:")
                  for (g in groups_with_data) {
                    g_data <- plot_data[plot_data$category == cat_val & plot_data$group == g, "value"]
                    result_text <- paste0(result_text, sprintf("\\n  %s: n=%d, mean=%.2f, sd=%.2f",
                                          g, length(g_data), mean(g_data, na.rm=TRUE), sd(g_data, na.rm=TRUE)))
                  }

                  # Add normality testing results
                  if (length(normality_text) > 0) {
                    result_text <- paste0(result_text, "\\n\\nNormality Testing (Shapiro-Wilk):\\n  ", paste(normality_text, collapse="\\n  "))
                  }

                  # Add variance test results
                  if (nchar(variance_text) > 0) {
                    result_text <- paste0(result_text, "\\n", variance_text)
                  }

                  # Add main test result
                  result_text <- paste0(result_text, "\\n\\n", sprintf("%s: p=%.4f (%s)",
                                        test_name, p_val, sig_label))
                  stat_text_results <- c(stat_text_results, result_text)

                  # Add ggpubr bracket for this comparison
                  # Calculate x position (find position in all unique_categories, not just selected ones)
                  cat_index <- which(unique_categories == cat_val)

                  # Calculate y position (above the highest bar/point in this category)
                  # Check if custom position is provided (check both possible orderings)
                  comp_key1 <- paste0(groups_with_data[1], "-", groups_with_data[2], "@", cat_val)
                  comp_key2 <- paste0(groups_with_data[2], "-", groups_with_data[1], "@", cat_val)

                  custom_y_pos <- NULL
                  if (!is.null(custom_y_positions[[comp_key1]])) {
                    custom_y_pos <- custom_y_positions[[comp_key1]]
                    cat("Using custom Y position for", comp_key1, ":", custom_y_pos, "\\n")
                  } else if (!is.null(custom_y_positions[[comp_key2]])) {
                    custom_y_pos <- custom_y_positions[[comp_key2]]
                    cat("Using custom Y position for", comp_key2, ":", custom_y_pos, "\\n")
                  }

                  if (!is.null(custom_y_pos)) {
                    # Use custom Y position
                    y_pos <- custom_y_pos
                  } else {
                    # Calculate default position
                    # Get the summary data for this category to find the bar heights
                    data_at_cat_all <- plot_data[plot_data$category == cat_val, ]
                    max_value <- max(data_at_cat_all$value, na.rm = TRUE)
                    # No error bars in box plots

                    # Position symbol above the tallest bar + error bar with some spacing
                    # Increased spacing to avoid overlapping with dots (0.055 -> 0.15)
                    y_pos <- max_value + (max_value * 0.20)  # 20% above highest value
                  }

                  # Add to bracket data based on comparison mode
                  should_add_bracket <- FALSE
                  if (comparison_mode == "significant") {
                    # Only add if significant
                    should_add_bracket <- (sig_label != "ns")
                  } else if (comparison_mode == "all") {
                    # Add all comparisons
                    should_add_bracket <- TRUE
                  } else if (comparison_mode == "custom") {
                    # Only add if this comparison is selected
                    # Check if this comparison is in the custom selections
                    comp_key1 <- paste0(groups_with_data[1], "-", groups_with_data[2], "@", cat_val)
                    comp_key2 <- paste0(groups_with_data[2], "-", groups_with_data[1], "@", cat_val)
                    should_add_bracket <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                  }

                  if (should_add_bracket) {
                    bracket_data <- rbind(bracket_data, data.frame(
                      group1 = as.character(groups_with_data[1]),
                      group2 = as.character(groups_with_data[2]),
                      p.signif = sig_label,
                      x.position = cat_index,
                      y.position = y_pos,
                      stringsAsFactors = FALSE
                    ))
                  }
                }
              }
            } else if (n_groups_at_cat >= 3) {
              # THREE OR MORE GROUPS: Use ANOVA/Kruskal-Wallis + post-hoc tests
              cat("  Performing ANOVA/Kruskal-Wallis for", n_groups_at_cat, "groups at category", cat_val, "\\n")

              # Calculate x position (find position in all unique_categories, not just selected ones)
              cat_index <- which(unique_categories == cat_val)

              # Prepare data for ANOVA
              anova_data <- data.frame(
                group = factor(data_at_cat$group),
                value = as.numeric(data_at_cat$value)
              )
              anova_data <- anova_data[complete.cases(anova_data), ]

              if (nrow(anova_data) >= 3) {
                # Perform normality test for auto-selection
                normality_text <- c()
                all_normal <- TRUE

                for (grp in groups_with_data) {
                  grp_data <- data_at_cat[data_at_cat$group == grp, "value"]
                  if (length(grp_data) >= 3 && length(grp_data) <= 5000) {
                    shapiro_result <- tryCatch(
                      shapiro.test(grp_data),
                      error = function(e) NULL
                    )
                    if (!is.null(shapiro_result)) {
                      is_normal <- shapiro_result$p.value >= 0.05
                      norm_status <- if (is_normal) "normal" else "non-normal"
                      normality_text <- c(normality_text,
                        sprintf("  %s: p=%.4f (%s)", grp, shapiro_result$p.value, norm_status))
                      if (!is_normal) all_normal <- FALSE
                    }
                  }
                }

                # Auto-select test
                test_to_use <- statistical_test
                if (statistical_test == "auto") {
                  test_to_use <- if (all_normal) "anova" else "kruskal"
                  cat("  Auto-selected:", test_to_use, "(normality:", all_normal, ")\\n")
                }

                # Perform omnibus test
                omnibus_p <- NA
                test_name <- ""
                if (test_to_use == "kruskal") {
                  kruskal_result <- kruskal.test(value ~ group, data = anova_data)
                  omnibus_p <- kruskal_result$p.value
                  test_name <- "Kruskal-Wallis"
                } else {
                  anova_result <- aov(value ~ group, data = anova_data)
                  anova_summary <- summary(anova_result)
                  omnibus_p <- anova_summary[[1]][["Pr(>F)"]][1]
                  test_name <- "ANOVA"
                }

                cat("  ", test_name, "p-value:", omnibus_p, "\\n")

                # Build result text - ORDER: Normality, ANOVA, Post-hoc
                result_text <- sprintf("Category %s:", cat_val)
                if (length(normality_text) > 0) {
                  result_text <- paste0(result_text, "\\nNormality (Shapiro-Wilk):\\n", paste(normality_text, collapse="\\n"))
                }
                # Add ANOVA result with significance symbol
                anova_sig_label <- if (!is.na(omnibus_p)) {
                  if (omnibus_p < 0.001) "***" else if (omnibus_p < 0.01) "**" else if (omnibus_p < 0.05) "*" else "ns"
                } else "ns"
                result_text <- paste0(result_text, "\\n", test_name, " p=", sprintf("%.4f", omnibus_p), " (", anova_sig_label, ")")

                # If significant, perform post-hoc tests
                if (!is.na(omnibus_p) && omnibus_p < 0.05) {
                  cat("  Omnibus test significant, performing post-hoc tests\\n")

                  # Perform appropriate post-hoc test based on omnibus test type
                  posthoc_result <- NULL
                  posthoc_text <- c()

                  if (test_to_use == "kruskal") {
                    # For Kruskal-Wallis, use Dunn test
                    cat("  Performing Dunn post-hoc test...\\n")
                    posthoc_result <- tryCatch({
                      if (!requireNamespace("dunn.test", quietly = TRUE)) {
                        cat("Installing dunn.test package...\\n")
                        webr::install("dunn.test")
                      }
                      library(dunn.test)

                      dunn_result <- dunn.test(anova_data$value, anova_data$group, method = "bonferroni")

                      # Format results into a data frame
                      data.frame(
                        Comparison = dunn_result$comparisons,
                        P.adj = dunn_result$P.adjusted,
                        stringsAsFactors = FALSE
                      )
                    }, error = function(e) {
                      cat("  Dunn test error:", e$message, "\\n")
                      # Add error as annotation text for debugging
                      result_text <<- paste0(result_text, "\\n[DUNN ERROR: ", e$message, "]")
                      NULL
                    })

                    if (!is.null(posthoc_result)) {
                      cat(sprintf("  Dunn found %d pairwise comparisons\\n", nrow(posthoc_result)))
                      for (i in 1:nrow(posthoc_result)) {
                        comparison <- posthoc_result$Comparison[i]
                        p_adj <- posthoc_result$P.adj[i]

                        sig_label <- ""
                        if (!is.na(p_adj)) {
                          sig_label <- get_sig_symbol(p_adj)
                        } else {
                          sig_label <- "ns"
                        }

                        cat(sprintf("    %s: p=%.4f (%s)\\n", comparison, p_adj, sig_label))
                        posthoc_text <- c(posthoc_text, sprintf("  %s: p=%.4f (%s)",
                                                                comparison, ifelse(is.na(p_adj), 1.0, p_adj), sig_label))

                        # Add comparisons to bracket data based on comparison mode
                        should_add_posthoc <- FALSE
                        if (comparison_mode == "significant") {
                          should_add_posthoc <- (sig_label != "ns")
                        } else if (comparison_mode == "all") {
                          should_add_posthoc <- TRUE
                        } else if (comparison_mode == "custom") {
                          # Parse comparison to check if selected
                          comp_parts_temp <- strsplit(comparison, " - ")[[1]]
                          if (length(comp_parts_temp) == 2) {
                            comp_key1 <- paste0(comp_parts_temp[1], "-", comp_parts_temp[2], "@", cat_val)
                            comp_key2 <- paste0(comp_parts_temp[2], "-", comp_parts_temp[1], "@", cat_val)
                            should_add_posthoc <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                          }
                        }

                        if (should_add_posthoc) {
                          # Parse comparison string (format: "group1 - group2")
                          comp_parts <- strsplit(comparison, " - ")[[1]]
                          if (length(comp_parts) == 2) {
                            bracket_data <- rbind(bracket_data, data.frame(
                              group1 = comp_parts[1],
                              group2 = comp_parts[2],
                              p.signif = sig_label,
                              x.position = cat_index,
                              y.position = NA,  # Will be calculated later
                              stringsAsFactors = FALSE
                            ))
                          }
                        }
                      }
                    }
                  } else {
                    # For ANOVA, use selected post-hoc test
                    cat(sprintf("  Performing %s post-hoc test...\\n", selected_posthoc_test))
                    posthoc_result <- NULL
                    posthoc_summary <- NULL

                    if (selected_posthoc_test == "tukey") {
                      # Tukey HSD test
                      posthoc_result <- tryCatch(
                        TukeyHSD(anova_result),
                        error = function(e) {
                          cat("  Tukey error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        posthoc_summary <- posthoc_result$group
                      }
                    } else if (selected_posthoc_test == "bonferroni") {
                      # Bonferroni correction using pairwise t-tests
                      posthoc_result <- tryCatch(
                        pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "bonferroni"),
                        error = function(e) {
                          cat("  Bonferroni error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        # Convert to Tukey-like format
                        groups <- levels(anova_data$group)
                        comparisons <- combn(groups, 2, simplify = FALSE)
                        diff_values <- c()
                        p_adj_values <- c()
                        comparison_names <- c()

                        for (comp in comparisons) {
                          g1 <- comp[1]
                          g2 <- comp[2]
                          g1_idx <- which(rownames(posthoc_result$p.value) == g1)
                          g2_idx <- which(colnames(posthoc_result$p.value) == g2)

                          if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                            p_val <- posthoc_result$p.value[g1_idx, g2_idx]
                          } else {
                            g1_idx <- which(rownames(posthoc_result$p.value) == g2)
                            g2_idx <- which(colnames(posthoc_result$p.value) == g1)
                            p_val <- if (length(g1_idx) > 0 && length(g2_idx) > 0) posthoc_result$p.value[g1_idx, g2_idx] else NA
                          }

                          g1_values <- anova_data$value[anova_data$group == g1]
                          g2_values <- anova_data$value[anova_data$group == g2]
                          diff <- mean(g2_values, na.rm = TRUE) - mean(g1_values, na.rm = TRUE)

                          comparison_names <- c(comparison_names, paste0(g2, "-", g1))
                          diff_values <- c(diff_values, diff)
                          p_adj_values <- c(p_adj_values, p_val)
                        }

                        posthoc_summary <- data.frame(
                          diff = diff_values,
                          lwr = rep(NA, length(diff_values)),
                          upr = rep(NA, length(diff_values)),
                          "p adj" = p_adj_values,
                          row.names = comparison_names,
                          check.names = FALSE
                        )
                      }
                    } else if (selected_posthoc_test == "holm") {
                      # Holm correction using pairwise t-tests
                      posthoc_result <- tryCatch(
                        pairwise.t.test(anova_data$value, anova_data$group, p.adjust.method = "holm"),
                        error = function(e) {
                          cat("  Holm error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        # Convert to Tukey-like format (same as Bonferroni)
                        groups <- levels(anova_data$group)
                        comparisons <- combn(groups, 2, simplify = FALSE)
                        diff_values <- c()
                        p_adj_values <- c()
                        comparison_names <- c()

                        for (comp in comparisons) {
                          g1 <- comp[1]
                          g2 <- comp[2]
                          g1_idx <- which(rownames(posthoc_result$p.value) == g1)
                          g2_idx <- which(colnames(posthoc_result$p.value) == g2)

                          if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                            p_val <- posthoc_result$p.value[g1_idx, g2_idx]
                          } else {
                            g1_idx <- which(rownames(posthoc_result$p.value) == g2)
                            g2_idx <- which(colnames(posthoc_result$p.value) == g1)
                            p_val <- if (length(g1_idx) > 0 && length(g2_idx) > 0) posthoc_result$p.value[g1_idx, g2_idx] else NA
                          }

                          g1_values <- anova_data$value[anova_data$group == g1]
                          g2_values <- anova_data$value[anova_data$group == g2]
                          diff <- mean(g2_values, na.rm = TRUE) - mean(g1_values, na.rm = TRUE)

                          comparison_names <- c(comparison_names, paste0(g2, "-", g1))
                          diff_values <- c(diff_values, diff)
                          p_adj_values <- c(p_adj_values, p_val)
                        }

                        posthoc_summary <- data.frame(
                          diff = diff_values,
                          lwr = rep(NA, length(diff_values)),
                          upr = rep(NA, length(diff_values)),
                          "p adj" = p_adj_values,
                          row.names = comparison_names,
                          check.names = FALSE
                        )
                      }
                    } else if (selected_posthoc_test == "dunnett") {
                      # Dunnett test (not typically used for grouped bar, default to Tukey)
                      cat("  Warning: Dunnett test not implemented for grouped bar, using Tukey\\n")
                      posthoc_result <- tryCatch(
                        TukeyHSD(anova_result),
                        error = function(e) {
                          cat("  Tukey error:", e$message, "\\n")
                          NULL
                        }
                      )
                      if (!is.null(posthoc_result)) {
                        posthoc_summary <- posthoc_result$group
                      }
                    }

                    if (!is.null(posthoc_summary)) {
                      cat(sprintf("  Post-hoc found %d pairwise comparisons\\n", nrow(posthoc_summary)))

                      for (i in 1:nrow(posthoc_summary)) {
                        comparison <- rownames(posthoc_summary)[i]
                        p_adj <- posthoc_summary[i, "p adj"]
                        diff <- posthoc_summary[i, "diff"]

                        sig_label <- ""
                        if (!is.na(p_adj)) {
                          sig_label <- get_sig_symbol(p_adj)
                        } else {
                          sig_label <- "ns"
                        }

                        cat(sprintf("    %s: diff=%.2f, p=%.4f (%s)\\n", comparison, diff, p_adj, sig_label))
                        posthoc_text <- c(posthoc_text, sprintf("  %s: diff=%.2f, p=%.4f (%s)",
                                                                comparison, diff, ifelse(is.na(p_adj), 1.0, p_adj), sig_label))

                        # Add comparisons to bracket data based on comparison mode
                        should_add_posthoc <- FALSE
                        if (comparison_mode == "significant") {
                          should_add_posthoc <- (sig_label != "ns")
                        } else if (comparison_mode == "all") {
                          should_add_posthoc <- TRUE
                        } else if (comparison_mode == "custom") {
                          # Parse comparison to check if selected
                          comp_parts_temp <- strsplit(comparison, "-")[[1]]
                          if (length(comp_parts_temp) == 2) {
                            comp_key1 <- paste0(comp_parts_temp[1], "-", comp_parts_temp[2], "@", cat_val)
                            comp_key2 <- paste0(comp_parts_temp[2], "-", comp_parts_temp[1], "@", cat_val)
                            should_add_posthoc <- (comp_key1 %in% custom_comps) || (comp_key2 %in% custom_comps)
                          }
                        }

                        if (should_add_posthoc) {
                          # Parse comparison string (format: "group2-group1")
                          comp_parts <- strsplit(comparison, "-")[[1]]
                          if (length(comp_parts) == 2) {
                            bracket_data <- rbind(bracket_data, data.frame(
                              group1 = comp_parts[1],
                              group2 = comp_parts[2],
                              p.signif = sig_label,
                              x.position = cat_index,
                              y.position = NA,  # Will be calculated later
                              stringsAsFactors = FALSE
                            ))
                          }
                        }
                      }
                    }
                  }

                  # Combine omnibus and post-hoc results (sig_label already added above)
                  if (length(posthoc_text) > 0) {
                    # Add post-hoc test type header with correct test name
                    posthoc_header <- if (test_to_use == "kruskal") {
                      "Post-hoc pairwise comparisons (Dunn test):"
                    } else {
                      # For ANOVA, show the actual selected post-hoc test
                      # Note: Bonferroni and Holm are p-value adjustments, not post-hoc tests
                      if (selected_posthoc_test == "tukey") {
                        "Post-hoc pairwise comparisons (Tukey HSD):"
                      } else if (selected_posthoc_test == "bonferroni") {
                        "Pairwise t-test with Bonferroni correction:"
                      } else if (selected_posthoc_test == "holm") {
                        "Pairwise t-test with Holm correction:"
                      } else if (selected_posthoc_test == "dunnett") {
                        "Post-hoc pairwise comparisons (Dunnett):"
                      } else {
                        "Post-hoc pairwise comparisons (Tukey HSD):"
                      }
                    }
                    result_text <- paste0(result_text, "\\n", posthoc_header, "\\n", paste(posthoc_text, collapse="\\n"))
                  }

                  # Add to stat text results
                  stat_text_results <- c(stat_text_results, result_text)
                } else {
                  # Not significant - sig_label already added above
                  stat_text_results <- c(stat_text_results, result_text)
                }
              }
            }
          }
        }

        # Calculate Y positions for brackets using proper stacking logic (same as regular bar)
        if (nrow(bracket_data) > 0) {
          # Calculate Y range for unit_step
          y_range <- diff(range(plot_data$value,
                                 na.rm = TRUE))

          # Calculate unit_step (spacing between brackets) based on symbol size and Y range
          # This matches the logic from regular bar chart
          unit_step <- y_range * 0.08  # 8% of Y range as default spacing

          # Group brackets by category for proper stacking
          for (cat_idx in unique(bracket_data$x.position)) {
            na_positions <- which(bracket_data$x.position == cat_idx & is.na(bracket_data$y.position))

            if (length(na_positions) > 0) {
              cat_val <- unique_categories[cat_idx]

              # Stack brackets at this category
              for (j in seq_along(na_positions)) {
                bracket_idx <- na_positions[j]

                # Get the two groups being compared
                g1 <- bracket_data$group1[bracket_idx]
                g2 <- bracket_data$group2[bracket_idx]

                # Check for custom Y position first (both possible orderings)
                comp_key1 <- paste0(g1, "-", g2, "@", cat_val)
                comp_key2 <- paste0(g2, "-", g1, "@", cat_val)

                custom_y_found <- FALSE
                if (!is.null(custom_y_positions[[comp_key1]])) {
                  bracket_data$y.position[bracket_idx] <- as.numeric(custom_y_positions[[comp_key1]])
                  cat("Using custom Y position for", comp_key1, ":", custom_y_positions[[comp_key1]], "\\n")
                  custom_y_found <- TRUE
                } else if (!is.null(custom_y_positions[[comp_key2]])) {
                  bracket_data$y.position[bracket_idx] <- as.numeric(custom_y_positions[[comp_key2]])
                  cat("Using custom Y position for", comp_key2, ":", custom_y_positions[[comp_key2]], "\\n")
                  custom_y_found <- TRUE
                }

                # If custom position found, skip auto-calculation
                if (custom_y_found) {
                  next
                }

                # Calculate baseHeight for this specific comparison
                # baseHeight = max of (mean + error) for the two groups
                data_at_cat_all <- plot_data[plot_data$category == cat_val, ]
                g1_data <- data_at_cat_all[data_at_cat_all$group == g1, ]
                g2_data <- data_at_cat_all[data_at_cat_all$group == g2, ]

                if (nrow(g1_data) > 0 && nrow(g2_data) > 0) {
                  # Match regular bar's baseHeight calculation:
                  # baseHeight = max of (error_bar_top, max_data_point) for each group
                  # No error bars in box/violin plots
                  # No error bars in box/violin plots

                  # For box/violin plots, calculate max values from data
                  g1_max_val <- max(g1_data$value, na.rm = TRUE)
                  g2_max_val <- max(g2_data$value, na.rm = TRUE)
                  
                  g1_top <- g1_max_val  # Use max value directly
                  g2_top <- g2_max_val  # Use max value directly

                  baseHeight <- max(g1_top, g2_top, na.rm = TRUE)
                } else {
                  # Fallback: use max of all groups at this category
                  baseHeight <- max(data_at_cat_all$value, na.rm = TRUE)
                }

                # Calculate Y position with stacking - same for all scales
                # ggplot2 handles log transformation internally
                if (j == 1) {
                  # First bracket at this category (higher to avoid overlap with dots)
                  bracket_data$y.position[bracket_idx] <- baseHeight + (unit_step * 4.0)
                } else {
                  # Stack subsequent brackets (more spacing to avoid overlap)
                  prev_bracket_idx <- na_positions[j-1]
                  bracket_data$y.position[bracket_idx] <- max(
                    baseHeight + (unit_step * 5.0),
                    bracket_data$y.position[prev_bracket_idx] + (unit_step * 5.0)
                  )
                }
              }
            }
          }
        }

        # Add statistical brackets using stat_pvalue_manual (same as regular bar)
        if (nrow(bracket_data) > 0) {
          # Calculate X positions for each group within each category
          # IMPORTANT: Use only groups that actually have data, not all factor levels
          unique_groups <- levels(droplevels(plot_data$group))
          n_groups <- length(unique_groups)

          # Calculate xmin and xmax for each bracket
          for (i in 1:nrow(bracket_data)) {
            if (!is.na(bracket_data$group1[i]) && bracket_data$group1[i] != "" &&
                !is.na(bracket_data$group2[i]) && bracket_data$group2[i] != "") {

              cat_idx <- bracket_data$x.position[i]
              g1_idx <- which(unique_groups == bracket_data$group1[i])
              g2_idx <- which(unique_groups == bracket_data$group2[i])

              if (length(g1_idx) > 0 && length(g2_idx) > 0) {
                # Calculate X positions to match the CENTER of each dodged bar
                # ggplot2's position_dodge formula for bar centers:
                # x = category_x + (index - 1 - (n-1)/2) * (dodge_width / n)
                # where index is 1-based

                # Calculate the center position of each bar
                g1_x <- cat_idx + (g1_idx - 1 - (n_groups - 1) / 2) * (dodge_width / n_groups)
                g2_x <- cat_idx + (g2_idx - 1 - (n_groups - 1) / 2) * (dodge_width / n_groups)

                # Store xmin/xmax for stat_pvalue_manual
                # IMPORTANT: Ensure xmin is leftmost and xmax is rightmost
                bracket_data$xmin[i] <- min(g1_x, g2_x)
                bracket_data$xmax[i] <- max(g1_x, g2_x)
              }
            }
          }

          # Store bracket data globally for educational R code export
          bracket_export_data <<- bracket_data

          # Get parameter values (same as regular bar)
          bracket_size <- symbol_size
          line_size <- ggpubr_line_size
          tip_length <- ggpubr_tip_length
          v_just <- ggpubr_vjust + 0.6

          # Split data by symbol type to apply different sizes (same as regular bar)
          asterisk_data <- bracket_data[bracket_data$p.signif != "ns", ]
          ns_data <- bracket_data[bracket_data$p.signif == "ns", ]

          # Add asterisk symbols with full size
          if (nrow(asterisk_data) > 0) {
            p <- p + stat_pvalue_manual(asterisk_data,
                                       label = "p.signif",
                                       xmin = "xmin",
                                       xmax = "xmax",
                                       y.position = "y.position",
                                       size = bracket_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = v_just,
                                       hjust = 0.5,
                                       step.increase = 0,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
          }

          # Add n.s. symbols with smaller size and higher position
          if (nrow(ns_data) > 0) {
            smaller_size <- bracket_size * 0.7
            ns_vjust <- v_just - 0.5
            p <- p + stat_pvalue_manual(ns_data,
                                       label = "p.signif",
                                       xmin = "xmin",
                                       xmax = "xmax",
                                       y.position = "y.position",
                                       size = smaller_size,
                                       size.line = line_size,
                                       tip.length = tip_length,
                                       vjust = ns_vjust,
                                       hjust = 0.5,
                                       step.increase = 0,
                                       bracket.nudge.y = 0,
                                       bracket.shorten = 0,
                                       remove.bracket = FALSE)
          }
        }

        # Store statistical results in global variable for JavaScript to retrieve
        if (length(stat_text_results) > 0) {
          grouped_bar_stat_results <<- paste(stat_text_results, collapse="\\n\\n")
          cat("\\n=== STATISTICAL RESULTS SUMMARY ===\\n")
          cat(grouped_bar_stat_results, "\\n")
          cat("===================================\\n")
          cat("DEBUG: Stored", nchar(grouped_bar_stat_results), "characters in grouped_bar_stat_results\\n")
        } else {
          cat("DEBUG: No stat_text_results to store\\n")
          grouped_bar_stat_results <<- ""
        }
        }, error = function(e) {
          grouped_bar_stat_results <<- paste(grouped_bar_stat_results, "\\nERROR in statistical analysis:", e$message)
        })
      }
      
      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text, 
                      show_title, show_x_label, show_y_label,
                      x_scale, y_scale,
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      add_statistics, statistical_test)
    }

    # IC50 Dose-Response curve fitting function
    sato_ic50 <- function(dat, x_col=1, y_col=2,
                          curve_color="#2563eb", curve_width=1.5,
                          point_size=3, point_color="#1f2937", point_alpha=1,
                          ic50_line_color="#dc2626",
                          show_ic50_value=TRUE, show_ic50_line=TRUE, show_confidence_band=FALSE,
                          data_display="all_points",
                          target_font="Arial", title_weight="plain", axis_title_weight="plain", axis_text_weight="plain",
                          title_size=14, x_axis_title_size=12, y_axis_title_size=12,
                          x_axis_text_size=10, y_axis_text_size=10, legend_text_size=16,
                          title_text="IC50 Dose-Response", x_text="Concentration", y_text="Response (%)",
                          show_title=TRUE, show_x_label=TRUE, show_y_label=TRUE,
                          theme_name="minimal",
                          x_axis_rotation=0, y_axis_rotation=0,
                          x_axis_hjust=0.5, x_axis_vjust=0.5,
                          y_axis_hjust=0.5, y_axis_vjust=0.5) {

      # Prepare data
      plot_data <- data.frame(
        conc = as.numeric(dat[[x_col]]),
        response = as.numeric(dat[[y_col]])
      )
      plot_data <- plot_data[complete.cases(plot_data), ]
      plot_data <- plot_data[plot_data$conc > 0, ]  # Remove zero/negative concentrations for log scale

      cat("IC50 Analysis: ", nrow(plot_data), " data points\\n")
      cat("Data display mode: ", data_display, "\\n")
      cat("Concentration range: ", min(plot_data$conc), " - ", max(plot_data$conc), "\\n")
      cat("Response range: ", min(plot_data$response), " - ", max(plot_data$response), "\\n")

      # 4-parameter logistic model (4PL)
      # y = Bottom + (Top - Bottom) / (1 + (x/IC50)^Hill)
      ic50_value <- NA
      fit <- NULL
      fit_success <- FALSE

      tryCatch({
        # Initial parameter estimates
        bottom_init <- min(plot_data$response)
        top_init <- max(plot_data$response)
        ic50_init <- median(plot_data$conc)
        hill_init <- 1

        # Try fitting with nls
        fit <- nls(
          response ~ Bottom + (Top - Bottom) / (1 + (conc / IC50)^Hill),
          data = plot_data,
          start = list(Bottom = bottom_init, Top = top_init, IC50 = ic50_init, Hill = hill_init),
          algorithm = "port",
          lower = c(0, 0, min(plot_data$conc)/10, 0.1),
          upper = c(max(plot_data$response), max(plot_data$response)*1.5, max(plot_data$conc)*10, 10),
          control = list(maxiter = 500, warnOnly = TRUE)
        )

        ic50_value <- coef(fit)["IC50"]
        fit_success <- TRUE
        cat("IC50 fitted successfully: ", ic50_value, "\\n")

      }, error = function(e) {
        cat("4PL fitting failed:", e$message, "\\n")
        cat("Trying simpler model...\\n")
      })

      # If 4PL fails, try 3-parameter model (fixed Bottom=0)
      if (!fit_success) {
        tryCatch({
          top_init <- max(plot_data$response)
          ic50_init <- median(plot_data$conc)
          hill_init <- 1

          fit <- nls(
            response ~ Top / (1 + (conc / IC50)^Hill),
            data = plot_data,
            start = list(Top = top_init, IC50 = ic50_init, Hill = hill_init),
            algorithm = "port",
            lower = c(0, min(plot_data$conc)/10, 0.1),
            upper = c(max(plot_data$response)*1.5, max(plot_data$conc)*10, 10),
            control = list(maxiter = 500, warnOnly = TRUE)
          )

          ic50_value <- coef(fit)["IC50"]
          fit_success <- TRUE
          cat("3PL IC50 fitted: ", ic50_value, "\\n")

        }, error = function(e) {
          cat("3PL fitting also failed:", e$message, "\\n")
        })
      }

      # Generate curve data for plotting
      curve_data <- NULL
      if (fit_success && !is.null(fit)) {
        conc_seq <- 10^seq(log10(min(plot_data$conc)), log10(max(plot_data$conc)), length.out = 500)
        curve_data <- data.frame(conc = conc_seq)
        curve_data$response <- predict(fit, newdata = curve_data)
      }

      # Create base plot based on data_display mode
      if (data_display == "mean_sd" || data_display == "mean_se") {
        # Calculate mean and error for each concentration
        summary_data <- aggregate(response ~ conc, data = plot_data, FUN = function(x) {
          c(mean = mean(x), sd = sd(x), se = sd(x)/sqrt(length(x)), n = length(x))
        })
        summary_data <- do.call(data.frame, summary_data)
        colnames(summary_data) <- c("conc", "mean", "sd", "se", "n")

        # Determine which error type to use
        if (data_display == "mean_sd") {
          summary_data$error <- summary_data$sd
          cat("Showing Mean +/- SD\\n")
        } else {
          summary_data$error <- summary_data$se
          cat("Showing Mean +/- SE\\n")
        }

        cat("Unique concentrations: ", nrow(summary_data), "\\n")

        # Create plot with mean points and error bars
        p <- ggplot(summary_data, aes(x = conc, y = mean)) +
          geom_errorbar(aes(ymin = mean - error, ymax = mean + error),
                       width = 0.1, linewidth = 0.5, color = point_color) +
          geom_point(size = point_size, color = point_color, alpha = point_alpha)

        # Use summary data for curve fitting as well (weighted by n)
        fit_data <- plot_data  # Still use all data for fitting
      } else {
        # Show all data points
        p <- ggplot(plot_data, aes(x = conc, y = response)) +
          geom_point(size = point_size, color = point_color, alpha = point_alpha)
        fit_data <- plot_data
      }

      # Add fitted curve if successful
      if (!is.null(curve_data)) {
        p <- p + geom_line(data = curve_data, aes(x = conc, y = response),
                          color = curve_color, linewidth = curve_width)

        # Add confidence band if requested
        if (show_confidence_band && fit_success) {
          tryCatch({
            pred_se <- predict(fit, newdata = curve_data, se.fit = TRUE)
            if (!is.null(pred_se$se.fit)) {
              curve_data$lower <- curve_data$response - 1.96 * pred_se$se.fit
              curve_data$upper <- curve_data$response + 1.96 * pred_se$se.fit
              p <- p + geom_ribbon(data = curve_data, aes(ymin = lower, ymax = upper),
                                  alpha = 0.2, fill = curve_color)
            }
          }, error = function(e) {
            cat("Could not compute confidence band\\n")
          })
        }
      }

      # Add IC50 reference lines if requested and IC50 was calculated
      if (show_ic50_line && fit_success && !is.na(ic50_value)) {
        # Calculate response at IC50 (should be ~50% between Bottom and Top)
        y_at_ic50 <- predict(fit, newdata = data.frame(conc = ic50_value))

        p <- p +
          geom_vline(xintercept = ic50_value, linetype = "dashed", color = ic50_line_color, linewidth = 0.8) +
          geom_hline(yintercept = y_at_ic50, linetype = "dashed", color = "gray50", linewidth = 0.5)
      }

      # Add IC50 value annotation if requested
      if (show_ic50_value && fit_success && !is.na(ic50_value)) {
        ic50_label <- sprintf("IC50 = %.3g", ic50_value)
        p <- p + annotate("text", x = max(plot_data$conc) * 0.7, y = max(plot_data$response) * 0.95,
                         label = ic50_label, hjust = 1, vjust = 1, size = 5, fontface = "bold")
      }

      # Log scale for x-axis
      p <- p + scale_x_log10()

      # Store IC50 results globally for export to Excel
      ic50_result <<- ifelse(fit_success, ic50_value, NA)

      # Store curve data for export
      if (fit_success && !is.null(curve_data)) {
        ic50_curve_data <<- curve_data
        ic50_model_params <<- if (!is.null(fit)) as.list(coef(fit)) else list()
        cat("Stored curve data with", nrow(curve_data), "points for export\\n")
      } else {
        ic50_curve_data <<- NULL
        ic50_model_params <<- list()
      }

      sato_apply_theme(p, target_font, title_weight, axis_title_weight, axis_text_weight,
                      title_size, x_axis_title_size, y_axis_title_size, x_axis_text_size, y_axis_text_size, legend_text_size,
                      title_text, x_text, y_text,
                      show_title, show_x_label, show_y_label,
                      "linear", "linear",  # Don't apply log transform in theme (already done)
                      theme_name,
                      x_axis_rotation, y_axis_rotation,
                      x_axis_hjust, x_axis_vjust,
                      y_axis_hjust, y_axis_vjust,
                      FALSE, "auto")  # No statistics for IC50 plot
    }
  `);

  webrReady = true;
  setStatus("Ready");
}

// ======================= データ取得・ピン留め・webR保存 完全版 =======================

// 直近で成功したセル範囲の「位置」を保存（シート名とアドレス）
let pinnedRange = null; // { sheetName, address, addressLocal }

// --- ヘルパー: 2D配列 -> CSV（RFC準拠の最低限: ダブルクォート/改行/カンマをエスケープ）
function toCSV(values) {
  const esc = (v) => {
    if (v === null || v === undefined) return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    const s = String(v);
    // Quote if contains: comma, quote, newline, OR special chars that might break R parsing
    // Include: leading hyphen, parentheses, brackets, special symbols
    const needsQuoting = /[",\n()[\]{}|<>@#$%&*+=;:\\/?]/.test(s) || /^[-+]/.test(s);
    if (needsQuoting) {
      const escaped = s.replace(/"/g, '""');
      const result = `"${escaped}"`;
      if (s.includes('"')) {
        console.log(`CSV escape: "${s}" -> ${result}`);
      }
      return result;
    }
    return s;
  };
  const csv = (values || []).map(row => row.map(esc).join(",")).join("\n") + "\n";
  console.log("📄 Generated CSV (first 500 chars):", csv.substring(0, 500));
  return csv;
}

// --- ヘルパー: 数値判定（"1e-3" や ".5" なども OK にする）
function isNumericLike(x) {
  if (x === null || x === undefined) return false;
  const s = String(x).trim();
  if (s === "") return false;
  // 先頭に=が来るExcelの式やテキストっぽい値は除外
  if (s.startsWith("=")) return false;
  return /^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/.test(s);
}

// SVG文字列 → PNG(Base64, prefixなし) へ変換（静的 import 版）
async function svgToBase64Png(svgText, wpx, hpx, fontName) {
  // Webフォントがある場合は読み込み完了を待つ
  if (document.fonts?.ready) { try { await document.fonts.ready; } catch {} }

  const canvas = document.createElement("canvas");
  canvas.width = wpx;
  canvas.height = hpx;

  const ctx = canvas.getContext("2d");
  // ★ canvg の既定フォントに落ちないよう、単一ファミリで固定
  if (fontName) {
    const hard = fontStack(fontName, { strict: true });
    try { if (document.fonts?.load) await document.fonts.load(`12px ${hard}`); } catch {}
    ctx.font = `12px ${hard}`;
    ctx.textBaseline = "alphabetic";
    ctx.textAlign = "left";
  }

  const v = await Canvg.fromString(ctx, svgText, {
    ignoreMouse: true,
    ignoreAnimation: true,
  });
  await v.render();

  // "data:image/png;base64,..." の prefix を落として返す
  return canvas.toDataURL("image/png").split(",")[1];
}

// --- ヘッダー自動判定（第一行が非数値寄り && 第二行が数値寄り なら header=TRUE）
function detectHeader(rows) {
  if (!rows || rows.length === 0) return false;
  const r1 = rows[0] || [];
  const r2 = rows[1] || [];
  const cols = Math.max(r1.length, r2.length, 1);
  let r1NonNum = 0, r2Num = 0, r2NonNum = 0;
  for (let i = 0; i < cols; i++) {
    if (!isNumericLike(r1[i])) r1NonNum++;
    if (isNumericLike(r2[i])) r2Num++;
    if (!isNumericLike(r2[i])) r2NonNum++;
  }
  const r1Ratio = r1NonNum / cols;
  const r2Ratio = r2Num / cols;

  // Standard check: row1 mostly non-numeric AND row2 mostly numeric
  if (r1Ratio >= 0.5 && r2Ratio >= 0.5) return true;

  // Additional check for categorical datasets:
  // If row1 is ALL non-numeric, likely a header (data rows usually have SOME variation)
  if (r1Ratio === 1.0 && rows.length >= 2) {
    // Extra safety: check if row2 is different from row1 (not just a duplicate)
    const isDifferent = r1.some((val, i) => val !== r2[i]);
    if (isDifferent) return true;
  }

  return false;
}

// --- 現在の選択 or ピン留め範囲を読み、CSVとメタを返す（図形選択でもOK）
async function readRangeValuesWithFallback() {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;

    // 1) まず現在の選択を試す
    try {
      const sel = wb.getSelectedRange();
      sel.load(["values", "rowCount", "columnCount", "worksheet/name", "address", "addressLocal"]);
      await ctx.sync();
      
      console.log("Selected range debug:", {
        address: sel.address,
        rowCount: sel.rowCount,
        columnCount: sel.columnCount,
        valuesLength: sel.values?.length || 0
      });
      
      if (!sel.values || sel.values.length === 0) {
        throw new Error("No data in selected range. Please select a cell range with data.");
      }
      
      return {
        values: sel.values,
        sheetName: sel.worksheet.name,
        address: sel.address,
        addressLocal: sel.addressLocal,
      };
    } catch (error) {
      console.log("Selection failed, trying pinned range:", error.message);
      // 2) 図形選択などで取れない時はピン留めを使う
      if (!pinnedRange) throw new Error("Please select a cell range with data and click 'Load Data'.");
      
      const ws  = wb.worksheets.getItem(pinnedRange.sheetName);
      const rng = ws.getRange(pinnedRange.addressLocal || pinnedRange.address);
      rng.load(["values", "worksheet/name", "address", "addressLocal"]);
      await ctx.sync();
      
      if (!rng.values || rng.values.length === 0) {
        throw new Error("No data in pinned range. Please select a new data range and click 'Load Data'.");
      }
      
      return {
        values: rng.values,
        sheetName: rng.worksheet.name,
        address: rng.address,
        addressLocal: rng.addressLocal,
      };
    }
  });
}

/**
 * 現在の選択（または ピン留め範囲）を読み取り、
 * - ヘッダーを自動/固定で判定
 * - CSVにして webR の .GlobalEnv に dat として保存
 * - pinnedRange を更新
 * @param {"auto"|"yes"|"no"} headerMode  既定 "auto"
 */
async function cacheDataFromSelection(headerMode = "auto") {
  // Excelから値とアドレスを取得
  const { values, sheetName, address, addressLocal } = await readRangeValuesWithFallback();
  if (!values || values.length === 0) throw new Error("Selected range is empty.");

  // Save the values for use in other functions
  window.lastProcessedData = values;
  console.log("🔥 Saved lastProcessedData from cacheDataFromSelection:", values);
  console.log("🔥 Header[2] at load time:", values[0][2]);
  console.log("🔥 Header[2] char codes:", Array.from(String(values[0][2])).map((c,i)=>`[${i}]='${c}'(${c.charCodeAt(0)})`).slice(0,5));

  // Header detection - always treat first row as header
  let header;
  if (headerMode === "yes") header = true;
  else if (headerMode === "no") header = false;
  else header = true;  // Always treat first row as header (no auto-detection)

  // CSV 化
  const csv = toCSV(values);

  // Debug the data before sending to R
  console.log("Data loading debug:");
  console.log("Values array:", values);
  console.log("Values[0] (headers):", values[0]);
  console.log("Values length:", values.length);
  console.log("First few rows:", values.slice(0, 3));
  console.log("Generated CSV:", csv);
  console.log("JSON.stringify(csv):", JSON.stringify(csv));
  console.log("CSV length:", csv.length);
  console.log("Header detected:", header);

  // Validate CSV is not empty
  if (!csv || csv.trim().length === 0) {
    throw new Error("CSV data is empty. Please select a cell range with valid data.");
  }

  // webR に dat を保存（stringsAsFactors=FALSE, check.names=FALSE, 空文字は NA として扱う）
  await initWebR();
  
  // Check if WebR is ready after initialization
  if (!webR || !webrReady) {
    throw new Error("Initialization failed. Please refresh and try again.");
  }
  
  const r = `
    cat("=== R DATA LOADING DEBUG ===\\n")
    cat("Received CSV length:", nchar(${JSON.stringify(csv)}), "\\n")
    cat("First 200 chars of CSV:", substr(${JSON.stringify(csv)}, 1, 200), "\\n")
    
    con <- textConnection(${JSON.stringify(csv)})
    dat <- utils::read.csv(con,
                           header = ${header ? "TRUE" : "FALSE"},
                           check.names = FALSE,
                           stringsAsFactors = FALSE,
                           na.strings = c("", "NA"))
    close(con)
    assign("dat", dat, envir = .GlobalEnv)

    cat("Data loaded successfully. Rows:", nrow(dat), "Cols:", ncol(dat), "\\n")
    cat("Column names in R:\\n")
    for(i in 1:ncol(dat)) {
      cat(sprintf("  [%d] '%s' (length=%d, first char code=%d)\\n",
                  i, colnames(dat)[i], nchar(colnames(dat)[i]),
                  utf8ToInt(substr(colnames(dat)[i], 1, 1))[1]))
    }
    c(nrow(dat), ncol(dat))
  `;
  const rc = await webR.evalR(r);
  const dims = (await rc.toJs()).values || [0, 0];

  // ピン留め更新
  pinnedRange = { sheetName, address, addressLocal };

  // Optional: Notify UI
  setStatus(`Data loaded & pinned: ${address}  |  header=${header ? "TRUE" : "FALSE"}  |  rows=${dims[0]}, cols=${dims[1]}`);

  return { rows: dims[0], cols: dims[1], header, pin: pinnedRange };
}

// --- Ensure dat exists in R (if not, create from current selection → pinned range → A1)
async function ensureDat() {
  await initWebR();
  const exists = await webR.evalR(`exists("dat", envir=.GlobalEnv) && is.data.frame(get("dat", envir=.GlobalEnv)) && nrow(dat) > 0`);
  const ok = (await exists.toJs()).values?.[0] ?? false;
  if (ok) return;

  // If not present, load from selection (or pinned range)
  try {
    await cacheDataFromSelection("auto");
  } catch (e) {
    // Final fallback: Read at least one cell from A1
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const a1 = ws.getRange("A1");
      a1.load(["values", "worksheet/name", "address", "addressLocal"]);
      await ctx.sync();
      const csv = toCSV(a1.values);
      await webR.evalRVoid(`
        con <- textConnection(${JSON.stringify(csv)})
        dat <- utils::read.csv(con, header=FALSE, check.names=FALSE, stringsAsFactors=FALSE, na.strings=c("", "NA"))
        close(con)
        assign("dat", dat, envir=.GlobalEnv)
      `);
      pinnedRange = { sheetName: a1.worksheet.name, address: a1.address, addressLocal: a1.addressLocal };
      setStatus("Data not found; loaded A1 as fallback.");
    });
  }
}

// ========= フォント問題完全修正版のメインレンダリング関数 =========

async function renderPlotPngFromSelectionFixed(kind = "histogram") {
  await ensureDat();
  await initWebR();

  // Font debug (uncomment to test new fonts)
  // if (window.location.hostname === 'localhost' || window.location.href.includes('dev')) {
  //   await debugWebRFonts();
  // }

const fontFamElement = document.getElementById("fontFam");
const uiFontName = fontFamElement?.value || "Arial";

// フォント検証（UIで既に適切な名前が選択されているため）
const chosen = validateFont(uiFontName);

// ローカルフォント宣言
ensureLocalFontFace(chosen);
const loaded = await ensureFontLoaded(chosen);

// 実際に使うフォントを一本化（ロード失敗時だけ安全フォールバック）
const effectiveFont = loaded ? chosen : "Arial";

// R へはそのまま（family 名）を渡す
const rFontName = effectiveFont;

// CSS/Canvas 用スタックも effective から作る
const fontStack = buildCompleteFontStack(effectiveFont); 

  console.log({ uiFontName, effectiveFont, rFontName, loaded });
  console.log(`=== FINAL FONT DECISION: UI(${uiFontName}) -> R(${rFontName}) ===`);

  // その他のUI値取得
  const o = uiOpts();

  // Get chart type from UI FIRST (needed for column index calculation)
  const chartType = document.getElementById("chartType")?.value || "histogram";
  console.log(`Chart type selected: ${chartType}`);

  // Chart type dependent default titles
  const getDefaultTitle = (type) => {
    switch(type) {
      case 'histogram': return 'Histogram';
      case 'box': return 'Box Plot';
      case 'box_dot': return 'Box Plot + Data Points';
      case 'dot': return 'Dot Plot';
      case 'bar': return 'Bar Plot';
      case 'bar_error': return 'Bar Plot with Error Bars';
      case 'bar_error_dot': return 'Bar + Error + Data Points';
      case 'bar_grouped': return 'Grouped Bar Plot';
      case 'bar_grouped_error': return 'Grouped Bar Plot with Error';
      case 'bar_grouped_error_dot': return 'Grouped Bar + Dot Plot with Error';
      case 'line': return 'Line Plot';
      case 'ic50_dose_response': return 'Dose-Response Curve';
      default: return 'Chart';
    }
  };
  
  // Get actual column names from UI
  const selectedXColumn = document.getElementById("xColumn")?.value || "";
  const selectedYColumn = document.getElementById("yColumn")?.value || "";
  const selectedGroupColumn = document.getElementById("groupColumn")?.value || "";
  const selectedErrorColumn = document.getElementById("errorColumn")?.value || "";

  // Helper function to escape quotes in column names for R strings
  const escapeColumnName = (colName) => {
    if (!colName) return "";
    // Replace double quotes with escaped quotes, and escape backslashes
    return colName.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  };

  // Calculate column indices for ALL charts (not just grouped)
  // The CSV data contains all columns in their original Excel order
  // We need to find which column index corresponds to each selection
  let groupColIndex = 1, xColIndex = 1, yColIndex = 2, errorColIndex = 4;

  console.log("🔍 Column index calculation:");
  console.log("  chartType:", chartType);
  console.log("  Is grouped chart?", GROUPED_CHART_TYPES.includes(chartType));
  console.log("  window.lastProcessedData exists?", !!window.lastProcessedData);
  console.log("  window.lastProcessedData length:", window.lastProcessedData?.length);
  console.log("  Selected columns - Group:", selectedGroupColumn, "X:", selectedXColumn, "Y:", selectedYColumn, "Error:", selectedErrorColumn);

  if (window.lastProcessedData && window.lastProcessedData.length > 0) {
    const headers = window.lastProcessedData[0];
    console.log("  Headers from data:", headers);

    // Find column indices (R uses 1-based indexing)
    const groupIdx = headers.findIndex(h => h === selectedGroupColumn);
    const xIdx = headers.findIndex(h => h === selectedXColumn);
    const yIdx = headers.findIndex(h => h === selectedYColumn);
    const errorIdx = headers.findIndex(h => h === selectedErrorColumn);

    console.log("  Found indices (0-based) - Group:", groupIdx, "X:", xIdx, "Y:", yIdx, "Error:", errorIdx);

    if (groupIdx >= 0) groupColIndex = groupIdx + 1;  // Convert to 1-based
    if (xIdx >= 0) xColIndex = xIdx + 1;
    if (yIdx >= 0) yColIndex = yIdx + 1;
    if (errorIdx >= 0) errorColIndex = errorIdx + 1;

    console.log(`  ✅ Final indices (1-based for R): group=${groupColIndex}, x=${xColIndex}, y=${yColIndex}, error=${errorColIndex}`);
  } else {
    console.log("  ⚠️ Using default indices because lastProcessedData not available");
  }

  // For grouped bar chart, we need to get all three column names
  // This will be filled when data is loaded, but we need fallbacks
  const columnNames = {
    group: selectedGroupColumn || "Group",     // Group column
    category: selectedXColumn || "Category",    // X-axis category
    value: selectedYColumn || "Value"           // Y-axis value
  };
  
  const getDefaultYLabel = (type) => {
    switch(type) {
      case 'histogram': return 'Frequency';
      case 'box': return selectedYColumn || 'Value';
      case 'box_dot': return selectedYColumn || '';           // Use column name when available
      case 'dot': return selectedYColumn || 'Value';
      case 'bar': return selectedYColumn || 'Value';
      case 'bar_error': return selectedYColumn || '';
      case 'bar_error_dot': return selectedYColumn || '';     // Use column name when available
      case 'bar_grouped': return selectedYColumn || 'Value';  // For grouped data, Y is the numeric value
      case 'bar_grouped_error': return selectedYColumn || 'Mean';  // With error bars, usually showing means
      case 'bar_grouped_error_dot': return selectedYColumn || 'Value';  // With dots, showing individual values
      case 'line': return selectedYColumn || 'Value';
      case 'ic50_dose_response': return selectedYColumn || 'Response (%)';
      default: return selectedYColumn || 'Y';
    }
  };

  const getDefaultXLabel = (type) => {
    switch(type) {
      case 'bar_error_dot': return selectedXColumn || '';  // Use column name when available
      case 'box_dot': return selectedXColumn || '';        // Use column name when available
      case 'bar_grouped': return selectedXColumn || 'Treatment';  // For grouped data, X is usually treatment/category
      case 'bar_grouped_error': return selectedXColumn || 'Treatment';  // For grouped data with error
      case 'bar_grouped_error_dot': return selectedXColumn || 'Treatment';  // For grouped data with error and dots
      case 'ic50_dose_response': return selectedXColumn || 'Concentration';
      default: return selectedXColumn || 'Value';
    }
  };
  
  const title = o.title || getDefaultTitle(chartType);
  const xlab = o.xlab || getDefaultXLabel(chartType);  
  const ylab = o.ylab || getDefaultYLabel(chartType);
  const bins = Math.max(1, Math.min(100, o.bins || 20));
  
  // テキスト書式設定ヘルパー（フォーマット対応版）
  const formatR = (s) => formatTextForR(s);
  
  // (Removed duplicate variable definitions)
  
  // 出力サイズ設定（適切なアスペクト比に修正）
  const { width, height, units, dpi } = getExportSettings();
  const toInch = (v, u) => u === "in" ? v : u === "cm" ? v / 2.54 : u === "mm" ? v / 25.4 : v;
  const wIn = toInch(width, units);
  const hIn = toInch(height, units);
  const wpx = Math.round(wIn * dpi);
  const hpx = Math.round(hIn * dpi);
  
  console.log(`Output size: ${wIn}"×${hIn}" (${wpx}×${hpx}px at ${dpi}dpi)`);

  // chartType already defined earlier (needed for column index calculation)
  console.log(`Original labels from UI: xlab="${o.xlab}", ylab="${o.ylab}"`);
  console.log(`Final labels used: xlab="${xlab}", ylab="${ylab}"`);
  
  // Extract all needed values before template to avoid scope issues
  const titleWeight = o.titleWeight === 'bold' ? 'bold' : 'plain';
  const axisTitleWeight = o.axisTitleWeight === 'bold' ? 'bold' : 'plain';
  const axisTextWeight = o.axisTextWeight === 'bold' ? 'bold' : 'plain';
  const fillColor = o.fill || "#4C78A8";
  const strokeColor = o.stroke || "#1f2937";
  const fillAlpha = o.fillAlpha || 0.9;
  const lineWidth = o.lineWidth || 0.7;
  // Different default bar width for grouped bar charts
  const defaultBarWidth = GROUPED_CHART_TYPES.includes(chartType) ? 0.9 : 0.4;
  const barWidth = o.barWidth || defaultBarWidth;
  const dodgeWidth = o.dodgeWidth || 0.9;
  const rotation = o.rotation || '0';
  const tableStyleLabels = o.tableStyleLabels || false;
  const errorBarType = o.errorBarType || 'sd';
  const dotSize = o.dotSize || 4;
  const dotColor = o.dotColor || '#333333';
  const dotShape = o.dotShape || '16';
  const dotAlpha = o.dotAlpha || 1.0;
  const titleSize = o.titleSize || 24;
  const xAxisTitleSize = o.xAxisTitleSize || 20;
  const yAxisTitleSize = o.yAxisTitleSize || 20;
  const xAxisTextSize = o.xAxisTextSize || 18;
  const yAxisTextSize = o.yAxisTextSize || 18;
  const legendTextSize = o.legendTextSize || 16;
  const showTitle = o.showTitle ? 'TRUE' : 'FALSE';
  const showXLabel = o.showXLabel ? 'TRUE' : 'FALSE';
  const showYLabel = o.showYLabel ? 'TRUE' : 'FALSE';
  const xScale = o.xScale || 'linear';
  const yScale = o.yScale || 'linear';
  const themeName = o.theme || 'minimal';
  const dataOrder = o.dataOrder || 'default';
  const customOrderGroup = o.customOrderGroup || '';
  const customOrderCategory = o.customOrderCategory || '';
  const numGroups = o.numGroups || 2;
  const groupColors = o.groupColors || ["#4C78A8", "#E15759", "#57C4AD", "#E9C46A", "#F76C6C", "#A8DADC"];
  
  // Axis text positioning (with debugging)
  const xAxisRotation = o.xAxisRotation !== undefined ? o.xAxisRotation : 0;
  const yAxisRotation = o.yAxisRotation !== undefined ? o.yAxisRotation : 0;  
  const xAxisHjust = o.xAxisHjust !== undefined ? o.xAxisHjust : 0.5;
  const xAxisVjust = o.xAxisVjust !== undefined ? o.xAxisVjust : 0.5;
  const yAxisHjust = o.yAxisHjust !== undefined ? o.yAxisHjust : 0.5;
  const yAxisVjust = o.yAxisVjust !== undefined ? o.yAxisVjust : 0.5;
  
  // Statistical analysis parameters
  const addStatistics = o.addStatistics || false;
  const statisticalTest = o.statisticalTest || 'auto';
  const varianceTest = o.varianceTest || 'levene';
  const postHocTest = o.postHocTest || 'tukey';
  const dunnettControl = o.dunnettControl || '';
  const statSymbolType = o.statSymbolType || 'stars';
  const customSymbol05 = o.customSymbol05 || '*';
  const customSymbol01 = o.customSymbol01 || '**';
  const customSymbol001 = o.customSymbol001 || '***';
  const customSymbolNS = o.customSymbolNS || 'ns';

  // Debug logging for custom symbols being passed to R
  console.log('🔥🔥 R code generation - stat_symbol_type:', statSymbolType);
  console.log('🔥🔥 R code generation - customSymbol05:', customSymbol05);
  console.log('🔥🔥 R code generation - customSymbol01:', customSymbol01);
  console.log('🔥🔥 R code generation - customSymbol001:', customSymbol001);
  console.log('🔥🔥 R code generation - customSymbolNS:', customSymbolNS);
  const showMainStatSymbol = o.showMainStatSymbol || false;
  const showPairwiseComparisons = o.showPairwiseComparisons || false;
  
  // Use the symbol size control
  let statSymbolSize = Number(o.statSymbolSize) || 7;
  if (!statSymbolSize || isNaN(statSymbolSize) || statSymbolSize < 6 || statSymbolSize > 20) {
    statSymbolSize = 7; // Safe default
  }
  
  // Additional safety check - ensure it's a plain number
  statSymbolSize = Math.round(statSymbolSize);
  console.log("Final statSymbolSize after validation:", statSymbolSize, "type:", typeof statSymbolSize);

  // ggpubr statistical bracket parameters (Symbol size uses existing statSymbolSize)
  const statLineSize = Number(o.statLineSize) || 1.0;
  const statTipLength = Number(o.statTipLength) || 0.04;
  const statVjust = Number(o.statVjust) || -0.3;

  console.log("🔥 GGPUBR Parameters:", {
    statSymbolSize: statSymbolSize,
    statLineSize: statLineSize,
    statTipLength: statTipLength,
    statVjust: statVjust
  });

  // Comparison mode settings
  const comparisonMode = o.comparisonMode || "significant";
  const customComparisons = o.customComparisons || [];
  const customPositions = o.customPositions || {};
  console.log("Comparison mode:", comparisonMode, "Custom comparisons:", customComparisons, "Custom positions:", customPositions);

  // VBracket legend settings (for 3+ groups line plots)
  const vbracketTimepoint = o.vbracketTimepoint || "";
  const vbracketPosition = "custom";  // Always use manual X/Y positioning
  const vbracketX = Number(o.vbracketX) || 0.05;
  const vbracketY = Number(o.vbracketY) || 0.99;
  const vbracketTextSize = Number(o.vbracketTextSize) || 14;
  const vbracketSigSize = Number(o.vbracketSigSize) || 20;
  const vbracketMargin = Number(o.vbracketMargin) || 0.06;
  const vbracketLineWidth = Number(o.vbracketLineWidth) || 3;
  // Manual override parameters with defaults
  const vbracketLegendLineLength = o.vbracketLegendLineLength ? Number(o.vbracketLegendLineLength) : 0.05;
  const vbracketLegendLineWidth = o.vbracketLegendLineWidth ? Number(o.vbracketLegendLineWidth) : 2;
  const vbracketItemSpacing = o.vbracketItemSpacing ? Number(o.vbracketItemSpacing) : 0.1;
  const vbracketBracketLayerSpacing = o.vbracketBracketLayerSpacing ? Number(o.vbracketBracketLayerSpacing) : null;
  console.log("VBracket settings:", { vbracketTimepoint, vbracketPosition, vbracketX, vbracketY, vbracketTextSize, vbracketSigSize, vbracketMargin, vbracketLineWidth, vbracketLegendLineLength, vbracketLegendLineWidth, vbracketItemSpacing, vbracketBracketLayerSpacing });

  // IC50 Analysis settings
  const enableIC50Analysis = document.getElementById("enableIC50Analysis")?.checked || false;
  const showIC50Value = document.getElementById("showIC50Value")?.checked ?? true;
  const showIC50Line = document.getElementById("showIC50Line")?.checked ?? true;
  const showConfidenceBand = document.getElementById("showConfidenceBand")?.checked || false;
  const ic50CurveColor = document.getElementById("ic50CurveColor")?.value || "#2563eb";
  const ic50CurveWidth = Number(document.getElementById("ic50CurveWidth")?.value) || 1.5;
  const ic50LineColor = document.getElementById("ic50LineColor")?.value || "#dc2626";
  const ic50PointSize = Number(document.getElementById("ic50PointSize")?.value) || 3;
  const ic50PointColor = document.getElementById("ic50PointColor")?.value || "#1f2937";
  const ic50PointAlpha = Number(document.getElementById("ic50PointAlpha")?.value) ?? 1;
  const ic50DataDisplay = document.getElementById("ic50DataDisplay")?.value || "all_points";
  console.log("IC50 settings:", { enableIC50Analysis, showIC50Value, showIC50Line, showConfidenceBand, ic50CurveColor, ic50CurveWidth, ic50LineColor, ic50PointSize, ic50PointColor, ic50PointAlpha, ic50DataDisplay });

  // Convert to JSON strings for R
  const customComparisonsJSON = JSON.stringify(customComparisons);
  const customPositionsJSON = JSON.stringify(customPositions);

  // Convert JavaScript boolean to R boolean format
  const addStatisticsR = addStatistics ? 'TRUE' : 'FALSE';
  
  // Debug log
  console.log("Axis positioning:", { xAxisRotation, yAxisRotation, xAxisHjust, xAxisVjust, yAxisHjust, yAxisVjust });
  console.log("Statistical analysis:", { addStatistics, statisticalTest, statSymbolSize });
  console.log("DEBUG: statSymbolSize type:", typeof statSymbolSize, "value:", statSymbolSize);
  console.log("Checkbox element:", el("addStatistics"));
  console.log("Checkbox checked state:", el("addStatistics")?.checked);
  console.log("Chart type:", chartType);
  
  // Add debug message that will appear in R output
  if (addStatistics) {
    console.log("*** JAVASCRIPT: Statistical analysis is ENABLED - should see debug messages in R ***");
  } else {
    console.log("*** JAVASCRIPT: Statistical analysis is DISABLED ***");
  }
  
  // Create R vector string for colors
  const rColorVector = `c("${groupColors.slice(0, numGroups).join('", "')}")`;
  
  // Format labels before template
  const formattedTitle = formatR(title);
  const formattedXlab = formatR(xlab);
  const formattedYlab = formatR(ylab);

  // Critical debugging before template creation
  console.log("=== TEMPLATE DEBUG ===");
  console.log("statSymbolSize value:", statSymbolSize);
  console.log("statSymbolSize type:", typeof statSymbolSize);
  console.log("statSymbolSize stringified:", JSON.stringify(statSymbolSize));
  console.log("addStatistics:", addStatistics);
  console.log("addStatisticsR:", addStatisticsR);
  console.log("chartType:", chartType);
  
  // Ensure statSymbolSize is definitely a number
  if (typeof statSymbolSize !== 'number' || isNaN(statSymbolSize)) {
    console.error("CRITICAL: statSymbolSize is not a valid number:", statSymbolSize);
    statSymbolSize = 15;
  }
  
  // Create a safe string representation for template substitution
  const statSymbolSizeStr = String(statSymbolSize);
  
  // Create a guaranteed safe numeric value for template substitution
  const statSymbolSizeValue = Math.max(4, Math.min(30, Math.round(Number(statSymbolSize) || 15)));
  console.log("Final statSymbolSize value for template:", statSymbolSizeValue);
  
  // Set the symbol size as a global R variable before plotting
  const symbolSizeValue = statSymbolSizeValue;
  console.log("Setting symbol size in R global environment:", symbolSizeValue);
  

  // Add debug info to persistent panel
  addDebugInfo("🔥 Loading R code - Stats: " + addStatistics + ", Chart: " + chartType);
  
  const plotCode = `
    # Create debug log file in taskpane directory (accessible from host filesystem)
    debug_log_file <- "/Users/yoshiakisato/My Office Add-in/src/taskpane/debug.txt"

    # Clear previous log
    if (file.exists(debug_log_file)) {
      file.remove(debug_log_file)
    }

    debug_log <- function(msg) {
      tryCatch({
        cat(msg, "\\n", file = debug_log_file, append = TRUE)
      }, error = function(e) {
        cat("Log write error:", e$message, "\\n")
      })
      cat(msg, "\\n")  # Also print to console
    }

    debug_log("\\n🔥🔥🔥🔥🔥 R PLOTTING CODE STARTED - VERSION 2024-10-12-15:00 🔥🔥🔥🔥🔥")
    debug_log(paste("🔥🔥🔥 Debug log file:", debug_log_file))
    debug_log(paste("🔥🔥🔥 JS Statistical analysis:", "${addStatistics}"))
    debug_log(paste("🔥🔥🔥 JS Chart type:", "${chartType}"))
    debug_log("🔥🔥🔥 FILE VERSION TIMESTAMP: 2024-10-12-15:00 🔥🔥🔥")
    cat("🔥🔥🔥 CHECKBOX TEST: showMainStatSymbol=${showMainStatSymbol ? 'TRUE' : 'FALSE'}, showPairwiseComparisons=${showPairwiseComparisons ? 'TRUE' : 'FALSE'} 🔥🔥🔥\\n")
    
    # Set symbol size as global variable
    sato_symbol_size <- ${symbolSizeValue}
    cat("JS Symbol size:", sato_symbol_size, "\\n")

    # Set post-hoc test and control group for statistics
    selected_posthoc_test <- "${postHocTest}"
    dunnett_control <- "${dunnettControl}"
    cat("JS Post-hoc test:", selected_posthoc_test, "\\n")
    cat("JS Dunnett control:", dunnett_control, "\\n")
    
    # データ確認
    if (!exists("dat") || !is.data.frame(dat) || nrow(dat) == 0) {
      stop("データが読み込まれていません")
    }
    
    # データ順序の制御
    data_order <- "${dataOrder}"
    cat("Data ordering method:", data_order, "\\n")
    
    if (data_order == "original") {
      # 元の順序を保持（最初に出現した順序でfactor化）
      chart_type <- "${chartType}"
      is_grouped <- grepl("grouped", chart_type, fixed = TRUE)

      # Always use xColIndex for X-axis column (user-selected column)
      x_col_idx <- ${xColIndex}
      cat("Original order mode. Chart type:", chart_type, "X column index:", x_col_idx, "\\n")

      # Apply original order to X-axis column
      if (ncol(dat) >= x_col_idx && (is.character(dat[[x_col_idx]]) || is.factor(dat[[x_col_idx]]))) {
        unique_levels <- unique(dat[[x_col_idx]])
        dat[[x_col_idx]] <- factor(dat[[x_col_idx]], levels = unique_levels)
        cat("Applied original order for X-axis column", x_col_idx, ". Levels:", paste(levels(dat[[x_col_idx]]), collapse = ", "), "\\n")
      }

      # For grouped charts, also apply original order to group column
      if (is_grouped) {
        group_col_idx <- ${groupColIndex}
        if (ncol(dat) >= group_col_idx && (is.character(dat[[group_col_idx]]) || is.factor(dat[[group_col_idx]]))) {
          unique_group_levels <- unique(dat[[group_col_idx]])
          dat[[group_col_idx]] <- factor(dat[[group_col_idx]], levels = unique_group_levels)
          cat("Applied original order for Group column", group_col_idx, ". Levels:", paste(levels(dat[[group_col_idx]]), collapse = ", "), "\\n")
        }
      }
    } else if (data_order == "custom") {
      # カスタム順序を適用
      chart_type <- "${chartType}"
      is_grouped <- grepl("grouped", chart_type, fixed = TRUE)
      x_col_idx <- ${xColIndex}

      cat("Custom order mode. Chart type:", chart_type, "X column index:", x_col_idx, "\\n")

      # For single-group charts: customOrderGroup contains X-axis values
      # For grouped charts: customOrderGroup contains group values, customOrderCategory contains X-axis values
      custom_order_group <- "${customOrderGroup}"
      custom_order_category <- "${customOrderCategory}"

      cat("customOrderGroup:", custom_order_group, "\\n")
      cat("customOrderCategory:", custom_order_category, "\\n")

      if (is_grouped) {
        # GROUPED CHARTS: Group column + X-axis column
        group_col_idx <- ${groupColIndex}

        # Apply custom order to Group column
        if (nchar(custom_order_group) > 0) {
          custom_levels_group <- trimws(strsplit(custom_order_group, ",")[[1]])
          cat("Custom group order requested:", paste(custom_levels_group, collapse = ", "), "\\n")

          if (ncol(dat) >= group_col_idx && (is.character(dat[[group_col_idx]]) || is.factor(dat[[group_col_idx]]))) {
            existing_levels <- unique(as.character(dat[[group_col_idx]]))
            cat("Existing group levels:", paste(existing_levels, collapse = ", "), "\\n")

            missing_levels <- setdiff(existing_levels, custom_levels_group)
            final_levels <- c(custom_levels_group[custom_levels_group %in% existing_levels], missing_levels)
            dat[[group_col_idx]] <- factor(dat[[group_col_idx]], levels = final_levels)

            cat("Applied custom order to Group column", group_col_idx, ". Final levels:", paste(levels(dat[[group_col_idx]]), collapse = ", "), "\\n")
          }
        }

        # Apply custom order to X-axis column
        if (nchar(custom_order_category) > 0) {
          custom_levels_category <- trimws(strsplit(custom_order_category, ",")[[1]])
          cat("Custom X-axis order requested:", paste(custom_levels_category, collapse = ", "), "\\n")

          if (ncol(dat) >= x_col_idx) {
            dat[[x_col_idx]] <- as.character(dat[[x_col_idx]])
            existing_levels <- unique(dat[[x_col_idx]])
            cat("Existing X-axis levels:", paste(existing_levels, collapse = ", "), "\\n")

            missing_levels <- setdiff(existing_levels, custom_levels_category)
            final_levels <- c(custom_levels_category[custom_levels_category %in% existing_levels], missing_levels)
            dat[[x_col_idx]] <- factor(dat[[x_col_idx]], levels = final_levels)

            cat("Applied custom order to X-axis column", x_col_idx, ". Final levels:", paste(levels(dat[[x_col_idx]]), collapse = ", "), "\\n")
          }
        }
      } else {
        # SINGLE-GROUP CHARTS: Only X-axis column (customOrderGroup contains X-axis values)
        if (nchar(custom_order_group) > 0) {
          custom_levels_x <- trimws(strsplit(custom_order_group, ",")[[1]])
          cat("Custom X-axis order requested (from customOrderGroup):", paste(custom_levels_x, collapse = ", "), "\\n")

          if (ncol(dat) >= x_col_idx) {
            dat[[x_col_idx]] <- as.character(dat[[x_col_idx]])
            existing_levels <- unique(dat[[x_col_idx]])
            cat("Existing X-axis levels:", paste(existing_levels, collapse = ", "), "\\n")

            missing_levels <- setdiff(existing_levels, custom_levels_x)
            final_levels <- c(custom_levels_x[custom_levels_x %in% existing_levels], missing_levels)
            dat[[x_col_idx]] <- factor(dat[[x_col_idx]], levels = final_levels)

            cat("Applied custom order to X-axis column", x_col_idx, ". Final levels:", paste(levels(dat[[x_col_idx]]), collapse = ", "), "\\n")
          }
        }
      }
    } else {
      # default: アルファベット順（Alphabetical order）
      chart_type <- "${chartType}"
      is_grouped <- grepl("grouped", chart_type, fixed = TRUE)
      x_col_idx <- ${xColIndex}

      cat("Alphabetical order mode. Chart type:", chart_type, "X column index:", x_col_idx, "\\n")

      # Apply alphabetical order to X-axis column
      if (ncol(dat) >= x_col_idx && (is.character(dat[[x_col_idx]]) || is.factor(dat[[x_col_idx]]))) {
        sorted_levels <- sort(unique(as.character(dat[[x_col_idx]])))
        dat[[x_col_idx]] <- factor(dat[[x_col_idx]], levels = sorted_levels)
        cat("Applied alphabetical order for X-axis column", x_col_idx, ". Levels:", paste(levels(dat[[x_col_idx]]), collapse = ", "), "\\n")
      }

      # For grouped charts, also apply alphabetical order to group column
      if (is_grouped) {
        group_col_idx <- ${groupColIndex}
        if (ncol(dat) >= group_col_idx && (is.character(dat[[group_col_idx]]) || is.factor(dat[[group_col_idx]]))) {
          sorted_group_levels <- sort(unique(as.character(dat[[group_col_idx]])))
          dat[[group_col_idx]] <- factor(dat[[group_col_idx]], levels = sorted_group_levels)
          cat("Applied alphabetical order for Group column", group_col_idx, ". Levels:", paste(levels(dat[[group_col_idx]]), collapse = ", "), "\\n")
        }
      }
    }
    
    # 直接的なフォント設定（構文エラー回避）
    cat("CRITICAL R DEBUG - Setting font to:", "${rFontName}", "\\n")
    cat("UI Font was:", "${uiFontName}", "\\n")
    cat("Normalized Font was:", "${effectiveFont}", "\\n")
    cat("Chart type:", "${chartType}", "\\n")
    
    # フォント変数を設定（エスケープ問題回避）
    target_font <- "${rFontName}"
    title_weight <- "${titleWeight}"
    axis_title_weight <- "${axisTitleWeight}"
    axis_text_weight <- "${axisTextWeight}"
    
    # フォント名をそのまま使用（マッピングなし）
    cat("Using font as-is:", target_font, "\\n")
    # WebRでも可能な限り元のフォント名を維持
    
    cat("Final target_font:", target_font, "\\n")
    cat("Title weight:", title_weight, "\\n")
    cat("Axis title weight:", axis_title_weight, "\\n")
    cat("Axis text weight:", axis_text_weight, "\\n")
    
    # システムフォントの強制設定（if systemfonts is available）
    if (requireNamespace("systemfonts", quietly = TRUE)) {
      cat("systemfonts package available, attempting font registration...\\n")
      tryCatch({
        systemfonts::register_font(
          name = target_font,
          plain = target_font
        )
        cat("Font registered successfully\\n")
      }, error = function(e) {
        cat("Font registration failed (this is normal in WebR):", e$message, "\\n")
      })
    } else {
      cat("systemfonts package not available, using system default\\n")
    }
    
    # チャートタイプに基づく分岐処理
    chart_type <- "${chartType}"
    cat("Creating", chart_type, "chart\\n")
    
    # Debug statistical analysis parameters
    cat("=== STATISTICAL ANALYSIS DEBUG ===\\n")
    cat("Statistical analysis enabled:", ${addStatistics ? 'TRUE' : 'FALSE'}, "\\n")
    cat("Statistical test method:", "${statisticalTest}", "\\n")
    cat("Data dimensions:", nrow(dat), "rows x", ncol(dat), "columns\\n")
    cat("Chart type:", chart_type, "\\n")
    if (${addStatistics ? 'TRUE' : 'FALSE'} == TRUE) {
      cat("*** STATISTICAL ANALYSIS IS ENABLED - SHOULD SEE RESULTS ***\\n")
    } else {
      cat("*** STATISTICAL ANALYSIS IS DISABLED ***\\n")
    }
    
    # Define add_statistics for use throughout the template
    add_statistics <- ${addStatistics ? 'TRUE' : 'FALSE'}

    # Call appropriate sato function based on chart type
    if (chart_type == "histogram") {
      p <- sato_hist(
        dat = dat,
        x_col = 1,
        bins = ${bins},
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}"
      )
    } else if (chart_type == "box") {
      p <- sato_box(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust}
      )

      # Add statistics if enabled by user
      if (${addStatistics ? 'TRUE' : 'FALSE'}) {
        p <- sato_add_statistics_to_plot(p, dat, ${xColIndex}, ${yColIndex}, '${statisticalTest}', sato_symbol_size, show_main_symbol=${showMainStatSymbol ? 'TRUE' : 'FALSE'}, show_pairwise=${showPairwiseComparisons ? 'TRUE' : 'FALSE'}, ggpubr_symbol_size=${statSymbolSize}, ggpubr_line_size=${statLineSize}, ggpubr_tip_length=${statTipLength}, ggpubr_vjust=${statVjust}, comparison_mode="${comparisonMode}", custom_comparisons='${JSON.stringify(customComparisons).replace(/'/g, "\\'")}', custom_positions='${JSON.stringify(customPositions).replace(/'/g, "\\'")}', posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale="${yScale}", stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }
    } else if (chart_type == "box_dot") {
      p <- sato_box_dot(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dot_size = ${dotSize},
        dot_alpha = ${dotAlpha},
        dot_color = "${dotColor}",
        dot_shape = ${dotShape},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust}
      )

      # Add statistics if enabled by user
      if (${addStatistics ? 'TRUE' : 'FALSE'}) {
        p <- sato_add_statistics_to_plot(p, dat, ${xColIndex}, ${yColIndex}, '${statisticalTest}', sato_symbol_size, show_main_symbol=${showMainStatSymbol ? 'TRUE' : 'FALSE'}, show_pairwise=${showPairwiseComparisons ? 'TRUE' : 'FALSE'}, ggpubr_symbol_size=${statSymbolSize}, ggpubr_line_size=${statLineSize}, ggpubr_tip_length=${statTipLength}, ggpubr_vjust=${statVjust}, comparison_mode="${comparisonMode}", custom_comparisons='${JSON.stringify(customComparisons).replace(/'/g, "\\'")}', custom_positions='${JSON.stringify(customPositions).replace(/'/g, "\\'")}', posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale="${yScale}", stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }
    } else if (chart_type == "violin") {
      p <- sato_violin(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust}
      )

      # ADD STATISTICS FOR VIOLIN (only if enabled)
      if (${addStatistics ? 'TRUE' : 'FALSE'}) {
        p <- sato_add_statistics_to_plot(p, dat, ${xColIndex}, ${yColIndex}, '${statisticalTest}', sato_symbol_size, show_main_symbol=${showMainStatSymbol ? 'TRUE' : 'FALSE'}, show_pairwise=${showPairwiseComparisons ? 'TRUE' : 'FALSE'}, ggpubr_symbol_size=${statSymbolSize}, ggpubr_line_size=${statLineSize}, ggpubr_tip_length=${statTipLength}, ggpubr_vjust=${statVjust}, comparison_mode="${comparisonMode}", custom_comparisons='${JSON.stringify(customComparisons).replace(/'/g, "\\'")}', custom_positions='${JSON.stringify(customPositions).replace(/'/g, "\\'")}', posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale="${yScale}", stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }
    } else if (chart_type == "violin_dot") {
      p <- sato_violin_dot(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dot_size = ${dotSize},
        dot_alpha = ${dotAlpha},
        dot_color = "${dotColor}",
        dot_shape = ${dotShape},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust}
      )

      # ADD STATISTICS FOR VIOLIN_DOT (only if enabled)
      if (${addStatistics ? 'TRUE' : 'FALSE'}) {
        p <- sato_add_statistics_to_plot(p, dat, ${xColIndex}, ${yColIndex}, '${statisticalTest}', sato_symbol_size, show_main_symbol=${showMainStatSymbol ? 'TRUE' : 'FALSE'}, show_pairwise=${showPairwiseComparisons ? 'TRUE' : 'FALSE'}, ggpubr_symbol_size=${statSymbolSize}, ggpubr_line_size=${statLineSize}, ggpubr_tip_length=${statTipLength}, ggpubr_vjust=${statVjust}, comparison_mode="${comparisonMode}", custom_comparisons='${JSON.stringify(customComparisons).replace(/'/g, "\\'")}', custom_positions='${JSON.stringify(customPositions).replace(/'/g, "\\'")}', posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale="${yScale}", stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }
    } else if (chart_type == "violin_grouped") {
      # Grouped violin chart - requires 3 columns: Group, Category, Value
      p <- sato_violin_grouped(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column
        fill_colors = ${rColorVector},  # Use flexible group colors
        stroke_color = "${strokeColor}",  # Add stroke color control
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dodge_width = ${dodgeWidth},
        # Pass actual column names from user selection
        group_name = "${escapeColumnName(selectedGroupColumn) || 'Group'}",
        category_name = "${escapeColumnName(selectedXColumn) || 'Category'}",
        value_name = "${escapeColumnName(selectedYColumn) || 'Value'}",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}",
        variance_test = "${varianceTest}",
        symbol_size = ${statSymbolSize},
        ggpubr_line_size = ${statLineSize},
        ggpubr_tip_length = ${statTipLength},
        ggpubr_vjust = ${statVjust},
        comparison_mode = "${comparisonMode}",
        custom_comparisons = '${JSON.stringify(customComparisons).replace(/'/g, "\\'")}',
        custom_positions = '${JSON.stringify(customPositions).replace(/'/g, "\\'")}',
        selected_posthoc_test = selected_posthoc_test,
        dunnett_control = dunnett_control,
        sato_symbol_size = ${statSymbolSize}
      )
    } else if (chart_type == "violin_grouped_dot") {
      # Grouped violin + dot chart - requires 3 columns: Group, Category, Value
      p <- sato_violin_grouped_dot(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column
        fill_colors = ${rColorVector},  # Use flexible group colors
        stroke_color = "${strokeColor}",  # Add stroke color control
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dodge_width = ${dodgeWidth},
        dot_size = ${dotSize},
        dot_alpha = ${dotAlpha},
        dot_color = "${dotColor}",
        dot_shape = ${dotShape},
        # Pass actual column names from user selection
        group_name = "${escapeColumnName(selectedGroupColumn) || 'Group'}",
        category_name = "${escapeColumnName(selectedXColumn) || 'Category'}",
        value_name = "${escapeColumnName(selectedYColumn) || 'Value'}",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}",
        variance_test = "${varianceTest}",
        symbol_size = ${statSymbolSize},
        ggpubr_line_size = ${statLineSize},
        ggpubr_tip_length = ${statTipLength},
        ggpubr_vjust = ${statVjust},
        comparison_mode = "${comparisonMode}",
        custom_comparisons = '${customComparisonsJSON}',
        custom_positions = '${customPositionsJSON}',
        selected_posthoc_test = "${postHocTest}",
        dunnett_control = "${dunnettControl}",
        sato_symbol_size = ${statSymbolSize},
        stat_symbol_type = "${statSymbolType}",
        custom_symbol_05 = "${customSymbol05}",
        custom_symbol_01 = "${customSymbol01}",
        custom_symbol_001 = "${customSymbol001}",
        custom_symbol_ns = "${customSymbolNS}"
      )
    } else if (chart_type == "dot") {
      p <- sato_dot(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        fill = "${dotColor}",
        color = "${dotColor}",
        alpha = ${dotAlpha},
        size = ${dotSize},
        shape = ${dotShape},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}"
      )
    } else if (chart_type == "bar") {
      p <- sato_bar(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}"
      )

      # Table-style Y-axis labels (left-aligned) - only for rotation=90
      cat("Table style labels:", ${tableStyleLabels ? 'TRUE' : 'FALSE'}, "Rotation:", ${rotation}, "\\n")
      if (${tableStyleLabels ? 'TRUE' : 'FALSE'} && ${rotation} == 90) {
        cat("Applying table-style labels...\\n")

        # Calculate dynamic offset based on data range
        # Use ~1-2% of max value for proper spacing regardless of scale
        y_max <- max(dat[[${yColIndex}]], na.rm=TRUE)
        label_offset <- y_max * 0.02  # 2% of maximum value

        cat("Y max:", y_max, "Label offset:", label_offset, "\\n")

        # Remove default Y-axis text and title
        p <- p + theme(
          axis.text.y = element_blank(),
          axis.title.y = element_blank(),
          plot.margin = margin(5, 5, 5, 5, "mm")  # Minimal margins
        )

        # Add custom left-aligned text labels
        # With coord_flip: original x=category (shows on Y), original y=value (shows on X)
        tryCatch({
          p <- p + geom_text(
            aes(x = factor(dat[[${xColIndex}]]), y = label_offset, label = dat[[${xColIndex}]]),
            hjust = 0,
            size = ${yAxisTextSize} / 3,  # Convert to geom_text size units
            family = target_font
          )

          # Start axis from 0 with minimal left padding to show labels at 0.1
          p <- p + scale_y_continuous(expand = expansion(mult = c(0, 0.05)))

          cat("Table-style labels added successfully\\n")
        }, error = function(e) {
          cat("Error adding table-style labels:", e$message, "\\n")
        })
      }
    } else if (chart_type == "bar_error") {
      p <- sato_bar_error(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        error_col = if(ncol(dat) >= ${errorColIndex}) ${errorColIndex} else NULL,
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}"
      )
    } else if (chart_type == "bar_error_dot") {
      cat("🔥🔥🔥 CALLING BAR ERROR DOT FUNCTION 🔥🔥🔥\\n")
      p <- sato_bar_error_dot(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        fill = "${fillColor}",
        color = "${strokeColor}",
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        error_type = "${errorBarType}",
        dot_size = ${dotSize},
        dot_alpha = ${dotAlpha},
        dot_color = "${dotColor}",
        dot_shape = ${dotShape},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = FALSE,  # DISABLE BUILT-IN STATS
        statistical_test = "${statisticalTest}"
      )
      
      # Add statistics if enabled by user
      if (${addStatistics ? 'TRUE' : 'FALSE'}) {
        cat("\\n🔥 Adding statistics to bar_error_dot plot\\n")
        p <- sato_add_statistics_to_plot(p, dat, ${xColIndex}, ${yColIndex}, '${statisticalTest}', sato_symbol_size, show_main_symbol=${showMainStatSymbol ? 'TRUE' : 'FALSE'}, show_pairwise=${showPairwiseComparisons ? 'TRUE' : 'FALSE'}, ggpubr_symbol_size=${statSymbolSize}, ggpubr_line_size=${statLineSize}, ggpubr_tip_length=${statTipLength}, ggpubr_vjust=${statVjust}, comparison_mode="${comparisonMode}", custom_comparisons='${JSON.stringify(customComparisons).replace(/'/g, "\\'")}', custom_positions='${JSON.stringify(customPositions).replace(/'/g, "\\'")}', posthoc_test=selected_posthoc_test, dunnett_control=dunnett_control, y_scale="${yScale}", stat_symbol_type="${statSymbolType}", custom_symbol_05="${customSymbol05}", custom_symbol_01="${customSymbol01}", custom_symbol_001="${customSymbol001}", custom_symbol_ns="${customSymbolNS}")
      }
    } else if (chart_type == "box_grouped") {
      # Grouped box chart - requires 3 columns: Group, Category, Value
      p <- sato_box_grouped(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column
        fill_colors = ${rColorVector},  # Use flexible group colors
        stroke_color = "${strokeColor}",  # Add stroke color control
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dodge_width = ${dodgeWidth},
        # Pass actual column names from user selection
        group_name = "${escapeColumnName(selectedGroupColumn) || 'Group'}",
        category_name = "${escapeColumnName(selectedXColumn) || 'Category'}",
        value_name = "${escapeColumnName(selectedYColumn) || 'Value'}",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}",
        variance_test = "${varianceTest}",
        symbol_size = ${statSymbolSize},
        ggpubr_line_size = ${statLineSize},
        ggpubr_tip_length = ${statTipLength},
        ggpubr_vjust = ${statVjust},
        comparison_mode = "${comparisonMode}",
        custom_comparisons = '${JSON.stringify(customComparisons).replace(/'/g, "\\'")}',
        custom_positions = '${JSON.stringify(customPositions).replace(/'/g, "\\'")}',
        selected_posthoc_test = selected_posthoc_test,
        dunnett_control = dunnett_control,
        sato_symbol_size = ${statSymbolSize}
      )
    } else if (chart_type == "box_grouped_dot") {
      # Grouped box + dot chart - requires 3 columns: Group, Category, Value
      p <- sato_box_grouped_dot(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column
        fill_colors = ${rColorVector},  # Use flexible group colors
        stroke_color = "${strokeColor}",  # Add stroke color control
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dodge_width = ${dodgeWidth},
        dot_size = ${dotSize},
        dot_alpha = ${dotAlpha},
        dot_color = "${dotColor}",
        dot_shape = ${dotShape},
        # Pass actual column names from user selection
        group_name = "${escapeColumnName(selectedGroupColumn) || 'Group'}",
        category_name = "${escapeColumnName(selectedXColumn) || 'Category'}",
        value_name = "${escapeColumnName(selectedYColumn) || 'Value'}",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}",
        variance_test = "${varianceTest}",
        symbol_size = ${statSymbolSize},
        ggpubr_line_size = ${statLineSize},
        ggpubr_tip_length = ${statTipLength},
        ggpubr_vjust = ${statVjust},
        comparison_mode = "${comparisonMode}",
        custom_comparisons = '${customComparisonsJSON}',
        custom_positions = '${customPositionsJSON}',
        selected_posthoc_test = "${postHocTest}",
        dunnett_control = "${dunnettControl}",
        sato_symbol_size = ${statSymbolSize},
        stat_symbol_type = "${statSymbolType}",
        custom_symbol_05 = "${customSymbol05}",
        custom_symbol_01 = "${customSymbol01}",
        custom_symbol_001 = "${customSymbol001}",
        custom_symbol_ns = "${customSymbolNS}"
      )
    } else if (chart_type == "line") {
      p <- sato_line(
        dat = dat,
        x_col = ${xColIndex},
        y_col = if(ncol(dat) >= ${yColIndex}) ${yColIndex} else NULL,
        color = "${strokeColor}",
        linewidth = ${lineWidth},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}",
        variance_test = "${varianceTest}"
      )
    } else if (chart_type == "line_grouped") {
      # Grouped line chart - requires 3 columns: Group, X, Value
      p <- sato_line_grouped(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column
        line_colors = ${rColorVector},
        linewidth = ${lineWidth},
        alpha = ${fillAlpha},
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}"
      )
    } else if (chart_type == "line_grouped_error") {
      # Grouped line chart with error - for PRE-CALCULATED mean + error
      # Requires 4 columns: Group, X, Mean, Error
      p <- sato_line_grouped_error(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column (Mean)
        error_col = ${errorColIndex},    # User-selected Error column
        line_colors = ${rColorVector},
        linewidth = ${lineWidth},
        alpha = ${fillAlpha},
        ribbon_alpha = 0.2,               # Transparency for error ribbon
        use_ribbon = FALSE,               # Use error bars instead of ribbon
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}"
      )
    } else if (chart_type == "line_grouped_error_raw") {
      # Grouped line chart with error - for RAW DATA with auto-calculation
      # Requires 3 columns: Group, X, Value (with replicates)
      p <- sato_line_grouped_error_raw(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column (Value)
        line_colors = ${rColorVector},
        linewidth = ${lineWidth},
        alpha = ${fillAlpha},
        ribbon_alpha = 0.2,               # Transparency for error ribbon
        error_type = "${errorBarType}",  # SD, SE, or CI95
        use_ribbon = FALSE,               # Use error bars instead of ribbon
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}",
        variance_test = "${varianceTest}",
        stat_symbol_size = ${statSymbolSize},  # User-specified symbol size
        comparison_mode = "${comparisonMode}",
        custom_comparisons = '${customComparisonsJSON}',
        custom_positions = '${customPositionsJSON}',
        # Custom symbol settings
        stat_symbol_type = "${statSymbolType}",
        custom_symbol_05 = "${customSymbol05}",
        custom_symbol_01 = "${customSymbol01}",
        custom_symbol_001 = "${customSymbol001}",
        custom_symbol_ns = "${customSymbolNS}",
        # VBracket legend settings (for 3+ groups)
        vbracket_timepoint = "${vbracketTimepoint}",
        vbracket_position = "${vbracketPosition}",
        vbracket_x = ${vbracketX},
        vbracket_y = ${vbracketY},
        vbracket_text_size = ${vbracketTextSize},
        vbracket_sig_size = ${vbracketSigSize},
        vbracket_margin = ${vbracketMargin},
        vbracket_line_width = ${vbracketLineWidth},
        vbracket_legend_line_length = ${vbracketLegendLineLength},
        vbracket_legend_line_width = ${vbracketLegendLineWidth},
        vbracket_item_spacing = ${vbracketItemSpacing},
        vbracket_bracket_layer_spacing = ${vbracketBracketLayerSpacing === null ? 'NULL' : vbracketBracketLayerSpacing},
        output_width = ${wIn},
        output_height = ${hIn}
      )
    } else if (chart_type == "bar_grouped") {
      # Grouped bar chart - requires 3 columns: Group, Category, Value
      p <- sato_bar_grouped(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column
        fill_colors = ${rColorVector},  # Use flexible group colors
        stroke_color = "${strokeColor}",  # Add stroke color control
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dodge_width = ${dodgeWidth},
        # Pass meaningful column names (will be improved by R function logic)
        group_name = "Group",  # Will use data column name if available, fallback to this
        category_name = "Treatment",
        value_name = "Value",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}"
      )
    } else if (chart_type == "bar_grouped_error") {
      # Grouped bar chart with error - requires 4 columns: Group, Category, Mean, Error
      p <- sato_bar_grouped_error(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column (Category)
        y_col = ${yColIndex},             # User-selected Y column (Mean)
        error_col = ${errorColIndex},    # User-selected Error column
        fill_colors = ${rColorVector},  # Use flexible group colors
        stroke_color = "${strokeColor}",  # Add stroke color control
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dodge_width = ${dodgeWidth},
        # Pass actual column names from user selection
        group_name = "${escapeColumnName(selectedGroupColumn) || 'Group'}",
        category_name = "${escapeColumnName(selectedXColumn) || 'Category'}",
        value_name = "${escapeColumnName(selectedYColumn) || 'Mean'}",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        title_text = ${formatR(title)},
        x_text = ${formatR(xlab)},
        y_text = ${formatR(ylab)},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust}
      )
    } else if (chart_type == "bar_grouped_error_dot") {
      # Grouped bar + dot chart with error - requires 3 columns: Group, Category, Value
      p <- sato_bar_grouped_error_dot(
        dat = dat,
        group_col = ${groupColIndex},    # User-selected Group column
        x_col = ${xColIndex},             # User-selected X column
        y_col = ${yColIndex},             # User-selected Y column
        error_type = "${errorBarType}",  # SD or SE
        fill_colors = ${rColorVector},  # Use flexible group colors
        stroke_color = "${strokeColor}",  # Add stroke color control
        alpha = ${fillAlpha},
        linewidth = ${lineWidth},
        width = ${barWidth},
        dodge_width = ${dodgeWidth},
        dot_size = ${dotSize},
        dot_alpha = ${dotAlpha},
        dot_color = "${dotColor}",
        dot_shape = ${dotShape},
        # Pass actual column names from user selection
        group_name = "${escapeColumnName(selectedGroupColumn) || 'Group'}",
        category_name = "${escapeColumnName(selectedXColumn) || 'Category'}",
        value_name = "${escapeColumnName(selectedYColumn) || 'Value'}",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        x_scale = "${xScale}",
        y_scale = "${yScale}",
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust},
        add_statistics = ${addStatistics ? 'TRUE' : 'FALSE'},
        statistical_test = "${statisticalTest}",
        variance_test = "${varianceTest}",
        symbol_size = ${statSymbolSize},
        ggpubr_line_size = ${statLineSize},
        ggpubr_tip_length = ${statTipLength},
        ggpubr_vjust = ${statVjust},
        comparison_mode = "${comparisonMode}",
        custom_comparisons = '${JSON.stringify(customComparisons).replace(/'/g, "\\'")}',
        custom_positions = '${JSON.stringify(customPositions).replace(/'/g, "\\'") }',
        stat_symbol_type = "${statSymbolType}",
        custom_symbol_05 = "${customSymbol05}",
        custom_symbol_01 = "${customSymbol01}",
        custom_symbol_001 = "${customSymbol001}",
        custom_symbol_ns = "${customSymbolNS}"
      )
    } else if (chart_type == "ic50_dose_response") {
      # IC50 Dose-Response curve fitting and plotting
      p <- sato_ic50(
        dat = dat,
        x_col = ${xColIndex},
        y_col = ${yColIndex},
        curve_color = "${ic50CurveColor}",
        curve_width = ${ic50CurveWidth},
        point_size = ${ic50PointSize},
        point_color = "${ic50PointColor}",
        point_alpha = ${ic50PointAlpha},
        ic50_line_color = "${ic50LineColor}",
        show_ic50_value = ${showIC50Value ? 'TRUE' : 'FALSE'},
        show_ic50_line = ${showIC50Line ? 'TRUE' : 'FALSE'},
        show_confidence_band = ${showConfidenceBand ? 'TRUE' : 'FALSE'},
        data_display = "${ic50DataDisplay}",
        target_font = target_font,
        title_weight = title_weight,
        axis_title_weight = axis_title_weight,
        axis_text_weight = axis_text_weight,
        title_size = ${titleSize},
        x_axis_title_size = ${xAxisTitleSize},
        y_axis_title_size = ${yAxisTitleSize},
        x_axis_text_size = ${xAxisTextSize},
        y_axis_text_size = ${yAxisTextSize},
        legend_text_size = ${legendTextSize},
        show_title = ${showTitle},
        show_x_label = ${showXLabel},
        show_y_label = ${showYLabel},
        theme_name = "${themeName}",
        x_axis_rotation = ${xAxisRotation},
        y_axis_rotation = ${yAxisRotation},
        x_axis_hjust = ${xAxisHjust},
        x_axis_vjust = ${xAxisVjust},
        y_axis_hjust = ${yAxisHjust},
        y_axis_vjust = ${yAxisVjust}
      )
    } else {
      stop("Unsupported chart type:", chart_type)
    }
    
    # Apply formatted labels after plot creation
    if (${showTitle} == TRUE) {
      p <- p + ggtitle(${formattedTitle})
    }
    if (${showXLabel} == TRUE) {
      p <- p + xlab(${formattedXlab})
    }
    if (${showYLabel} == TRUE) {
      p <- p + ylab(${formattedYlab})
    }
    
    # Apply rotation if specified
    if ("${rotation}" == "90") {
      p <- p + coord_flip(clip = "off")
    }

    # フォント確認
    cat("Plot created with font:", target_font, "\\n")

    # Dump debug log to console at the end
    if (exists("debug_log_file") && file.exists(debug_log_file)) {
      cat("\\n\\n========== DEBUG LOG START ==========\\n")
      log_contents <- readLines(debug_log_file, warn = FALSE)
      cat(paste(log_contents, collapse = "\\n"), "\\n")
      cat("========== DEBUG LOG END ==========\\n\\n")
    }

    print(p)
  `;

  // Store R code globally for export functionality
  window.lastRCode = plotCode;
  window.lastChartType = chartType;  // Store chart type for extraction
  console.log("✅ Stored R code for export (length:", plotCode.length, "chars)");
  console.log("✅ Stored chart type:", chartType);

  // Store all plot settings for accurate R code generation
  window.lastPlotSettings = {
    chartType, title, xlab, ylab,
    showTitle: showTitle === 'TRUE',
    showXLabel: showXLabel === 'TRUE',
    showYLabel: showYLabel === 'TRUE',
    fontFamily: rFontName, titleSize, xAxisTitleSize, yAxisTitleSize,
    xAxisTextSize, yAxisTextSize, legendTextSize, titleWeight, axisTitleWeight, axisTextWeight,
    themeName, fillColor, strokeColor, fillAlpha, barWidth, dodgeWidth, lineWidth,
    groupColors, dotSize, dotAlpha, dotColor, dotShape,
    xAxisRotation, yAxisRotation, xAxisHjust, xAxisVjust, yAxisHjust, yAxisVjust,
    xScale, yScale, rotation, tableStyleLabels, groupColIndex, xColIndex, yColIndex, errorColIndex,
    selectedGroupColumn, selectedXColumn, selectedYColumn, selectedErrorColumn,
    addStatistics, errorBarType, statisticalTest, varianceTest, postHocTest,
    statSymbolType, customSymbol05, customSymbol01, customSymbol001, customSymbolNS,
    statSymbolSize, statLineSize, statTipLength, statVjust,
    comparisonMode, customComparisons, customPositions,
    vbracketTimepoint, vbracketPosition, vbracketX, vbracketY, vbracketTextSize, vbracketSigSize, vbracketMargin, vbracketLineWidth,
    dataOrder, customOrderGroup, customOrderCategory, numGroups, bins, expWidth: wIn, expHeight: hIn,
    // IC50 settings
    enableIC50Analysis, showIC50Value, showIC50Line, showConfidenceBand,
    ic50CurveColor, ic50CurveWidth, ic50LineColor, ic50PointSize, ic50PointColor, ic50PointAlpha,
    ic50DataDisplay
  };
  console.log("✅ Stored all plot settings for R code generation");

  // Create subset R code (data frame + ggplot only) for Excel cell output
  window.lastRCodeSubset = generateSubsetRCodeFromData(chartType, {
    title, xlab, ylab,
    fillColor, strokeColor, fillAlpha,
    barWidth, dodgeWidth, lineWidth,
    groupColIndex, xColIndex, yColIndex,
    rColorVector, themeName
  });
  console.log("✅ Stored subset R code (length:", window.lastRCodeSubset?.length || 0, "chars)");

  let b64;
  
  // SVG経由でcanvgを使用（サイズ指定を確実に）
  try {
    const rSvg = `
      if (!requireNamespace("svglite", quietly = TRUE)) {
        webr::install("svglite")
        library(svglite)
      }
      
      tf <- tempfile(fileext = ".svg")
      
      # SVGサイズを明示的に指定（シンプル版）
      svglite::svglite(
        tf, 
        width = ${wIn}, 
        height = ${hIn}, 
        bg = "white",
        fix_text_size = FALSE,
        pointsize = 12
        # system_fonts を削除してデフォルトを使用（XMLエラー回避）
      )
      
      # フォント設定確認ログ
      cat("SVG device initialized with font: ${rFontName}\\n")
      
      tryCatch({
        ${plotCode}
      }, error = function(e) {
        cat("Plot generation error:", e$message, "\\n")
        plot.new()
        text(0.5, 0.5, paste("エラー:", e$message), cex = 1.2, family = "Arial")
      })
      
      grDevices::dev.off()
      
      # SVG読み込み
      svg_lines <- readLines(tf, warn = FALSE)
      svg_content <- paste(svg_lines, collapse = "\\n")
      
      cat("SVG generated successfully, size:", ${wIn}, "x", ${hIn}, "inches, length:", nchar(svg_content), "\\n")
      svg_content
    `;
    
    const svgResult = await webR.evalR(rSvg);
    const svgJs = await svgResult.toJs();
    let svgText = String(Array.isArray(svgJs?.values) ? svgJs.values[0] : svgJs);
    
    if (!svgText || svgText.length < 100) {
      throw new Error(`SVG generation failed: length ${svgText.length}`);
    }
    
    // Get debug info from R
    try {
      const debugResult = await webR.evalR('if(exists("manual_annotation_debug")) manual_annotation_debug else "No debug info"');
      const debugJs = await debugResult.toJs();
      const debugMsg = Array.isArray(debugJs?.values) ? debugJs.values[0] : debugJs;
      addDebugInfo("🔥 From R: " + debugMsg);
    } catch (e) {
      addDebugInfo("🔥 Could not retrieve R debug info: " + e.message);
    }
    
    console.log("Original SVG preview:", svgText.substring(0, 500));
    
    // 安全なフォント強制適用
    svgText = forceSvgFontSafe(svgText, effectiveFont);
    
    console.log("Modified SVG preview:", svgText.substring(0, 500));

    // Extract bracket positions from R for educational code generation (AFTER plot is rendered)
    try {
      // First check if bracket_export_data exists
      const existsCheck = await webR.evalR(`exists("bracket_export_data")`);
      const existsJS = await existsCheck.toJs();
      const exists = Array.isArray(existsJS?.values) ? existsJS.values[0] : existsJS;
      console.log("🔍 bracket_export_data exists in R:", exists);

      if (exists) {
        // Show what's in bracket_export_data
        const printBracketData = await webR.evalR(`
          cat("🔥 R bracket_export_data:\\n")
          print(bracket_export_data)
          bracket_export_data
        `);

        const bracketDataR = await webR.evalR(`
          jsonlite::toJSON(bracket_export_data, dataframe = "rows")
        `);
        const bracketDataJS = await bracketDataR.toJs();
        const bracketJSON = String(Array.isArray(bracketDataJS?.values) ? bracketDataJS.values[0] : bracketDataJS);
        window.lastBracketData = JSON.parse(bracketJSON);
        console.log("✅ Extracted bracket data from R:", window.lastBracketData);
        console.log("✅ Number of brackets:", window.lastBracketData.length);
        if (window.lastBracketData.length > 0) {
          console.log("✅ First bracket:", window.lastBracketData[0]);
        }
      } else {
        console.warn("⚠️ bracket_export_data does not exist in R environment");
        window.lastBracketData = null;
      }
    } catch (e) {
      console.error("❌ Error extracting bracket data:", e);
      console.error("❌ Error stack:", e.stack);
      window.lastBracketData = null;
    }

    // Final check - log what we got
    console.log("🔍 FINAL CHECK - window.lastBracketData:", window.lastBracketData);

    // canvgでPNG化（サイズを確実に指定）
    const canvas = document.createElement("canvas");
    canvas.width = wpx;
    canvas.height = hpx;
    const ctx = canvas.getContext("2d");
    
    console.log(`Canvas size set to: ${canvas.width}×${canvas.height}`);
    
    // canvg実行前にフォント確認と設定
    console.log(`Preparing canvas rendering with font: ${effectiveFont}`);
    
    // フォントスタックをcanvasコンテキストに設定
    // const fontStack = buildCompleteFontStack(effectiveFont);
    // ctx.font = `12px ${fontStack}`;
    const hard = (f) => f.includes(" ") ? `'${f}'` : f;
    ctx.font = `12px ${hard(effectiveFont)}`;
    
    // フォント確認（複数の方法で）
    if (document.fonts && document.fonts.check) {
      const testConfigs = [
        `12px "${effectiveFont}"`,
        `12px ${effectiveFont}`,
        `normal 12px "${effectiveFont}"`
      ];
      
      let fontAvailable = false;
      for (const config of testConfigs) {
        try {
          if (document.fonts.check(config)) {
            console.log(`Font "${effectiveFont}" verified for canvg with: ${config}`);
            fontAvailable = true;
            break;
          }
        } catch (e) {
          continue;
        }
      }
      
      if (!fontAvailable) {
        console.warn(`Font "${effectiveFont}" not available for canvg, using fallback`);
      }
    }
    
    // canvgにフォント設定を含めて実行（フォント優先設定）
    const v = await Canvg.fromString(ctx, svgText, {
      ignoreMouse: true,
      ignoreAnimation: true,
      DOMParser: typeof DOMParser !== 'undefined' ? DOMParser : null,
      fetch: typeof fetch !== 'undefined' ? fetch : null,
      useCORS: true,
      enableRedraw: false,
      // フォント設定を明示的に指定
      fontCallback: (family, weight, style) => {
        console.log(`DEBUG: canvg fontCallback called with: ${family}, ${weight}, ${style}`);
        // 常に選択されたフォントを返す（canvgの既定フォール  バックを無効化）
        return effectiveFont;
      },
      // 追加のcanvg設定でフォントを強制
      rootEmSize: 16,
      enableRedraw: false,
      ignoreStylesheet: false
    });
    
    // レンダリング前にcanvasのフォント設定を強制
    const fontSetting = effectiveFont.includes(" ") ? `"${effectiveFont}"` : effectiveFont;
    ctx.font = `12px ${fontSetting}`;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = 'left';
    
    console.log(`DEBUG: Canvas font set to: ${ctx.font}`);
    console.log(`DEBUG: Canvas textBaseline: ${ctx.textBaseline}`);
    
    await v.render();
    b64 = canvas.toDataURL("image/png").split(",")[1];
    
  } catch (e) {
    console.error("SVG rendering failed:", e);
    setStatus("SVG rendering error: " + e.message);
    throw e;
  }
  
  return {
    b64,
    wPt: wIn * 72,
    hPt: hIn * 72,
    meta: {
      width, height, units, dpi, wpx, hpx,
      font: {
        ui: uiFontName,
        effective: effectiveFont,
        r: rFontName,
        stack: fontStack,
        loaded: loaded
      }
    }
  };
}

// ========= テキスト書式設定サポート関数群 =========

// テキスト書式の検出と変換（R expression 対応版）
function formatTextForR(text) {
  if (!text || typeof text !== 'string') return `"${text || ''}"`;

  const trimmed = text.trim();
  if (!trimmed) return '""';

  console.log(`DEBUG formatTextForR: Input text = "${trimmed}"`);

  // 1) expression() 形式がそのまま入力された場合
  if (trimmed.startsWith('expression(') && trimmed.endsWith(')')) {
    console.log(`DEBUG: Using raw R expression: ${trimmed}`);
    return trimmed;
  }
  
  // 2) Unicode symbols (preserve as-is)
  const hasUnicodeSuper = /[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾]/.test(trimmed);
  const hasUnicodeSub = /[₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎]/.test(trimmed);
  
  if ((hasUnicodeSuper || hasUnicodeSub) && !trimmed.includes('*') && !trimmed.includes('^') && !trimmed.includes('~')) {
    console.log(`DEBUG: Unicode super/subscript detected, using as-is: ${trimmed}`);
    const escapedText = trimmed.replace(/"/g, '\\"');
    return `"${escapedText}"`;
  }
  
  // 3) Check for formatting markers
  const hasFormatting = 
    trimmed.includes('*') || 
    trimmed.includes('^') || 
    trimmed.includes('~');
  
  if (hasFormatting) {
    console.log(`DEBUG: Formatting detected, creating R expression`);
    const rExpression = parseFormattingToRExpression(trimmed);
    console.log(`DEBUG: Generated R expression: ${rExpression}`);
    return rExpression;
  }
  
  // 4) Plain text - properly escape for R string
  const escapedText = trimmed
    .replace(/\\/g, '\\\\')   // Escape backslashes first
    .replace(/"/g, '\\"')      // Escape quotes
    .replace(/\r?\n/g, '\\n'); // Escape newlines

  console.log(`DEBUG: Plain text escaped: "${escapedText}"`);
  return `"${escapedText}"`;
}

// 書式設定マーカーをR expression に変換（フォント対応版）
function parseFormattingToRExpression(text) {
  if (!text) return 'NULL';
  
  // 単純なイタリックのみの場合
  const italicOnlyMatch = text.match(/^\*([^*]+)\*$/);
  if (italicOnlyMatch && !text.includes('^') && !text.includes('~')) {
    return `expression(italic("${escapeForR(italicOnlyMatch[1])}"))`;
  }
  
  // 書式マーカーがあるかチェック
  const hasFormatting = text.includes('*') || text.includes('^') || text.includes('~');
  
  if (!hasFormatting) {
    // 書式なし - 普通の文字列
    return `"${escapeForR(text)}"`;
  }
  
  // 複数の書式を含む場合 - パーサーで処理
  const parts = [];
  let i = 0;
  
  while (i < text.length) {
    // *italic* をチェック
    if (text[i] === '*' && text.substr(i, 2) !== '**') {
      const endPos = text.indexOf('*', i + 1);
      if (endPos !== -1) {
        const italicText = text.substring(i + 1, endPos);
        parts.push(`italic("${escapeForR(italicText)}")`);
        i = endPos + 1;
        continue;
      }
    }
    
    // ^superscript^ をチェック
    if (text[i] === '^') {
      const endPos = text.indexOf('^', i + 1);
      if (endPos !== -1) {
        const superText = text.substring(i + 1, endPos);
        parts.push(`""^{${escapeForR(superText)}}`);
        i = endPos + 1;
        continue;
      }
    }
    
    // ~subscript~ をチェック
    if (text[i] === '~') {
      const endPos = text.indexOf('~', i + 1);
      if (endPos !== -1) {
        const subText = text.substring(i + 1, endPos);
        parts.push(`""[${escapeForR(subText)}]`);
        i = endPos + 1;
        continue;
      }
    }
    
    // 通常の文字を収集
    let plainText = '';
    let j = i;
    while (j < text.length && 
           text[j] !== '*' && 
           text[j] !== '^' && 
           text[j] !== '~') {
      plainText += text[j];
      j++;
    }
    
    if (plainText) {
      parts.push(`"${escapeForR(plainText)}"`);
    }
    
    i = j;
  }
  
  // 結果を結合
  if (parts.length === 0) {
    return 'NULL';
  } else if (parts.length === 1) {
    return `expression(${parts[0]})`;
  } else {
    // * 演算子を使って結合（plotmathでは * が連結演算子）
    return `expression(${parts.join(' * ')})`;
  }
}

// Unicode数学用イタリック文字への変換（限定的）
function convertToItalic(text) {
  // Unicode Mathematical Italic symbols (for single letters)
  const italicMap = {
    'A': '𝐴', 'B': '𝐵', 'C': '𝐶', 'D': '𝐷', 'E': '𝐸', 'F': '𝐹', 'G': '𝐺',
    'H': '𝐻', 'I': '𝐼', 'J': '𝐽', 'K': '𝐾', 'L': '𝐿', 'M': '𝑀', 'N': '𝑁',
    'O': '𝑂', 'P': '𝑃', 'Q': '𝑄', 'R': '𝑅', 'S': '𝑆', 'T': '𝑇', 'U': '𝑈',
    'V': '𝑉', 'W': '𝑊', 'X': '𝑋', 'Y': '𝑌', 'Z': '𝑍',
    'a': '𝑎', 'b': '𝑏', 'c': '𝑐', 'd': '𝑑', 'e': '𝑒', 'f': '𝑓', 'g': '𝑔',
    'h': 'ℎ', 'i': '𝑖', 'j': '𝑗', 'k': '𝑘', 'l': '𝑙', 'm': '𝑚', 'n': '𝑛',
    'o': '𝑜', 'p': '𝑝', 'q': '𝑞', 'r': '𝑟', 's': '𝑠', 't': '𝑡', 'u': '𝑢',
    'v': '𝑣', 'w': '𝑤', 'x': '𝑥', 'y': '𝑦', 'z': '𝑧'
  };
  
  // 全てのテキストをUnicodeイタリック文字に変換
  return text.split('').map(char => italicMap[char] || char).join('');
}

// Unicode上付き文字への変換
function convertToSuperscript(text) {
  const superscriptMap = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
    'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
    'f': 'ᶠ', 'g': 'ᵍ', 'h': 'ʰ', 'i': 'ⁱ', 'j': 'ʲ',
    'k': 'ᵏ', 'l': 'ˡ', 'm': 'ᵐ', 'n': 'ⁿ', 'o': 'ᵒ',
    'p': 'ᵖ', 'r': 'ʳ', 's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ',
    'v': 'ᵛ', 'w': 'ʷ', 'x': 'ˣ', 'y': 'ʸ', 'z': 'ᶻ'
  };
  
  return text.split('').map(char => superscriptMap[char] || char).join('');
}

// Unicode下付き文字への変換  
function convertToSubscript(text) {
  const subscriptMap = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
    'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ⱼ',
    'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ',
    'p': 'ₚ', 'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ',
    'v': 'ᵥ', 'x': 'ₓ'
  };
  
  return text.split('').map(char => subscriptMap[char] || char).join('');
}


// R 文字列用エスケープ
function escapeForR(str) {
  return str
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/\r?\n/g, '\\n');
}

// よく使われる科学記号のヘルパー関数
function getCommonScientificExpressions() {
  return {
    // よく使われるパターンの例
    'R²': 'expression(R^{2})',
    'R2': 'expression(R^{2})',
    'X²': 'expression(X^{2})',
    'H₂O': 'expression(H[2]*O)',
    'CO₂': 'expression(CO[2])',
    'pH': 'expression(italic(pH))',
    '°C': 'expression(degree*C)',
    'μg/L': 'expression(mu*g/L)',
    'mg/m²': 'expression(mg/m^{2})',
    'Temperature (°C)': 'expression(paste("Temperature (", degree, "C)"))',
  };
}

// UI入力フィールドにヒント表示を追加する関数
function addFormattingHints() {
  const titleField = document.getElementById("titleText");
  const xLabelField = document.getElementById("xLabel");  
  const yLabelField = document.getElementById("yLabel");
  
  const hint = "Formatting tips:\n" +
    "• Use buttons below for formatting\n" +
    "• Or type: *italic*, **bold**, ^super^, ~sub~\n" +
    "• Direct Unicode: R², H₂O, °C\n" +
    "• Advanced: expression(R^{2})";
    
  [titleField, xLabelField, yLabelField].forEach(field => {
    if (field) {
      field.title = hint;
      // Remove the placeholder modification as we now have buttons
    }
  });
}

// 書式設定ボタン用の関数群（グローバルに公開）
window.formatText = function(fieldId, formatType) {
  const field = document.getElementById(fieldId);
  if (!field) return;
  
  const start = field.selectionStart;
  const end = field.selectionEnd;
  const selectedText = field.value.substring(start, end);
  
  if (!selectedText) {
    // 何も選択されていない場合はプレースホルダーを挿入
    let placeholder, wrapStart, wrapEnd;
    
    switch(formatType) {
      case 'italic':
        placeholder = 'italic text';
        wrapStart = '*';
        wrapEnd = '*';
        break;
      case 'bold':
        placeholder = 'bold text';
        wrapStart = '**';
        wrapEnd = '**';
        break;
      case 'super':
        placeholder = '2';
        wrapStart = '^';
        wrapEnd = '^';
        break;
      case 'sub':
        placeholder = '2';
        wrapStart = '~';
        wrapEnd = '~';
        break;
      default:
        return;
    }
    
    const newText = wrapStart + placeholder + wrapEnd;
    insertAtCursor(field, newText);
    
    // プレースホルダーテキストを選択状態にする
    field.selectionStart = start + wrapStart.length;
    field.selectionEnd = start + wrapStart.length + placeholder.length;
  } else {
    // 選択されたテキストを書式設定で囲む
    let wrapStart, wrapEnd;
    
    switch(formatType) {
      case 'italic':
        wrapStart = '*';
        wrapEnd = '*';
        break;
      case 'bold':
        wrapStart = '**';
        wrapEnd = '**';
        break;
      case 'super':
        wrapStart = '^';
        wrapEnd = '^';
        break;
      case 'sub':
        wrapStart = '~';
        wrapEnd = '~';
        break;
      default:
        return;
    }
    
    const newText = wrapStart + selectedText + wrapEnd;
    replaceSelection(field, newText);
  }
  
  field.focus();
}

window.insertSymbol = function(fieldId, symbol) {
  const field = document.getElementById(fieldId);
  if (!field) return;
  
  insertAtCursor(field, symbol);
  field.focus();
}

function insertAtCursor(field, text) {
  const start = field.selectionStart;
  const end = field.selectionEnd;
  
  field.value = field.value.substring(0, start) + text + field.value.substring(end);
  field.selectionStart = field.selectionEnd = start + text.length;
}

function replaceSelection(field, text) {
  const start = field.selectionStart;
  const end = field.selectionEnd;
  
  field.value = field.value.substring(0, start) + text + field.value.substring(end);
  field.selectionStart = start;
  field.selectionEnd = start + text.length;
}


// ========= 既存の関数群（ヘルパー関数など）=========

function applyTemplate(o){
  const set = (id, v)=>{ if(el(id)!=null && v!=null) el(id).value = v; };
  set("titleText", o.title);
  set("xLabel", o.xlab);
  set("yLabel", o.ylab);
  set("ggtheme", o.theme);
  set("fontFam", o.family);
  set("fillColor", o.fill);
  set("fillPicker", o.fill || o.fillPicker);
  set("fillAlpha", o.fillAlpha);
  set("strokeColor", o.stroke);
  set("strokePicker", o.stroke || o.strokePicker);
  set("lineWidth", o.lineWidth);
  set("titleSize", o.titleSize);
  // Handle both old format (axisTitleSize/axisTextSize) and new format (4 separate values)
  set("xAxisTitleSize",  o.xAxisTitleSize || o.axisTitleSize);
  set("yAxisTitleSize",  o.yAxisTitleSize || o.axisTitleSize);
  set("xAxisTextSize",   o.xAxisTextSize || o.axisTextSize);
  set("yAxisTextSize",   o.yAxisTextSize || o.axisTextSize);
  set("legendTextSize", o.legendTextSize);
  set("titleWeight", o.titleWeight);
  set("axisTitleWeight", o.axisTitleWeight || o.axisWeight);
  set("axisTextWeight", o.axisTextWeight || o.axisWeight);
  set("dataOrder", o.dataOrder || "default");
  set("customOrder", o.customOrder || "");
  set("numGroups", o.numGroups || 2);
  for (let i = 0; i < 6; i++) {
    set(`groupColor${i+1}`, (o.groupColors && o.groupColors[i]) || ["#4C78A8", "#E15759", "#57C4AD", "#E9C46A", "#F76C6C", "#A8DADC"][i]);
  }
  if(o.xMin!=null) set("xMin", o.xMin);
  if(o.xMax!=null) set("xMax", o.xMax);
  if(o.yMin!=null) set("yMin", o.yMin);
  if(o.yMax!=null) set("yMax", o.yMax);
  if(o.bins!=null) set("bins", o.bins);
}

async function tplLoadFromStorage(){
  const name = (el("tplName")?.value || el("template")?.value || "").trim();
  if(!name){ setStatus("Please specify template name"); return; }
  try{
    const raw = await OfficeRuntime.storage.getItem(`sato:tpl:${name}`);
    if(!raw){ setStatus(`Template "${name}" not found`); return; }
    const o = JSON.parse(raw);
    applyTemplate(o);
    setStatus(`Template loaded (storage): "${name}" applied`);
  }catch(e){
    console.error(e); setStatus("Template load error: " + (e?.message || e));
  }
}

// Safely read numeric input
function numVal(id) {
  const el = document.getElementById(id);
  if (!el) return null;
  const v = parseFloat(el.value);
  return Number.isFinite(v) ? v : null;
}

// ---- Excel range → values ----
async function getSelectedValues() {
  return Excel.run(async (ctx) => {
    const rng = ctx.workbook.getSelectedRange();
    rng.load(["values","worksheet/name"]);
    await ctx.sync();

    // Save the values for use in other functions
    window.lastProcessedData = rng.values;
    console.log("🔥 Saved lastProcessedData:", rng.values);

    return { values: rng.values, sheetName: rng.worksheet.name };
  });
}
function headersFromValues(values){ return values?.[0]?.map(v=>String(v ?? "")) ?? []; }
function toCsv(values){ return values.map(r=>r.map(v=>String(v).replace(/\r?\n/g," ")).join(",")).join("\n"); }

// Clear old statistical results from webR to prevent stale data issues
async function clearOldStatisticalResults() {
  try {
    if (!window.webR) return;

    await window.webR.evalR(`
      # Clear all statistical result variables
      if (exists("stat_results")) rm(stat_results)
      if (exists("grouped_bar_stat_results")) rm(grouped_bar_stat_results)
      if (exists("line_plot_stat_results")) rm(line_plot_stat_results)
      if (exists("stat_text_results")) rm(stat_text_results)
      cat("Statistical results cleared\\n")
    `);
    console.log("✅ Cleared old statistical results from webR");
  } catch (e) {
    console.log("Note: Could not clear statistical results (webR may not be initialized yet)");
  }
}

// ---- Load headers ----
async function loadHeadersFromSelection(){
  try {
    // Clear old statistical results to prevent stale data issues
    await clearOldStatisticalResults();

    // Force refresh the data cache first
    await cacheDataFromSelection("auto");
    
    const { values } = await getSelectedValues();
    const headers = headersFromValues(values);
    if (!headers.length) { setStatus("Headers not found"); return; }
    const xSel = el("xColumn"), ySel = el("yColumn"), groupSel = el("groupColumn"), errorSel = el("errorColumn");
    xSel.innerHTML=""; ySel.innerHTML="";
    if (groupSel) groupSel.innerHTML="";
    if (errorSel) errorSel.innerHTML="";
    headers.forEach(h => {
      xSel.add(new Option(h,h));
      ySel.add(new Option(h,h));
      if (groupSel) groupSel.add(new Option(h,h));
      if (errorSel) errorSel.add(new Option(h,h));
    });

    // Numeric column to Y, other column to X
    const rows = values.slice(1);
    const isNum = (i)=>rows.every(r=>r[i]==="" || Number.isFinite(+r[i]));
    let yIdx = headers.findIndex((_,i)=>isNum(i)); if (yIdx<0) yIdx=0;
    let xIdx = yIdx===0 && headers.length>1 ? 1 : 0;

    // For grouped charts: Auto-detect columns intelligently
    const chartType = el("chartType")?.value;

    if (GROUPED_CHART_TYPES.includes(chartType) && headers.length >= 3) {
      // For grouped charts with 3+ columns
      // Find the numeric column for Y axis (use the last numeric column if multiple exist)
      let yIdx = -1;
      for (let i = headers.length - 1; i >= 0; i--) {
        if (isNum(i)) {
          yIdx = i;
          break;
        }
      }
      if (yIdx < 0) yIdx = headers.length - 1; // Fallback to last column

      ySel.value = headers[yIdx];  // Numeric column = Value

      // For Group and X columns, use the first two non-Y columns
      const nonYIndices = headers.map((_, i) => i).filter(i => i !== yIdx);
      if (groupSel && nonYIndices.length >= 2) {
        groupSel.value = headers[nonYIndices[1]];  // 2nd non-numeric column = Group
      } else if (groupSel && nonYIndices.length >= 1) {
        groupSel.value = headers[nonYIndices[0]];  // 1st non-numeric column = Group
      }
      if (nonYIndices.length >= 2) {
        xSel.value = headers[nonYIndices[nonYIndices.length - 1]];  // Last non-numeric column before Y = X/Category
      } else if (nonYIndices.length >= 1) {
        xSel.value = headers[nonYIndices[0]];  // 1st non-numeric column = X
      }

      // For charts with error column (3+ or 4+ columns): Error=last column
      if ((chartType === "line_grouped_error" || chartType === "bar_grouped_error") && headers.length >= 4) {
        if (errorSel) errorSel.value = headers[3];  // 4th column = Error (Group, X, Y, Error)
      } else if (chartType === "bar_error" && headers.length >= 3) {
        if (errorSel) errorSel.value = headers[2];  // 3rd column = Error (X, Y, Error)
      }
    } else {
      // For regular charts
      xSel.value=headers[xIdx];
      ySel.value=headers[yIdx];
      if (groupSel && headers.length > 0) groupSel.value = headers[0];  // Default to first column
    }

    // Auto-detect number of groups and update UI (after column selectors are populated)
    const actualGroups = getActualGroupNames();
    if (actualGroups && actualGroups.length >= 2) {
      const numGroupsInput = document.getElementById("numGroups");
      if (numGroupsInput && parseInt(numGroupsInput.value) !== actualGroups.length) {
        numGroupsInput.value = actualGroups.length;
        // Trigger change event to show/hide color inputs
        numGroupsInput.dispatchEvent(new Event('change'));
        console.log("🔥 Auto-updated numGroups to:", actualGroups.length);
      }
    }

    // Update comparison checkboxes when new data is loaded
    populateComparisonCheckboxes();
    console.log("🔥 Updated comparison checkboxes after loading new data 🔥");

    // Update Dunnett/Steel control group dropdown
    const dataType = document.getElementById("dataTypeSelect")?.value || "parametric";
    const postHocTest = dataType === "nonparametric"
      ? (document.getElementById("postHocTestNonparam")?.value || "dunn")
      : (document.getElementById("postHocTest")?.value || "tukey");
    if (postHocTest === "dunnett" || postHocTest === "steel") {
      populateDunnettControl();
    }

    // Detect and store groups for custom order
    await detectAndStoreGroups();

    // Populate vbracket timepoint dropdown with unique X values
    populateVbracketTimepoints();

    setStatus(`Load complete: n=${rows.length}`);
  } catch (error) {
    setStatus(`Data load error: ${error.message}`);
  }
}

// Normalize any color format (color name/red, #rgb/#rrggbb, rgb()/rgba()) → #RRGGBB
function toHexColor(input) {
  const s = (input || "").trim();
  if (!s) return null;

  // すでに #RGB / #RRGGBB
  if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(s)) {
    if (s.length === 4) return ("#" + s[1]+s[1] + s[2]+s[2] + s[3]+s[3]).toUpperCase();
    return s.toUpperCase();
  }

  // キャンバスで CSS 色解決（色名・rgb/rgba など全部これでOK）
  const c = document.createElement("canvas").getContext("2d");
  c.fillStyle = "#000";            // いったん既知の色
  c.fillStyle = s;                 // 不正なら変化しない
  const v = c.fillStyle;           // 例: "rgb(255, 0, 0)" または "#rrggbb"
  if (/^#[0-9a-f]{6}$/i.test(v)) return v.toUpperCase();

  const m = v.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/i);
  if (m) {
    return (
      "#" +
      [m[1], m[2], m[3]]
        .map(n => Math.max(0, Math.min(255, +n)).toString(16).padStart(2, "0"))
        .join("")
    ).toUpperCase();
  }
  return null;
}

// 1組の text/picker を #RRGGBB で双方向同期
function attachPickerSync(textId, pickerId) {
  const txt  = document.getElementById(textId);
  const pick = document.getElementById(pickerId);
  if (!txt || !pick) return;

  // 初期同期（text が空なら picker → text、常に #RRGGBB に）
  const init = toHexColor(txt.value || pick.value) || "#000000";
  txt.value = init; pick.value = init;

  // picker → text
  pick.addEventListener("input", () => {
    const hex = toHexColor(pick.value);
    if (hex) txt.value = hex;
  });

  // text（色名/rgba/…OK）→ 正規化して picker に反映
  const syncFromText = () => {
    const hex = toHexColor(txt.value);
    if (hex) { txt.value = hex; pick.value = hex; }
  };
  txt.addEventListener("input",  syncFromText);
  txt.addEventListener("change", syncFromText);
}

// 置き換え：両方のペアを同期
function wireColorSync() {
  attachPickerSync("fillColor",   "fillPicker");
  attachPickerSync("strokeColor", "strokePicker");
}
// R はPNGを書くだけ；JSでBase64化（jsonlite不要）
async function runRtoPng(csvText, rCode, _opts = {}) {
  await initWebR();

  // --- 寸法計算（px/インチ）---
  const { width, height, units, dpi } = getExportSettings();
  const toInch = (v, u) => (u === "in" ? v : u === "cm" ? v / 2.54 : u === "mm" ? v / 25.4 : v);
  const wIn = toInch(width, units), hIn = toInch(height, units);
  const wpx = Math.round(wIn * dpi), hpx = Math.round(hIn * dpi);

  // --- フォント（UI → R用単一名 / SVG用スタック）---
  const uiFamily = _opts.family ??
                   (typeof uiOpts === "function" ? (uiOpts().family || "Arial") : "Arial");

  const primaryFam = validateFont(uiFamily);    // R の base_family 想定

  // --- R：CSV を dat に読み込み → svglite で SVG 出力（PNGは使わない）---
  await webR.evalRVoid(`if (!requireNamespace("svglite", quietly=TRUE)) webr::install("svglite")`);

  const r = `
    dat <- utils::read.csv(textConnection(${JSON.stringify(csvText)}),
                           header = TRUE, check.names = FALSE)
    tf <- tempfile(fileext = ".svg")
    svglite::svglite(tf, width = ${wIn}, height = ${hIn}, bg = "white")
    try({ ${rCode} }, silent = TRUE)   # ★ rCode 内では family="${primaryFam}" を使うのが望ましい
    grDevices::dev.off()
    paste(readLines(tf, warn = FALSE), collapse = "\\n")
  `;

  const res = await webR.evalR(r);
  const js  = await res.toJs();
  const svgText = String(Array.isArray(js?.values) ? js.values[0] : js);

  // --- SVG にフォントを強制埋め込み → canvg で PNG 化 ---
  const svgFixed = forceSvgFontAggressive(svgText, primaryFam);

  // 選択フォントを明示的にロード（フォールバック回避）
  try {
    const fam = primaryFam.includes(" ") ? `'${primaryFam}'` : primaryFam;
    if (document.fonts?.load) await document.fonts.load(`12px ${fam}`);
    if (document.fonts?.ready) await document.fonts.ready;
  } catch { /* 失敗しても続行 */ }

  // ここはグローバルの svgToBase64Png を使用
  const b64 = await svgToBase64Png(svgFixed, wpx, hpx, primaryFam);
  return b64;
}

// ---- UI → オプション ----
function uiOpts(){
  // ← pickerの値をテキストのフォールバックとして使う
  const fillRaw   = (el("fillColor")?.value || el("fillPicker")?.value || "").trim();
  const strokeRaw = (el("strokeColor")?.value || el("strokePicker")?.value || "").trim();

  return {
    title:  (el("titleText")?.value || "").trim(),
    xlab:   (el("xLabel")?.value || "").trim(),
    ylab:   (el("yLabel")?.value || "").trim(),
    theme:  el("ggtheme")?.value || "minimal",
    family: el("fontFam")?.value || "Arial",

    // 色（←ここだけ変更）
    fill:   fillRaw,
    fillPicker: el("fillPicker")?.value || "",
    fillAlpha:  (+el("fillAlpha")?.value) || 1,
    stroke: strokeRaw,
    strokePicker: el("strokePicker")?.value || "",

    lineWidth:      (+el("lineWidth")?.value)      || 0.8,
    barWidth:       (+el("barWidth")?.value)       || (function() {
      const chartType = el("chartType")?.value;
      return GROUPED_CHART_TYPES.includes(chartType) ? 0.9 : 0.4;
    })(),
    dodgeWidth:     (+el("dodgeWidth")?.value)      || 0.9,
    rotation:       el("rotation")?.value           || "0",
    tableStyleLabels: el("tableStyleLabels")?.checked || false,
    errorBarType:   el("errorBarType")?.value       || "sd",
    
    // Dot settings
    dotSize:        (+el("dotSize")?.value)        || 4,
    dotColor:       (el("dotColor")?.value || el("dotColorPicker")?.value || "").trim() || "#333333",
    dotShape:       el("dotShape")?.value          || "16",
    dotAlpha:       (+el("dotAlpha")?.value)       || 1.0,
    
    titleSize:       (+el("titleSize")?.value)       || 24,
    xAxisTitleSize:  (+el("xAxisTitleSize")?.value)  || 20,
    yAxisTitleSize:  (+el("yAxisTitleSize")?.value)  || 20,
    xAxisTextSize:   (+el("xAxisTextSize")?.value)   || 18,
    yAxisTextSize:   (+el("yAxisTextSize")?.value)   || 18,
    legendTextSize:  (+el("legendTextSize")?.value)  || 16,

    titleWeight: (el("titleWeight")?.value || "plain"),
    axisTitleWeight: (el("axisTitleWeight")?.value || "plain"),
    axisTextWeight:  (el("axisTextWeight")?.value  || "plain"),
    
    // データ順序設定
    dataOrder: el("dataOrder")?.value || "default",
    customOrderGroup: getCustomOrderGroupString(),
    customOrderCategory: getCustomOrderCategoryString(),
    
    // グループ色設定（Grouped Bar Chart用）
    numGroups: (function() {
      const chartType = el("chartType")?.value || "";
      const isGroupedLineChart = ["line_grouped", "line_grouped_error", "line_grouped_error_raw"].includes(chartType);

      // For grouped line charts, auto-detect number of groups from data
      if (isGroupedLineChart && window.lastProcessedData && window.lastProcessedData.length > 1) {
        const groupColumn = el("groupColumn")?.value;
        if (groupColumn) {
          const headers = window.lastProcessedData[0];
          const groupIndex = headers.findIndex(h => h === groupColumn);
          if (groupIndex >= 0) {
            const groupData = window.lastProcessedData.slice(1).map(row => row[groupIndex]);
            const uniqueGroups = [...new Set(groupData)].filter(g => g && g.trim() !== '');
            return uniqueGroups.length || 2;
          }
        }
      }

      // For other grouped charts, use the manual UI control
      return parseInt(el("numGroups")?.value) || 2;
    })(),
    groupColors: [
      el("groupColor1")?.value || "#4C78A8",
      el("groupColor2")?.value || "#E15759",
      el("groupColor3")?.value || "#57C4AD",
      el("groupColor4")?.value || "#E9C46A",
      el("groupColor5")?.value || "#F76C6C",
      el("groupColor6")?.value || "#A8DADC"
    ],

    // 軸スケール設定
    xScale: el("xScale")?.value || "linear",
    yScale: el("yScale")?.value || "linear",

    // 表示・非表示オプション
    showTitle: el("showTitle")?.checked !== false,
    showXLabel: el("showXLabel")?.checked !== false,
    showYLabel: el("showYLabel")?.checked !== false,

    xMin: numVal("xMin"), xMax: numVal("xMax"),
    yMin: numVal("yMin"), yMax: numVal("yMax"),

    bins: Math.max(1, Math.min(100, +(el("bins")?.value || 20))),
    
    // Data ordering settings
    dataOrder: el("dataOrder")?.value || "default",
    customOrderGroup: getCustomOrderGroupString(),
    customOrderCategory: getCustomOrderCategoryString(),
    
    // Axis text positioning
    xAxisRotation: (+el("xAxisRotation")?.value) || 0,
    yAxisRotation: (+el("yAxisRotation")?.value) || 0,
    xAxisHjust: (+el("xAxisHjust")?.value) || 0.5,
    xAxisVjust: (+el("xAxisVjust")?.value) || 0.5,
    yAxisHjust: (+el("yAxisHjust")?.value) || 0.5,
    yAxisVjust: (+el("yAxisVjust")?.value) || 0.5,
    
    // Statistical analysis
    addStatistics: el("addStatistics")?.checked || false,
    statisticalTestMode: el("statisticalTestMode")?.value || "auto",
    // When manual mode, use dataType to determine test; when auto, use "auto"
    statisticalTest: (el("statisticalTestMode")?.value || "auto") === "auto"
      ? "auto"
      : (el("dataTypeSelect")?.value === "nonparametric" ? "nonparametric" : "parametric"),
    varianceTest: el("varianceTest")?.value || "levene",
    postHocTest: (el("dataTypeSelect")?.value === "nonparametric" ? el("postHocTestNonparam")?.value : el("postHocTest")?.value) || "tukey",
    dataType: el("dataTypeSelect")?.value || "parametric",
    dunnettControl: el("dunnettControl")?.value || "",
    statSymbolSize: Number(el("statSymbolSize")?.value) || 7,
    statSymbolType: el("statSymbolType")?.value || "stars",
    customSymbol05: el("customSymbol05")?.value || "*",
    customSymbol01: el("customSymbol01")?.value || "**",
    customSymbol001: el("customSymbol001")?.value || "***",
    customSymbolNS: el("customSymbolNS")?.value || "ns",
    showMainStatSymbol: false, // Removed from UI - always disabled
    showPairwiseComparisons: el("showPairwiseComparisons")?.checked !== false, // Default to TRUE

    // ggpubr statistical bracket parameters (Symbol size uses existing statSymbolSize)
    statLineSize: Number(el("statLineSize")?.value) || 1.0,
    statTipLength: Number(el("statTipLength")?.value) || 0.04,
    statVjust: Number(el("statVjust")?.value) || -0.3,

    // Comparison mode settings
    comparisonMode: document.querySelector('input[name="comparisonMode"]:checked')?.value || "significant",
    customComparisons: getSelectedCustomComparisons(),
    customPositions: getCustomBracketPositions(),

    // VBracket legend settings (for 3+ groups line plots)
    vbracketTimepoint: el("vbracketTimepoint")?.value || "",
    vbracketPosition: "custom",
    vbracketX: Number(el("vbracketX")?.value) || 0.05,
    vbracketY: Number(el("vbracketY")?.value) || 0.99,
    vbracketTextSize: Number(el("vbracketTextSize")?.value) || 14,
    vbracketSigSize: Number(el("vbracketSigSize")?.value) || 20,
    vbracketMargin: Number(el("vbracketMargin")?.value) || 0.06,
    vbracketLineWidth: Number(el("vbracketLineWidth")?.value) || 3,
    vbracketLegendLineLength: el("vbracketLegendLineLength")?.value || "0.05",
    vbracketLegendLineWidth: el("vbracketLegendLineWidth")?.value || "2",
    vbracketItemSpacing: el("vbracketItemSpacing")?.value || "0.1",
    vbracketBracketLayerSpacing: el("vbracketBracketLayerSpacing")?.value || "",
  };
}

// 直近のレンダリングをキャッシュ（挿入でも使う）
let lastRender = null;

// ========= 修正版のプレビューとExcel挿入関数 =========

// Preview with debug functionality (canvas size corrected version)
async function previewPlotWithDebug() {
  try {
    setStatus("Generating preview...");

    lastRender = await renderPlotPngFromSelectionFixed();

    // Update comparison checkboxes with actual data after rendering
    populateComparisonCheckboxes();
    
    const cv = document.getElementById("plot");
    if (!cv) {
      setStatus("Preview canvas not found");
      return;
    }
    
    const ctx = cv.getContext("2d");
    const img = new Image();
    
    img.onload = () => {
      // キャンバスの表示サイズを固定（CSS）
      const maxDisplayWidth = 500;
      const maxDisplayHeight = 400;
      
      // 元画像のアスペクト比を保持
      const imgAspectRatio = img.width / img.height;
      let displayWidth = maxDisplayWidth;
      let displayHeight = maxDisplayWidth / imgAspectRatio;
      
      // 高さが制限を超える場合は高さ基準で調整
      if (displayHeight > maxDisplayHeight) {
        displayHeight = maxDisplayHeight;
        displayWidth = maxDisplayHeight * imgAspectRatio;
      }
      
      // CSS表示サイズを設定
      cv.style.width = displayWidth + "px";
      cv.style.height = displayHeight + "px";
      
      // キャンバスの実際の解像度は元画像と同じに
      cv.width = img.width;
      cv.height = img.height;
      
      // 画像を描画
      ctx.clearRect(0, 0, cv.width, cv.height);
      ctx.drawImage(img, 0, 0);
      
      const m = lastRender.meta;
      console.log(`Canvas size: ${cv.width}x${cv.height}, Display: ${displayWidth}x${displayHeight}`);
      console.log(`Image size: ${img.width}x${img.height}`);
      
      setStatus(`✅ Preview complete: ${m.width}${m.units} × ${m.height}${m.units}`);

      // Auto-display statistical results in UI if statistics are enabled
      // Small delay to ensure R has finished storing the results
      setTimeout(() => {
        displayStatisticalResultsInUI().catch(err => {
          console.log("Statistical auto-display skipped:", err.message || err);
        });
      }, 100);

      // Auto-export statistical results to new sheet if statistics are enabled AND auto-export is enabled
      // Longer delay to ensure R results are fully stored
      setTimeout(() => {
        const addStatistics = document.getElementById("addStatistics")?.checked;
        const autoExportStats = document.getElementById("autoExportStats")?.checked;
        if (addStatistics && autoExportStats) {
          exportStatisticalResultsToNewSheet().catch(err => {
            console.log("Statistical auto-export to sheet skipped:", err.message || err);
          });
        }
      }, 300);

      // Auto-export IC50 curve data to new sheet for IC50 chart type (if enabled)
      setTimeout(() => {
        const chartType = document.getElementById("chartType")?.value || "";
        const exportIC50ToSheet = document.getElementById("exportIC50ToSheet")?.checked;
        console.log("IC50 export check - chartType:", chartType, "exportIC50ToSheet:", exportIC50ToSheet);
        if (chartType === "ic50_dose_response" && exportIC50ToSheet) {
          console.log("Starting IC50 export...");
          exportIC50CurveDataToExcel().catch(err => {
            console.error("IC50 curve data export failed:", err.message || err);
            setStatus("❌ IC50 export failed: " + (err.message || err));
          });
        }
      }, 500);
    };

    img.onerror = (e) => {
      console.error("Image load error:", e);
      setStatus("❌ Image load error");
    };

    img.src = "data:image/png;base64," + lastRender.b64;

  } catch (e) {
    console.error("Preview generation error:", e);
    // Provide helpful error message
    let errorMsg = e?.message || String(e);
    if (errorMsg.includes("trim is not a function")) {
      errorMsg = "Data mismatch: Loaded data doesn't match the preset chart type. Please load appropriate data or change chart type.";
    }
    setStatus("❌ Preview error: " + errorMsg);
  }
}

// 修正版のExcel挿入関数
async function insertIntoExcelFixed() {
  try {
    setStatus("Generating image...");

    // 最新の設定で再レンダリング
    const renderResult = await renderPlotPngFromSelectionFixed();
    const { b64, wPt, hPt, meta } = renderResult;

    // Convert base64 to blob
    const dataUrl = "data:image/png;base64," + b64;
    const response = await fetch(dataUrl);
    const pngBlob = await response.blob();

    // Collect metadata
    const metadata = {
      version: 1,
      generator: "Figra",
      created_utc: new Date().toISOString(),

      data: {
        headers: window.lastProcessedData ? window.lastProcessedData[0] : [],
        rows: window.lastProcessedData ? window.lastProcessedData.slice(1) : []
      },

      chart: {
        chartType: document.getElementById("chartType")?.value || "",
        xColumn: document.getElementById("xColumn")?.value || "",
        yColumn: document.getElementById("yColumn")?.value || "",
        groupColumn: document.getElementById("groupColumn")?.value || "",
        errorColumn: document.getElementById("errorColumn")?.value || "",
        bins: document.getElementById("bins")?.value || "20"
      },

      settings: collectCurrentSettings()
    };

    // Retrieve statistical results if applicable
    const STATS_SUPPORTED_CHART_TYPES = [
      "box", "box_dot", "violin", "violin_dot", "bar_error_dot",
      "bar_grouped_error_dot", "box_grouped", "box_grouped_dot",
      "violin_grouped", "violin_grouped_dot",
      "line_grouped_error_raw"
    ];

    const chartType = document.getElementById("chartType")?.value || "";
    const addStatistics = document.getElementById("addStatistics")?.checked;
    const isChartTypeSupported = STATS_SUPPORTED_CHART_TYPES.includes(chartType);

    if (addStatistics && isChartTypeSupported) {
      try {
        let statResults = null;

        if (chartType === "line_grouped_error_raw") {
          statResults = await getLineStatisticalResultsText();
        } else if (chartType === "bar_grouped_error_dot" || chartType === "box_grouped" || chartType === "box_grouped_dot" || chartType === "violin_grouped" || chartType === "violin_grouped_dot") {
          statResults = await getGroupedBarStatisticalResultsText();
        } else {
          statResults = await getStatisticalResultsText();
        }

        if (statResults &&
            !statResults.includes("No statistical") &&
            !statResults.includes("ERROR") &&
            !statResults.includes("Statistics disabled") &&
            statResults.trim() !== "") {
          metadata.statisticalResults = statResults;
        } else {
          metadata.statisticalResults = null;
        }
      } catch (e) {
        console.error("Error retrieving statistical results for metadata:", e);
        metadata.statisticalResults = null;
      }
    } else {
      metadata.statisticalResults = null;
    }

    // Embed metadata into PNG
    const pngWithMetadata = await embedPngMetadata(pngBlob, metadata);

    // Convert back to base64
    const reader = new FileReader();
    const b64WithMetadata = await new Promise((resolve) => {
      reader.onloadend = () => {
        const result = reader.result.split(',')[1]; // Remove data:image/png;base64, prefix
        resolve(result);
      };
      reader.readAsDataURL(pngWithMetadata);
    });

    // Excel挿入処理
    const anchor = await getSafeAnchorLeftTop();

    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const img = ws.shapes.addImage(b64WithMetadata);

      img.lockAspectRatio = false;
      img.left = anchor.left;
      img.top = anchor.top;
      img.width = wPt;
      img.height = hPt;

      await ctx.sync();
    });

    setStatus(`✅ Inserted to Excel: ${meta.width}${meta.units} × ${meta.height}${meta.units}`);

  } catch (e) {
    console.error("Insert error:", e);
    setStatus("❌ Insert error: " + (e?.message || e));
  }
}