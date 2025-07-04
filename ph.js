(function () {
  const COLUMN_NAMES = [
    "hero_image_with_info",
    "info_with_image_section",
    "info_with_images",
    "grid_info_cards",
    "info_with_image_instance_2",
    "info_and_image",
    "info_with_image",
    "megamenu_desc",
    "meta_desc",
    "excerpt"
  ];

  const UNIQUE_SEPARATOR = "|||---PASTE-SEPARATOR---|||";
  let pasteQueue = [];
  let currentIndex = 0;

  const styles = {
    success: 'color: #28a745; font-weight: bold; font-size: 18px;',
    pasted: 'color: #007bff; font-weight: bold; font-size: 16px;',
    next: 'color: #6c757d; font-size: 14px;',
    finished: 'color: #17a2b8; font-weight: bold; font-size: 18px;',
  };

  function cleanExcelText(text) {
    if (typeof text !== 'string') return '';
    let cleaned = text.trim();
    if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
      cleaned = cleaned.slice(1, -1);
    }
    return cleaned.replace(/""/g, '"');
  }

  function processPastedData(rawData) {
    if (!rawData || typeof rawData !== 'string') {
      console.log("No valid data pasted. Aborting.");
      return;
    }

    const processedData = rawData
      .replace(/\}\s*\{/g, `}${UNIQUE_SEPARATOR}{`)
      .replace(/\t/g, UNIQUE_SEPARATOR);

    const columnsData = processedData.split(UNIQUE_SEPARATOR);

    pasteQueue = columnsData.map((data, index) => ({
      columnName: COLUMN_NAMES[index] || `Column ${index + 1}`,
      data: cleanExcelText(data)
    })).filter(item => item.data);

    if (pasteQueue.length > 0) {
      currentIndex = 0;
      console.clear();
      console.log(`%c✅ Success! Processed ${pasteQueue.length} columns.`, styles.success);
      console.log(`%cNext up: ${pasteQueue[currentIndex].columnName}`, styles.next);

      document.removeEventListener('keydown', handlePasteEvent, true);
      document.addEventListener('keydown', handlePasteEvent, true);
    } else {
      console.error("Could not find any columns to process in the pasted data.");
    }
  }

  async function handlePasteEvent(event) {
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'v') {
      if (pasteQueue.length > 0 && currentIndex < pasteQueue.length) {
        event.preventDefault();
        event.stopPropagation();

        const item = pasteQueue[currentIndex];
        const itemToPaste = item.data;

        try {
          await navigator.clipboard.writeText(itemToPaste);

          setTimeout(() => {
            document.execCommand('insertText', false, itemToPaste);
            console.log(`%cPasted column ${currentIndex + 1}/${pasteQueue.length}: ${item.columnName}`, styles.pasted);
            currentIndex++;

            if (currentIndex < pasteQueue.length) {
              console.log(`%cNext up: ${pasteQueue[currentIndex].columnName}`, styles.next);
            } else {
              console.log(`%c✅ Queue finished! Normal paste is restored.`, styles.finished);
              document.removeEventListener('keydown', handlePasteEvent, true);
            }
          }, 300); // Delay each paste to ease processing
        } catch (err) {
          console.error("Clipboard error: ", err);
          alert("Clipboard paste failed. Check browser focus and permissions.");
        }
      }
    }
  }

  function createOverlayInput() {
    const textarea = document.createElement('textarea');
    textarea.placeholder = "Paste your Excel raw data here...";
    textarea.style.position = "fixed";
    textarea.style.top = "10%";
    textarea.style.left = "10%";
    textarea.style.width = "80%";
    textarea.style.height = "300px";
    textarea.style.zIndex = 9999;
    textarea.style.fontSize = "16px";
    textarea.style.padding = "10px";
    textarea.style.border = "2px solid #007bff";
    textarea.style.borderRadius = "8px";
    textarea.style.boxShadow = "0 4px 10px rgba(0,0,0,0.2)";
    textarea.style.background = "#fefefe";
    textarea.style.color = "#333";
    textarea.style.lineHeight = "1.5";
    textarea.style.fontFamily = "monospace";

    const button = document.createElement('button');
    button.textContent = "✅ Process";
    button.style.marginTop = "10px";
    button.style.padding = "10px 16px";
    button.style.background = "#007bff";
    button.style.color = "#fff";
    button.style.border = "none";
    button.style.cursor = "pointer";
    button.style.borderRadius = "4px";
    button.style.fontSize = "14px";

    const container = document.createElement('div');
    container.style.position = "fixed";
    container.style.top = "5%";
    container.style.left = "5%";
    container.style.width = "90%";
    container.style.background = "#fff";
    container.style.padding = "20px";
    container.style.borderRadius = "12px";
    container.style.boxShadow = "0 8px 24px rgba(0,0,0,0.2)";
    container.style.zIndex = 9999;
    container.style.display = "flex";
    container.style.flexDirection = "column";
    container.appendChild(textarea);
    container.appendChild(button);
    document.body.appendChild(container);

    button.onclick = () => {
      const value = textarea.value.trim();
      document.body.removeChild(container);
      setTimeout(() => {
        processPastedData(value);
      }, 0); // defer processing to next event loop
    };
  }

  // Start
  createOverlayInput();
})();
