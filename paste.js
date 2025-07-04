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
          const fieldSelector = `#ARTICLE\\.metafields\\.${item.columnName}-anchor`;
          const targetField = document.querySelector(fieldSelector)?.closest('[role="button"]');

          if (targetField) {
            targetField.click();
            await new Promise(res => setTimeout(res, 300));
          }

          await navigator.clipboard.writeText(itemToPaste);
          document.execCommand('insertText', false, itemToPaste);

          console.log(`%cPasted ${currentIndex + 1}/${pasteQueue.length}: ${item.columnName}`, styles.pasted);
          currentIndex++;

          if (currentIndex < pasteQueue.length) {
            console.log(`%cNext up: ${pasteQueue[currentIndex].columnName}`, styles.next);
          } else {
            console.log(`%c✅ Queue finished!`, styles.finished);
            document.removeEventListener('keydown', handlePasteEvent, true);
          }
        } catch (err) {
          console.error("Clipboard error: ", err);
          alert("Clipboard write failed. Ensure this tab has focus and clipboard permission.");
        }
      }
    }
  }

  function createOverlayInput() {
    const textarea = document.createElement('textarea');
    textarea.placeholder = "Paste your Excel raw data here...";
    Object.assign(textarea.style, {
      position: "fixed",
      top: "10%",
      left: "10%",
      width: "80%",
      height: "300px",
      zIndex: 9999,
      fontSize: "16px",
      padding: "10px",
      border: "2px solid #007bff",
      borderRadius: "8px",
      boxShadow: "0 4px 10px rgba(0,0,0,0.2)",
      background: "#fefefe",
      color: "#333",
      lineHeight: "1.5",
      fontFamily: "monospace"
    });

    const button = document.createElement('button');
    button.textContent = "✅ Process";
    Object.assign(button.style, {
      marginTop: "10px",
      padding: "10px 16px",
      background: "#007bff",
      color: "#fff",
      border: "none",
      cursor: "pointer",
      borderRadius: "4px",
      fontSize: "14px"
    });

    const container = document.createElement('div');
    Object.assign(container.style, {
      position: "fixed",
      top: "5%",
      left: "5%",
      width: "90%",
      background: "#fff",
      padding: "20px",
      borderRadius: "12px",
      boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
      zIndex: 9999,
      display: "flex",
      flexDirection: "column"
    });

    container.appendChild(textarea);
    container.appendChild(button);
    document.body.appendChild(container);

    button.onclick = () => {
      const value = textarea.value.trim();
      document.body.removeChild(container);
      setTimeout(() => processPastedData(value), 0);
    };
  }

  createOverlayInput();
})();
