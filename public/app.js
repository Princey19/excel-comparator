// Get references to DOM elements
const file1Input = document.getElementById("file1Input");
const file2Input = document.getElementById("file2Input");
const uploadForm = document.getElementById("uploadForm");
const crossCheckBtn = document.getElementById("crossCheckBtn");
const messageBox = document.getElementById("messageBox");
const resultsSection = document.getElementById("resultsSection");
const resultsSummary = document.getElementById("resultsSummary");
const downloadMatchedSection = document.getElementById(
  "downloadMatchedSection"
);
const downloadMatchedBtn = document.getElementById("downloadMatchedBtn");
const downloadMissingSection = document.getElementById(
  "downloadMissingSection"
);
const downloadMissingBtn = document.getElementById("downloadMissingBtn");
const missingContentsDisplay = document.getElementById(
  "missingContentsDisplay"
);
const buttonText = document.getElementById("buttonText");
const loadingSpinner = document.getElementById("loadingSpinner");
// Removed: const headerSelect = document.getElementById("headerSelect");

/**
 * Displays a message in the message box.
 * @param {string} message The message to display.
 * @param {boolean} isError True if it's an error message, false for success/info.
 */
function showMessage(message, isError = true) {
  messageBox.textContent = message;
  messageBox.classList.remove(
    "hidden",
    "bg-green-50",
    "text-green-700",
    "border-green-400",
    "bg-red-50",
    "text-red-700",
    "border-red-400"
  );
  if (isError) {
    messageBox.classList.add("bg-red-50", "text-red-700", "border-red-400");
  } else {
    messageBox.classList.add(
      "bg-green-50",
      "text-green-700",
      "border-green-400"
    );
  }
  messageBox.classList.add("show"); // Use 'show' class for display
}

/**
 * Hides the message box.
 */
function hideMessage() {
  messageBox.classList.remove("show");
  messageBox.textContent = "";
}

/**
 * Shows the loading indicator.
 */
function showLoading(text = "Processing...") {
  buttonText.textContent = text;
  loadingSpinner.classList.remove("hidden");
  loadingSpinner.classList.add("show");
  crossCheckBtn.disabled = true;
  file1Input.disabled = true;
  file2Input.disabled = true;
}

/**
 * Hides the loading indicator.
 */
function hideLoading() {
  buttonText.textContent = "Cross-Check Files";
  loadingSpinner.classList.remove("show");
  loadingSpinner.classList.add("hidden");
  crossCheckBtn.disabled = false;
  file1Input.disabled = false;
  file2Input.disabled = false;
}

// Removed: updateHeaderDropdown function and its event listeners
// file1Input.addEventListener("change", updateHeaderDropdown);
// file2Input.addEventListener("change", updateHeaderDropdown);

// Event listener for the form submission
uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault(); // Prevent default form submission

  hideMessage();
  resultsSection.classList.add("hidden");
  downloadMatchedSection.classList.add("hidden");
  downloadMissingSection.classList.add("hidden");
  missingContentsDisplay.innerHTML = "";

  const file1 = file1Input.files[0];
  const file2 = file2Input.files[0];

  if (!file1 || !file2) {
    showMessage("Please select both File A and File B.");
    return;
  }

  // Client-side check for Excel files
  const isFile1Excel =
    file1.name.endsWith(".xlsx") || file1.name.endsWith(".xls");
  const isFile2Excel =
    file2.name.endsWith(".xlsx") || file2.name.endsWith(".xls");

  if (!(isFile1Excel && isFile2Excel)) {
    showMessage("Both files must be Excel files for comparison.");
    return;
  }

  showLoading("Cross-Checking...");

  const formData = new FormData();
  formData.append("fileA", file1);
  formData.append("fileB", file2);
  // Removed: formData.append("selectedColumn", selectedColumn); as it's no longer selected by user

  try {
    const response = await fetch("/cross-check", {
      method: "POST",
      body: formData,
    });
    const result = await response.json();

    if (result.success) {
      // Display results summary
      resultsSummary.innerHTML = `
              <p class="text-lg font-semibold">File A ('${result.file1Name}') contains ${result.totalFile1Rows} rows.</p>
              <p class="text-lg font-semibold">Out of these, <span class="text-green-600">${result.foundCount}</span> rows were found in File B ('${result.file2Name}') based on the <b>'${result.comparisonColumn}'</b> column.</p>
              <p class="text-lg font-semibold">The remaining <span class="text-red-600">${result.missingCount}</span> rows were NOT found in File B based on the <b>'${result.comparisonColumn}'</b> column.</p>
            `;

      if (result.matchedDownloadUrl) {
        downloadMatchedSection.classList.remove("hidden");
        downloadMatchedBtn.href = result.matchedDownloadUrl;
      } else {
        downloadMatchedSection.classList.add("hidden");
      }

      if (result.missingDownloadUrl) {
        downloadMissingSection.classList.remove("hidden");
        downloadMissingBtn.href = result.missingDownloadUrl;
        // Display missing contents on screen, showing a sample of all columns
        if (result.missingContents && result.missingContents.length > 0) {
          missingContentsDisplay.innerHTML = `
                  <pre class="p-4 bg-gray-50 rounded-md border border-gray-200">${result.missingContents
                    .map((item) => JSON.stringify(item, null, 2)) // Pretty print JSON
                    .join("\n\n")}</pre>
                `;
        } else {
          missingContentsDisplay.innerHTML = `
                  <p class="text-green-700 font-semibold">All values from File 1 were found in File 2.</p>
                `;
        }
      } else {
        downloadMissingSection.classList.add("hidden");
        missingContentsDisplay.innerHTML = `
                <p class="text-green-700 font-semibold">All rows from File 1 were found in File 2.</p>
              `;
      }
      resultsSection.classList.remove("hidden");
      showMessage("Cross-check completed successfully!", false);
    } else {
      showMessage(result.message);
    }
  } catch (error) {
    console.log("Cross-check error:", error);
    showMessage(`An error occurred: ${error.message}`);
  } finally {
    hideLoading();
  }
});
