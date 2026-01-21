/**
 * DYMO Label Web Editor - Main Application
 * Supports .label (Legacy) and .dymo (Connect) files
 */

// Application state
const state = {
    label: null,           // Current loaded label object
    labelXml: null,        // Raw XML content
    objectNames: [],       // List of editable field names
    objectTypes: {},       // Map of name -> type (e.g., "AddressObject", "ImageObject")
    initialValues: {},     // Map of name -> original value (for restoring images)
    isServiceRunning: false,
    selectedPrinter: null  // Currently selected printer
};

// DOM Elements
const elements = {
    statusContainer: document.getElementById('status-container'),
    statusText: document.getElementById('status-text'),
    labelFile: document.getElementById('label-file'),
    fileError: document.getElementById('file-error'),
    formCard: document.getElementById('form-card'),
    labelForm: document.getElementById('label-form'),
    noFieldsMessage: document.getElementById('no-fields-message'),
    previewCard: document.getElementById('preview-card'),
    previewImage: document.getElementById('preview-image'),
    printCard: document.getElementById('print-card'),
    printerSelect: document.getElementById('printer-select'),
    printQuantity: document.getElementById('print-quantity'),
    previewBtn: document.getElementById('preview-btn'),
    printBtn: document.getElementById('print-btn'),
    clearBtn: document.getElementById('clear-btn')
};

/**
 * Initialize the application
 */
async function init() {
    updateStatus('checking', 'Initializing DYMO Framework...');

    // Set up event listeners
    elements.labelFile.addEventListener('change', handleFileUpload);
    elements.previewBtn.addEventListener('click', updatePreview);
    elements.printBtn.addEventListener('click', printLabels);
    elements.clearBtn.addEventListener('click', clearForm);
    elements.printerSelect.addEventListener('change', (e) => {
        state.selectedPrinter = e.target.value;
    });

    try {
        // Initialize the framework
        // The DYMO Connect Framework usually requires this init call
        if (dymo && dymo.label && dymo.label.framework) {
            await new Promise((resolve, reject) => {
                dymo.label.framework.init(resolve);
                // Note: dymo.label.framework.init takes a callback, not always a promise in older versions, 
                // but dymo-connect-framework usually supports standardized startup.
                // We will fallback to a timeout if it stalls, but the library should handle it.
            });

            // Check environment
            const env = dymo.label.framework.checkEnvironment();
            if (env.isFrameworkInstalled && env.isWebServicePresent) {
                state.isServiceRunning = true;
                updateStatus('connected', 'DYMO Web Service connected');
                await loadPrinters();
            } else {
                throw new Error('DYMO Web Service not found or not installed.');
            }
        } else {
            throw new Error('DYMO Framework script not loaded improperly.');
        }
    } catch (error) {
        console.error('Initialization failed:', error);
        state.isServiceRunning = false;
        updateStatus('disconnected', `Service Error: ${error.message}. Is DYMO Connect installed?`);
    }
}

/**
 * Update status indicator
 */
function updateStatus(type, message) {
    elements.statusContainer.className = `alert mb-4 status-${type}`;
    let icon = 'bi-hourglass-split';
    if (type === 'connected') icon = 'bi-check-circle-fill';
    if (type === 'disconnected') icon = 'bi-exclamation-triangle-fill';
    elements.statusText.innerHTML = `<i class="bi ${icon} me-2"></i>${message}`;
}

/**
 * Load available DYMO printers
 */
async function loadPrinters() {
    try {
        console.log('Loading printers...');
        // getPrintersAsync is the preferred method in the Connect Framework
        const printers = await dymo.label.framework.getPrintersAsync();

        // Filter for LabelWriter printers
        const labelPrinters = printers.filter(p => p.printerType === 'LabelWriterPrinter');

        elements.printerSelect.innerHTML = '<option value="">Select printer...</option>';

        if (labelPrinters.length === 0) {
            const option = document.createElement('option');
            option.text = "No DYMO printers found";
            elements.printerSelect.add(option);
            return;
        }

        labelPrinters.forEach(printer => {
            const option = document.createElement('option');
            option.value = printer.name;
            option.textContent = printer.name;
            // Auto-select 450 Turbo or similar if found
            if (printer.name.includes('450') || printer.name.includes('LabelWriter')) {
                if (!state.selectedPrinter) {
                    option.selected = true;
                    state.selectedPrinter = printer.name;
                }
            }
            elements.printerSelect.appendChild(option);
        });

    } catch (error) {
        console.error('Error loading printers:', error);
        elements.printerSelect.innerHTML = '<option value="">Error loading printers</option>';
    }
}

/**
 * Handle file upload
 */
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();
    const isLabel = fileName.endsWith('.label');
    const isDymo = fileName.endsWith('.dymo');

    if (!isLabel && !isDymo) {
        showFileError('Invalid file type. Please upload a .label or .dymo file.');
        return;
    }

    hideFileError();

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            loadLabel(e.target.result);
        } catch (error) {
            showFileError(`Error reading file: ${error.message}`);
        }
    };
    reader.onerror = () => showFileError('Error reading file. Please try again.');
    reader.readAsText(file);
}

/**
 * Load label from XML content
 */
function loadLabel(xml) {
    try {
        // openLabelXml supports both .label and .dymo XML contents in the unified framework
        state.label = dymo.label.framework.openLabelXml(xml);
        state.labelXml = xml;

        // object names retrieval works for both formats through the framework abstraction
        state.objectNames = state.label.getObjectNames();

        // Parse XML to identify object types
        state.objectTypes = parseObjectTypes(xml);

        // Store initial values
        state.initialValues = {};
        state.objectNames.forEach(name => {
            try {
                state.initialValues[name] = state.label.getObjectText(name) || '';
            } catch (e) { }
        });

        if (state.objectNames.length === 0) {
            // Some .dymo files might not have "objects" in the same way if they use different structure,
            // but for standard variable filling, this is the API. 
            // If empty, it might still be printable, just not editable via this simple form.
            elements.noFieldsMessage.textContent = "No editable fields found, or this file type structure allows no variable editing.";
            elements.noFieldsMessage.style.display = 'block';

            // Still show print/preview even if no fields
            elements.formCard.style.display = 'none';
            elements.previewCard.style.display = 'block';
            elements.printCard.style.display = 'block';
            updatePreview();
        } else {
            generateForm(state.objectNames);
            elements.formCard.style.display = 'block';
            elements.noFieldsMessage.style.display = 'none';
            elements.printCard.style.display = 'block';
            updatePreview();
        }
    } catch (error) {
        showFileError(`Error parsing label file: ${error.message}`);
        console.error('Label parsing error:', error);
    }
}

/**
 * Helper to parse XML and map object names to their types
 */
function parseObjectTypes(xmlString) {
    const types = {};
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "text/xml");

        // Look for all elements that have a "Name" child
        // This covers DLS identity structure usually
        const objects = xmlDoc.getElementsByTagName("*"); // simplistic iteration
        for (let i = 0; i < objects.length; i++) {
            const node = objects[i];
            // Check if this node is likely an Object definition (ends in Object usually, e.g. TextObject)
            if (node.tagName.endsWith("Object")) {
                const nameNode = node.getElementsByTagName("Name")[0];
                if (nameNode) {
                    types[nameNode.textContent] = node.tagName;
                }
            }
        }
    } catch (e) {
        console.warn("Failed to parse label XML types", e);
    }
    return types;
}

/**
 * Generate form fields from object names
 */
function generateForm(names) {
    elements.labelForm.innerHTML = '';

    // Filter out fields starting with "IGNORE" (case-insensitive) and sort alphabetically
    const editableNames = names
        .filter(name => !name.toUpperCase().startsWith('IGNORE'))
        .sort((a, b) => a.localeCompare(b));

    editableNames.forEach(name => {
        const type = state.objectTypes[name] || 'Unknown';
        const isImage = type === 'ImageObject' || type === 'GraphicObject';

        let currentValue = '';
        // For images, we use the initial value state to track "checked" status logic if needed, 
        // but for text boxes we read current.
        // Actually, for generation, we can just use the state.label's current state.
        try {
            currentValue = state.label.getObjectText(name) || '';
        } catch (e) { /* Object might not support getText */ }

        const div = document.createElement('div');
        div.className = 'mb-3';

        if (isImage) {
            // Render as Checkbox
            // If currentValue is empty, it's unchecked. If it has data, it's checked.
            const isChecked = currentValue.length > 0;

            div.className = 'form-check mb-3';
            div.innerHTML = `
                <input class="form-check-input" type="checkbox" id="field-${name}" 
                       data-field-name="${name}" data-is-image="true"
                       ${isChecked ? 'checked' : ''}>
                <label class="form-check-label" for="field-${name}">
                    Include ${formatFieldName(name)}
                </label>
            `;
        } else {
            // Render as Textarea
            div.innerHTML = `
                <label for="field-${name}" class="form-label">${formatFieldName(name)}</label>
                <textarea class="form-control" id="field-${name}" rows="3"
                       data-field-name="${name}" 
                       placeholder="Enter ${formatFieldName(name).toLowerCase()}">${escapeHtml(currentValue)}</textarea>
            `;
        }
        elements.labelForm.appendChild(div);
    });

    // Real-time preview
    elements.labelForm.querySelectorAll('input, textarea').forEach(input => {
        input.addEventListener('change', debounce(updatePreview, 300)); // Checkboxes trigger change
        input.addEventListener('input', debounce(updatePreview, 300));  // Text inputs trigger input
    });
}

/**
 * Format field name for display
 */
function formatFieldName(name) {
    return name
        .replace(/([A-Z])/g, ' $1')
        .replace(/^./, str => str.toUpperCase())
        .trim();
}

/**
 * Update label preview
 */
async function updatePreview() {
    if (!state.label) return;

    try {
        // Update label with form values
        const inputs = elements.labelForm.querySelectorAll('[data-field-name]');
        inputs.forEach(input => {
            try {
                const name = input.dataset.fieldName;
                if (input.type === 'checkbox') {
                    // If checked, restore initial value (image data). If unchecked, set to empty string (hides image)
                    if (input.checked) {
                        // Restore original value if we have it
                        const val = state.initialValues[name];
                        // Only set if we really have something to restore, otherwise keep as is or empty
                        if (val) state.label.setObjectText(name, val);
                    } else {
                        state.label.setObjectText(name, '');
                    }
                } else {
                    state.label.setObjectText(name, input.value);
                }
            } catch (e) {
                console.warn(`Could not set text for: ${input.dataset.fieldName}`);
            }
        });

        // Render preview using Async method
        const pngBase64 = await state.label.renderAsync();

        elements.previewImage.src = `data:image/png;base64,${pngBase64}`;
        elements.previewCard.style.display = 'block';
    } catch (error) {
        console.error('Preview error:', error);
        elements.previewImage.src = '';
        elements.previewImage.alt = 'Preview unavailable: ' + error.message;
        elements.previewCard.style.display = 'block';
    }
}

/**
 * Print labels
 */
async function printLabels() {
    if (!state.label) {
        alert('No label loaded');
        return;
    }

    const printerName = elements.printerSelect.value;
    if (!printerName) {
        alert('Please select a printer');
        return;
    }

    const quantity = parseInt(elements.printQuantity.value, 10);
    if (isNaN(quantity) || quantity < 1) {
        alert('Please enter a valid quantity');
        return;
    }

    updateStatus('checking', 'Printing...');

    try {
        // Update label with form values before printing
        const inputs = elements.labelForm.querySelectorAll('[data-field-name]');
        inputs.forEach(input => {
            try {
                const name = input.dataset.fieldName;
                if (input.type === 'checkbox') {
                    if (input.checked) {
                        const val = state.initialValues[name];
                        if (val) state.label.setObjectText(name, val);
                    } else {
                        state.label.setObjectText(name, '');
                    }
                } else {
                    state.label.setObjectText(name, input.value);
                }
            } catch (e) { /* ignore */ }
        });

        // Create a print parameters object if needed, but simple print works usually
        // Construct print params XML if we wanted to be fancy, but standard print is fine.

        // NOTE: The framework does not have a "printAsync" method in all versions, 
        // sometimes it's just 'print' which fires and forgets. 
        // However, dymo.label.framework.printLabel is the method. 
        // We will loop for quantity or use print parameters. 

        // Printing multiple copies is efficient if we build a print param xml, 
        // but looping print calls is also acceptable for small quantities.

        // Let's use printLabel which takes (printerName, printParamsXml, labelXml)
        // OR standard label.print(printerName, printParamsXml, labelSetXml)

        // Create print params for multiple copies
        let printParamsXml = '';
        if (dymo.label.framework.createLabelWriterPrintParamsXml) {
            printParamsXml = dymo.label.framework.createLabelWriterPrintParamsXml({ copies: quantity });
        }

        // Print label with params
        // print(printerName, printParamsXml, labelSetXml)
        // We don't have a labelSet (variable data for multiple different labels), so we leave it empty.
        state.label.print(printerName, printParamsXml, '');

        updateStatus('connected', `Successfully sent print job (${quantity} copies) to ${printerName}`);
        setTimeout(() => {
            updateStatus('connected', 'DYMO Web Service connected');
        }, 3000);

    } catch (error) {
        alert(`Print error: ${error.message}`);
        console.error('Print error:', error);
        updateStatus('connected', 'DYMO Web Service connected'); // Reset status
    }
}

/**
 * Clear form and reset state
 */
function clearForm() {
    elements.labelFile.value = '';
    state.label = null;
    state.labelXml = null;
    state.objectNames = [];

    elements.formCard.style.display = 'none';
    elements.noFieldsMessage.style.display = 'none';
    elements.previewCard.style.display = 'none';
    elements.printCard.style.display = 'none';

    hideFileError();
    elements.previewImage.src = '';
}

function showFileError(message) {
    elements.fileError.textContent = message;
    elements.fileError.style.display = 'block';
}

function hideFileError() {
    elements.fileError.style.display = 'none';
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function debounce(func, wait) {
    let timeout;
    return function (...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func(...args), wait);
    };
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', init);
