// ============================================================
// Core logic to read a Google Form and export its
// structure (items, types, options) to JSON.
// Page breaks and section navigation are excluded.
// ============================================================

// Exports a Google Form to JSON, saves to Drive and returns content for client-side download
function exportFormToJson(formId, fileName, destinationFolder) {
    try {

        var form = FormApp.openById(formId);

        var formJson = {
            formId: form.getId(),
            title: form.getTitle(),
            description: form.getDescription(),
            publishedUrl: form.getPublishedUrl(),
            items: []
        };

        var items = form.getItems();

        if (items.length === 0) {
            Logger.log('Warning: The form has no items. Nothing to export.');
            return { success: false, error: 'The form has no items.' };
        }

        items.forEach(function (item) {
            // Skip page break items entirely
            if (item.getType() === FormApp.ItemType.PAGE_BREAK) return;
            var itemJson = buildItemJson(item);
            formJson.items.push(itemJson);
        });

        var formJsonString = JSON.stringify(formJson, null, 2);

        // File saved as .json
        //var fileName = fileName + '.json';
        var formFile = destinationFolder.createFile(fileName + '.json', formJsonString, MimeType.PLAIN_TEXT);

        Logger.log('Successfully created JSON for form "' + formJson.title + '" as: ' + formFile.getName());

        // Returns formJsonContent for client-side download, consistent with codeExporter pattern
        return {
            success: true,
            formUrl: formFile.getUrl(),
            formName: formFile.getName(),
            formJsonContent: formJsonString
        };

    } catch (e) {
        Logger.log('CRITICAL ERROR exporting Google Form with ID ' + formId + '. Error: ' + e.toString());
        Logger.log('Stack Trace: ' + e.stack);
        return { success: false, error: e.message };
    }
}

/**
 * Builds a JSON object for a single form item.
 * Handles common properties + type-specific properties.
 *
 * @param {GoogleAppsScript.Forms.Item} item - A raw form item
 * @returns {Object} Structured JSON for this item
 */
function buildItemJson(item) {
    // Common properties shared by all item types
    const itemJson = {
        index: item.getIndex(),
        itemId: item.getId(),
        type: item.getType().toString(),
        title: item.getTitle(),
        helpText: item.getHelpText()
    };

    // Append type-specific properties
    const type = item.getType();
    const ItemType = FormApp.ItemType;

    switch (type) {

        case ItemType.MULTIPLE_CHOICE:
            appendMultipleChoiceData(item.asMultipleChoiceItem(), itemJson);
            break;

        case ItemType.CHECKBOX:
            appendCheckboxData(item.asCheckboxItem(), itemJson);
            break;

        case ItemType.DROPDOWN:
            appendDropdownData(item.asDropdownItem(), itemJson);
            break;

        case ItemType.LINEAR_SCALE:
            appendLinearScaleData(item.asScaleItem(), itemJson);
            break;

        case ItemType.GRID:
            appendGridData(item.asGridItem(), itemJson);
            break;

        case ItemType.CHECKBOX_GRID:
            appendCheckboxGridData(item.asCheckboxGridItem(), itemJson);
            break;

        case ItemType.DATE:
            appendDateData(item.asDateItem(), itemJson);
            break;

        case ItemType.DATETIME:
            appendDateTimeData(item.asDateTimeItem(), itemJson);
            break;

        case ItemType.TIME:
            itemJson.isRequired = item.asTimeItem().isRequired();
            break;

        case ItemType.TEXT:
            itemJson.isRequired = item.asTextItem().isRequired();
            break;

        case ItemType.PARAGRAPH_TEXT:
            itemJson.isRequired = item.asParagraphTextItem().isRequired();
            break;

        case ItemType.SECTION_HEADER:
            // Section headers have no extra fields; title/helpText already captured
            break;

        case ItemType.IMAGE:
            appendImageData(item.asImageItem(), itemJson);
            break;

        case ItemType.VIDEO:
            appendVideoData(item.asVideoItem(), itemJson);
            break;

        default:
            itemJson.note = "Type not explicitly handled: " + type.toString();
            break;
    }

    return itemJson;
}


// ============================================================
// Type-Specific Data Builders
// Each function mutates the itemJson object in place
// ============================================================

/**
 * Appends Multiple Choice item data.
 * Navigation fields are excluded.
 */
function appendMultipleChoiceData(mcItem, itemJson) {
    itemJson.isRequired = mcItem.isRequired();
    itemJson.hasOtherOption = mcItem.hasOtherOption();
    itemJson.choices = mcItem.getChoices().map(function (choice) {
        return {
            value: choice.getValue(),
            isCorrect: safeCall(function () { return choice.isCorrectAnswer(); }, null)
        };
    });
}

/**
 * Appends Checkbox item data.
 */
function appendCheckboxData(cbItem, itemJson) {
    itemJson.isRequired = cbItem.isRequired();
    itemJson.hasOtherOption = cbItem.hasOtherOption();
    itemJson.choices = cbItem.getChoices().map(function (choice) {
        return {
            value: choice.getValue(),
            isCorrect: safeCall(function () { return choice.isCorrectAnswer(); }, null)
        };
    });
}

/**
 * Appends Dropdown item data.
 */
function appendDropdownData(ddItem, itemJson) {
    itemJson.isRequired = ddItem.isRequired();
    itemJson.choices = ddItem.getChoices().map(function (choice) {
        return {
            value: choice.getValue(),
            isCorrect: safeCall(function () { return choice.isCorrectAnswer(); }, null)
        };
    });
}

/**
 * Appends Linear Scale item data.
 */
function appendLinearScaleData(scaleItem, itemJson) {
    itemJson.isRequired = scaleItem.isRequired();
    itemJson.lowerBound = scaleItem.getLowerBound();
    itemJson.upperBound = scaleItem.getUpperBound();
    itemJson.leftLabel = scaleItem.getLeftLabel();
    itemJson.rightLabel = scaleItem.getRightLabel();
}

/**
 * Appends Grid item data (rows and columns).
 */
function appendGridData(gridItem, itemJson) {
    itemJson.isRequired = gridItem.isRequired();
    itemJson.rows = gridItem.getRows();
    itemJson.columns = gridItem.getColumns();
}

/**
 * Appends Checkbox Grid item data.
 */
function appendCheckboxGridData(cbGridItem, itemJson) {
    itemJson.isRequired = cbGridItem.isRequired();
    itemJson.rows = cbGridItem.getRows();
    itemJson.columns = cbGridItem.getColumns();
}

/**
 * Appends Date item data.
 */
function appendDateData(dateItem, itemJson) {
    itemJson.isRequired = dateItem.isRequired();
    itemJson.includesYear = dateItem.includesYear();
}

/**
 * Appends DateTime item data.
 */
function appendDateTimeData(dtItem, itemJson) {
    itemJson.isRequired = dtItem.isRequired();
    itemJson.includesYear = dtItem.includesYear();
}

/**
 * Appends Image item data.
 */
function appendImageData(imgItem, itemJson) {
    itemJson.alignment = imgItem.getAlignment().toString();
    itemJson.width = imgItem.getWidth();
    itemJson.imageTitle = imgItem.getTitle();
}

/**
 * Appends Video item data.
 */
function appendVideoData(vidItem, itemJson) {
    itemJson.alignment = vidItem.getAlignment().toString();
    itemJson.width = vidItem.getWidth();
    itemJson.videoUrl = vidItem.getVideoUrl();
}

// ============================================================
// Utility
// ============================================================

/**
 * Safely executes a function and returns a fallback value on error.
 * Useful for quiz-only properties like isCorrectAnswer().
 *
 * @param {Function} fn       - Function to attempt
 * @param {*}        fallback - Value to return if fn throws
 * @returns {*}
 */
function safeCall(fn, fallback) {
    try {
        return fn();
    } catch (e) {
        return fallback;
    }
}