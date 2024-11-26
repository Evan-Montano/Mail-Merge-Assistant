/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
  }
});


// Function to get all merge fields
async function getAllMergefields() {
  Word.run(async (context) => {
      // Get all fields in the document
      const fields = context.document.body.fields;
      fields.load("items/code");

      await context.sync();

      // Filter for merge fields based on their code
      const mergeFields = fields.items.filter(field =>
          field.code.toLowerCase().includes("mergefield")
      );

      console.log(`Found ${mergeFields.length} merge fields.`);

      mergeFields.forEach((field, index) => {
          console.log(`Merge Field ${index + 1}: ${field.code}`);
      });

      return context.sync();
  }).catch((error) => {
      console.error("Error in getAllMergefields:", error);
  });
}


async function insertParagraph(contentToInsert) {
  await Word.run(async (context) => {
      const docBody = context.document.body;
      docBody.insertParagraph(contentToInsert, Word.InsertLocation.start);
      await context.sync();
  });
}

// Function to read the uploaded data as text
document.addEventListener("DOMContentLoaded", function () {
  const xmlUpload = document.getElementById("xml-upload");
  const container = document.getElementById("xml-tree-container");

  if (!xmlUpload || !container) {
    console.error("Required elements are missing from the DOM.");
    return;
  }

  xmlUpload.addEventListener("change", function (event) {
    const file = event.target.files[0];

    if (!file) {
      container.innerHTML = "<p>No file selected. Please choose an XML file.</p>";
      return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const content = e.target.result; // Read the file content as text
        if (!content.trim()) {
          throw new Error("The file is empty.");
        }

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(content, "application/xml");
        const parseError = xmlDoc.getElementsByTagName("parsererror");
        if (parseError.length > 0) {
          throw new Error("Error parsing XML document.");
        }

        // Clear the container and create the collapsible tree
        container.innerHTML = "";
        const collapsibleTree = createCollapsibleTree(xmlDoc.documentElement);
        container.appendChild(collapsibleTree);
      } catch (error) {
        console.error("Error:", error.message);
        container.innerHTML = `<p>${error.message}</p>`;
      }
    };

    reader.onerror = function () {
      console.error("File could not be read.");
      container.innerHTML = "<p>Error reading file.</p>";
    };

    reader.readAsText(file);
  });
});

// ... (other code remains the same)

// Updated createCollapsibleTree function
function createCollapsibleTree(node, depth = 0, path = []) {
  const details = document.createElement('details');
  const summary = document.createElement('summary');

  // Create a span for the node text
  const nodeTextSpan = document.createElement('span');
  nodeTextSpan.textContent = node.nodeName;

  // Append the node text span to the summary
  summary.appendChild(nodeTextSpan);

  details.appendChild(summary);

  // Apply indentation
  details.style.marginLeft = `${depth * 20}px`;

  let hasElementChild = false;

  const currentPath = [...path, node.nodeName];

  // Process attributes as child nodes
  if (node.attributes && node.attributes.length > 0) {
    hasElementChild = true; // Since we have attributes, consider it as non-leaf node
    Array.from(node.attributes).forEach(attr => {
      const attrDetails = document.createElement('div');
      attrDetails.style.marginLeft = `${(depth + 1) * 20}px`;
      const attrSpan = document.createElement('span');
      attrSpan.textContent = `${attr.name}="${attr.value}"`;
      attrDetails.appendChild(attrSpan);
      attrSpan.classList.add('leaf-node');
      attrSpan.addEventListener('click', function(event) {
        event.stopPropagation();
        event.preventDefault();
        // Insert only the attribute name prefixed with '@'
        insertContentIntoWord(`@${attr.name}`);
      });
      details.appendChild(attrDetails);
    });
  }

  node.childNodes.forEach(child => {
    if (child.nodeType === Node.ELEMENT_NODE) {
      hasElementChild = true;
      details.appendChild(createCollapsibleTree(child, depth + 1, currentPath));
    } else if (child.nodeType === Node.TEXT_NODE && child.nodeValue.trim()) {
      const textNode = document.createElement('div');
      textNode.textContent = child.nodeValue.trim();
      textNode.style.marginLeft = `${(depth + 1) * 20}px`;
      textNode.classList.add('xml-text-node');
      details.appendChild(textNode);
    }
  });

  if (hasElementChild) {
    // It's a non-leaf node, make the node text clickable for inserting TableStart
    nodeTextSpan.classList.add('non-leaf-node');
    nodeTextSpan.addEventListener('click', function (event) {
      event.stopPropagation();
      event.preventDefault();
      // Exclude the root node from the fullPath
      const fullPath = currentPath.slice(1).join('/');
      // Insert TableStart into the Word document
      insertTableStartIntoWord(fullPath);
    });
  } else {
    // It's a leaf node (no attributes, no element children), make the node text clickable for inserting merge field
    nodeTextSpan.classList.add('leaf-node');
    nodeTextSpan.addEventListener('click', function (event) {
      event.stopPropagation();
      event.preventDefault();
      // Insert only the node's name
      insertContentIntoWord(node.nodeName);
    });
  }

  return details;
}

// Updated insertContentIntoWord function
function insertContentIntoWord(nodeName) {
  Word.run(async (context) => {
    console.log('Inserting merge field for node:', nodeName);
    const range = context.document.getSelection();
    const sanitizedNodeName = nodeName.replace(/[^a-zA-Z0-9_@]/g, '');
    console.log('Sanitized node name:', sanitizedNodeName);
    console.log(`Attempting Insert: ${sanitizedNodeName}`);
    range.insertField(Word.InsertLocation.replace, Word.FieldType.mergeField, sanitizedNodeName, false);
    await context.sync();
    console.log('Merge field inserted successfully.');
  }).catch(function (error) {
    console.error('Error inserting merge field:', error);
  });
}

function insertTableStartIntoWord(fullPath) {
  Word.run(async (context) => {
    console.log('Inserting TableStart and TableEnd for path:', fullPath);
    const range = context.document.getSelection();
    const sanitizedPath = fullPath.replace(/[^a-zA-Z0-9_\/]/g, '');
    const fieldCodeStart = `TableStart:${sanitizedPath}`;
    const fieldCodeEnd = `TableEnd:${sanitizedPath}`;

    // Insert TableStart field
    range.insertField(Word.InsertLocation.replace, Word.FieldType.mergeField, fieldCodeStart, false);
    await context.sync();

    const endRange = range.getRange(Word.RangeLocation.end);
    range.insertField(Word.InsertLocation.after, Word.FieldType.mergeField, fieldCodeEnd, false);
    await context.sync();

    console.log('TableStart and TableEnd fields inserted successfully.');
  }).catch(function (error) {
    console.error('Error inserting TableStart and TableEnd fields:', error);
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      console.error(error);
  }
}
