/**
 * Code Block Formatter — Apps Script Backend
 * Modernized "Greendom" styling engine for Google Docs.
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu("Code Formatter")
    .addItem("Open Code Formatter", "openSidebar")
    .addSeparator()
    .addItem("Remove Code Block Under Cursor", "removeCodeBlockUnderCursor")
    .addToUi();
}

function openSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile("Sidebar")
    .setTitle("Code Formatter")
    .setWidth(280);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Robustly fetches selected text or returns null if no selection.
 */
function getSelectedText() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) return null;

  const rangeElements = selection.getRangeElements();
  let textParts = [];

  for (let i = 0; i < rangeElements.length; i++) {
    const rangeEl = rangeElements[i];
    const el = rangeEl.getElement();
    
    // Check if the element can be treated as text directly
    if (el.editAsText) {
      const textObj = el.asText();
      let txt = textObj.getText();
      
      // If it's a partial selection within a paragraph, crop it accurately
      if (rangeEl.isPartial()) {
        txt = txt.substring(rangeEl.getStartOffset(), rangeEl.getEndOffsetInclusive() + 1);
      }
      textParts.push(txt);
    } else if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      textParts.push(el.asParagraph().getText());
    } else if (el.getType() === DocumentApp.ElementType.LIST_ITEM) {
      textParts.push(el.asListItem().getText());
    }
  }

  const joined = textParts.join("\n");
  return joined.trim() ? joined : null;
}

/**
 * Inserts an ultra-modern, high-contrast dark code block layout.
 * Accepts rich token objects to print exact matching colors into the Doc lines.
 */
function applyCodeBlock(tokenLines, langLabel) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const selection = doc.getSelection();
  if (!selection) throw new Error("Please select some text first.");

  const rangeEls = selection.getRangeElements();
  
  // Find top-level structural anchor to replace cleanly
  let firstEl = rangeEls[0].getElement();
  while (firstEl.getParent() && firstEl.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    firstEl = firstEl.getParent();
  }
  const insertIndex = body.getChildIndex(firstEl);
  if (insertIndex < 0) throw new Error("Could not map insertion anchor point.");

  // Identify all structural elements to clean up later
  let toDelete = [];
  for (let i = 0; i < rangeEls.length; i++) {
    let el = rangeEls[i].getElement();
    while (el.getParent() && el.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
      el = el.getParent();
    }
    if (body.getChildIndex(el) >= 0 && toDelete.indexOf(el) === -1) {
      toDelete.push(el);
    }
  }

  // Modernized Greendom Color Scheme Tokens
  const BG_HEADER = "#002F1E"; 
  const BG_CODE   = "#012116"; // Deeper rich dark black-green finish
  const FG_GUTTER = "#67E387"; // Vibrant line tags
  const FONT      = "Consolas"; 

  // 1. Build a sleek container table structure (2 Rows, 1 Column)
  const table = body.insertTable(insertIndex, [[" "], [" "]]);
  
  // Strip Google's default heavy borders aggressively
  const noBorder = {};
  noBorder[DocumentApp.Attribute.BORDER_WIDTH] = 0;
  noBorder[DocumentApp.Attribute.BORDER_COLOR] = BG_CODE;
  table.setAttributes(noBorder);

  // Style Header Cell
  const headerCell = table.getRow(0).getCell(0);
  headerCell.setBackgroundColor(BG_HEADER);
  headerCell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(12).setPaddingRight(12);
  headerCell.setAttributes(noBorder);
  
  const headerPara = headerCell.getChild(0).asParagraph();
  headerPara.setText("● ● ●   " + langLabel);
  
  const headText = headerPara.editAsText();
  headText.setFontFamily(FONT).setFontSize(9).setBold(true);
  headText.setForegroundColor(0, 0, "#00C859");  // Dot 1
  headText.setForegroundColor(2, 2, "#B4F5B5");  // Dot 2
  headText.setForegroundColor(4, 4, "#67E387");  // Dot 3
  headText.setForegroundColor(6, headerPara.getText().length - 1, "#8C9DA1"); // Window label muted

  // Style Code block Cell
  const codeCell = table.getRow(1).getCell(0);
  codeCell.setBackgroundColor(BG_CODE);
  codeCell.setPaddingTop(12).setPaddingBottom(12).setPaddingLeft(12).setPaddingRight(12);
  codeCell.setAttributes(noBorder);

  const totalLines = tokenLines.length;
  const padLen = String(totalLines).length;

  // 2. Iterate tokens array passed down from front-end to paint lines
  for (let i = 0; i < totalLines; i++) {
    const para = (i === 0) ? codeCell.getChild(0).asParagraph() : codeCell.appendParagraph("");
    para.setLineSpacing(1.2).setSpacingBefore(0).setSpacingAfter(0);
    
    // Construct line numbers gutter text
    const lineNumStr = String(i + 1).padStart(padLen, ' ') + "  ";
    para.setText(lineNumStr);
    para.editAsText().setFontFamily(FONT).setFontSize(9.5).setForegroundColor(FG_GUTTER).setBold(false);
    
    // Append code tokens side-by-side cleanly inside the structural line paragraph
    const tokens = tokenLines[i];
    for (let j = 0; j < tokens.length; j++) {
      const token = tokens[j];
      if (token.text) {
        const offsetStart = para.getText().length;
        para.appendText(token.text);
        const offsetEnd = para.getText().length - 1;
        
        // Dynamically style token spans directly inside the paragraph document slice
        const textTarget = para.editAsText();
        textTarget.setFontFamily(offsetStart, offsetEnd, FONT);
        textTarget.setFontSize(offsetStart, offsetEnd, 9.5);
        textTarget.setForegroundColor(offsetStart, offsetEnd, token.color || "#E2F5DB");
      }
    }
  }

  // 3. Delete original code text cleanly safely backwards to prevent offset collisions
  for (let k = toDelete.length - 1; k >= 0; k--) {
    try { toDelete[k].removeFromParent(); } catch(e) {}
  }
}

/**
 * Context-aware code removal utility
 */
function removeCodeBlockUnderCursor() {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  if (!cursor) {
    DocumentApp.getUi().alert("Place your cursor inside the target code block framework.");
    return;
  }
  let el = cursor.getElement();
  while (el && el.getType() !== DocumentApp.ElementType.TABLE) {
    el = el.getParent ? el.getParent() : null;
  }
  if (el) {
    el.removeFromParent();
  } else {
    DocumentApp.getUi().alert("No code block structural table detected at current cursor vector.");
  }
}
