/**
 * Client-side PDF to Excel converter
 * Uses pdf.js for PDF parsing and ExcelJS for Excel generation
 * All processing happens in the browser - no server required
 */

/**
 * Extract text from PDF using pdf.js
 * Reconstructs text in reading order, preserving line structure
 */
async function extractPdfText(arrayBuffer) {
  // Access pdf.js library from window global
  if (typeof window.pdfjsLib === 'undefined') {
    throw new Error('PDF.js library not loaded. Please check your internet connection and refresh the page.');
  }
  
  const pdfjs = window.pdfjsLib;
  
  // Set worker source on first use
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
  
  const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
  let fullText = '';
  
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    
    // Get all text items with their positions
    const items = textContent.items;
    if (items.length === 0) continue;
    
    // Sort by Y position (top to bottom), then X (left to right)
    // PDF Y coordinates increase upward, so we sort descending
    const sortedItems = items.slice().sort((a, b) => {
      const yDiff = b.transform[5] - a.transform[5];
      if (Math.abs(yDiff) > 5) return yDiff; // Different lines (5 unit threshold)
      return a.transform[4] - b.transform[4]; // Same line, sort by X
    });
    
    // Group into lines based on Y position
    const lines = [];
    let currentLine = [];
    let currentY = sortedItems[0]?.transform[5];
    
    for (const item of sortedItems) {
      const y = item.transform[5];
      if (Math.abs(y - currentY) > 5) {
        // New line
        if (currentLine.length > 0) {
          // Sort items in current line by X position
          currentLine.sort((a, b) => a.transform[4] - b.transform[4]);
          lines.push(currentLine.map(it => it.str).join(' '));
        }
        currentLine = [item];
        currentY = y;
      } else {
        currentLine.push(item);
      }
    }
    // Don't forget last line
    if (currentLine.length > 0) {
      currentLine.sort((a, b) => a.transform[4] - b.transform[4]);
      lines.push(currentLine.map(it => it.str).join(' '));
    }
    
    fullText += lines.join('\n') + '\n';
  }
  
  return fullText;
}

/**
 * Extract PO metadata from PDF text
 */
async function getPoMetadata(text) {
  const metadata = {
    'VENDOR': '',
    'PO #': '',
    'DEPT #': '',
    'STORE #': '',
    'STATE': '',
    'ORDER DATE': '',
    'SHIP DATE': '',
    'CANCEL DATE': ''
  };
  
  const lines = text.split('\n');
  const allText = text;
  
  // Extract DEPT NUMBER and ORDER NUMBER
  // Format: "DEPT. NUMBER:ORDER NUMBER:7761665367" (concatenated)
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Try concatenated format
    const concatMatch = line.match(/DEPT\.\s*NUMBER:\s*ORDER\s*NUMBER:\s*(\d{3,4}?)(\d{6,7})/i);
    if (concatMatch) {
      const fullNum = concatMatch[1] + concatMatch[2];
      let dept, po;
      if (fullNum.length === 10) {
        dept = fullNum.slice(0, 3);
        po = fullNum.slice(3);
      } else if (fullNum.length === 9) {
        dept = fullNum.slice(0, 3);
        po = fullNum.slice(3);
      } else if (fullNum.length === 11) {
        dept = fullNum.slice(0, 4);
        po = fullNum.slice(4);
      } else {
        dept = concatMatch[1];
        po = concatMatch[2];
      }
      metadata['DEPT #'] = parseInt(dept);
      metadata['PO #'] = po;
      break;
    }
    
    // Try spaced format
    const spacedMatch = line.match(/DEPT\.\s*NUMBER:\s*(\d+)\s*ORDER\s*NUMBER:\s*(\d+)/i);
    if (spacedMatch) {
      metadata['DEPT #'] = parseInt(spacedMatch[1]);
      metadata['PO #'] = spacedMatch[2];
      break;
    }
  }
  
  // Fallback: look for patterns in full text
  if (!metadata['PO #']) {
    // First try separate lines approach - DEPT. NUMBER: and ORDER NUMBER: are labels
    // followed by the actual values on subsequent lines
    let deptLineFound = false;
    let orderLineFound = false;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line.includes('DEPT. NUMBER:')) {
        deptLineFound = true;
      } else if (line.includes('ORDER NUMBER:') && deptLineFound) {
        orderLineFound = true;
      } else if (deptLineFound && orderLineFound && /^\d+$/.test(line)) {
        // First number after both labels = DEPT#
        // Second number = PO#
        if (!metadata['DEPT #']) {
          metadata['DEPT #'] = parseInt(line);
        } else if (!metadata['PO #']) {
          metadata['PO #'] = line;
          break;
        }
      }
    }
  }
  
  // Final fallback: look for patterns in full text
  if (!metadata['PO #']) {
    // Try finding DEPT and ORDER separately
    const deptMatch = allText.match(/DEPT\.?\s*(?:NUMBER|#)?:?\s*(\d{3,4})/i);
    const poMatch = allText.match(/ORDER\s*NUMBER:?\s*(\d{6,7})/i) || allText.match(/PO\s*#?\s*:?\s*(\d{6,7})/i);
    if (deptMatch) metadata['DEPT #'] = parseInt(deptMatch[1]);
    if (poMatch) metadata['PO #'] = poMatch[1];
  }
  
  // Extract dates
  const cancelDateMatch = allText.match(/Cancel\s*Date:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
  if (cancelDateMatch) metadata['CANCEL DATE'] = cancelDateMatch[1];
  
  const shipDateMatch = allText.match(/Ship\s*Date:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
  if (shipDateMatch) metadata['SHIP DATE'] = shipDateMatch[1];
  
  const orderDateMatch = allText.match(/Order\s*Date:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
  if (orderDateMatch) metadata['ORDER DATE'] = orderDateMatch[1];
  
  // Look for dates after Cancel Date line (they often appear on consecutive lines)
  if (metadata['CANCEL DATE'] && (!metadata['SHIP DATE'] || !metadata['ORDER DATE'])) {
    for (let i = 0; i < lines.length; i++) {
      if (/Cancel\s*Date/i.test(lines[i])) {
        // Look at next lines for dates
        for (let j = i + 1; j < Math.min(i + 5, lines.length); j++) {
          const dateMatch = lines[j].trim().match(/^(\d{1,2}\/\d{1,2}\/\d{4})$/);
          if (dateMatch) {
            if (!metadata['SHIP DATE']) {
              metadata['SHIP DATE'] = dateMatch[1];
            } else if (!metadata['ORDER DATE']) {
              metadata['ORDER DATE'] = dateMatch[1];
              break;
            }
          }
        }
        break;
      }
    }
  }
  
  // Extract store number
  const storePatterns = [
    /Store:?\s*(\d+)/i,
    /Mark\s*For:?\s*(\d+)/i,
    /DIST\s*CENTER\s*#?\s*(\d+)/i,
    /Support\s*Center\s*#?\s*(\d+)/i
  ];
  
  for (const pattern of storePatterns) {
    const match = allText.match(pattern);
    if (match) {
      metadata['STORE #'] = parseInt(match[1]);
      break;
    }
  }
  
  // Extract vendor name
  const vendorMatch = allText.match(/([\w\s]+)\s+Outlet\s+Stores/i) || 
                      allText.match(/([\w\s]+)\s+DIST\s*CENTER/i);
  if (vendorMatch) {
    metadata['VENDOR'] = vendorMatch[1].trim().toUpperCase();
  }
  
  // Extract state from shipping address
  const stateMatch = allText.match(/,\s*([A-Z]{2})\s+\d{5}/);
  if (stateMatch) {
    metadata['STATE'] = stateMatch[1];
  }
  
  return metadata;
}

/**
 * Parse date string to Date object
 */
function parseDate(dateStr) {
  if (!dateStr) return null;
  try {
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      const month = parseInt(parts[0]);
      const day = parseInt(parts[1]);
      const year = parseInt(parts[2]);
      return new Date(year, month - 1, day);
    }
  } catch (e) {
    console.error('Date parse error:', e);
  }
  return null;
}

/**
 * Get column letter from 1-based index
 */
function getColumnLetter(colNum) {
  let letter = '';
  while (colNum > 0) {
    const remainder = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    colNum = Math.floor((colNum - 1) / 26);
  }
  return letter;
}

/**
 * Parse SKU table from PDF text
 * Complete 4-strategy parser ported from backend
 */
function parseSkuTable(pdfText) {
  const lines = pdfText
    .split(/\r?\n/)
    .map(l => l.trim());

  const items = [];

  let headerIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    const upper = lines[i].toUpperCase();
    if (upper.includes('SKU')) {
      const windowLines = lines.slice(i, i + 6).join(' ').toUpperCase();
      if (windowLines.includes('MFG') && windowLines.includes('STYLE')) {
        headerIdx = i;
        break;
      }
    }
  }

  if (headerIdx === -1) {
    return items;
  }
  
  // Strategy 1: stacked columns (each field on its own line following headers)
  const firstSkuIdx = lines.findIndex((line, idx) => idx > headerIdx && /^\d{8,9}$/.test(line));
  
  if (firstSkuIdx !== -1) {
    let idx = firstSkuIdx;
    
    const isUpc = (text) => /^upc:/i.test(text) || /^\d{12,14}$/.test(text);
    const isPrice = (text) => /^\$[\d,]+\.\d{2}$/.test(text);
    const isSmallQty = (text) => /^\d+$/.test(text) && parseInt(text) >= 1 && parseInt(text) <= 9999;
    const isPackLetter = (text) => /^[A-Z]$/.test(text);
    
    while (idx < lines.length) {
      // Skip pack letters (A, B, C) that appear before SKUs in ComplexDomestic format
      while (idx < lines.length && isPackLetter(lines[idx])) {
        idx++;
      }
      if (idx >= lines.length) break;
      
      const skuLine = lines[idx];
      if (!/^\d{8,9}$/.test(skuLine)) {
        break;
      }
      if (/^number\s+of\s+packs\s*$/i.test(skuLine) || /^total\s+(cost|qty)/i.test(skuLine)) break;
      
      const item = {
        'SKU': skuLine,
        'MFG Style': '',
        'MFG Color': '',
        'Size Desc.': '',
        'Description': '',
        'Cost/Unit': '',
        'Retail': '',
        'Comp': '',
        'Pack Qty.': '',
        'Qty': ''
      };
      
      idx++;
      let fieldStage = 0;
      const prices = [];
      const quantities = [];
      
      while (idx < lines.length) {
        const line = lines[idx].trim();
        if (!line) { idx++; continue; }
        
        if (/^\d{8,9}$/.test(line)) break;
        if (/^number\s+of\s+packs\s+ordered:/i.test(line)) { idx++; break; }
        if (/^number\s+of\s+packs\s*$/i.test(line)) break;
        if (/^total\s+(cost|qty)/i.test(line)) break;
        if (/^page/i.test(line)) break;
        
        if (isUpc(line)) { idx++; continue; }
        
        if (isPrice(line)) {
          prices.push(line);
          idx++;
          continue;
        }
        
        if (prices.length > 0 && isSmallQty(line)) {
          quantities.push(parseInt(line));
          idx++;
          continue;
        }
        
        if (prices.length === 0) {
          if (fieldStage === 0) {
            item['MFG Style'] = line;
            fieldStage = 1;
          } else if (fieldStage === 1) {
            item['MFG Color'] = line;
            fieldStage = 2;
          } else if (fieldStage === 2) {
            item['Size Desc.'] = line;
            fieldStage = 3;
          } else if (fieldStage === 3) {
            if (item['Description']) {
              item['Description'] += ' ' + line;
            } else {
              item['Description'] = line;
            }
          }
        }
        idx++;
      }
      
      if (prices.length >= 1) item['Cost/Unit'] = prices[0];
      if (prices.length >= 2) item['Comp'] = prices[1];
      if (prices.length >= 3) item['Retail'] = prices[2];
      
      if (quantities.length >= 1) item['Pack Qty.'] = quantities[0];
      if (quantities.length >= 3) {
        item['Qty'] = quantities[2];
      } else if (quantities.length >= 2) {
        item['Qty'] = quantities[1];
      } else if (quantities.length === 1) {
        item['Qty'] = quantities[0];
      }
      
      items.push(item);
    }
  }

  // Strategy 2: multi-line blocks (SKU/style/desc on first line, UPC, prices)
  // Also handles inline format where pack letter precedes SKU on same line
  if (items.length === 0) {
    const colorRegex = /(White|Black|Red|Blue|Green|Yellow|Grey|Gray|Pink|Brown|Purple|Navy|Silver|Gold|Orange|Multi|Ivory|Cream|Beige|Khaki|Tan|Bone|Natural|Royal|Teal|Turquoise|Maroon|Olive|Charcoal|Burgundy|No Color)/i;

    for (let i = headerIdx + 1; i < lines.length; i++) {
      const line = lines[i];

      if (!line || /^\s*$/u.test(line)) continue;
      if (/^total\s+(qty|cost)/i.test(line)) break;
      if (/^number\s+of\s+packs\s*$/i.test(line)) break;

      // Allow optional pack letter (A, B, C) before SKU
      if (!/^[A-Z]?\s*\d{8,9}/.test(line)) continue;

      // Extract SKU, handling optional pack letter prefix
      // Use 's' flag to make . match newlines
      const skuMatch = line.match(/^([A-Z]?\s*)(\d{8,9})(.*?)$/s);
      if (!skuMatch) continue;

      const sku = skuMatch[2];
      let remainder = (skuMatch[3] || '').trim();
      
      // Normalize newlines to spaces (pdf.js combines text with newlines)
      remainder = remainder.replace(/\n+/g, ' ').replace(/\s+/g, ' ');
      
      const hasInlinePrice = /\$\d/.test(remainder);
      
      if (hasInlinePrice) {
        const allPrices = remainder.match(/\$[\d,]+\.\d{1,2}/g) || [];
        if (allPrices.length > 0) {
          const normalizedPrices = allPrices.map(p => {
            const parts = p.split('.');
            return parts[1] && parts[1].length === 1 ? parts[0] + '.' + parts[1] + '0' : p;
          });
          
          const firstPriceMatch = remainder.match(/\$[\d,]+\.\d{1,2}/);
          const left = remainder.slice(0, firstPriceMatch.index).trim();
          
          let style = '', color = '', sizeDesc = '', description = '', packQtyFromLeft = '';
          let working = left;
          const colorMatch = working.match(colorRegex);
          if (colorMatch) {
            style = working.slice(0, colorMatch.index).trim();
            color = colorMatch[1];
            working = working.slice(colorMatch.index + colorMatch[1].length).trim();
          } else {
            const firstToken = working.split(/\s+/)[0] || '';
            style = firstToken.trim();
            working = working.slice(firstToken.length).trim();
          }
          
          // Size can be: "10X8X5", "10X8", ".", "NO SIZE", or other descriptors
          // Try to match dimension format first
          const sizeMatch = working.match(/^(\d+X\d+X\d+|\d+X\d+|\.)/i);
          if (sizeMatch) {
            sizeDesc = sizeMatch[1];
            working = working.slice(sizeMatch.index + sizeMatch[1].length).trim();
          } else {
            // Match patterns like "NO SIZE", "ONE SIZE", "SIZE", etc. - words ending with SIZE
            const wordSizeMatch = working.match(/^([A-Z\s]*SIZE)\s+/i);
            if (wordSizeMatch) {
              sizeDesc = wordSizeMatch[1].trim();
              working = working.slice(wordSizeMatch[0].length).trim();
            }
          }
          
          // Check for Pack Qty (decimal like 3.0 or small integer 1-99) before description
          const packQtyMatch = working.match(/^(\d+\.?\d*)\s+/);
          if (packQtyMatch) {
            const num = parseFloat(packQtyMatch[1]);
            if (num >= 1 && num <= 99) {
              packQtyFromLeft = String(Math.round(num));
              working = working.slice(packQtyMatch[0].length).trim();
            }
          }
          
          description = working.replace(/^[\.\-\s]+/, '').trim();
          
          let lastPriceEndIdx = 0;
          for (const p of allPrices) {
            const idx = remainder.indexOf(p, lastPriceEndIdx);
            if (idx >= 0) lastPriceEndIdx = idx + p.length;
          }
          const right = remainder.slice(lastPriceEndIdx).trim();
          
          const parseQty = (text) => {
            const parts = (text || '').split(/\s+/).filter(Boolean);
            if (parts.length >= 3) {
              return { packQty: parts[0], totalPacks: parts[1], qty: parts[2] };
            }
            if (parts.length === 2) {
              return { packQty: parts[0], totalPacks: '', qty: parts[1] };
            }
            if (parts.length === 1) {
              const digits = parts[0].replace(/\D/g, '');
              if (digits.length >= 4) {
                for (let pqLen = 1; pqLen <= 2 && pqLen < digits.length - 2; pqLen++) {
                  const remaining = digits.slice(pqLen);
                  const halfLen = Math.floor(remaining.length / 2);
                  if (halfLen >= 2) {
                    const pq = digits.slice(0, pqLen);
                    const tp = remaining.slice(0, halfLen);
                    const q = remaining.slice(halfLen);
                    if (parseInt(tp) > 0 && parseInt(q) > 0) {
                      return { packQty: pq, totalPacks: tp, qty: q };
                    }
                  }
                }
              }
              return { packQty: '', totalPacks: '', qty: digits };
            }
            return { packQty: '', totalPacks: '', qty: '' };
          };
          
          const qtyParsed = parseQty(right);
          let packQty = qtyParsed.packQty;
          let qty = qtyParsed.qty;
          
          let packQtyFromNumberLine = '';
          let qtyFromNumberLine = '';
          for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
            const checkLine = lines[j] || '';
            if (/^\d{8,9}/.test(checkLine)) break;
            if (/^total\s+(qty|cost)/i.test(checkLine)) break;
            
            const numPacksFullMatch = checkLine.match(/number\s+of\s+packs\s+ordered:\s*(\d+)\s*units:\s*(\d+)/i);
            if (numPacksFullMatch) {
              const totalPacks = parseInt(numPacksFullMatch[1], 10);
              const unitsPerPack = parseInt(numPacksFullMatch[2], 10);
              packQtyFromNumberLine = numPacksFullMatch[2];
              qtyFromNumberLine = String(totalPacks * unitsPerPack);
              break;
            }
            const numPacksMatch = checkLine.match(/number\s+of\s+packs\s+ordered:\s*\d*\s*units:\s*(\d+)/i);
            if (numPacksMatch) {
              packQtyFromNumberLine = numPacksMatch[1];
              break;
            }
            if (/units:\s*$/i.test(checkLine)) {
              const nextLine = (lines[j + 1] || '').trim();
              if (/^\d{1,2}$/.test(nextLine)) {
                packQtyFromNumberLine = nextLine;
                const packsMatch = checkLine.match(/ordered:\s*(\d+)/i);
                if (packsMatch) {
                  const totalPacks = parseInt(packsMatch[1], 10);
                  const unitsPerPack = parseInt(nextLine, 10);
                  qtyFromNumberLine = String(totalPacks * unitsPerPack);
                }
              }
              break;
            }
          }
          
          const finalPackQty = packQtyFromNumberLine || packQtyFromLeft || packQty;
          const finalQty = qtyFromNumberLine || qty;
          
          items.push({
            'SKU': sku,
            'MFG Style': style,
            'MFG Color': color,
            'Size Desc.': sizeDesc,
            'Description': description,
            'Cost/Unit': normalizedPrices[0] || '',
            'Comp': normalizedPrices[1] || '',
            'Retail': normalizedPrices[2] || '',
            'Pack Qty.': finalPackQty,
            'Qty': finalQty,
            'UPC': ''
          });
          continue;
        }
      }

      let style = '';
      let color = '';
      let sizeDesc = '';
      let description = '';
      let packQtyFromLine = '';

      const colorMatch = remainder.match(colorRegex);
      if (colorMatch) {
        style = remainder.slice(0, colorMatch.index).trim();
        color = colorMatch[1];
        remainder = remainder.slice(colorMatch.index + colorMatch[1].length);
      } else {
        const firstToken = remainder.split(/\s+/)[0] || '';
        style = firstToken.trim();
        remainder = remainder.slice(firstToken.length);
      }

      const packQtyPattern = remainder.match(/^\.(\d{1,2})\.(.*)$/);
      const textSizePackQtyPattern = remainder.match(/^([A-Z\s]+)(\d{1,2})\.(.*)$/i);
      const packQtyInlinePattern = remainder.match(/^\.(\d)(\d[A-Z].*)$/i);
      const looksLikeSize = remainder.match(/^\.\d+X\d+/i);
      
      if (packQtyPattern) {
        sizeDesc = '.';
        packQtyFromLine = packQtyPattern[1];
        remainder = packQtyPattern[2];
      } else if (textSizePackQtyPattern) {
        sizeDesc = textSizePackQtyPattern[1].trim();
        packQtyFromLine = textSizePackQtyPattern[2];
        remainder = textSizePackQtyPattern[3];
      } else if (packQtyInlinePattern && !looksLikeSize) {
        sizeDesc = '.';
        packQtyFromLine = packQtyInlinePattern[1];
        remainder = packQtyInlinePattern[2];
      } else {
        const sizeMatch = remainder.match(/(\d+X\d+X\d+|\d+X\d+|\d+\.\d+|\d+)/i);
        if (sizeMatch) {
          sizeDesc = sizeMatch[1];
          remainder = remainder.slice(sizeMatch.index + sizeMatch[1].length);
        }
      }
      description = remainder.replace(/^[\.\-\s]+/, '').trim();

      let upcLine = '';
      let priceLine = '';
      let searchEnd = Math.min(i + 10, lines.length);
      let packQtyFromNumberLine = '';
      let qtyFromNumberLine = '';
      
      for (let j = i + 1; j < searchEnd; j++) {
        const checkLine = lines[j] || '';
        
        if (/^\d{8,9}/.test(checkLine)) break;
        if (/^total\s+(qty|cost)/i.test(checkLine)) break;
        
        const numPacksFullMatch = checkLine.match(/number\s+of\s+packs\s+ordered:\s*(\d+)\s*units:\s*(\d+)/i);
        if (numPacksFullMatch) {
          const totalPacks = parseInt(numPacksFullMatch[1], 10);
          const unitsPerPack = parseInt(numPacksFullMatch[2], 10);
          packQtyFromNumberLine = numPacksFullMatch[2];
          qtyFromNumberLine = String(totalPacks * unitsPerPack);
          break;
        }
        const numPacksMatch = checkLine.match(/number\s+of\s+packs\s+ordered:\s*\d*\s*units:\s*(\d+)/i);
        if (numPacksMatch) {
          packQtyFromNumberLine = numPacksMatch[1];
          break;
        }
        if (/units:\s*$/i.test(checkLine)) {
          const nextLine = (lines[j + 1] || '').trim();
          if (/^\d{1,2}$/.test(nextLine)) {
            packQtyFromNumberLine = nextLine;
            const packsMatch = checkLine.match(/ordered:\s*(\d+)/i);
            if (packsMatch) {
              const totalPacks = parseInt(packsMatch[1], 10);
              const unitsPerPack = parseInt(nextLine, 10);
              qtyFromNumberLine = String(totalPacks * unitsPerPack);
            }
          }
          break;
        }
        if (/^number\s+of\s+packs\s+ordered:/i.test(checkLine)) {
          break;
        }
        
        if (/^upc:/i.test(checkLine)) {
          upcLine = checkLine;
          continue;
        }
        
        if (/\$\d/.test(checkLine)) {
          priceLine += ' ' + checkLine;
          continue;
        }
        
        if (priceLine && /^[\d\$\.]/.test(checkLine)) {
          const hasCompleteQty = /\$[\d,]+\.\d{2}\d+/.test(priceLine);
          const isStandaloneNumber = /^\d{2,4}$/.test(checkLine.trim());
          
          if (!hasCompleteQty && !isStandaloneNumber) {
            priceLine += checkLine;
            continue;
          }
          if (hasCompleteQty && isStandaloneNumber) {
            break;
          }
        }
        
        if (priceLine && /^number\s+of\s+packs/i.test(checkLine)) {
          const numPacksMatch = checkLine.match(/units:\s*(\d+)/i);
          if (numPacksMatch) {
            packQtyFromNumberLine = numPacksMatch[1];
          }
          const totalPacksMatch = checkLine.match(/ordered:\s*(\d+)/i);
          if (totalPacksMatch && numPacksMatch) {
            const totalPacks = parseInt(totalPacksMatch[1], 10);
            const unitsPerPack = parseInt(numPacksMatch[1], 10);
            qtyFromNumberLine = String(totalPacks * unitsPerPack);
          }
          i = j - 1;
          break;
        }
        
        if (!upcLine && !priceLine && checkLine && !/^upc:/i.test(checkLine)) {
          description += ' ' + checkLine.trim();
        }
      }

      let priceText = priceLine.trim();
      
      priceText = priceText.replace(/(\$\d+\.\d)(\d)(\$)/g, '$1$2 $3');
      priceText = priceText.replace(/(\$\d+\.\d)(\d)(\d)/g, '$1$2 $3');
      
      const priceMatches = priceText.match(/\$[\d,]+\.\d{1,2}/g) || [];
      if (priceMatches.length === 0) continue;
      
      const normalizedPrices = priceMatches.map(p => {
        const parts = p.split('.');
        if (parts[1].length === 1) {
          return parts[0] + '.' + parts[1] + '0';
        }
        return p;
      });

      const cost = normalizedPrices[0] || '';
      const comp = normalizedPrices[1] || '';
      const retail = normalizedPrices[2] || '';
      
      let afterPrices = priceText;
      for (const p of priceMatches) {
        afterPrices = afterPrices.replace(p, ' ');
      }
      const qtyDigits = afterPrices.replace(/\D/g, '');
      let packQty = '';
      let qty = '';
      
      if (qtyDigits.length >= 4) {
        const pq1 = parseInt(qtyDigits[0], 10);
        const rest1 = qtyDigits.slice(1);
        const halfLen1 = Math.floor(rest1.length / 2);
        const tp1 = parseInt(rest1.slice(0, halfLen1), 10);
        const tu1 = parseInt(rest1.slice(halfLen1), 10);
        
        if (tu1 === tp1 * pq1 && pq1 > 0 && pq1 <= 12) {
          packQty = String(pq1);
          qty = String(tu1);
        } else {
          const pq2 = parseInt(qtyDigits.slice(0, 2), 10);
          const rest2 = qtyDigits.slice(2);
          if (rest2.length >= 2) {
            const halfLen2 = Math.floor(rest2.length / 2);
            const tp2 = parseInt(rest2.slice(0, halfLen2), 10);
            const tu2 = parseInt(rest2.slice(halfLen2), 10);
            
            if (tu2 === tp2 * pq2 && pq2 > 0 && pq2 <= 24) {
              packQty = String(pq2);
              qty = String(tu2);
            } else {
              qty = qtyDigits.slice(Math.floor(qtyDigits.length / 2));
            }
          } else {
            qty = qtyDigits;
          }
        }
      } else if (qtyDigits.length > 0) {
        qty = qtyDigits;
      }

      const finalPackQty = packQtyFromLine || packQtyFromNumberLine || packQty;
      const finalQty = qtyFromNumberLine || qty;

      items.push({
        'SKU': sku,
        'MFG Style': style,
        'MFG Color': color,
        'Size Desc.': sizeDesc,
        'Description': description.trim(),
        'Cost/Unit': cost,
        'Comp': comp,
        'Retail': retail,
        'Pack Qty.': finalPackQty,
        'Qty': finalQty,
        'UPC': upcLine.match(/(\d{12})/)?.[1] || ''
      });
    }
  }

  // Strategy 3: single-line compressed rows
  {
    const colorRegex = /(White|Black|Red|Blue|Green|Yellow|Grey|Gray|Pink|Brown|Purple|Navy|Silver|Gold|Orange|Multi|Ivory|Cream|Beige|Khaki|Tan|Bone|Natural|Royal|Teal|Turquoise|Maroon|Olive|Charcoal|Burgundy|No Color)/i;

    const existingSkus = new Set(items.map(item => item.SKU));

    const parseLeftSide = (text) => {
      let style = '';
      let color = '';
      let sizeDesc = '';
      let description = '';
      let packQtyFromLeft = '';

      let working = text.trim();
      const colorMatch = working.match(colorRegex);
      if (colorMatch) {
        style = working.slice(0, colorMatch.index).trim();
        color = colorMatch[1];
        working = working.slice(colorMatch.index + colorMatch[1].length).trim();
      } else {
        const firstToken = working.split(/\s+/)[0] || '';
        style = firstToken;
        working = working.slice(firstToken.length).trim();
      }

      // Size is typically like "10X8X5" or a dot "." or empty
      const sizeMatch = working.match(/^(\d+X\d+X\d+|\d+X\d+|\.)/i);
      if (sizeMatch) {
        sizeDesc = sizeMatch[1];
        working = working.slice(sizeMatch.index + sizeMatch[1].length).trim();
      }
      
      // Check for Pack Qty (decimal like 3.0 or small integer 1-99) before description
      const packQtyMatch = working.match(/^(\d+\.?\d*)\s+/);
      if (packQtyMatch) {
        const num = parseFloat(packQtyMatch[1]);
        if (num >= 1 && num <= 99) {
          packQtyFromLeft = String(Math.round(num));
          working = working.slice(packQtyMatch[0].length).trim();
        }
      }

      description = working.replace(/^[\.\-\s]+/, '').trim();
      return { style, color, sizeDesc, description, packQtyFromLeft };
    };

    const parseQuantities = (text) => {
      const parts = (text || '').split(/\s+/).filter(Boolean);
      if (parts.length >= 3) {
        return { packQty: parts[0], totalPacks: parts[1], qty: parts[2] };
      }
      if (parts.length === 2) {
        return { packQty: parts[0], totalPacks: '', qty: parts[1] };
      }
      if (parts.length === 1) {
        const digits = parts[0].replace(/\D/g, '');
        if (digits.length >= 4) {
          for (let pqLen = 1; pqLen <= 2 && pqLen < digits.length - 2; pqLen++) {
            const remaining = digits.slice(pqLen);
            const halfLen = Math.floor(remaining.length / 2);
            if (halfLen >= 2) {
              const packQty = digits.slice(0, pqLen);
              const totalPacks = remaining.slice(0, halfLen);
              const qty = remaining.slice(halfLen);
              if (parseInt(totalPacks) > 0 && parseInt(qty) > 0) {
                return { packQty, totalPacks, qty };
              }
            }
          }
        }
        return { packQty: '', totalPacks: '', qty: digits };
      }

      return { packQty: '', totalPacks: '', qty: '' };
    };

    for (let i = headerIdx + 1; i < lines.length; i++) {
      const line = lines[i];
      if (!line) continue;
      if (/^total\s+(cost|qty)/i.test(line)) break;
      if (/^number\s+of\s+packs\s*$/i.test(line)) break;
      if (/^ordered:\s*$/i.test(line)) break;
      if (/^no\.\s+of\s+\w+\s+packs\s+ordered:\s*$/i.test(line)) break;
      if (!/^[A-Z]?\s*\d{8,9}/.test(line)) continue;

      const skuMatch = line.match(/^([A-Z]?\s*)(\d{8,9})(.*)$/);
      if (!skuMatch) continue;
      const packPrefix = skuMatch[1] || '';
      const sku = skuMatch[2];
      
      if (existingSkus.has(sku)) continue;
      
      const afterSku = (skuMatch[3] || '').trim();

      const allPrices = afterSku.match(/\$[\d,]+\.\d{1,2}/g) || [];
      if (allPrices.length === 0) continue;
      
      const normalizedPrices = allPrices.map(p => {
        const parts = p.split('.');
        if (parts[1] && parts[1].length === 1) {
          return parts[0] + '.' + parts[1] + '0';
        }
        return p;
      });

      const firstPriceMatch = afterSku.match(/\$[\d,]+\.\d{1,2}/);
      if (!firstPriceMatch) continue;
      
      const left = afterSku.slice(0, firstPriceMatch.index).trim();
      
      let lastPriceEndIdx = 0;
      for (const p of allPrices) {
        const idx = afterSku.indexOf(p, lastPriceEndIdx);
        if (idx !== -1) {
          lastPriceEndIdx = idx + p.length;
        }
      }
      const right = afterSku.slice(lastPriceEndIdx).trim();

      const { style, color, sizeDesc, description, packQtyFromLeft } = parseLeftSide(left);
      const { packQty, totalPacks, qty } = parseQuantities(right);
      
      const cost = normalizedPrices[0] || '';
      const comp = normalizedPrices[1] || '';
      const retail = normalizedPrices[2] || (normalizedPrices.length === 2 ? normalizedPrices[1] : '');

      items.push({
        'SKU': sku,
        'MFG Style': style,
        'MFG Color': color,
        'Size Desc.': sizeDesc,
        'Description': description,
        'Cost/Unit': cost,
        'Retail': retail,
        'Pack Qty.': packQtyFromLeft || packQty,
        'Qty': qty || totalPacks || ''
      });
      
      existingSkus.add(sku);
    }
  }

  // Strategy 4: Multi-line format (SKU on one line, prices on later line)
  {
    const colorRegex = /(White|Black|Red|Blue|Green|Yellow|Grey|Gray|Pink|Brown|Purple|Navy|Silver|Gold|Orange|Multi|Ivory|Cream|Beige|Khaki|Tan|Bone|Natural|Royal|Teal|Turquoise|Maroon|Olive|Charcoal|Burgundy|No Color)/i;
    
    const existingSkus = new Set(items.map(item => item.SKU));

    for (let i = headerIdx + 1; i < lines.length; i++) {
      const line = lines[i];
      if (!line) continue;
      if (/^total\s+(cost|qty)/i.test(line)) continue;
      if (/^no\.\s+of\s+\w+\s+packs/i.test(line)) continue;
      if (/^total\s+pack\s+/i.test(line)) continue;
      if (/^Pack\s*SKU/i.test(line)) continue;
      if (/^page/i.test(line)) continue;
      
      if (!/^[A-Z]?\s*\d{8,9}/.test(line)) continue;
      
      const skuMatch = line.match(/^([A-Z]?\s*)(\d{8,9})(.*)$/);
      if (!skuMatch) continue;
      
      const packPrefix = (skuMatch[1] || '').trim();
      const sku = skuMatch[2];
      let afterSku = (skuMatch[3] || '').trim();
      
      if (existingSkus.has(sku)) continue;
      
      if (/\$\d/.test(afterSku)) continue;
      
      let style = '';
      let color = '';
      let sizeDesc = '';
      let descParts = [];
      let dataLineIdx = i;
      
      if (!afterSku && i + 1 < lines.length) {
        const nextLine = lines[i + 1].trim();
        if (nextLine && !/^UPC:/i.test(nextLine) && !/^\$/.test(nextLine) && 
            !/^[A-Z]?\s*\d{8,9}/.test(nextLine) && !/^number/i.test(nextLine)) {
          afterSku = nextLine;
          dataLineIdx = i + 1;
        }
      }
      
      let packQtyFromLine = '';
      
      const colorMatch = afterSku.match(colorRegex);
      if (colorMatch) {
        style = afterSku.slice(0, colorMatch.index).trim();
        color = colorMatch[1];
        afterSku = afterSku.slice(colorMatch.index + colorMatch[1].length).trim();
      } else {
        const firstToken = afterSku.split(/\s+/)[0] || '';
        style = firstToken;
        afterSku = afterSku.slice(firstToken.length).trim();
      }
      
      const packQtyPattern = afterSku.match(/^\.(\d{1,2})\.(.*)$/);
      const textSizePackQtyPattern = afterSku.match(/^([A-Z\s]+)(\d{1,2})\.(.*)$/i);
      const packQtyInlinePattern = afterSku.match(/^\.(\d)(\d[A-Z].*)$/i);
      const looksLikeSize = afterSku.match(/^\.\d+X\d+/i);
      
      if (packQtyPattern) {
        sizeDesc = '.';
        packQtyFromLine = packQtyPattern[1];
        afterSku = packQtyPattern[2].trim();
      } else if (textSizePackQtyPattern) {
        sizeDesc = textSizePackQtyPattern[1].trim();
        packQtyFromLine = textSizePackQtyPattern[2];
        afterSku = textSizePackQtyPattern[3].trim();
      } else if (packQtyInlinePattern && !looksLikeSize) {
        sizeDesc = '.';
        packQtyFromLine = packQtyInlinePattern[1];
        afterSku = packQtyInlinePattern[2].trim();
      } else {
        const sizeMatch = afterSku.match(/^\.?(\d+(?:\.\d+)?[A-Z]?)\s*/i);
        if (sizeMatch) {
          sizeDesc = sizeMatch[1];
          afterSku = afterSku.slice(sizeMatch[0].length).trim();
        }
      }
      
      if (afterSku) descParts.push(afterSku);
      
      let packQty = '';
      let qty = '';
      
      let priceLineIdx = -1;
      for (let j = dataLineIdx + 1; j < Math.min(dataLineIdx + 6, lines.length); j++) {
        const checkLine = lines[j].trim();
        if (!checkLine) continue;
        if (/^number\s+of\s+packs\s*$/i.test(checkLine)) break;
        if (/^ordered:\s*$/i.test(checkLine)) break;
        if (/^no\.\s+of\s+\w+\s+packs\s+ordered:\s*$/i.test(checkLine)) break;
        if (/^total\s+(cost|qty|pack)/i.test(checkLine)) break;
        if (/^[A-Z]?\s*\d{8,9}/.test(checkLine)) break;
        
        if (/\$\d/.test(checkLine)) {
          priceLineIdx = j;
          break;
        }
        
        if (/^UPC:/i.test(checkLine)) continue;
        
        descParts.push(checkLine);
      }
      
      if (priceLineIdx === -1) continue;
      
      let priceText = '';
      let foundPriceLine = false;
      let lastPriceLineIdx = priceLineIdx;
      let packQtyFromNumberLine = '';
      
      for (let j = priceLineIdx; j < Math.min(priceLineIdx + 8, lines.length); j++) {
        const checkLine = lines[j].trim();
        
        if (/^number\s+of\s+packs/i.test(checkLine)) {
          const unitsMatch = checkLine.match(/Units:\s*(\d+)/i);
          if (unitsMatch) {
            packQtyFromNumberLine = unitsMatch[1];
          } else if (/units:\s*$/i.test(checkLine)) {
            const nextLine = (lines[j + 1] || '').trim();
            if (/^\d{1,2}$/.test(nextLine)) {
              packQtyFromNumberLine = nextLine;
            }
          }
          const packsMatch = checkLine.match(/Number Of Packs Ordered:\s*(\d+)/i);
          if (packsMatch && packQtyFromNumberLine) {
            const totalPacks = parseInt(packsMatch[1], 10);
            const unitsPerPack = parseInt(packQtyFromNumberLine, 10);
            qty = String(totalPacks * unitsPerPack);
          } else if (packsMatch) {
            qty = packsMatch[1];
          }
          lastPriceLineIdx = j;
          break;
        }
        if (/^total\s+(cost|qty|pack)/i.test(checkLine)) break;
        if (/^[A-Z]?\s*\d{8,9}/.test(checkLine)) break;
        if (/^UPC:/i.test(checkLine)) continue;
        
        if (/\$\d/.test(checkLine)) {
          const hasCompleteQty = /\$[\d,]+\.\d{2}\d+/.test(priceText);
          if (hasCompleteQty) {
            break;
          }
          foundPriceLine = true;
          priceText += ' ' + checkLine;
          lastPriceLineIdx = j;
          continue;
        }
        
        if (foundPriceLine && /^\d$/.test(checkLine)) {
          const hasCompleteQty = /\$[\d,]+\.\d{2}\d+/.test(priceText);
          if (!hasCompleteQty) {
            priceText += checkLine;
            lastPriceLineIdx = j;
          }
          continue;
        }
        
        if (foundPriceLine && /^\d{2,3}$/.test(checkLine)) {
          continue;
        }
        
        if (foundPriceLine && /^\d{4,}$/.test(checkLine)) {
          const hasCompleteQty = /\$[\d,]+\.\d{2}\d+/.test(priceText);
          if (!hasCompleteQty) {
            priceText += ' ' + checkLine;
            lastPriceLineIdx = j;
          }
          continue;
        }
        
        if (foundPriceLine && /^\d{1,4}$/.test(checkLine)) {
          continue;
        }
        
        priceText += ' ' + checkLine;
        lastPriceLineIdx = j;
      }
      
      let cleanPriceText = priceText.trim();
      
      cleanPriceText = cleanPriceText.replace(/(\$[\d,]+\.\d)(\d)(?=[^.\d]|$)/g, '$1$2 ');
      
      const priceMatches = cleanPriceText.match(/\$[\d,]+\.\d{1,2}/g) || [];
      
      const normalizedPrices = priceMatches.map(p => {
        const parts = p.split('.');
        if (parts[1] && parts[1].length === 1) {
          return parts[0] + '.' + parts[1] + '0';
        }
        return p;
      });
      
      const cost = normalizedPrices[0] || '';
      const comp = normalizedPrices[1] || '';
      const retail = normalizedPrices[2] || (normalizedPrices.length === 2 ? normalizedPrices[1] : '');
      
      if (!qty) {
        let afterPrices = cleanPriceText;
        for (const p of priceMatches) {
          afterPrices = afterPrices.replace(p, ' ');
        }
        const qtyDigits = afterPrices.replace(/\D/g, '');
        
        if (qtyDigits.length >= 4) {
          packQty = qtyDigits.slice(0, 1);
          const rest = qtyDigits.slice(1);
          qty = rest.slice(Math.floor(rest.length / 2));
        } else if (qtyDigits.length > 0) {
          qty = qtyDigits;
        }
      }
      
      const finalPackQty = packQtyFromLine || packQtyFromNumberLine || packQty;
      
      items.push({
        'SKU': sku,
        'MFG Style': style,
        'MFG Color': color,
        'Size Desc.': sizeDesc,
        'Description': descParts.join(' ').trim(),
        'Cost/Unit': cost,
        'Retail': retail,
        'Pack Qty.': finalPackQty,
        'Qty': qty
      });
      
      i = priceLineIdx;
    }
  }

  // Post-processing: Filter invalid entries
  const filteredItems = items.filter(item => {
    const style = (item['MFG Style'] || '').trim();
    const sku = (item['SKU'] || '').trim();
    
    if (/^Units\d*/i.test(style)) {
      return false;
    }
    
    if (/^\d+$/.test(style)) {
      return false;
    }
    
    const cost = (item['Cost/Unit'] || '').trim();
    if (!cost && !item['Qty']) {
      return false;
    }
    
    return true;
  });

  return filteredItems;
}

/**
 * Convert multiple PDFs to merged Excel
 */
async function convertMultiplePdfsToExcel(files) {
  const allRecords = [];
  let firstMetadata = null;
  
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const text = await extractPdfText(arrayBuffer);
      
      const metadata = await getPoMetadata(text);
      
      if (i === 0) firstMetadata = metadata;
      
      const skuItems = parseSkuTable(text);
      
      const records = skuItems.map(item => ({
        ...metadata,
        ...item,
        'SourceFile': file.name
      }));
      
      allRecords.push(...records);
    } catch (error) {
      console.error(`Error processing ${file.name}:`, error);
    }
  }
  
  if (allRecords.length === 0) {
    throw new Error('No data extracted from any PDF. Check browser console for details.');
  }
  
  // Count unique source files that actually produced records
  const uniqueSourceFiles = [...new Set(allRecords.map(r => r.SourceFile))];
  
  // Merge records by SKU/MFG Style
  const mergedDict = {};
  const keysOrder = [];
  
  for (const record of allRecords) {
    const sku = record['SKU'] || '';
    const style = record['MFG Style'] || '';
    const key = `${sku}|${style}`;
    
    if (!mergedDict[key]) {
      mergedDict[key] = {
        'SKU': sku,
        'MFG Style': style,
        'Cost/Unit': record['Cost/Unit'] || '',
        'Retail': record['Retail'] || '',
        'Pack Qty.': null,
        'PO_Quantities': {},
        'SourceFiles': new Set()
      };
      keysOrder.push(key);
    }
    
    const merged = mergedDict[key];
    
    // Pack Qty - take first non-zero
    if (!merged['Pack Qty.'] || merged['Pack Qty.'] === 0) {
      const packQty = parseInt(record['Pack Qty.'] || 0);
      if (packQty > 0) merged['Pack Qty.'] = packQty;
    }
    
    // Sum quantities by PO
    const poNum = record['PO #'] || '';
    if (poNum) {
      const qty = parseInt(record['Qty'] || 0);
      merged['PO_Quantities'][poNum] = (merged['PO_Quantities'][poNum] || 0) + qty;
    }
    
    if (record['SourceFile']) {
      merged['SourceFiles'].add(record['SourceFile']);
    }
  }
  
  // Get unique PO numbers
  const uniquePOs = [...new Set(allRecords.map(r => r['PO #']).filter(Boolean))];
  uniquePOs.sort((a, b) => (parseInt(a) || 0) - (parseInt(b) || 0));
  
  // Create workbook
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('PO Data');
  
  const vendorName = firstMetadata?.['VENDOR'] || 'VENDOR';
  
  // Header info in column G
  ws.getCell('G1').value = `${vendorName} PO INFO`;
  ws.getCell('G2').value = 'ORDER DATE';
  ws.getCell('G3').value = 'SHIP DATE';
  ws.getCell('G4').value = 'CANCEL DATE';
  ws.getCell('G5').value = 'DEPT#';
  ws.getCell('G6').value = 'DC';
  ws.getCell('G7').value = 'STORE #';
  ws.getCell('G8').value = `${vendorName} PO#`;
  
  // PO columns starting at H (column 8)
  for (let i = 0; i < uniquePOs.length; i++) {
    const po = uniquePOs[i];
    const col = 8 + i;
    
    // Find a record with this PO to get its metadata
    const record = allRecords.find(r => r['PO #'] === po);
    if (record) {
      // Row 8: PO number
      ws.getCell(8, col).value = `PO# ${po}`;
      
      // Row 2: Order Date
      if (record['ORDER DATE']) {
        const cell = ws.getCell(2, col);
        cell.value = parseDate(record['ORDER DATE']);
        if (cell.value) cell.numFmt = 'mm/dd/yyyy';
      }
      
      // Row 3: Ship Date
      if (record['SHIP DATE']) {
        const cell = ws.getCell(3, col);
        cell.value = parseDate(record['SHIP DATE']);
        if (cell.value) cell.numFmt = 'mm/dd/yyyy';
      }
      
      // Row 4: Cancel Date
      if (record['CANCEL DATE']) {
        const cell = ws.getCell(4, col);
        cell.value = parseDate(record['CANCEL DATE']);
        if (cell.value) cell.numFmt = 'mm/dd/yyyy';
      }
      
      // Row 5: Dept#
      if (record['DEPT #']) ws.getCell(5, col).value = record['DEPT #'];
      
      // Row 6: DC/State
      if (record['STATE']) ws.getCell(6, col).value = record['STATE'];
      
      // Row 7: Store#
      if (record['STORE #']) ws.getCell(7, col).value = record['STORE #'];
    }
    
    // Row 9: TTL UNITS header
    ws.getCell(9, col).value = 'TTL UNITS';
    ws.getCell(9, col).font = { bold: true };
  }
  
  // Column headers in row 9
  const headers = ['SKU #', 'MFG STYLE', 'COST / UNIT', 'RETAIL', 'TTL AMT', 'TTL UNITS', 'PACK QTY'];
  headers.forEach((h, idx) => {
    ws.getCell(9, idx + 1).value = h;
    ws.getCell(9, idx + 1).font = { bold: true };
  });
  
  // Source file column
  const srcCol = 8 + uniquePOs.length;
  ws.getCell(9, srcCol).value = 'SOURCE FILE';
  ws.getCell(9, srcCol).font = { bold: true };
  
  // Data rows starting at row 10
  let row = 10;
  for (const key of keysOrder) {
    const merged = mergedDict[key];
    
    ws.getCell(row, 1).value = merged['SKU'];
    ws.getCell(row, 2).value = merged['MFG Style'];
    
    // Cost and Retail as numbers
    const costStr = (merged['Cost/Unit'] || '').replace(/[$,]/g, '');
    const retailStr = (merged['Retail'] || '').replace(/[$,]/g, '');
    const costNum = parseFloat(costStr) || null;
    const retailNum = parseFloat(retailStr) || null;
    
    if (costNum) ws.getCell(row, 3).value = costNum;
    if (retailNum) ws.getCell(row, 4).value = retailNum;
    
    // TTL AMT formula
    ws.getCell(row, 5).value = { formula: `C${row}*F${row}` };
    
    // TTL UNITS formula
    if (uniquePOs.length > 0) {
      const firstLetter = getColumnLetter(8);
      const lastLetter = getColumnLetter(8 + uniquePOs.length - 1);
      ws.getCell(row, 6).value = { formula: `SUM(${firstLetter}${row}:${lastLetter}${row})` };
    }
    
    // Pack Qty
    ws.getCell(row, 7).value = merged['Pack Qty.'] || '';
    
    // PO quantities
    uniquePOs.forEach((po, idx) => {
      const qty = merged['PO_Quantities'][po];
      if (qty && qty > 0) {
        ws.getCell(row, 8 + idx).value = qty;
      }
    });
    
    // Source files
    ws.getCell(row, srcCol).value = [...merged['SourceFiles']].sort().join(', ');
    
    row++;
  }
  
  // Totals in row 8
  const lastRow = row - 1;
  if (lastRow >= 10) {
    ws.getCell(8, 5).value = { formula: `SUM(E10:E${lastRow})` };
    ws.getCell(8, 6).value = { formula: `SUM(F10:F${lastRow})` };
  }
  
  // Formatting
  for (let r = 9; r <= lastRow; r++) {
    ws.getCell(r, 3).numFmt = '$#,##0.00';
    ws.getCell(r, 4).numFmt = '$#,##0.00';
    ws.getCell(r, 5).numFmt = '$#,##0.00';
  }
  ws.getCell(8, 5).numFmt = '$#,##0.00';
  
  // Column widths
  ws.getColumn(1).width = 12;
  ws.getColumn(2).width = 15;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 10;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 10;
  ws.getColumn(7).width = 10;
  
  return workbook;
}

/**
 * Parse CSV with proper quote handling
 */
function parseCSV(content) {
  const rows = [];
  let currentRow = [];
  let currentField = '';
  let inQuotes = false;
  
  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    const nextChar = content[i + 1];
    
    if (char === '"') {
      if (!inQuotes) {
        inQuotes = true;
      } else if (nextChar === '"') {
        currentField += '"';
        i++;
      } else {
        inQuotes = false;
      }
    } else if (char === ',' && !inQuotes) {
      currentRow.push(currentField.trim());
      currentField = '';
    } else if ((char === '\n' || (char === '\r' && nextChar === '\n')) && !inQuotes) {
      if (char === '\r') i++;
      currentRow.push(currentField.trim());
      if (currentRow.some(f => f)) rows.push(currentRow);
      currentRow = [];
      currentField = '';
    } else if (char === '\r' && !inQuotes) {
      currentRow.push(currentField.trim());
      if (currentRow.some(f => f)) rows.push(currentRow);
      currentRow = [];
      currentField = '';
    } else {
      currentField += char;
    }
  }
  
  currentRow.push(currentField.trim());
  if (currentRow.some(f => f)) rows.push(currentRow);
  
  if (rows.length === 0) return { headers: [], rows: [] };
  
  const headers = rows[0];
  const dataRows = rows.slice(1).map(values => {
    const row = {};
    headers.forEach((h, idx) => row[h] = values[idx] || '');
    return row;
  });
  
  return { headers, rows: dataRows };
}

/**
 * Merge vendor data into customer Excel
 */
async function mergeVendorData(customerArrayBuffer, vendorArrayBuffer, vendorFileName) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(customerArrayBuffer);
  
  const customerWs = workbook.worksheets[0];
  if (!customerWs) throw new Error('Customer Excel has no worksheets');
  
  const mergedWorkbook = new ExcelJS.Workbook();
  const mergedWs = mergedWorkbook.addWorksheet('Merged Data');
  
  const maxRow = customerWs.rowCount;
  const maxCol = customerWs.columnCount;
  
  // Copy columns A-B as-is
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= 2; c++) {
      const src = customerWs.getCell(r, c);
      const dest = mergedWs.getCell(r, c);
      dest.value = src.value;
      if (src.style) dest.style = JSON.parse(JSON.stringify(src.style));
    }
  }
  
  // Copy columns C+ shifted right by 4
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 3; c <= maxCol; c++) {
      const src = customerWs.getCell(r, c);
      const dest = mergedWs.getCell(r, c + 4);
      dest.value = src.value;
      if (src.style) dest.style = JSON.parse(JSON.stringify(src.style));
    }
  }
  
  // New column headers
  mergedWs.getCell(9, 3).value = 'VEND #';
  mergedWs.getCell(9, 3).font = { bold: true };
  mergedWs.getCell(9, 4).value = 'BASE COST';
  mergedWs.getCell(9, 4).font = { bold: true };
  mergedWs.getCell(9, 5).value = 'Box/Case';
  mergedWs.getCell(9, 5).font = { bold: true };
  mergedWs.getCell(9, 6).value = 'Unit/Case';
  mergedWs.getCell(9, 6).font = { bold: true };
  
  // Build vendor lookup
  const vendorLookup = {};
  const isCSV = vendorFileName.toLowerCase().endsWith('.csv');
  
  if (isCSV) {
    const content = new TextDecoder().decode(vendorArrayBuffer);
    const { rows } = parseCSV(content);
    
    for (const row of rows) {
      const itemNum = (row['Item#'] || '').toString().trim();
      const whsNum = parseInt(row['WHS#']) || 0;
      
      if (itemNum && !vendorLookup[itemNum] && whsNum === 1) {
        const baseCost = parseFloat((row['Base Cost'] || '').replace(/[$,]/g, '')) || null;
        vendorLookup[itemNum] = {
          vend: (row['Vend#'] || '').trim(),
          baseCost,
          unitCase: (row['Unit/Case'] || '').trim(),
          boxCase: (row['Box/Case'] || '').trim()
        };
      }
    }
  } else {
    const vendorWb = new ExcelJS.Workbook();
    await vendorWb.xlsx.load(vendorArrayBuffer);
    const vendorWs = vendorWb.worksheets[0];
    
    const headerRow = vendorWs.getRow(1);
    let cols = {};
    headerRow.eachCell((cell, num) => {
      const val = (cell.value || '').toString().trim();
      if (val === 'Item#') cols.item = num;
      else if (val === 'WHS#') cols.whs = num;
      else if (val === 'Vend#') cols.vend = num;
      else if (val === 'Base Cost') cols.baseCost = num;
      else if (val === 'Unit/Case') cols.unitCase = num;
      else if (val === 'Box/Case') cols.boxCase = num;
    });
    
    vendorWs.eachRow((row, num) => {
      if (num === 1) return;
      const itemNum = (row.getCell(cols.item).value || '').toString().trim();
      const whsNum = parseInt(row.getCell(cols.whs).value) || 0;
      
      if (itemNum && !vendorLookup[itemNum] && whsNum === 1) {
        const bcVal = row.getCell(cols.baseCost).value;
        const baseCost = typeof bcVal === 'number' ? bcVal : parseFloat((bcVal || '').toString().replace(/[$,]/g, '')) || null;
        vendorLookup[itemNum] = {
          vend: (row.getCell(cols.vend).value || '').toString().trim(),
          baseCost,
          unitCase: cols.unitCase ? (row.getCell(cols.unitCase).value || '').toString().trim() : '',
          boxCase: cols.boxCase ? (row.getCell(cols.boxCase).value || '').toString().trim() : ''
        };
      }
    });
  }
  
  // Apply vendor data
  for (let r = 10; r <= maxRow; r++) {
    const mfgStyle = (mergedWs.getCell(r, 2).value || '').toString().trim();
    if (vendorLookup[mfgStyle]) {
      const data = vendorLookup[mfgStyle];
      mergedWs.getCell(r, 3).value = data.vend;
      mergedWs.getCell(r, 4).value = data.baseCost;
      mergedWs.getCell(r, 5).value = data.boxCase;
      mergedWs.getCell(r, 6).value = data.unitCase;
    }
  }
  
  // Fix formulas
  mergedWs.getCell(8, 9).value = { formula: `SUM(I10:I${maxRow})` };
  for (let r = 10; r <= maxRow; r++) {
    mergedWs.getCell(r, 9).value = { formula: `G${r}*J${r}` };
  }
  mergedWs.getCell(8, 10).value = { formula: `SUM(J10:J${maxRow})` };
  
  const lastColLetter = getColumnLetter(maxCol + 4);
  for (let r = 10; r <= maxRow; r++) {
    mergedWs.getCell(r, 10).value = { formula: `SUM(L${r}:${lastColLetter}${r})` };
  }
  
  // Formatting
  for (let r = 8; r <= maxRow; r++) {
    mergedWs.getCell(r, 4).numFmt = '$#,##0.00';
    mergedWs.getCell(r, 9).numFmt = '$#,##0.00';
  }
  
  return mergedWorkbook;
}

// Export for browser
window.POConverter = {
  extractPdfText,
  getPoMetadata,
  parseSkuTable,
  convertMultiplePdfsToExcel,
  mergeVendorData,
  parseCSV
};
